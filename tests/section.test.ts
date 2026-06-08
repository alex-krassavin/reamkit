import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { convertDocxToPdfSync } from '@/converter';
import { parseTtf } from '@/font';
import { OpcPackage } from '@/opc';
import { parseSection, parseSections } from '@/ooxml/wordproc';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
};

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

function sectionOf(bodyInner: string) {
  const docx = buildDocxFromBody(bodyInner);
  const pkg = OpcPackage.open(docx);
  return parseSection(pkg.getMainDocument().data);
}

describe('parseSections (multi-section)', () => {
  it('returns one section spanning entire body when no sectPr', () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>a</w:t></w:r></w:p>');
    const pkg = OpcPackage.open(docx);
    const sections = parseSections(pkg.getMainDocument().data);
    expect(sections).toHaveLength(1);
    expect(sections[0]!.endIndex).toBe(1);
  });

  it('collects mid-document sectPr inside pPr and final body-level sectPr', () => {
    const body = `
      <w:p><w:r><w:t>p1</w:t></w:r></w:p>
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
          </w:sectPr>
        </w:pPr>
        <w:r><w:t>p2 (ends section 1)</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>p3</w:t></w:r></w:p>
      <w:p><w:r><w:t>p4</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838" w:orient="landscape"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body);
    const pkg = OpcPackage.open(docx);
    const sections = parseSections(pkg.getMainDocument().data);

    expect(sections).toHaveLength(2);
    // First section ends after paragraph 2 (1-indexed body[0..2)).
    expect(sections[0]!.endIndex).toBe(2);
    expect(sections[0]!.properties.pageSize?.widthTwips).toBe(12240);
    // Second section covers remaining paragraphs.
    expect(sections[1]!.endIndex).toBe(4);
    expect(sections[1]!.properties.pageSize?.orientation).toBe('landscape');
  });
});

describe('parseSection', () => {
  it('returns empty section when no sectPr present', () => {
    const s = sectionOf('<w:p><w:r><w:t>Hi</w:t></w:r></w:p>');
    expect(s.pageSize).toBeUndefined();
    expect(s.margins).toBeUndefined();
    expect(s.headers).toEqual([]);
    expect(s.footers).toEqual([]);
  });

  it('parses pgSz and pgMar', () => {
    const s = sectionOf(`
      <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
        <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720"/>
      </w:sectPr>`);
    expect(s.pageSize).toEqual({
      widthTwips: 11906,
      heightTwips: 16838,
      orientation: 'portrait',
    });
    expect(s.margins).toEqual({
      topTwips: 1440,
      rightTwips: 1800,
      bottomTwips: 1440,
      leftTwips: 1800,
      headerTwips: 720,
      footerTwips: 720,
    });
  });

  it('parses titlePg toggle from sectPr', () => {
    const noTitle = sectionOf(`
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`);
    expect(noTitle.titlePg).toBeUndefined();

    const withTitle = sectionOf(`
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:titlePg/>
      </w:sectPr>`);
    expect(withTitle.titlePg).toBe(true);
  });

  it('parses header and footer references with types', () => {
    const s = sectionOf(`
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:headerReference r:id="rId12" w:type="first"/>
        <w:footerReference r:id="rId11" w:type="default"/>
      </w:sectPr>`);
    expect(s.headers).toHaveLength(2);
    expect(s.headers[0]).toEqual({ type: 'default', relationshipId: 'rId10' });
    expect(s.headers[1]).toEqual({ type: 'first', relationshipId: 'rId12' });
    expect(s.footers).toEqual([{ type: 'default', relationshipId: 'rId11' }]);
  });
});

describe('Headers and footers in rendered PDF', () => {
  it('renders header text on the page when document references a header part', () => {
    const headerXml = '<w:p><w:r><w:t>HEADER-MARK</w:t></w:r></w:p>';
    const footerXml = '<w:p><w:r><w:t>FOOTER-MARK</w:t></w:r></w:p>';
    const body = `
      <w:p><w:r><w:t>Body paragraph</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:footerReference r:id="rId11" w:type="default"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, { headerXml, footerXml });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');

    expect(text).toContain(`<${hexOf('HEADER-MARK')}> Tj`);
    expect(text).toContain(`<${hexOf('FOOTER-MARK')}> Tj`);
    expect(text).toContain(`<${hexOf('Body')}> Tj`);
  });

  it('uses first-page header on page 1 when titlePg is set', () => {
    // Two long paragraphs → exactly 2 pages. FIRST-MARK must render once (on
    // page 1), DEFAULT-MARK must render once (on page 2).
    const longText = 'X '.repeat(2000);
    const body = `
      <w:p><w:r><w:t>${longText}</w:t></w:r></w:p>
      <w:p><w:r><w:t>${longText}</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:headerReference r:id="rId12" w:type="first"/>
        <w:titlePg/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>DEFAULT-MARK</w:t></w:r></w:p>',
      firstHeaderXml: '<w:p><w:r><w:t>FIRST-MARK</w:t></w:r></w:p>',
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    const countOf = (s: string) => (text.match(new RegExp(`<${hexOf(s)}> Tj`, 'g')) ?? []).length;

    // FIRST-MARK exactly once (page 1), DEFAULT-MARK exactly once (page 2).
    expect(countOf('FIRST-MARK')).toBe(1);
    expect(countOf('DEFAULT-MARK')).toBe(1);
  });

  it('starts a new page on a forced page break (w:br w:type="page")', () => {
    // Short doc whose only reason to span 2 pages is the explicit page break
    // in its own paragraph (the POI header/footer-suite pattern).
    const body = `
      <w:p><w:r><w:t>FIRSTPAGE</w:t></w:r></w:p>
      <w:p><w:r><w:br w:type="page"/></w:r></w:p>
      <w:p><w:r><w:t>SECONDPAGE</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>HMARK</w:t></w:r></w:p>',
    });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    const countOf = (s: string) => (text.match(new RegExp(`<${hexOf(s)}> Tj`, 'g')) ?? []).length;

    // Header renders once per page → exactly 2 pages means the break worked.
    expect(countOf('HMARK')).toBe(2);
    expect(countOf('FIRSTPAGE')).toBe(1); // before the break (page 1)
    expect(countOf('SECONDPAGE')).toBe(1); // after the break (page 2, not lost)
  });

  it('renders the header/footer band when the body is empty', () => {
    // A header/footer-only document (empty body) must still emit one page that
    // carries the bands — otherwise the only text in the file is lost (POI
    // headerFooter.docx pattern).
    const body = `
      <w:p></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>ONLYHEADER</w:t></w:r></w:p>',
    });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    expect(text).toContain(`<${hexOf('ONLYHEADER')}> Tj`);
  });

  it('applies per-section orientation (portrait section 1, landscape section 2)', () => {
    // Section 1 (portrait) ends at the paragraph carrying its sectPr; section 2
    // (landscape via w:orient) is the body-level sectPr. Each page must get its
    // own section's MediaBox. (Reduction of bug65649's multi-section case.)
    const body =
      `<w:p><w:r><w:t>SEC1</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>SEC2</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    const boxes = [...text.matchAll(/\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/g)].map(
      (m) => [Math.round(Number(m[1])), Math.round(Number(m[2]))] as const,
    );
    expect(boxes).toHaveLength(2);
    expect(boxes[0]![0]).toBeLessThan(boxes[0]![1]); // page 1 portrait (w < h)
    expect(boxes[1]![0]).toBeGreaterThan(boxes[1]![1]); // page 2 landscape (w > h)
  });

  it('uses even-page header on page 2 when evenAndOddHeaders is set', () => {
    // Two long paragraphs → ~2 pages. EVEN-MARK on page 2, ODD-MARK on page 1.
    const longText = 'X '.repeat(2000);
    const body = `
      <w:p><w:r><w:t>${longText}</w:t></w:r></w:p>
      <w:p><w:r><w:t>${longText}</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:headerReference r:id="rId14" w:type="even"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>ODD-MARK</w:t></w:r></w:p>',
      evenHeaderXml: '<w:p><w:r><w:t>EVEN-MARK</w:t></w:r></w:p>',
      settingsXml: '<w:evenAndOddHeaders/>',
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    const countOf = (s: string) => (text.match(new RegExp(`<${hexOf(s)}> Tj`, 'g')) ?? []).length;

    expect(countOf('ODD-MARK')).toBe(1);
    expect(countOf('EVEN-MARK')).toBe(1);
  });

  it('falls back to default header when first variant is missing', () => {
    // titlePg is set but no firstHeaderReference → default header on every page.
    const body = `
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference r:id="rId10" w:type="default"/>
        <w:titlePg/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>ONLY-MARK</w:t></w:r></w:p>',
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    expect(text).toContain(`<${hexOf('ONLY-MARK')}> Tj`);
  });

  it('emits per-section MediaBox when sections have different page sizes', () => {
    // Section 1: A4 portrait (596×842). Section 2: Letter landscape (792×612).
    // Section break is on para 1 → page 1 is portrait, page 2 is landscape.
    const body = `
      <w:p>
        <w:pPr>
          <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          </w:sectPr>
        </w:pPr>
        <w:r><w:t>Page 1 portrait</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>Page 2 landscape</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
        <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // Both MediaBoxes must be present in the output.
    // A4 portrait: 11906/20=595.3 × 16838/20=841.9.
    expect(text).toMatch(/\/MediaBox \[0 0 595\.3 841\.9\]/);
    // Letter landscape: 15840/20=792 × 12240/20=612.
    expect(text).toMatch(/\/MediaBox \[0 0 792 612\]/);
  });

  it('section break forces a page break between sections', () => {
    const body = `
      <w:p>
        <w:pPr>
          <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
        </w:pPr>
        <w:r><w:t>SECTION-ONE</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>SECTION-TWO</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const pageCount = (text.match(/\/Type \/Page\b/g) ?? []).filter(
      (m) => !m.includes('Pages'),
    ).length;
    expect(pageCount).toBe(2);
  });

  it('uses pgSz from sectPr to set MediaBox', () => {
    const body = `
      <w:p><w:r><w:t>x</w:t></w:r></w:p>
      <w:sectPr>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      </w:sectPr>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // 12240 twips = 612 pt (US Letter width); 15840 twips = 792 pt (US Letter height)
    expect(text).toMatch(/\/MediaBox \[0 0 612 792\]/);
  });
});
