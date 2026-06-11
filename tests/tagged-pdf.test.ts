import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { convertDocxToPdfSync } from '@/core/converter';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

// Extract the raw (uncompressed) content streams referenced by page /Contents.
function pageContentStreams(pdf: Uint8Array): Array<string> {
  const text = asLatin1(pdf);
  const ids = [...text.matchAll(/\/Contents (\d+) 0 R/g)].map((m) => Number(m[1]));
  const out: Array<string> = [];
  for (const id of ids) {
    const objIdx = text.indexOf(`\n${id} 0 obj`);
    if (objIdx < 0) continue;
    const s = text.indexOf('stream\n', objIdx);
    const e = text.indexOf('\nendstream', s);
    if (s >= 0 && e >= 0) out.push(text.slice(s + 'stream\n'.length, e));
  }
  return out;
}

// PDF/A-1a §6.3.2 / ISO 32000 §14.8: every painting or text-showing operator
// must sit inside a marked-content sequence (a tagged BDC, or an /Artifact
// BMC/BDC). Walk the stream tracking marked-content depth; throw if a paint
// happens at depth 0 or the BDC/EMC brackets are unbalanced.
function assertAllContentMarked(stream: string): void {
  let depth = 0;
  for (const raw of stream.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    if (/\bBDC$/.test(line) || /\bBMC$/.test(line)) {
      depth++;
    } else if (line === 'EMC') {
      depth--;
      if (depth < 0) throw new Error('unbalanced EMC');
    } else if (
      (/ (Tj|TJ|Do)$/.test(line) || /^(f|F|f\*|S|s|B|B\*|b|b\*)$/.test(line)) &&
      depth === 0
    ) {
      throw new Error(`unmarked painting operator outside marked content: "${line}"`);
    }
  }
  if (depth !== 0) throw new Error(`unbalanced marked content (depth ${depth})`);
}

// A paragraph long enough to wrap onto several lines on A4 (so the structure
// element collects more than one MCID).
const LONG =
  'The quick brown fox jumps over the lazy dog and then keeps on running across ' +
  'the wide open meadow under a bright blue sky while the sun slowly sets behind ' +
  'the distant rolling hills far away on the horizon line.';

describe('tagged PDF — structure tree (M6.1)', () => {
  const oneParagraph = () =>
    buildDocxFromBody('<w:p><w:r><w:t>Hello tagged world</w:t></w:r></w:p>');

  it('marks the document as tagged in the catalog', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, tagged: true }));
    expect(text).toContain('/MarkInfo <</Marked true>>');
    expect(text).toContain('/StructTreeRoot');
    expect(text).toContain('/Lang (en-US)');
  });

  it('tags a paragraph in a non-default language with a per-element /Lang', () => {
    // Document default = en-US (the first w:lang); the Russian paragraph differs
    // and gets its own /Lang on its StructElem.
    const body =
      '<w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>English paragraph</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:lang w:val="ru-RU"/></w:rPr><w:t>Russkij abzac</w:t></w:r></w:p>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS, tagged: true }),
    );
    expect(text).toContain('/Lang (en-US)'); // catalog default
    expect(text).toContain('/Lang (ru-RU)'); // the Russian paragraph's StructElem
    expect((text.match(/\/Lang \(ru-RU\)/g) ?? []).length).toBe(1); // not the English one
  });

  it('wraps body text in /P marked content with an MCID', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, tagged: true }));
    expect(text).toContain('/P <</MCID 0>> BDC');
    expect(text).toContain('EMC');
    // The single text line sits inside exactly one marked-content sequence.
    // (Match the full operator — a bare "BDC" also occurs in binary font bytes.)
    expect((text.match(/\/P <<\/MCID \d+>> BDC/g) ?? []).length).toBe(1);
  });

  it('builds a StructTreeRoot → Document → P element chain', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, tagged: true }));
    expect(text).toContain('/Type /StructTreeRoot');
    expect(text).toContain('/Type /StructElem /S /Document');
    expect(text).toContain('/Type /StructElem /S /P');
    // The P element references its marked content via an MCR dict.
    expect(text).toMatch(/\/Type \/MCR \/Pg \d+ 0 R \/MCID 0/);
  });

  it('emits a ParentTree and gives the page /StructParents + /Tabs', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, tagged: true }));
    expect(text).toContain('/ParentTree');
    expect(text).toContain('/ParentTreeNextKey 1');
    expect(text).toMatch(/\/Nums \[0 \[\d+ 0 R\]\]/);
    expect(text).toContain('/StructParents 0');
    expect(text).toContain('/Tabs /S');
  });

  it('collects every wrapped line of a paragraph under one P (multiple MCIDs)', () => {
    const docx = buildDocxFromBody(`<w:p><w:r><w:t>${LONG}</w:t></w:r></w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS, tagged: true }));
    // The paragraph wrapped, so there is more than one MCID...
    expect(text).toContain('/P <</MCID 0>> BDC');
    expect(text).toContain('/P <</MCID 1>> BDC');
    // ...but still a single P structure element owning them.
    expect((text.match(/\/StructElem \/S \/P\b/g) ?? []).length).toBe(1);
    // ParentTree array length tracks the number of MCIDs on the page.
    const mcidCount = (text.match(/\/MCID \d+>> BDC/g) ?? []).length;
    const numsRefs = (text.match(/\/Nums \[0 \[([^\]]*)\]\]/)?.[1] ?? '').match(/\d+ 0 R/g) ?? [];
    expect(numsRefs.length).toBe(mcidCount);
  });

  it('maps outline levels and heading styles to H1–H6 (M6.2)', () => {
    const body =
      '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Chapter</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Subsection</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:outlineLvl w:val="8"/></w:pPr><w:r><w:t>Deep level</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>The Title</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Ordinary body text.</w:t></w:r></w:p>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS, tagged: true }),
    );
    expect(text).toContain('/Type /StructElem /S /H1'); // outline level 0
    expect(text).toContain('/Type /StructElem /S /H3'); // outline level 2
    expect(text).toContain('/Type /StructElem /S /H6'); // outline level 8 → clamped
    // The Title style id is also recognised as a heading (H1), so two H1s total.
    expect((text.match(/\/StructElem \/S \/H1\b/g) ?? []).length).toBe(2);
    // The plain paragraph stays a P.
    expect((text.match(/\/StructElem \/S \/P\b/g) ?? []).length).toBe(1);
  });

  it('marks header/footer text as a pagination artifact, not structure (M6.3)', () => {
    const body =
      '<w:p><w:r><w:t>Body paragraph.</w:t></w:r></w:p>' +
      '<w:sectPr>' +
      '<w:headerReference r:id="rId10" w:type="default"/>' +
      '<w:footerReference r:id="rId11" w:type="default"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>' +
      '</w:sectPr>';
    const docx = buildDocxFromBody(body, {
      headerXml: '<w:p><w:r><w:t>Running header</w:t></w:r></w:p>',
      footerXml: '<w:p><w:r><w:t>Page footer</w:t></w:r></w:p>',
    });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS, tagged: true }));
    // Header and footer text are pagination artifacts (one each)...
    expect((text.match(/\/Artifact <<\/Type \/Pagination>> BDC/g) ?? []).length).toBe(2);
    // ...while only the body paragraph is in the structure tree.
    expect((text.match(/\/StructElem \/S \/P\b/g) ?? []).length).toBe(1);
    expect(text).toContain('/P <</MCID');
  });

  const table2x2 =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

  it('builds Table → TR → TD → P structure for tables (M6.4)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(table2x2), { fonts: FONTS, tagged: true }),
    );
    expect((text.match(/\/StructElem \/S \/Table\b/g) ?? []).length).toBe(1);
    expect((text.match(/\/StructElem \/S \/TR\b/g) ?? []).length).toBe(2);
    expect((text.match(/\/StructElem \/S \/TD\b/g) ?? []).length).toBe(4);
    expect((text.match(/\/StructElem \/S \/P\b/g) ?? []).length).toBe(4);
    // Each cell's text is /P marked content (tag matches the P struct type).
    expect(text).toContain('/P <</MCID');
    // Every MCID resolves through the ParentTree (one line per cell → 4).
    const mcidCount = (text.match(/\/MCID \d+>> BDC/g) ?? []).length;
    const numsRefs = (text.match(/\/Nums \[0 \[([^\]]*)\]\]/)?.[1] ?? '').match(/\d+ 0 R/g) ?? [];
    expect(mcidCount).toBe(4);
    expect(numsRefs.length).toBe(mcidCount);
  });

  it('emits TH with /Scope /Column for a w:tblHeader row (§14.8.5.2)', () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>D1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>D2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(table), { fonts: FONTS, tagged: true }),
    );
    // Header row → 2 TH; data row → 2 TD.
    expect((text.match(/\/StructElem \/S \/TH\b/g) ?? []).length).toBe(2);
    expect((text.match(/\/StructElem \/S \/TD\b/g) ?? []).length).toBe(2);
    // Each TH carries a /Table attribute object with /Scope /Column.
    expect((text.match(/\/A <<\/O \/Table \/Scope \/Column>>/g) ?? []).length).toBe(2);
  });

  it('emits a single TD for a column-spanning cell (M6.4)', () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Wide</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(table), { fonts: FONTS, tagged: true }),
    );
    // Three logical cells → three TDs (the spanning cell is one TD), two TRs.
    expect((text.match(/\/StructElem \/S \/TD\b/g) ?? []).length).toBe(3);
    expect((text.match(/\/StructElem \/S \/TR\b/g) ?? []).length).toBe(2);
    // The spanning cell carries /ColSpan 2 (§14.8.5.2).
    expect(text).toContain('/A <</O /Table /ColSpan 2>>');
  });

  it('emits /RowSpan for a vertically-merged cell (§14.8.5.2)', () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>RS</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(table), { fonts: FONTS, tagged: true }),
    );
    // Origin cell spans 2 rows → /RowSpan 2; the continuation cell is suppressed
    // (3 logical cells: RS origin + B1 + B2 → 3 TDs).
    expect(text).toContain('/A <</O /Table /RowSpan 2>>');
    expect((text.match(/\/StructElem \/S \/TD\b/g) ?? []).length).toBe(3);
  });

  it('builds nested L → LI → LBody → P structure for lists (M6.5)', () => {
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl>' +
      '<w:lvl w:ilvl="2"><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const li = (ilvl: number, t: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const body = li(0, 'One') + li(1, 'One-a') + li(1, 'One-b') + li(0, 'Two') + li(2, 'Two-deep');
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body, { numberingXml }), {
        fonts: FONTS,
        tagged: true,
      }),
    );
    // 5 items → 5 LI/LBody/P; lists opened at level 0, 1, and 2 → 3 L elements.
    expect((text.match(/\/StructElem \/S \/L\b/g) ?? []).length).toBe(3);
    expect((text.match(/\/StructElem \/S \/LI\b/g) ?? []).length).toBe(5);
    expect((text.match(/\/StructElem \/S \/LBody\b/g) ?? []).length).toBe(5);
    expect((text.match(/\/StructElem \/S \/P\b/g) ?? []).length).toBe(5);
    // §14.8.4.3.3: every item's marker glyphs sit in their own Lbl element.
    expect((text.match(/\/StructElem \/S \/Lbl\b/g) ?? []).length).toBe(5);
    // No stray top-level P (every list item's P is inside an LBody).
    expect(text).toContain('/StructElem /S /L ');
  });

  // A standalone picture (paragraph whose only content is a drawing → block
  // image). `descr` becomes the drawing's wp:docPr description.
  const pictureBody = (descr?: string) =>
    `<w:p><w:r><w:drawing>` +
    `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="914400" cy="914400"/>` +
    `<wp:docPr id="1" name="Picture 1"${descr !== undefined ? ` descr="${descr}"` : ''}/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  const pngOpts = () => ({
    images: {
      rId20: {
        contentType: 'image/png' as const,
        bytes: buildTinyPng(4, 4, [200, 50, 50, 255]),
        extension: 'png' as const,
      },
    },
  });
  const altHex = (s: string) =>
    'FEFF' +
    [...s].map((c) => c.codePointAt(0)!.toString(16).padStart(4, '0').toUpperCase()).join('');

  it('tags a picture as a Figure with /Alt from docPr descr (M6.6)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(pictureBody('A red square'), pngOpts()), {
        fonts: FONTS,
        tagged: true,
      }),
    );
    expect(text).toContain('/Type /StructElem /S /Figure');
    expect(text).toContain(`/Alt <${altHex('A red square')}>`);
    // The image XObject is inside /Figure marked content, not an artifact.
    expect(text).toMatch(/\/Figure <<\/MCID \d+>> BDC/);
  });

  it('falls back to a generic /Alt when a picture has no description (M6.6)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(pictureBody(), pngOpts()), {
        fonts: FONTS,
        tagged: true,
      }),
    );
    expect(text).toContain('/Type /StructElem /S /Figure');
    expect(text).toContain(`/Alt <${altHex('Image')}>`);
  });

  it('PDF/A-1a implies tagging and sets XMP conformance to A', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, pdfA: 'PDF/A-1a' }));
    // Tagged structures present...
    expect(text).toContain('/StructTreeRoot');
    expect(text).toContain('/P <</MCID 0>> BDC');
    // ...plus the inherited PDF/A-1b apparatus, now at conformance level A.
    expect(text).toContain('/OutputIntents');
    expect(text).toContain('pdfaid:conformance>A<');
    expect(
      asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, pdfA: 'PDF/A-1a' })),
    ).toMatch(/^%PDF-1\.4/);
  });

  it('uses options.language for the catalog /Lang (M6.7)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, tagged: true, language: 'ru-RU' }),
    );
    expect(text).toContain('/Lang (ru-RU)');
  });

  it('auto-detects the document language from w:lang (M6.7)', () => {
    const body =
      '<w:p><w:r><w:rPr><w:lang w:val="fr-FR"/></w:rPr><w:t>Bonjour le monde</w:t></w:r></w:p>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS, tagged: true }),
    );
    expect(text).toContain('/Lang (fr-FR)');
  });

  it('sets ViewerPreferences DisplayDocTitle when a title is present (M6.7)', () => {
    const text = asLatin1(
      convertDocxToPdfSync(oneParagraph(), {
        fonts: FONTS,
        tagged: true,
        info: { title: 'My Doc' },
      }),
    );
    expect(text).toContain('/ViewerPreferences <</DisplayDocTitle true>>');
  });

  it('PDF/A-1b stays level B and untagged', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, pdfA: 'PDF/A-1b' }));
    expect(text).toContain('pdfaid:conformance>B<');
    expect(text).not.toContain('/StructTreeRoot');
    // "<</MCID" is specific to our marked content (a bare "BDC" also lands in
    // binary font bytes, so it can't be used for absence checks).
    expect(text).not.toContain('<</MCID');
  });

  it('emits no structure when not tagging', () => {
    const text = asLatin1(convertDocxToPdfSync(oneParagraph(), { fonts: FONTS }));
    expect(text).not.toContain('/StructTreeRoot');
    expect(text).not.toContain('/MarkInfo');
    expect(text).not.toContain('<</MCID');
    expect(text).not.toContain('/StructParents');
  });

  it('is deterministic (byte-identical for identical tagged input)', () => {
    const a = convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, pdfA: 'PDF/A-1a' });
    const b = convertDocxToPdfSync(oneParagraph(), { fonts: FONTS, pdfA: 'PDF/A-1a' });
    expect(asLatin1(a)).toBe(asLatin1(b));
  });

  // A document exercising every structure path (heading, paragraph, table with
  // shading+borders, nested list, picture) plus a header/footer artifact band.
  const richBody = () => {
    const numberingXml =
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const li = (l: number, t: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${l}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const tbl =
      '<w:tbl><w:tblPr><w:tblBorders>' +
      '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>' +
      '<w:left w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>' +
      '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>' +
      '</w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const body =
      '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Body paragraph.</w:t></w:r></w:p>' +
      tbl +
      li(0, 'First item') +
      li(1, 'Nested item') +
      pictureBody('A diagram') +
      '<w:sectPr><w:headerReference r:id="rId10" w:type="default"/>' +
      '<w:footerReference r:id="rId11" w:type="default"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>';
    return buildDocxFromBody(body, {
      numberingXml,
      headerXml: '<w:p><w:r><w:t>Running header</w:t></w:r></w:p>',
      footerXml: '<w:p><w:r><w:t>Page footer</w:t></w:r></w:p>',
      ...pngOpts(),
    });
  };

  it('PDF/A-1a gate: no page content sits outside the structure or an artifact (M6.8)', () => {
    const pdf = convertDocxToPdfSync(richBody(), { fonts: FONTS, pdfA: 'PDF/A-1a' });
    const streams = pageContentStreams(pdf);
    expect(streams.length).toBeGreaterThan(0);
    // The core conformance invariant: every painting op is marked.
    for (const s of streams) assertAllContentMarked(s);
  });

  it('PDF/A-1a gate: full structure present and every Figure has /Alt (M6.8)', () => {
    const text = asLatin1(convertDocxToPdfSync(richBody(), { fonts: FONTS, pdfA: 'PDF/A-1a' }));
    // The whole range of structure + artifact paths was exercised.
    expect(text).toContain('/S /H1');
    expect(text).toContain('/S /Table');
    expect(text).toContain('/S /TD');
    expect(text).toContain('/S /L '); // a list
    expect(text).toContain('/S /Figure');
    expect(text).toContain('/Artifact <</Type /Pagination>> BDC'); // header/footer
    expect(text).toContain('pdfaid:conformance>A<');
    // Every Figure structure element carries non-empty alternate text.
    const figureLines = text.split('\n').filter((l) => l.includes('/S /Figure'));
    expect(figureLines.length).toBeGreaterThan(0);
    for (const line of figureLines) expect(line).toContain('/Alt <FEFF');
  });
});
