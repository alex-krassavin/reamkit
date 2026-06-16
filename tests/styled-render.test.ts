import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody, buildRichDocx } from './fixtures/build-docx';
import type { FamilyKey } from '@/core/fonts';
import { convertDocxToPdfSync } from '@/core/converter';
import { FontRegistry, parseTtf } from '@/core/font';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-BoldItalic.ttf'))),
};

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('Styled rendering: rPr + pPr → PDF', () => {
  it('embeds a separate font resource for each variant used', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Reg ' },
          { text: 'Bold ', rPrXml: '<w:rPr><w:b/></w:rPr>' },
          { text: 'Italic ', rPrXml: '<w:rPr><w:i/></w:rPr>' },
          { text: 'BoldItalic', rPrXml: '<w:rPr><w:b/><w:i/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // Subsetted fonts carry a 6-letter subset tag prefix (PDF §9.6.4).
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Bold\b/);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Italic\b/);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-BoldItalic\b/);
    expect(text.match(/\/Subtype \/Type0/g)?.length).toBeGreaterThanOrEqual(4);
  });

  it('emits an rg color operator with the run colour', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Plain ' },
          { text: 'Red', rPrXml: '<w:rPr><w:color w:val="ff0000"/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\b1 0 0 rg\b/);
  });

  it('falls back to bold-only when boldItalic font is missing', () => {
    const docx = buildRichDocx([{ runs: [{ text: 'BI', rPrXml: '<w:rPr><w:b/><w:i/></w:rPr>' }] }]);
    const pdf = convertDocxToPdfSync(docx, {
      fonts: { regular: FONTS.regular, bold: FONTS.bold },
    });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Bold\b/);
    expect(text).not.toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-BoldItalic\b/);
  });

  it('emits different /Tf font sizes for runs with different sizes', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Small ', rPrXml: '<w:rPr><w:sz w:val="20"/></w:rPr>' },
          { text: 'Big', rPrXml: '<w:rPr><w:sz w:val="48"/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\/F\d+ 10 Tf/);
    expect(text).toMatch(/\/F\d+ 24 Tf/);
  });

  it('renders table borders as path operators and places cell text', () => {
    const body = `
      <w:p><w:r><w:t>Before</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="3000"/>
          <w:gridCol w:w="3000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:r><w:t>After</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toMatch(/\d+(\.\d+)? w/);
    expect(text).toMatch(/m\n[^\n]*l\nS/);

    const tjMatches = [...text.matchAll(/<[0-9A-Fa-f]+> Tj/g)];
    expect(tjMatches.length).toBeGreaterThanOrEqual(6);
  });

  it('draws an internal separator defined on a cell right border (neighbour fallback)', () => {
    // No <w:tblBorders>; the table's only border is A1's RIGHT (the A|B
    // separator). Internal verticals are drawn on the right cell's LEFT side, so
    // without the neighbour fallback (cell.left ?? leftNeighbor.right) this
    // separator would vanish — the real-world bug where a contract table lost all
    // its column gridlines.
    const body = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcBorders><w:right w:val="single" w:sz="8" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    // The separator is emitted as a stroked path (m … l … S).
    expect(text).toMatch(/m\n[^\n]*l\nS/);
  });

  it('resolves a shared border by weight — the heavier one wins (§17.4)', () => {
    // The A|B edge is specified on BOTH cells: A.right = 2pt (sz 16), B.left =
    // 0.5pt (sz 4). Border-conflict resolution keeps the heavier (2pt) and
    // discards the lighter.
    const body = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcBorders><w:right w:val="single" w:sz="16" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcBorders><w:left w:val="single" w:sz="4" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    expect(text).toMatch(/(^|\s)2 w(\s|$)/); // the 2pt edge is drawn
    expect(text).not.toMatch(/(^|\s)0\.5 w(\s|$)/); // the 0.5pt one is discarded
  });

  it('justify lines use per-token Tm positioning (more Tms than left-aligned)', () => {
    const longText =
      'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor ' +
      'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis ' +
      'nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo.';
    const left = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="left"/></w:pPr>', runs: [{ text: longText }] },
    ]);
    const both = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="both"/></w:pPr>', runs: [{ text: longText }] },
    ]);
    const pdfLeft = asLatin1(convertDocxToPdfSync(left, { fonts: FONTS }));
    const pdfBoth = asLatin1(convertDocxToPdfSync(both, { fonts: FONTS }));

    const tmLeft = (pdfLeft.match(/Tm/g) ?? []).length;
    const tmBoth = (pdfBoth.match(/Tm/g) ?? []).length;
    expect(tmBoth).toBeGreaterThan(tmLeft);
  });

  it('renders a gridSpan=2 cell with the combined width of two columns', () => {
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2000"/>
          <w:gridCol w:w="2000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged header that spans both columns</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const tjMatches = [...text.matchAll(/<[0-9A-Fa-f]+> Tj/g)];
    expect(tjMatches.length).toBeGreaterThanOrEqual(3);
  });

  it('does not render content for vMerge=continue cells', () => {
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2000"/>
          <w:gridCol w:w="2000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
            <w:p><w:r><w:t>StartCell</w:t></w:r></w:p>
          </w:tc>
          <w:tc><w:p><w:r><w:t>RightTop</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge/></w:tcPr>
            <w:p><w:r><w:t>ShouldBeHidden</w:t></w:r></w:p>
          </w:tc>
          <w:tc><w:p><w:r><w:t>RightBottom</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });

    const parsed = parseTtf(FONTS.regular);
    const hidden = [...'ShouldBeHidden'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const start = [...'StartCell'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const text = asLatin1(pdf);

    const hiddenHex = hidden.map((g) => g.toString(16).padStart(4, '0').toUpperCase()).join('');
    const startHex = start.map((g) => g.toString(16).padStart(4, '0').toUpperCase()).join('');
    expect(text).toContain(`<${startHex}> Tj`);
    expect(text).not.toContain(`<${hiddenHex}> Tj`);
  });

  it('does not over-justify a short single-line paragraph (last line stays left)', () => {
    const docx = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="both"/></w:pPr>', runs: [{ text: 'Short.' }] },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const btMatch = text.match(/BT\n([\s\S]*?)\nET/);
    expect(btMatch).not.toBeNull();
    const tmInBt = (btMatch![1]!.match(/Tm/g) ?? []).length;
    expect(tmInBt).toBe(1);
  });

  it('splits a table row taller than the page into chunks across pages', () => {
    // 80 paragraphs in one cell exceeds A4 content height (~698pt) at typical
    // 14pt line height. Expect at least two pages with continuation borders.
    const paragraphs = Array.from(
      { length: 80 },
      (_, i) => `<w:p><w:r><w:t>Line${i}</w:t></w:r></w:p>`,
    ).join('');
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
        <w:tr>
          <w:tc>${paragraphs}</w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const pageCount = (text.match(/\/Type \/Page\b/g) ?? []).filter(
      (m) => !m.includes('Pages'),
    ).length;
    expect(pageCount).toBeGreaterThan(1);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    // First and last lines must both render (row split, not clipping).
    expect(text).toContain(`<${hexOf('Line0')}> Tj`);
    expect(text).toContain(`<${hexOf('Line79')}> Tj`);
  });

  it('honors paragraph alignment center and right', () => {
    const docx = buildRichDocx([
      {
        pPrXml: '<w:pPr><w:jc w:val="center"/></w:pPr>',
        runs: [{ text: 'Centered' }],
      },
      {
        pPrXml: '<w:pPr><w:jc w:val="right"/></w:pPr>',
        runs: [{ text: 'Right' }],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const tmMatches = [...text.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm/g)];
    expect(tmMatches.length).toBeGreaterThanOrEqual(2);
    const xPositions = tmMatches.map((m) => Number(m[1]));
    // First line is centered → x > marginLeft (72)
    expect(xPositions[0]!).toBeGreaterThan(72);
    // Second line is right-aligned → even further right than centered
    expect(xPositions[1]!).toBeGreaterThan(xPositions[0]!);
  });

  it('renders a nested table inside a cell (table-in-cell)', () => {
    // A w:tbl nested in a w:tc. Previously the cell layout skipped non-paragraph
    // content, so the nested table (and all its text) was dropped — the POI
    // 60329.docx pattern (0/4812 chars). Now it lays out and renders.
    const body = `
      <w:tbl>
        <w:tblPr><w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders></w:tblPr>
        <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:p><w:r><w:t>OUTERCELL</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr><w:tblBorders>
              <w:top w:val="single" w:sz="4" w:color="000000"/>
              <w:insideV w:val="single" w:sz="4" w:color="000000"/>
            </w:tblBorders></w:tblPr>
            <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:p><w:r><w:t>NESTEDA</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>NESTEDB</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:tc></w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) =>
          parsed.glyphForCodepoint(c.codePointAt(0)!).toString(16).padStart(4, '0').toUpperCase(),
        )
        .join('');
    expect(text).toContain(`<${hexOf('OUTERCELL')}> Tj`); // outer cell paragraph
    expect(text).toContain(`<${hexOf('NESTEDA')}> Tj`); // nested cell 1 (was lost)
    expect(text).toContain(`<${hexOf('NESTEDB')}> Tj`); // nested cell 2 (was lost)
  });

  it('measures table auto-layout with per-family fonts (was: bare-variant lookup crash)', () => {
    const reg = (b: Uint8Array): FontRegistry => FontRegistry.fromBytes({ regular: b });
    const registriesByFamily: ReadonlyMap<FamilyKey, FontRegistry> = new Map([
      ['arimo', reg(FONTS.regular)],
      ['tinos', reg(FONTS.bold)],
    ]);
    // Auto-layout table (no explicit grid widths) forces measureSingleLine,
    // which used to look fontResources up by bare variant and crash when the
    // keys are per-family ('roboto:regular', …).
    const body =
      `<w:tbl><w:tr>` +
      `<w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>CELLSANS</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>CELLSERIF</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`;
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body), {
      fonts: { regular: FONTS.regular },
      registriesByFamily,
    });
    // Pre-fix this crashed (fontResources.get('regular') is undefined when
    // keys are per-family). Surviving + embedding BOTH family stand-ins
    // proves each cell measured and rendered with its own family.
    const baseFonts = new Set(
      [...asLatin1(pdf).matchAll(/\/BaseFont \/[A-Z]{6}\+([A-Za-z-]+)/g)].map((m) => m[1]),
    );
    expect(baseFonts.has('Roboto-Regular')).toBe(true); // Arial → arimo
    expect(baseFonts.has('Roboto-Bold')).toBe(true); // Times → tinos stand-in
  });

  it('resolves the substitute font per run (sans / serif / mono families)', () => {
    // Distinct stand-in fonts per family (Roboto variants carry distinct
    // BaseFont names) so we can prove each run picked its OWN family by w:ascii.
    const reg = (b: Uint8Array): FontRegistry => FontRegistry.fromBytes({ regular: b });
    const registriesByFamily: ReadonlyMap<FamilyKey, FontRegistry> = new Map([
      ['arimo', reg(FONTS.regular)],
      ['tinos', reg(FONTS.bold)],
      ['cousine', reg(FONTS.italic)],
    ]);
    const body =
      `<w:p>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>SANS</w:t></w:r>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>SERIF</w:t></w:r>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>MONO</w:t></w:r>` +
      `</w:p>`;
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body), {
      fonts: { regular: FONTS.regular },
      registriesByFamily,
    });
    const baseFonts = new Set(
      [...asLatin1(pdf).matchAll(/\/BaseFont \/[A-Z]{6}\+([A-Za-z-]+)/g)].map((m) => m[1]),
    );
    // Each run resolved to its family's stand-in → three distinct fonts embedded.
    expect(baseFonts.has('Roboto-Regular')).toBe(true); // SANS → arimo
    expect(baseFonts.has('Roboto-Bold')).toBe(true); // SERIF → tinos stand-in
    expect(baseFonts.has('Roboto-Italic')).toBe(true); // MONO → cousine stand-in
  });
});
