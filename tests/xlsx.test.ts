import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { convertXlsxToPdfSync } from '@/core/converter';
import { parseTtf } from '@/core/font';
import {
  formatCellRef,
  parseAreaRef,
  parseCellRef,
  parseSharedStrings,
  parseTitleRowRange,
  parseWorkbook,
  parseWorksheet,
} from '@/excel';
import { OpcPackage } from '@/core/opc';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const BOLD = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf')));

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('cell reference (A1 ↔ {row, col})', () => {
  it('round-trips A1 / B2 / Z1 / AA1 / AB12', () => {
    const cases: Array<[string, { column: number; row: number }]> = [
      ['A1', { column: 0, row: 0 }],
      ['B2', { column: 1, row: 1 }],
      ['Z1', { column: 25, row: 0 }],
      ['AA1', { column: 26, row: 0 }],
      ['AB12', { column: 27, row: 11 }],
      ['BA100', { column: 52, row: 99 }],
    ];
    for (const [ref, expected] of cases) {
      expect(parseCellRef(ref)).toEqual(expected);
      expect(formatCellRef(expected)).toBe(ref);
    }
  });

  it('throws on invalid refs', () => {
    expect(() => parseCellRef('')).toThrow();
    expect(() => parseCellRef('A')).toThrow();
    expect(() => parseCellRef('1A')).toThrow();
  });
});

describe('xlsx parsers', () => {
  it('parses sheets from workbook.xml', () => {
    const xlsx = buildXlsx([['A', 'B']]);
    const pkg = OpcPackage.open(xlsx);
    const { sheets, date1904 } = parseWorkbook(pkg.requirePart('xl/workbook.xml'));
    expect(sheets).toHaveLength(1);
    expect(sheets[0]!.name).toBe('Sheet1');
    expect(sheets[0]!.relationshipId).toBe('rId1');
    expect(date1904).toBe(false);
  });

  it('parses shared strings', () => {
    const xlsx = buildXlsx([
      ['Hello', 'World'],
      ['Hello', 42],
    ]);
    const pkg = OpcPackage.open(xlsx);
    const strs = parseSharedStrings(pkg.requirePart('xl/sharedStrings.xml')).texts;
    expect(strs).toEqual(['Hello', 'World']);
  });

  it('parses row heights, pageMargins and pageSetup from worksheet.xml', () => {
    const xlsx = buildXlsx({
      rows: [['A'], ['B'], ['C']],
      rowHeights: [{ row: 0, heightPt: 30, customHeight: true }],
      pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      pageSetup: { paperSize: 1, orientation: 'landscape' },
    });
    const pkg = OpcPackage.open(xlsx);
    const ws = parseWorksheet(pkg.requirePart('xl/worksheets/sheet1.xml'));
    expect(ws.rowHeights).toEqual([{ row: 0, heightPt: 30, customHeight: true }]);
    expect(ws.pageMargins).toEqual({
      leftInches: 0.7,
      rightInches: 0.7,
      topInches: 0.75,
      bottomInches: 0.75,
      headerInches: 0.3,
      footerInches: 0.3,
    });
    expect(ws.pageSetup).toEqual({ paperSize: 1, orientation: 'landscape' });
  });

  it('parses workbookPr date1904 flag', () => {
    const xlsxA = buildXlsx({ rows: [['x']] });
    const a = parseWorkbook(OpcPackage.open(xlsxA).requirePart('xl/workbook.xml'));
    expect(a.date1904).toBe(false);

    const xlsxB = buildXlsx({ rows: [['x']], date1904: true });
    const b = parseWorkbook(OpcPackage.open(xlsxB).requirePart('xl/workbook.xml'));
    expect(b.date1904).toBe(true);
  });

  it('parses worksheet cells with types', () => {
    const xlsx = buildXlsx([
      ['Name', 'Value', 'Active'],
      ['Alice', 42, true],
      ['Bob', 3.14, false],
    ]);
    const pkg = OpcPackage.open(xlsx);
    const ws = parseWorksheet(pkg.requirePart('xl/worksheets/sheet1.xml'));
    expect(ws.maxRow).toBe(2);
    expect(ws.maxColumn).toBe(2);
    expect(ws.cells).toHaveLength(9);
    const a1 = ws.cells.find((c) => c.row === 0 && c.column === 0);
    const b2 = ws.cells.find((c) => c.row === 1 && c.column === 1);
    const c2 = ws.cells.find((c) => c.row === 1 && c.column === 2);
    expect(a1!.type).toBe('s');
    expect(b2!.type).toBe('n');
    expect(b2!.rawValue).toBe('42');
    expect(c2!.type).toBe('b');
    expect(c2!.rawValue).toBe('1');
  });

  // Regression: some producers (e.g. Haansoft HCell) write SpreadsheetML with
  // an explicit `x:` namespace prefix instead of the default namespace. Before
  // the parsers stripped prefixes, workbook.xml yielded 0 sheets → the whole
  // file failed with "xlsx has no sheets" (found across 6 POI corpus files).
  it('tolerates an explicit x: namespace prefix (workbook / sst / worksheet)', () => {
    const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
    const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // workbook.xml: x: on elements, r: on the relationship id.
    const { sheets } = parseWorkbook(
      enc(
        `<x:workbook xmlns:r="${R}" xmlns:x="${X}">` +
          `<x:sheets><x:sheet name="Features" sheetId="1" r:id="rId7"/></x:sheets></x:workbook>`,
      ),
    );
    expect(sheets).toHaveLength(1);
    expect(sheets[0]!.name).toBe('Features');
    expect(sheets[0]!.relationshipId).toBe('rId7'); // r:id resolved via prefix strip

    // sharedStrings.xml with x: prefix.
    expect(
      parseSharedStrings(
        enc(
          `<x:sst xmlns:x="${X}"><x:si><x:t>Hi</x:t></x:si><x:si><x:t>There</x:t></x:si></x:sst>`,
        ),
      ).texts,
    ).toEqual(['Hi', 'There']);

    // worksheet.xml with x: prefix: one shared-string cell + one number cell.
    const ws = parseWorksheet(
      enc(
        `<x:worksheet xmlns:x="${X}"><x:sheetData><x:row r="1">` +
          `<x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:v>42</x:v></x:c>` +
          `</x:row></x:sheetData></x:worksheet>`,
      ),
    );
    expect(ws.cells).toHaveLength(2);
    const a1 = ws.cells.find((c) => c.row === 0 && c.column === 0);
    const b1 = ws.cells.find((c) => c.row === 0 && c.column === 1);
    expect(a1!.type).toBe('s');
    expect(a1!.rawValue).toBe('0');
    expect(b1!.type).toBe('n');
    expect(b1!.rawValue).toBe('42');
  });
});

describe('convertXlsxToPdfSync end-to-end', () => {
  it('renders a small grid as a bordered table', () => {
    const xlsx = buildXlsx([
      ['Region', 'Revenue', 'Growth'],
      ['Moscow', 1200000, true],
      ['SPb', 720000, false],
    ]);
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text.startsWith('%PDF-1.7\n')).toBe(true);
    expect(text).toContain('/Type /Page');

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');

    expect(text).toContain(`<${hexOf('Region')}> Tj`);
    expect(text).toContain(`<${hexOf('Moscow')}> Tj`);
    expect(text).toContain(`<${hexOf('1200000')}> Tj`);
    expect(text).toContain(`<${hexOf('TRUE')}> Tj`);
  });

  it('skips empty cells gracefully', () => {
    const xlsx = buildXlsx([
      ['A', null, 'C'],
      [null, 'B', null],
    ]);
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.byteLength).toBeGreaterThan(100);
  });

  it('applies font properties (bold + colour) from cellXfs', () => {
    const stylesXml = `
      <fonts count="2">
        <font><sz val="11"/><name val="Calibri"/></font>
        <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
      </fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [
        [{ value: 'Title', styleIndex: 1 }, 'Plain'],
        ['Body', 'More'],
      ],
      stylesXml,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: { regular: FONTS.regular, bold: BOLD } });
    const text = asLatin1(pdf);

    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Bold\b/);
    // Red colour: 255/255 → "1" component for r in rg operator.
    expect(text).toMatch(/\b1 0 0 rg\b/);
  });

  it('applies built-in number format (thousands separator)', () => {
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="3" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [[{ value: 1234567, styleIndex: 1 }]],
      stylesXml,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    expect(text).toContain(`<${hexOf('1,234,567')}> Tj`);
  });

  it('applies custom number format (#,##0.00) to currency-like value', () => {
    const stylesXml = `
      <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0.00"/>
      </numFmts>
      <fonts count="1"><font/></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [[{ value: 50000.5, styleIndex: 1 }]],
      stylesXml,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    expect(text).toContain(`<${hexOf('50,000.50')}> Tj`);
  });

  it('paints cell background fill as a re + f rectangle', () => {
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [[{ value: 'Header', styleIndex: 1 }, 'Plain']],
      stylesXml,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // Filled rectangle with #4472C4 → r/g/b roughly (0.267, 0.447, 0.769).
    expect(text).toMatch(/0\.267 0\.447 0\.769 rg/);
    expect(text).toMatch(/[\d.]+ [\d.]+ [\d.]+ [\d.]+ re\nf/);
  });

  it('applies per-cell border style + colour from xf.borderId', () => {
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="2">
        <border/>
        <border>
          <left style="thick"><color rgb="FFD92020"/></left>
          <right style="thick"><color rgb="FFD92020"/></right>
          <top style="thick"><color rgb="FFD92020"/></top>
          <bottom style="thick"><color rgb="FFD92020"/></bottom>
        </border>
      </borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [[{ value: 'Box', styleIndex: 1 }]],
      stylesXml,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // Thick border ~12/8 pt = 1.5 pt width
    expect(text).toMatch(/1\.5 w/);
    // Stroke colour FFD92020 → ~(0.851, 0.125, 0.125)
    expect(text).toMatch(/0\.851 0\.125 0\.125 RG/);
  });

  it('renders every sheet on its own page', () => {
    const xlsx = buildXlsx({
      sheets: [
        {
          name: 'Income',
          rows: [
            ['Year', 'Total'],
            ['2024', 100],
          ],
        },
        {
          name: 'Expenses',
          rows: [
            ['Year', 'Total'],
            ['2024', 60],
          ],
        },
        {
          name: 'Summary',
          rows: [
            ['Year', 'Profit'],
            ['2024', 40],
          ],
        },
      ],
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const pageCount = (text.match(/\/Type \/Page\b/g) ?? []).filter(
      (m) => !m.includes('Pages'),
    ).length;
    expect(pageCount).toBeGreaterThanOrEqual(3);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    // Each sheet's distinct data renders (one sheet per page).
    expect(text).toContain(`<${hexOf('100')}> Tj`);
    expect(text).toContain(`<${hexOf('60')}> Tj`);
    expect(text).toContain(`<${hexOf('40')}> Tj`);
    // The sheet NAMES are NOT printed — Calc/Excel `--convert-to pdf` emit them
    // nowhere; a synthetic title is pure extra text vs the print golden.
    expect(text).not.toContain(`<${hexOf('Income')}> Tj`);
    expect(text).not.toContain(`<${hexOf('Summary')}> Tj`);
  });

  it('honours per-row heights from <row ht customHeight="1">', () => {
    // 6 rows total. Without explicit heights every row is roughly the body's
    // computed text height (~ font size * 1.2 = 13.2pt). Pinning row 1 to
    // 60pt should expand at least one row by ~47pt vs the unpinned variant.
    const rows = [['A'], ['B'], ['C'], ['D'], ['E'], ['F']];
    const baseline = convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS });
    const expanded = convertXlsxToPdfSync(
      buildXlsx({ rows, rowHeights: [{ row: 1, heightPt: 60, customHeight: true }] }),
      { fonts: FONTS },
    );
    const tallerBy = expanded.byteLength - baseline.byteLength;
    // The expanded PDF carries the same number of glyphs, so any byte-size
    // diff is irrelevant. Instead check that the actual text baselines have
    // shifted by parsing every Tm y-coordinate. The baseline of row 2+ should
    // sit ~47pt lower in the expanded variant.
    void tallerBy;
    const baselinesOf = (pdf: Uint8Array): Array<number> =>
      Array.from(asLatin1(pdf).matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/g)).map((m) => Number(m[1]));
    const baseY = baselinesOf(baseline);
    const expandedY = baselinesOf(expanded);
    // The last grid baseline (row F) — body grows top-down so we look at the
    // minimum y across cell text baselines (smallest = closest to page bottom).
    const minBase = Math.min(...baseY);
    const minExpanded = Math.min(...expandedY);
    // Expanded variant must place its bottom text lower (smaller y) by at least
    // ~30pt (allowing for slight padding adjustments).
    expect(minBase - minExpanded).toBeGreaterThan(30);
  });

  it('uses pageSetup paperSize=1 (Letter) for MediaBox', () => {
    const xlsx = buildXlsx({
      rows: [['x']],
      pageSetup: { paperSize: 1, orientation: 'portrait' },
      pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);
    // Letter portrait = 612 × 792 pt.
    expect(text).toMatch(/\/MediaBox \[0 0 612 792\]/);
  });

  it('uses landscape orientation by swapping width and height', () => {
    const xlsx = buildXlsx({
      rows: [['x']],
      pageSetup: { paperSize: 9, orientation: 'landscape' },
      pageMargins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 },
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);
    // A4 landscape: 16838twips × 11906twips = 841.9pt × 595.3pt (rounded).
    expect(text).toMatch(/\/MediaBox \[0 0 841\.9 595\.3\]/);
  });

  it('applies date1904 epoch when workbookPr date1904="1" is set', () => {
    // Excel 1900-mode serial 0 = 1900-01-01 (after the Lotus bug). The same
    // calendar date in 1904 mode is serial -1462. Picking serial 0 in each
    // mode demonstrates the offset directly:
    //   1900 mode: 1899-12-30 (rendered as "12/30/1899" by m/d/yyyy)
    //   1904 mode: 1904-01-01 (rendered as "1/1/1904")
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>`;
    const xlsx = buildXlsx({
      rows: [[{ value: 0, styleIndex: 1 }]],
      stylesXml,
      date1904: true,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');
    expect(text).toContain(`<${hexOf('1/1/1904')}> Tj`);
  });

  it('handles merged cells by emitting gridSpan on the origin', () => {
    const xlsx = buildXlsx({
      rows: [
        ['Title', null, null],
        ['A', 'B', 'C'],
      ],
      mergeRefs: ['A1:C1'],
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    // The merged cell should still be emitted (with text 'Title') and only
    // once per row. If we naively wrote 3 cells in row 0, we'd see two extra
    // empty cell rectangles. Easier check: just confirm the PDF builds and
    // renders the title text exactly once.
    const text = asLatin1(pdf);
    const parsed = parseTtf(FONTS.regular);
    const titleHex = [...'Title']
      .map((c) =>
        parsed.glyphForCodepoint(c.codePointAt(0)!).toString(16).padStart(4, '0').toUpperCase(),
      )
      .join('');
    const matches = text.match(new RegExp(`<${titleHex}> Tj`, 'g')) ?? [];
    expect(matches).toHaveLength(1);
  });
});

// A hand-built minimal xlsx (fflate) lets us seed cells the dense buildXlsx
// fixture can't — here a stray empty styled cell far out in the sheet.
function rawXlsx(sheetDataXml: string): Uint8Array {
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    'xl/workbook.xml': strToU8(
      `<?xml version="1.0"?><workbook xmlns="${main}" xmlns:r="${rel}">` +
        `<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<?xml version="1.0"?><worksheet xmlns="${main}"><sheetData>${sheetDataXml}</sheetData></worksheet>`,
    ),
  });
}

describe('xlsx robustness', () => {
  it('bounds the grid to the used range (stray far cells do not explode it)', () => {
    // Content in A1:B1 plus an EMPTY styled cell seeded at XFD1000 (column
    // 16384, row 1000) — the whole-row-styling artifact that OOM'd CVLKRA-KYC.
    // The grid must trim to the used range (1×2), not materialize 16384×1000.
    const CONTENT = `<row r="1"><c r="A1"><v>11</v></c><c r="B1"><v>22</v></c></row>`;
    const clean = convertXlsxToPdfSync(rawXlsx(CONTENT), { fonts: FONTS });
    const stray = convertXlsxToPdfSync(
      rawXlsx(CONTENT + `<row r="1000"><c r="XFD1000" s="0"/></row>`),
      { fonts: FONTS },
    );
    const text = asLatin1(stray);
    expect(text.startsWith('%PDF')).toBe(true);
    // The stray empty far cell is trimmed to the used range → output is
    // byte-for-byte identical to the no-stray render, not a 16384×1000 grid
    // (which would also exhaust memory).
    expect(stray.length).toBe(clean.length);
    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) =>
          parsed.glyphForCodepoint(c.codePointAt(0)!).toString(16).padStart(4, '0').toUpperCase(),
        )
        .join('');
    expect(text).toContain(`<${hexOf('11')}> Tj`); // real content kept
  });

  it('caps over-long cell strings to the 32 767-char limit (DoS guard)', () => {
    // A crafted ~1 MB string referenced by thousands of cells would be
    // shaped/measured per cell and hang the renderer (poc-shared-strings.xlsx).
    // Cap each cell string to Excel's hard limit on both code paths.
    const big = 'A'.repeat(100_000);
    const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
    const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const shared = parseSharedStrings(enc(`<sst xmlns="${M}"><si><t>${big}</t></si></sst>`)).texts;
    expect(shared[0]!.length).toBe(32_767);
    const ws = parseWorksheet(
      enc(
        `<worksheet xmlns="${M}"><sheetData><row r="1">` +
          `<c r="A1" t="inlineStr"><is><t>${big}</t></is></c>` +
          `</row></sheetData></worksheet>`,
      ),
    );
    expect(ws.cells[0]!.inlineText!.length).toBe(32_767);
  });
});

const M_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('xlsx print-model parsing', () => {
  it('parses pageSetup scale/fit, fitToPage, printOptions and row/col breaks', () => {
    const ws = parseWorksheet(
      enc(
        `<worksheet xmlns="${M_NS}">` +
          `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` +
          `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` +
          `<printOptions gridLines="1" horizontalCentered="1"/>` +
          `<pageSetup paperSize="9" orientation="landscape" scale="80" fitToWidth="1" fitToHeight="0"/>` +
          `<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks>` +
          `<colBreaks count="1"><brk id="3" man="1"/></colBreaks>` +
          `</worksheet>`,
      ),
    );
    expect(ws.fitToPage).toBe(true);
    expect(ws.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
    expect(ws.pageSetup).toEqual({
      paperSize: 9,
      orientation: 'landscape',
      scale: 80,
      fitToWidth: 1,
      fitToHeight: 0,
    });
    expect(ws.rowBreaks).toEqual([10]);
    expect(ws.colBreaks).toEqual([3]);
  });

  it('fitToHeight shrinks a tall sheet onto fewer pages', () => {
    const rows = Array.from({ length: 120 }, (_, i) => [`row ${i}`]);
    const countPages = (b: Uint8Array): number =>
      (asLatin1(b).match(/\/Type \/Page(?![s])/g) ?? []).length;
    const tall = convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS });
    const fit = convertXlsxToPdfSync(
      buildXlsx({ rows, fitToPage: true, pageSetup: { fitToWidth: 0, fitToHeight: 1 } }),
      { fonts: FONTS },
    );
    // The unscaled sheet spans several pages; fitToHeight="1" shrinks it down.
    expect(countPages(tall)).toBeGreaterThan(1);
    expect(countPages(fit)).toBeLessThan(countPages(tall));
  });

  it('leaves the print model undefined when no print elements are present', () => {
    const ws = parseWorksheet(
      enc(
        `<worksheet xmlns="${M_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
      ),
    );
    expect(ws.fitToPage).toBeUndefined();
    expect(ws.printOptions).toBeUndefined();
    expect(ws.rowBreaks).toBeUndefined();
    expect(ws.colBreaks).toBeUndefined();
  });

  it('parses workbook definedNames with localSheetId', () => {
    const wb = parseWorkbook(
      enc(
        `<workbook xmlns="${M_NS}" xmlns:r="${R_NS}">` +
          `<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>` +
          `<definedNames>` +
          `<definedName name="_xlnm.Print_Area" localSheetId="0">S1!$A$1:$B$2</definedName>` +
          `<definedName name="MyGlobal">S1!$Z$9</definedName>` +
          `</definedNames>` +
          `</workbook>`,
      ),
    );
    expect(wb.definedNames).toEqual([
      { name: '_xlnm.Print_Area', localSheetId: 0, value: 'S1!$A$1:$B$2' },
      { name: 'MyGlobal', value: 'S1!$Z$9' },
    ]);
  });
});

describe('parseAreaRef', () => {
  it('resolves qualified, absolute, single-cell and multi-area refs', () => {
    expect(parseAreaRef('Sheet1!$A$1:$D$20')).toEqual({
      startColumn: 0,
      startRow: 0,
      endColumn: 3,
      endRow: 19,
    });
    expect(parseAreaRef("'My Sheet'!$B$2")).toEqual({
      startColumn: 1,
      startRow: 1,
      endColumn: 1,
      endRow: 1,
    });
    // Bounding box across two disjoint areas.
    expect(parseAreaRef('S!$A$1:$B$2,S!$D$1:$E$3')).toEqual({
      startColumn: 0,
      startRow: 0,
      endColumn: 4,
      endRow: 2,
    });
  });

  it('returns undefined for empty / whole-row / whole-column refs', () => {
    expect(parseAreaRef('')).toBeUndefined();
    expect(parseAreaRef('S!$A:$A')).toBeUndefined();
    expect(parseAreaRef('S!$1:$1')).toBeUndefined();
  });
});

describe('implicit cell/row positions (§18.3.1.4 — r= optional)', () => {
  it('assigns positions by document order when rows/cells omit r=', () => {
    const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    // No r= anywhere: 2 rows × 2 cells → A1,B1,A2,B2 by order.
    const ws = parseWorksheet(
      enc(
        `<worksheet xmlns="${M}"><sheetData>` +
          `<row><c t="inlineStr"><is><t>A1</t></is></c><c t="inlineStr"><is><t>B1</t></is></c></row>` +
          `<row><c t="inlineStr"><is><t>A2</t></is></c><c t="inlineStr"><is><t>B2</t></is></c></row>` +
          `</sheetData></worksheet>`,
      ),
    );
    expect(ws.cells).toHaveLength(4);
    const at = (r: number, c: number) =>
      ws.cells.find((x) => x.row === r && x.column === c)?.inlineText;
    expect(at(0, 0)).toBe('A1');
    expect(at(0, 1)).toBe('B1');
    expect(at(1, 0)).toBe('A2');
    expect(at(1, 1)).toBe('B2');
    expect(ws.maxRow).toBe(1);
    expect(ws.maxColumn).toBe(1);
  });

  it('honours explicit r= and resumes implicit numbering after it', () => {
    const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    // Row pinned to r="3"; cell pinned to C → next cell resumes at D.
    const ws = parseWorksheet(
      enc(
        `<worksheet xmlns="${M}"><sheetData>` +
          `<row r="3"><c r="C3"><v>1</v></c><c><v>2</v></c></row>` +
          `</sheetData></worksheet>`,
      ),
    );
    const c1 = ws.cells.find((x) => x.rawValue === '1')!;
    const c2 = ws.cells.find((x) => x.rawValue === '2')!;
    expect([c1.row, c1.column]).toEqual([2, 2]); // C3 → row 2, col 2
    expect([c2.row, c2.column]).toEqual([2, 3]); // resumes at D3
  });
});

describe('parseTitleRowRange', () => {
  it('extracts the repeated row range, ignoring column-title parts', () => {
    expect(parseTitleRowRange('Sheet1!$1:$2')).toEqual({ startRow: 0, endRow: 1 });
    expect(parseTitleRowRange('Sheet1!$1:$1')).toEqual({ startRow: 0, endRow: 0 });
    expect(parseTitleRowRange("'My Sheet'!$A:$B,'My Sheet'!$1:$3")).toEqual({
      startRow: 0,
      endRow: 2,
    });
    expect(parseTitleRowRange('Sheet1!$A:$B')).toBeUndefined(); // columns only
    expect(parseTitleRowRange('')).toBeUndefined();
  });
});

describe('xlsx print-model rendering', () => {
  const parsed = parseTtf(FONTS.regular);
  const hexOf = (s: string): string =>
    [...s]
      .map((c) =>
        parsed.glyphForCodepoint(c.codePointAt(0)!).toString(16).padStart(4, '0').toUpperCase(),
      )
      .join('');

  it('suppresses cell gridlines by default; draws them with printOptions gridLines="1"', () => {
    const rows = [
      ['A', 'B'],
      ['C', 'D'],
    ];
    const plain = asLatin1(convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS }));
    const gridded = asLatin1(
      convertXlsxToPdfSync(buildXlsx({ rows, printOptions: { gridLines: true } }), {
        fonts: FONTS,
      }),
    );
    // The synthetic grid strokes thin black lines → "0 0 0 RG" (stroke colour).
    // Text fills use lowercase "rg", so its absence proves no grid was drawn.
    expect(plain).not.toContain('0 0 0 RG');
    expect(gridded).toContain('0 0 0 RG');
    // Cell text is identical either way.
    expect(plain).toContain(`<${hexOf('A')}> Tj`);
    expect(gridded).toContain(`<${hexOf('A')}> Tj`);
  });

  it('clips rendering to _xlnm.Print_Area (A1-anchored)', () => {
    const xlsx = buildXlsx({
      rows: [
        ['In1', 'In2', 'Out1'],
        ['In3', 'In4', 'Out2'],
        ['Out3', 'Out4', 'Out5'],
      ],
      definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Sheet1!$A$1:$B$2' }],
    });
    const text = asLatin1(convertXlsxToPdfSync(xlsx, { fonts: FONTS }));
    const has = (s: string): boolean => text.includes(`<${hexOf(s)}> Tj`);
    expect(has('In1')).toBe(true);
    expect(has('In4')).toBe(true);
    expect(has('Out1')).toBe(false);
    expect(has('Out3')).toBe(false);
    expect(has('Out5')).toBe(false);
  });

  it('renders a print area that does not start at A1 (offset window)', () => {
    const xlsx = buildXlsx({
      rows: [
        ['o', 'o', 'o', 'o'],
        ['o', 'B2', 'C2', 'o'],
        ['o', 'B3', 'C3', 'o'],
        ['o', 'o', 'o', 'o'],
      ],
      definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Sheet1!$B$2:$C$3' }],
    });
    const text = asLatin1(convertXlsxToPdfSync(xlsx, { fonts: FONTS }));
    const count = (s: string): number =>
      (text.match(new RegExp(`<${hexOf(s)}> Tj`, 'g')) ?? []).length;
    expect(count('B2')).toBe(1);
    expect(count('C2')).toBe(1);
    expect(count('B3')).toBe(1);
    expect(count('C3')).toBe(1);
    // Every 'o' filler lies outside B2:C3 → none rendered.
    expect(count('o')).toBe(0);
  });

  it('ignores a print area scoped to a different sheet', () => {
    // Print_Area is localSheetId=1 (a second sheet that does not exist here),
    // so sheet 0 renders in full.
    const xlsx = buildXlsx({
      rows: [
        ['P', 'Q'],
        ['R', 'S'],
      ],
      definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 1, value: 'Other!$A$1:$A$1' }],
    });
    const text = asLatin1(convertXlsxToPdfSync(xlsx, { fonts: FONTS }));
    for (const v of ['P', 'Q', 'R', 'S']) {
      expect(text.includes(`<${hexOf(v)}> Tj`)).toBe(true);
    }
  });

  // Font sizes emitted as "/Fn <sizePt> Tf" in the content stream.
  const tfSizes = (pdf: Uint8Array): Array<number> =>
    Array.from(asLatin1(pdf).matchAll(/\/[A-Za-z0-9]+ ([\d.]+) Tf/g)).map((m) => Number(m[1]));

  it('shrinks cell fonts for an explicit <pageSetup scale="50">', () => {
    const rows = [
      ['Hello', 'World'],
      ['Foo', 'Bar'],
    ];
    const plain = tfSizes(convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS }));
    const scaled = tfSizes(
      convertXlsxToPdfSync(buildXlsx({ rows, pageSetup: { scale: 50 } }), { fonts: FONTS }),
    );
    // Default body text is 11pt → 5.5pt at 50%. The sheet-title (14pt) is not
    // scaled, so 5.5 appears only in the scaled variant.
    expect(scaled).toContain(5.5);
    expect(plain).not.toContain(5.5);
  });

  it('shrinks fonts under fitToPage when the columns overflow the page width', () => {
    // 6 columns × 40 chars ≈ 25 200 twips ≫ ~9 026 twip content width → fit to
    // one page wide forces a strong uniform shrink.
    const columns = Array.from({ length: 6 }, (_, i) => ({
      min: i + 1,
      max: i + 1,
      widthChars: 40,
    }));
    const rows = [['a', 'b', 'c', 'd', 'e', 'f']];
    const plain = tfSizes(convertXlsxToPdfSync(buildXlsx({ rows, columns }), { fonts: FONTS }));
    const fitted = tfSizes(
      convertXlsxToPdfSync(
        buildXlsx({ rows, columns, fitToPage: true, pageSetup: { fitToWidth: 1, fitToHeight: 1 } }),
        { fonts: FONTS },
      ),
    );
    // The smallest font (the body cells) must shrink below the unscaled body.
    expect(Math.min(...fitted)).toBeLessThan(Math.min(...plain));
  });

  it('treats scale=100 as a true no-op (byte-identical output)', () => {
    const rows = [['X', 'Y']];
    const a = convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS });
    const b = convertXlsxToPdfSync(buildXlsx({ rows, pageSetup: { scale: 100 } }), {
      fonts: FONTS,
    });
    expect(asLatin1(b)).toBe(asLatin1(a));
  });

  // A tall sheet that overflows one page, with a distinctive header row.
  const tallRows = (): Array<Array<string | number>> => {
    const rows: Array<Array<string | number>> = [['HDR', 'COL2']];
    for (let i = 0; i < 90; i++) rows.push([`r${i}`, i]);
    return rows;
  };
  const hdrCount = (pdf: Uint8Array): number =>
    (asLatin1(pdf).match(new RegExp(`<${hexOf('HDR')}> Tj`, 'g')) ?? []).length;

  it('repeats _xlnm.Print_Titles header rows on every continuation page', () => {
    const pdf = convertXlsxToPdfSync(
      buildXlsx({
        rows: tallRows(),
        definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 0, value: 'Sheet1!$1:$1' }],
      }),
      { fonts: FONTS },
    );
    // The sheet spills onto ≥2 pages, so the header row's text repeats.
    const pages = (asLatin1(pdf).match(/\/Type \/Page\b/g) ?? []).filter(
      (m) => !m.includes('Pages'),
    ).length;
    expect(pages).toBeGreaterThanOrEqual(2);
    expect(hdrCount(pdf)).toBeGreaterThanOrEqual(2);
  });

  it('does not repeat the first row without Print_Titles', () => {
    const pdf = convertXlsxToPdfSync(buildXlsx({ rows: tallRows() }), { fonts: FONTS });
    expect(hdrCount(pdf)).toBe(1);
  });

  const pageCount = (pdf: Uint8Array): number =>
    (asLatin1(pdf).match(/\/Type \/Page\b/g) ?? []).filter((m) => !m.includes('Pages')).length;

  it('forces a page break at a manual <rowBreaks> entry', () => {
    // A 5-row sheet fits one page; a manual break before row index 2 splits it.
    const rows = [['r0'], ['r1'], ['r2'], ['r3'], ['r4']];
    const baseline = convertXlsxToPdfSync(buildXlsx({ rows }), { fonts: FONTS });
    const broken = convertXlsxToPdfSync(buildXlsx({ rows, rowBreaks: [2] }), { fonts: FONTS });
    expect(pageCount(baseline)).toBe(1);
    expect(pageCount(broken)).toBe(2);
    // Content is preserved across the break.
    expect(asLatin1(broken)).toContain(`<${hexOf('r2')}> Tj`);
  });

  it('centers the table with <printOptions horizontalCentered="1">', () => {
    // One styled (filled) cell so we can read the table's left edge from the
    // fill rectangle's x — unaffected by the always-centered sheet title.
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
      </cellXfs>`;
    const rows = [
      [
        { value: 'A', styleIndex: 1 },
        { value: 'B', styleIndex: 1 },
      ],
    ];
    const minRectX = (pdf: Uint8Array): number => {
      const xs = Array.from(asLatin1(pdf).matchAll(/([\d.]+) [\d.]+ [\d.]+ [\d.]+ re/g)).map((m) =>
        Number(m[1]),
      );
      return Math.min(...xs);
    };
    const left = convertXlsxToPdfSync(buildXlsx({ rows, stylesXml }), { fonts: FONTS });
    const centered = convertXlsxToPdfSync(
      buildXlsx({ rows, stylesXml, printOptions: { horizontalCentered: true } }),
      { fonts: FONTS },
    );
    // The narrow table shifts well right of the left margin when centered.
    expect(minRectX(centered)).toBeGreaterThan(minRectX(left) + 1);
  });

  it('clips cell text blocked by an occupied neighbour, overflows into empty ones', () => {
    // Column A is 5 chars wide. Row 1: A1 long, B1 occupied → A1 clips to ~5
    // chars. Row 2: A2 long, B2 empty → A2 overflows (full text kept). Repeated
    // narrow chars (no kerning/ligatures) render each run as a single Tj.
    const blocked = 'l'.repeat(18);
    const overflow = 'i'.repeat(18);
    const xlsx = buildXlsx({
      rows: [[blocked, 'X'], [overflow]],
      columns: [{ min: 1, max: 1, widthChars: 5 }],
    });
    const text = asLatin1(convertXlsxToPdfSync(xlsx, { fonts: FONTS }));
    // A2 overflows into the empty B2 → full text kept.
    expect(text).toContain(`<${hexOf(overflow)}> Tj`);
    // A1 is clipped — its full text is gone, only the fitting prefix (~5) remains.
    expect(text).not.toContain(`<${hexOf(blocked)}> Tj`);
    expect(text).toContain(`<${hexOf('l'.repeat(5))}> Tj`);
  });
});
