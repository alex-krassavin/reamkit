// E-SHEET SC3 — Excel tables: table parts resolved from the worksheet
// relationships, their named style mapped to header / band fills against the
// workbook accent, and banded shading overlaid onto the grid cells.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const fourByTwo = [
  [1, 2],
  [3, 4],
  [5, 6],
  [7, 8],
];

const shadingGrid = (
  flow: ReturnType<typeof Ream.parse>['flow'],
): Array<Array<string | undefined>> => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows.map((row) => row.cells.map((c) => c.properties.shading?.colorHex));
};

describe('Excel tables — parse + resolve (E-SHEET SC3)', () => {
  it('resolves a table part and its style colours onto the sheet', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: fourByTwo,
        tables: [
          { ref: 'A1:B4', name: 'Data', styleName: 'TableStyleMedium2', showRowStripes: true },
        ],
      }),
    );
    const tables = sheet.sheets[0]!.grid.tables;
    expect(tables).toHaveLength(1);
    const t = tables![0]!;
    expect(t.ref).toEqual({ startColumn: 0, startRow: 0, endColumn: 1, endRow: 3 });
    expect(t.styleName).toBe('TableStyleMedium2');
    expect(t.showRowStripes).toBe(true);
    expect(t.headerRowCount).toBe(1);
    // The accent resolves to a darker header and a lighter band, both defined.
    expect(t.headerHex).toMatch(/^[0-9A-F]{6}$/);
    expect(t.bandHex).toMatch(/^[0-9A-F]{6}$/);
    expect(t.headerHex).not.toBe(t.bandHex);
  });

  it('leaves a style-less (TableStyleNone) table uncoloured', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: fourByTwo, tables: [{ ref: 'A1:B4', styleName: 'TableStyleNone' }] }),
    );
    const t = sheet.sheets[0]!.grid.tables![0]!;
    expect(t.headerHex).toBeUndefined();
    expect(t.bandHex).toBeUndefined();
  });
});

describe('Excel tables — banding projection (E-SHEET SC3)', () => {
  it('shades the header row and every second data row, from the first', () => {
    // §18.8.40 — TableStyleMedium2 as the standard defines it: the header in
    // accent1, the first row stripe in accent1 lightened by 80%, the second
    // stripe nothing. The first stripe is the first data row: Excel's own PDF
    // shades it, and the row after it is white — in 156082 and C0E6F5, the
    // Office 2023 theme's, for a workbook that carries none.
    const flow = Ream.parse(
      buildXlsx({
        rows: fourByTwo,
        tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2', showRowStripes: true }],
      }),
    ).flow;
    const grid = shadingGrid(flow);
    expect(grid[0]![0]).toBe('156082'); // header row
    expect(grid[0]![1]).toBe('156082'); // the whole of it
    expect(grid[1]![0]).toBe('C0E6F5'); // 1st data row: the first stripe
    expect(grid[2]![0]).toBeUndefined(); // 2nd: the second stripe, unfilled
    expect(grid[3]![0]).toBe('C0E6F5'); // 3rd: the first stripe again
  });

  it('rules the rows apart and the totals off, as the style draws them', () => {
    // TableStyleMedium2's whole table is ruled in accent1 lightened by 40%,
    // round its edge and between its rows; its totals row is bold under a
    // double rule in accent1 itself.
    // A header, two data rows and the totals row.
    const flow = Ream.parse(
      buildXlsx({
        rows: fourByTwo,
        tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2', totalsRowCount: 1 }],
      }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const cell = (r: number, c: number) => table.table.rows[r]!.cells[c]!;
    expect(cell(1, 0).properties.borders?.bottom?.colorHex).toBe('44B3E1');
    expect(cell(2, 1).properties.borders?.right?.colorHex).toBe('44B3E1');
    // No rule between the columns: Medium2 draws none.
    expect(cell(2, 0).properties.borders?.right).toBeUndefined();
    const total = cell(3, 0);
    expect(total.properties.borders?.top).toMatchObject({ style: 'double', colorHex: '156082' });
    const run =
      total.content[0]?.kind === 'paragraph' ? total.content[0].paragraph.runs[0] : undefined;
    expect(run?.properties.bold).toBe(true);
    // The stripes run through the data rows only: the totals row, where the
    // first stripe would come round again, is not one.
    expect(shadingGrid(flow).map((row) => row[0])).toEqual([
      '156082',
      'C0E6F5',
      undefined,
      undefined,
    ]);
  });

  it('takes a style the workbook defines itself, region over region', () => {
    // §18.8.42 — a dxf gives a solid fill's colour as its background; the
    // first column (bold, its own fill) lies over the row stripe, the header
    // over both.
    const stylesXml = `
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
      <dxfs count="3">
        <dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFFE699"/></patternFill></fill></dxf>
        <dxf><fill><patternFill><bgColor rgb="FFDDEBF7"/></patternFill></fill></dxf>
        <dxf><font><color rgb="FFFFFFFF"/></font><fill><patternFill><bgColor rgb="FF203764"/></patternFill></fill></dxf>
      </dxfs>
      <tableStyles count="1"><tableStyle name="Own" pivot="0" count="3">
        <tableStyleElement type="headerRow" dxfId="2"/>
        <tableStyleElement type="firstColumn" dxfId="0"/>
        <tableStyleElement type="firstRowStripe" dxfId="1"/>
      </tableStyle></tableStyles>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: fourByTwo,
        stylesXml,
        tables: [{ ref: 'A1:B4', styleName: 'Own', showFirstColumn: true }],
      }),
    ).flow;
    expect(shadingGrid(flow)).toEqual([
      ['203764', '203764'],
      ['FFE699', 'DDEBF7'],
      ['FFE699', undefined],
      ['FFE699', 'DDEBF7'],
    ]);
  });

  it("whitens a header whose black is the Normal style's, not one the author set", () => {
    // Excel's own PDF (2026-10-01): four Medium2 tables whose header fonts
    // name their black as `theme="1"` (as Normal does), as `rgb="FF000000"`,
    // not at all, and the Normal style itself. The first and last turn white;
    // the rgb black and the automatic one stay black.
    const stylesXml = `
      <fonts count="4">
        <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
        <font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
        <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
        <font><sz val="11"/><name val="Calibri"/></font>
      </fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
        <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"/>
        <xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/>
      </cellXfs>`;
    const headerColour = (styleIndex: number): string | undefined => {
      const flow = Ream.parse(
        buildXlsx({
          rows: [[{ value: 'Head', styleIndex }], [1], [2]],
          stylesXml,
          tables: [{ ref: 'A1:A3', styleName: 'TableStyleMedium2' }],
        }),
      ).flow;
      const table = flow.body.find((el) => el.kind === 'table');
      if (table?.kind !== 'table') throw new Error('expected a table');
      const block = table.table.rows[0]!.cells[0]!.content[0];
      return block?.kind === 'paragraph' ? block.paragraph.runs[0]?.properties.colorHex : undefined;
    };
    expect(headerColour(1)).toBe('FFFFFF');
    expect(headerColour(0)).toBe('FFFFFF');
    expect(headerColour(2)).toBe('000000');
    expect(headerColour(3)).not.toBe('FFFFFF');
  });

  it('does not band when showRowStripes is off (header only)', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: fourByTwo,
        tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2', showRowStripes: false }],
      }),
    ).flow;
    const grid = shadingGrid(flow);
    expect(grid[0]![0]).toBeDefined(); // header still shaded
    expect(grid[1]![0]).toBeUndefined();
    expect(grid[2]![0]).toBeUndefined(); // no band stripes
    expect(grid[3]![0]).toBeUndefined();
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('Excel tables — per-style accents + header text (Tail TC1)', () => {
  const tableOf = (styleName: string) =>
    readXlsxToSheetDoc(buildXlsx({ rows: fourByTwo, tables: [{ ref: 'A1:B4', styleName }] }))
      .sheets[0]!.grid.tables![0]!;

  it('a Medium style gives a solid accent header with white text', () => {
    const t = tableOf('TableStyleMedium2');
    expect(t.headerTextHex).toBe('FFFFFF');
    expect(t.headerHex).toMatch(/^[0-9A-F]{6}$/);
  });

  it('a Light style keeps black header text (no override)', () => {
    const t = tableOf('TableStyleLight9');
    expect(t.headerTextHex).toBeUndefined();
    expect(t.headerHex).toMatch(/^[0-9A-F]{6}$/);
  });

  it('different style numbers resolve to different theme accents', () => {
    // Medium2 → accent1, Medium5 → accent4 ((N-1)%7).
    expect(tableOf('TableStyleMedium2').headerHex).not.toBe(tableOf('TableStyleMedium5').headerHex);
  });

  it('applies the white header font colour to the projected header cells', () => {
    const flow = Ream.parse(
      buildXlsx({ rows: fourByTwo, tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2' }] }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const headerBlock = table.table.rows[0]!.cells[0]!.content[0];
    if (headerBlock?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(headerBlock.paragraph.runs[0]?.properties.colorHex).toBe('FFFFFF');
    const dataBlock = table.table.rows[1]!.cells[0]!.content[0];
    if (dataBlock?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(dataBlock.paragraph.runs[0]?.properties.colorHex).not.toBe('FFFFFF');
  });
});

describe('Excel tables — render smoke (E-SHEET SC3)', () => {
  it('renders a banded table to a valid PDF', () => {
    const xlsx = buildXlsx({
      rows: fourByTwo,
      tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2', showRowStripes: true }],
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
