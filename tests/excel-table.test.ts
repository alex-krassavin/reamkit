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
  it('shades the header row and every second data row', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: fourByTwo,
        tables: [{ ref: 'A1:B4', styleName: 'TableStyleMedium2', showRowStripes: true }],
      }),
    ).flow;
    const grid = shadingGrid(flow);
    const header = grid[0]![0];
    const band = grid[2]![0];
    expect(header).toBeDefined(); // header row
    expect(grid[0]![1]).toBe(header); // whole header row, same colour
    expect(grid[1]![0]).toBeUndefined(); // 1st data row (band1) unfilled
    expect(band).toBeDefined(); // 2nd data row (band2) filled
    expect(band).not.toBe(header);
    expect(grid[3]![0]).toBeUndefined(); // 3rd data row (band1) unfilled
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
