// E-PIVOT PV1 — pivot tables. A pivot is referenced only by a worksheet
// relationship (no element in the sheet XML); the reader enumerates the
// pivotTable rels, parses each pivotTableN.xml, and records its location +
// named style in the sheet model. The pivot's OUTPUT cells are already cached
// in the worksheet, so they keep rendering as a normal grid — PV1 adds no
// styling (that is PV2), it only models the pivot.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

// A pivot's cached output grid (what Excel writes into the sheet cells).
const PIVOT_ROWS: ReadonlyArray<ReadonlyArray<string | number>> = [
  ['Count', 'X', 'Y'],
  ['A', 1, 2],
  ['B', 3, 4],
];

function pivotXlsx(): Uint8Array {
  return buildXlsx({
    rows: PIVOT_ROWS,
    pivotTables: [
      { ref: 'A1:C3', styleName: 'PivotStyleDark1', firstHeaderRow: 1, showRowStripes: true },
    ],
  });
}

describe('Excel pivot tables — parse + model (E-PIVOT PV1)', () => {
  it('resolves the pivot table part into the sheet model', () => {
    const sheet = readXlsxToSheetDoc(pivotXlsx());
    const pivots = sheet.sheets[0]!.grid.pivotTables;
    expect(pivots).toHaveLength(1);
    const p = pivots![0]!;
    expect(p.styleName).toBe('PivotStyleDark1');
    expect(p.firstHeaderRow).toBe(1);
    expect(p.showRowStripes).toBe(true);
    expect(p.showColStripes).toBe(false);
    // <location ref="A1:C3"> → a 3×3 region (0-indexed, inclusive).
    expect(p.ref).toEqual({ startColumn: 0, startRow: 0, endColumn: 2, endRow: 2 });
  });

  it('keeps rendering the cached pivot output cells (byte-stable, no styling yet)', async () => {
    // PV1 is model-only: the grid that Excel cached into the sheet still renders.
    const html = new TextDecoder().decode(await Ream.parse(pivotXlsx()).convert('html'));
    expect(html).toContain('Count');
    expect(html).toContain('A');
    expect(html).toContain('3'); // a data value
  });

  it('omits pivotTables when the sheet has none', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: PIVOT_ROWS }));
    expect(sheet.sheets[0]!.grid.pivotTables).toBeUndefined();
  });
});

// The projected sheet is one table; each cell carries its resolved shading.
const shadingGrid = (
  flow: ReturnType<typeof Ream.parse>['flow'],
): Array<Array<string | undefined>> => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows.map((row) => row.cells.map((c) => c.properties.shading?.colorHex));
};

// A pivot with two header rows (firstDataRow=2) and three data rows over A1:C5.
const BANDED_ROWS: ReadonlyArray<ReadonlyArray<string | number>> = [
  ['Count', 'X', 'Y'],
  ['Hdr', 'a', 'b'],
  ['A', 1, 2],
  ['B', 3, 4],
  ['C', 5, 6],
];

describe('Excel pivot tables — banding projection (E-PIVOT PV2)', () => {
  it('shades the header rows and every second data row from the pivot style', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: BANDED_ROWS,
        pivotTables: [
          { ref: 'A1:C5', styleName: 'PivotStyleDark2', firstDataRow: 2, showRowStripes: true },
        ],
      }),
    ).flow;
    const grid = shadingGrid(flow);
    expect(grid[0]![0]).toBeDefined(); // header row 0
    expect(grid[1]![0]).toBeDefined(); // header row 1 (firstDataRow=2)
    expect(grid[2]![0]).toBeUndefined(); // 1st data row (band1) unfilled
    expect(grid[3]![0]).toBeDefined(); // 2nd data row (band2) filled
    expect(grid[4]![0]).toBeUndefined(); // 3rd data row (band1) unfilled
    expect(grid[3]![0]).not.toBe(grid[0]![0]); // band colour ≠ header colour
  });

  it('shades only the header rows when showRowStripes is off', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: BANDED_ROWS,
        pivotTables: [
          { ref: 'A1:C5', styleName: 'PivotStyleDark2', firstDataRow: 2, showRowStripes: false },
        ],
      }),
    ).flow;
    const grid = shadingGrid(flow);
    expect(grid[0]![0]).toBeDefined(); // headers still shaded
    expect(grid[1]![0]).toBeDefined();
    expect(grid[2]![0]).toBeUndefined();
    expect(grid[3]![0]).toBeUndefined(); // no stripes
  });
});
