// E-SHEET SE2/SE3 — frozen panes. SE2 parses <sheetView><pane state="frozen">
// into the SheetDoc IR (a VIEW setting — no print/PDF effect, carried for the
// round-trip and HTML sticky panes). SE3's HTML tests live alongside.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const paneOf = (xlsx: Uint8Array) => readXlsxToSheetDoc(xlsx).sheets[0]!.grid.pane;

describe('frozen panes — parsing (E-SHEET SE2)', () => {
  it('reads frozen rows and columns from <pane state="frozen">', () => {
    const xlsx = buildXlsx({
      rows: [
        [1, 2],
        [3, 4],
      ],
      freeze: { rows: 1, cols: 2 },
    });
    expect(paneOf(xlsx)).toEqual({ frozenRows: 1, frozenCols: 2 });
  });

  it('reads a frozen top row only', () => {
    expect(paneOf(buildXlsx({ rows: [[1], [2]], freeze: { rows: 1 } }))).toEqual({
      frozenRows: 1,
      frozenCols: 0,
    });
  });

  it('reads frozen leading columns only', () => {
    expect(paneOf(buildXlsx({ rows: [[1, 2, 3]], freeze: { cols: 2 } }))).toEqual({
      frozenRows: 0,
      frozenCols: 2,
    });
  });

  it('leaves pane undefined when the sheet has no freeze', () => {
    expect(paneOf(buildXlsx({ rows: [[1]] }))).toBeUndefined();
  });
});
