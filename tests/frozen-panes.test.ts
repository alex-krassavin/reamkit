// E-SHEET SE2/SE3 — frozen panes. SE2 parses <sheetView><pane state="frozen">
// into the SheetDoc IR (a VIEW setting — no print/PDF effect, carried for the
// round-trip and HTML sticky panes). SE3's HTML tests live alongside.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const paneOf = (xlsx: Uint8Array) => readXlsxToSheetDoc(xlsx).sheets[0]!.grid.pane;

const htmlOf = async (xlsx: Uint8Array): Promise<string> =>
  new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));

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

describe('frozen panes — HTML sticky (E-SHEET SE3)', () => {
  it('pins a frozen top row and first column with position:sticky', async () => {
    const xlsx = buildXlsx({
      rows: [
        ['H1', 'H2'],
        ['a', 'b'],
      ],
      freeze: { rows: 1, cols: 1 },
    });
    const html = await htmlOf(xlsx);
    expect(html).toContain('position:sticky');
    expect(html).toContain('top:0pt'); // frozen top row pins to the top
    expect(html).toContain('left:0pt'); // frozen first column pins to the left
    expect(html).toContain('z-index:3'); // the corner cell sits above both
  });

  it('does not emit sticky styling for a sheet with no freeze', async () => {
    const html = await htmlOf(
      buildXlsx({
        rows: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      }),
    );
    expect(html).not.toContain('position:sticky');
  });
});
