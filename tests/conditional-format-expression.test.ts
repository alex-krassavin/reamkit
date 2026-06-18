// E-SHEET W9 — the last two conditional-format families, end to end:
//   • expression — an arbitrary formula evaluated per cell against the grid's
//     cached values (no recalc). Works through the plain projection (.flow), no
//     date needed unless the formula calls TODAY()/NOW().
//   • timePeriod — a clock-relative date window, resolved against an injected
//     reference date (options.now), never the wall clock. Absent ⇒ no-op.
// Both round-trip through the xlsx writer (they live in the normalised grid).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { serialFromDate } from '@/excel/formula';
import { convertXlsxToPdfSync } from '@/core/converter';

// dxf 0 → solid red, dxf 1 → solid green (mirrors the SC1/W5 fixtures).
const STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="2">
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>
  </dxfs>`;

const RED = 'FF0000';

// The resolved fill of every column-A cell, projecting with an optional
// reference date (E-SHEET W9). Uses the SheetDoc → FlowDoc projection directly so
// the injected `now` reaches the conditional-format formula engine.
function columnShadings(
  rows: ReadonlyArray<ReadonlyArray<number | string>>,
  cfXml: string,
  now?: Date,
): Array<string | undefined> {
  const sheet = readXlsxToSheetDoc(
    buildXlsx({
      rows: rows.map((r) => [...r]),
      stylesXml: STYLES,
      conditionalFormattingXml: cfXml,
    }),
  );
  const flow = projectSheetDoc(sheet, now ? { now } : {});
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a grid table');
  return table.table.rows.map((row) => row.cells[0]?.properties.shading?.colorHex);
}

describe('conditional formatting — expression (E-SHEET W9)', () => {
  it('highlights cells whose own value matches a relative formula', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;5</formula></cfRule>' +
      '</conditionalFormatting>';
    // A1>5 shifts per row → each cell tests its own value.
    expect(columnShadings([[3], [7], [5], [9], [1]], cf)).toEqual([
      undefined,
      RED,
      undefined,
      RED,
      undefined,
    ]);
  });

  it('evaluates a formula referencing another column (absolute range)', () => {
    // Highlight A when the same row's B equals "x". B is column 2.
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>B1="x"</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(
      columnShadings(
        [
          ['a', 'x'],
          ['b', 'y'],
          ['c', 'x'],
        ],
        cf,
      ),
    ).toEqual([RED, undefined, RED]);
  });

  it('uses a function (MOD) over the value', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A4">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>MOD(A1,2)=0</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(columnShadings([[1], [2], [3], [4]], cf)).toEqual([undefined, RED, undefined, RED]);
  });

  it('resolves ROW() to the current cell (row-banding idiom)', () => {
    // MOD(ROW(),2)=0 paints even rows — ROW() reports each covered cell's row.
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>MOD(ROW(),2)=0</formula></cfRule>' +
      '</conditionalFormatting>';
    const out = columnShadings([[1], [2], [3]], cf);
    expect(out[0]).toBeUndefined(); // row 1 (odd)
    expect(out[1]).toBeTruthy(); // row 2 (even) → painted
    expect(out[2]).toBeUndefined(); // row 3 (odd)
  });

  it('an unsupported construct simply does not apply (graceful loss)', () => {
    // INDIRECT is a dynamic, volatile reference — intentionally not modelled (it
    // cannot be deterministic) → #NAME? → the rule never paints.
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>INDIRECT("A2")=2</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(columnShadings([[1], [2], [3]], cf)).toEqual([undefined, undefined, undefined]);
  });

  it('evaluates one of the new functions (SUMPRODUCT) over two columns', () => {
    // Highlight A where the row's A·B + … cumulative? Keep it per-row: SUMPRODUCT of
    // the single-row pair A1:B1 = A1*B1; highlight when the product exceeds 10.
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1">' +
      '<formula>SUMPRODUCT(A1:A1,B1:B1)&gt;10</formula></cfRule></conditionalFormatting>';
    expect(
      columnShadings(
        [
          [2, 3],
          [4, 5],
          [1, 1],
        ],
        cf,
      ),
    ).toEqual([undefined, RED, undefined]); // 6, 20, 1
  });
});

// W9 tail — an `expression` rule that reaches outside its own sheet: a
// sheet-qualified reference (Sheet2!A1) or a workbook defined name. The whole
// workbook (every sheet's grid + the defined names) is wired into the formula
// engine through the projection.
describe('conditional formatting — cross-sheet refs & defined names', () => {
  // Sheet1 carries the rows + CF; Sheet2 the referenced data. Returns Sheet1's
  // column-A shadings.
  function crossShadings(
    sheet1: ReadonlyArray<ReadonlyArray<number | string>>,
    sheet2: ReadonlyArray<ReadonlyArray<number | string>>,
    cfXml: string,
    definedNames?: ReadonlyArray<{ name: string; value: string; localSheetId?: number }>,
  ): Array<string | undefined> {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [
          { name: 'Sheet1', rows: sheet1.map((r) => [...r]), conditionalFormattingXml: cfXml },
          { name: 'Sheet2', rows: sheet2.map((r) => [...r]) },
        ],
        stylesXml: STYLES,
        ...(definedNames ? { definedNames: [...definedNames] } : {}),
      }),
    );
    const flow = projectSheetDoc(sheet, {});
    const table = flow.body.find((el) => el.kind === 'table'); // Sheet1's grid is first
    if (table?.kind !== 'table') throw new Error('expected a grid table');
    return table.table.rows.map((row) => row.cells[0]?.properties.shading?.colorHex);
  }

  it('resolves a per-row sheet-qualified reference (Sheet2!A1)', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>Sheet2!A1&gt;5</formula></cfRule>' +
      '</conditionalFormatting>';
    // The unanchored Sheet2!A1 shifts per row → Sheet2 A1/A2/A3 = 10/3/7.
    expect(crossShadings([[0], [0], [0]], [[10], [3], [7]], cf)).toEqual([RED, undefined, RED]);
  });

  it('aggregates a sheet-qualified absolute range', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2">' +
      '<cfRule type="expression" dxfId="0" priority="1">' +
      '<formula>SUM(Sheet2!$A$1:$A$3)&gt;15</formula></cfRule></conditionalFormatting>';
    expect(crossShadings([[0], [0]], [[10], [3], [7]], cf)).toEqual([RED, RED]); // sum 20 > 15
    expect(crossShadings([[0], [0]], [[1], [2], [3]], cf)).toEqual([undefined, undefined]); // 6
  });

  it('resolves a defined name pointing at another sheet', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;Threshold</formula></cfRule>' +
      '</conditionalFormatting>';
    const names = [{ name: 'Threshold', value: 'Sheet2!$A$1' }]; // = 10
    expect(crossShadings([[5], [12], [20]], [[10]], cf, names)).toEqual([undefined, RED, RED]);
  });

  it('resolves a defined name bound to a literal constant', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;Limit</formula></cfRule>' +
      '</conditionalFormatting>';
    const names = [{ name: 'Limit', value: '8' }];
    expect(crossShadings([[5], [9], [20]], [[0]], cf, names)).toEqual([undefined, RED, RED]);
  });

  it('an unknown sheet qualifier is #REF! → the rule never paints', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>Nope!A1&gt;0</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(crossShadings([[0], [0]], [[10], [10]], cf)).toEqual([undefined, undefined]);
  });
});

describe('conditional formatting — timePeriod (E-SHEET W9)', () => {
  const now = new Date(Date.UTC(2026, 5, 17)); // 2026-06-17
  const serial = (y: number, m: number, d: number): number =>
    serialFromDate(new Date(Date.UTC(y, m - 1, d)), false);

  const todayCf =
    '<conditionalFormatting sqref="A1:A4">' +
    '<cfRule type="timePeriod" timePeriod="today" dxfId="0" priority="1">' +
    '<formula>FLOOR(A1,1)=TODAY()</formula></cfRule></conditionalFormatting>';

  it('highlights the cell whose date is today (relative to options.now)', () => {
    const rows = [
      [serial(2026, 6, 17)],
      [serial(2026, 6, 16)],
      [serial(2026, 6, 18)],
      [serial(2026, 5, 1)],
    ];
    expect(columnShadings(rows, todayCf, now)).toEqual([RED, undefined, undefined, undefined]);
  });

  it('is a no-op when no reference date is supplied (deterministic)', () => {
    const rows = [[serial(2026, 6, 17)], [serial(2026, 6, 16)]];
    expect(columnShadings(rows, todayCf)).toEqual([undefined, undefined]);
  });

  it('matches a whole-month window', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="timePeriod" timePeriod="thisMonth" dxfId="0" priority="1">' +
      '<formula>x</formula></cfRule></conditionalFormatting>';
    const rows = [[serial(2026, 6, 1)], [serial(2026, 6, 30)], [serial(2026, 7, 1)]];
    expect(columnShadings(rows, cf, now)).toEqual([RED, RED, undefined]);
  });
});

describe('expression + TODAY in an expression rule (E-SHEET W9)', () => {
  it('reads options.now through TODAY() inside an expression', () => {
    const now = new Date(Date.UTC(2026, 5, 17));
    const today = serialFromDate(now, false);
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1=TODAY()</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(columnShadings([[today], [today - 1], [today + 5]], cf, now)).toEqual([
      RED,
      undefined,
      undefined,
    ]);
  });
});

describe('write-back round-trip (E-SHEET W9)', () => {
  it('preserves expression and timePeriod rules through the xlsx writer', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;5</formula></cfRule>' +
      '<cfRule type="timePeriod" timePeriod="lastWeek" dxfId="1" priority="2">' +
      '<formula>helper</formula></cfRule></conditionalFormatting>';
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [[1]], stylesXml: STYLES, conditionalFormattingXml: cf }),
    );
    const reparsed = readXlsxToSheetDoc(writeXlsx(sheet).bytes);
    const rules = reparsed.sheets[0]!.grid.conditionalFormats?.[0]?.rules ?? [];
    // TS infers the type predicate from `r.type === 'expression'`, so the find
    // already narrows to the rule subtype — no assertion needed.
    const expr = rules.find((r) => r.type === 'expression');
    const period = rules.find((r) => r.type === 'timePeriod');
    expect(expr?.formula).toBe('A1>5');
    expect(period?.timePeriod).toBe('lastWeek');
    expect(period?.formula).toBe('helper');
  });
});

describe('PDF smoke (E-SHEET W9)', () => {
  it('renders a sheet with expression + timePeriod CF to a valid PDF', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;1</formula></cfRule>' +
      '<cfRule type="timePeriod" timePeriod="today" dxfId="1" priority="2">' +
      '<formula>FLOOR(A1,1)=TODAY()</formula></cfRule></conditionalFormatting>';
    const pdf = convertXlsxToPdfSync(
      buildXlsx({ rows: [[1], [2], [3]], stylesXml: STYLES, conditionalFormattingXml: cf }),
      {
        now: new Date(Date.UTC(2026, 5, 17)),
        fonts: {
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
        },
      },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
