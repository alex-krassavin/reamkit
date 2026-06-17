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

  it('an unsupported construct simply does not apply (graceful loss)', () => {
    // ROW() is not in the function library → #NAME? → the rule never paints.
    const cf =
      '<conditionalFormatting sqref="A1:A3">' +
      '<cfRule type="expression" dxfId="0" priority="1"><formula>ROW()=2</formula></cfRule>' +
      '</conditionalFormatting>';
    expect(columnShadings([[1], [2], [3]], cf)).toEqual([undefined, undefined, undefined]);
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
