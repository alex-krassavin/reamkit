// E-SHEET W5 — the extended conditional-format families the flat projection now
// resolves: top10 (top/bottom N or N%), aboveAverage (mean ± N·σ),
// duplicate/uniqueValues (value frequency, numbers and text) and the text tests
// (containsText / notContainsText / beginsWith / endsWith). Each reuses the SC1
// dxf → per-cell fill machinery; only the matching predicate is new. The
// non-deterministic families (expression needs a formula engine, timePeriod the
// wall clock) stay unparsed, so they are absent here by design.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { ConditionalFormat } from '@/core/spreadsheet-model';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

// dxf 0 → solid red fill, dxf 1 → solid green fill. The new rules all reference a
// dxf, so a match shows up as the cell's resolved shading colour.
const STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="2">
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>
  </dxfs>`;

// Build a single-column sheet with one conditional-format block and return the
// resolved fill colour of every cell in column A (undefined ⇒ no highlight).
function columnShadings(
  rows: ReadonlyArray<ReadonlyArray<number | string>>,
  cfXml: string,
): Array<string | undefined> {
  const flow = Ream.parse(
    buildXlsx({
      rows: rows.map((r) => [...r]),
      stylesXml: STYLES,
      conditionalFormattingXml: cfXml,
    }),
  ).flow;
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a grid table');
  return table.table.rows.map((row) => row.cells[0]?.properties.shading?.colorHex);
}

const RED = 'FF0000';
const GREEN = '00FF00';

describe('conditional formatting — top10 (E-SHEET W5)', () => {
  it('highlights the top N values (ties at the cutoff included)', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5"><cfRule type="top10" dxfId="0" priority="1" rank="2"/></conditionalFormatting>';
    // [3,1,5,2,4] → top 2 are 5 and 4 → cutoff 4, value ≥ 4 matches.
    expect(columnShadings([[3], [1], [5], [2], [4]], cf)).toEqual([
      undefined,
      undefined,
      RED,
      undefined,
      RED,
    ]);
  });

  it('flips to the bottom N with bottom="1"', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5"><cfRule type="top10" dxfId="0" priority="1" rank="2" bottom="1"/></conditionalFormatting>';
    // bottom 2 of [3,1,5,2,4] are 1 and 2.
    expect(columnShadings([[3], [1], [5], [2], [4]], cf)).toEqual([
      undefined,
      RED,
      undefined,
      RED,
      undefined,
    ]);
  });

  it('reads rank as a percentage with percent="1"', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5"><cfRule type="top10" dxfId="0" priority="1" rank="40" percent="1"/></conditionalFormatting>';
    // 40% of 5 cells = 2 → top 2 of [10,20,30,40,50] are 40 and 50.
    expect(columnShadings([[10], [20], [30], [40], [50]], cf)).toEqual([
      undefined,
      undefined,
      undefined,
      RED,
      RED,
    ]);
  });
});

describe('conditional formatting — aboveAverage (E-SHEET W5)', () => {
  it('highlights values strictly above the range mean', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A4"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting>';
    // mean of [2,4,6,8] = 5 → 6 and 8 are above.
    expect(columnShadings([[2], [4], [6], [8]], cf)).toEqual([undefined, undefined, RED, RED]);
  });

  it('flips to below-average with aboveAverage="0"', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A4"><cfRule type="aboveAverage" dxfId="0" priority="1" aboveAverage="0"/></conditionalFormatting>';
    expect(columnShadings([[2], [4], [6], [8]], cf)).toEqual([RED, RED, undefined, undefined]);
  });

  it('includes the mean itself with equalAverage="1"', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3"><cfRule type="aboveAverage" dxfId="0" priority="1" equalAverage="1"/></conditionalFormatting>';
    // mean of [1,2,3] = 2 → value ≥ 2 matches.
    expect(columnShadings([[1], [2], [3]], cf)).toEqual([undefined, RED, RED]);
  });

  it('shifts the threshold by N standard deviations with stdDev', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A5"><cfRule type="aboveAverage" dxfId="0" priority="1" stdDev="1"/></conditionalFormatting>';
    // [1..5]: mean 3, population σ ≈ 1.414 → threshold ≈ 4.414 → only 5 exceeds it.
    expect(columnShadings([[1], [2], [3], [4], [5]], cf)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      RED,
    ]);
  });
});

describe('conditional formatting — duplicate / unique (E-SHEET W5)', () => {
  it('highlights repeated numbers', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A6"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>';
    // 2 (×2) and 3 (×3) repeat; 1 is unique.
    expect(columnShadings([[1], [2], [2], [3], [3], [3]], cf)).toEqual([
      undefined,
      RED,
      RED,
      RED,
      RED,
      RED,
    ]);
  });

  it('highlights one-off values for uniqueValues', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A4"><cfRule type="uniqueValues" dxfId="0" priority="1"/></conditionalFormatting>';
    // 1 and 3 occur once (unique); the repeated 2 does not.
    expect(columnShadings([[1], [2], [2], [3]], cf)).toEqual([RED, undefined, undefined, RED]);
  });

  it('compares text case-insensitively', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>';
    // "Apple" and "apple" collapse to one value → both duplicates; "Banana" unique.
    expect(columnShadings([['Apple'], ['apple'], ['Banana']], cf)).toEqual([RED, RED, undefined]);
  });

  it('keeps a numeric 5 and the string "5" distinct', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>';
    // number 5 vs text "5" are different values in Excel → neither is a duplicate.
    expect(columnShadings([[5], ['5']], cf)).toEqual([undefined, undefined]);
  });
});

describe('conditional formatting — text tests (E-SHEET W5)', () => {
  it('containsText matches a substring, case-insensitively', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="err"><formula>NOT(ISERROR(SEARCH("err",A1)))</formula></cfRule></conditionalFormatting>';
    expect(columnShadings([['error'], ['ok'], ['ERR!']], cf)).toEqual([RED, undefined, RED]);
  });

  it('notContainsText inverts the test', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2"><cfRule type="notContainsText" dxfId="0" priority="1" operator="notContains" text="x"><formula>ISERROR(SEARCH("x",A1))</formula></cfRule></conditionalFormatting>';
    expect(columnShadings([['abc'], ['xyz']], cf)).toEqual([RED, undefined]);
  });

  it('beginsWith anchors to the start', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A3"><cfRule type="beginsWith" dxfId="0" priority="1" operator="beginsWith" text="Total"><formula>LEFT(A1,5)="Total"</formula></cfRule></conditionalFormatting>';
    expect(columnShadings([['Total sales'], ['Subtotal'], ['total cost']], cf)).toEqual([
      RED,
      undefined,
      RED,
    ]);
  });

  it('endsWith anchors to the end', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2"><cfRule type="endsWith" dxfId="0" priority="1" operator="endsWith" text=".txt"><formula>RIGHT(A1,4)=".txt"</formula></cfRule></conditionalFormatting>';
    expect(columnShadings([['file.txt'], ['file.csv']], cf)).toEqual([RED, undefined]);
  });
});

describe('conditional formatting — priority across families (E-SHEET W5)', () => {
  it('a lower-priority cellIs claims the fill over a higher-numbered top10', () => {
    const cf =
      '<conditionalFormatting sqref="A1">' +
      '<cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan"><formula>5</formula></cfRule>' +
      '<cfRule type="top10" dxfId="0" priority="2" rank="1"/>' +
      '</conditionalFormatting>';
    // A1=10 matches both; priority 1 (green cellIs) wins over priority 2 (red top10).
    expect(columnShadings([[10]], cf)).toEqual([GREEN]);
  });
});

describe('extended conditional formats — round-trip + render (E-SHEET W5)', () => {
  const ALL_RULES =
    '<conditionalFormatting sqref="A1:A4">' +
    '<cfRule type="top10" dxfId="0" priority="1" rank="2"/>' +
    '<cfRule type="aboveAverage" dxfId="0" priority="2" aboveAverage="0" equalAverage="1"/>' +
    '<cfRule type="duplicateValues" dxfId="0" priority="3"/>' +
    '<cfRule type="uniqueValues" dxfId="1" priority="4"/>' +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="B1:B2">' +
    '<cfRule type="containsText" dxfId="0" priority="5" operator="containsText" text="a"><formula>NOT(ISERROR(SEARCH("a",B1)))</formula></cfRule>' +
    '</conditionalFormatting>';

  const xlsx = buildXlsx({
    rows: [
      [1, 'cat'],
      [2, 'dog'],
      [2, 'x'],
      [9, 'y'],
    ],
    stylesXml: STYLES,
    conditionalFormattingXml: ALL_RULES,
  });

  it('preserves every new rule type across a read→write→read loop', () => {
    const s1 = readXlsxToSheetDoc(xlsx);
    const b1 = writeXlsx(s1).bytes;
    const s2 = readXlsxToSheetDoc(b1);
    const b2 = writeXlsx(s2).bytes;
    // Idempotent: the writer is a fixpoint on its own output.
    expect(b2).toEqual(b1);
    const cf1 = s1.sheets[0]!.grid.conditionalFormats as ReadonlyArray<ConditionalFormat>;
    const cf2 = s2.sheets[0]!.grid.conditionalFormats as ReadonlyArray<ConditionalFormat>;
    expect(cf2).toEqual(cf1);
    // Sanity: the parsed types survived the round-trip.
    const types = cf1.flatMap((cf) => cf.rules.map((r) => r.type)).sort();
    expect(types).toEqual([
      'aboveAverage',
      'containsText',
      'duplicateValues',
      'top10',
      'uniqueValues',
    ]);
  });

  it('renders a sheet carrying the new rules to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(xlsx, {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
