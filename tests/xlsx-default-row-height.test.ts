// ECMA-376 §18.3.1.81 — the height of a row the sheet declares none for. It is
// the line height of the workbook's Normal font, and a workbook written before
// Calibri names Arial 10 and gets 12.75pt, not the 15pt Calibri 11 gets.
// tdf100709.xlsx is such a file: twenty rows fit its page and only fourteen fit
// ours, the rest spilling onto a second.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { defaultRowHeightPtFor } from '@/excel/print-model';
import { Ream } from '@/core/converter/ream';

function stylesWithSize(sizePt: number): string {
  return `
    <fonts count="1"><font><sz val="${String(sizePt)}"/><name val="Arial"/></font></fonts>
    <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border/></borders>
    <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>`;
}

/** The height the first row lays out at, in points. */
function firstRowPt(sizePt: number): number | undefined {
  const { flow } = Ream.parse(
    buildXlsx({ rows: [['a'], ['b']], stylesXml: stylesWithSize(sizePt) }),
  );
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') return undefined;
  return table.table.rows[0]?.properties.height;
}

describe('default row height (§18.3.1.81)', () => {
  it('gives Excel’s own table for the sizes it names', () => {
    expect(defaultRowHeightPtFor(11)).toBe(15);
    expect(defaultRowHeightPtFor(10)).toBe(12.75);
    expect(defaultRowHeightPtFor(8)).toBe(11);
    expect(defaultRowHeightPtFor(12)).toBe(15.75);
  });

  it('scales a size outside the table onto the grid heights are stored on', () => {
    // 20pt → 25.5pt, already a multiple of 0.75.
    expect(defaultRowHeightPtFor(20)).toBe(25.5);
    // 13pt → 16.575 → up to the next 0.75.
    expect(defaultRowHeightPtFor(13)).toBeCloseTo(17.25, 5);
  });

  it('falls back to Calibri 11’s height when the size is unknown', () => {
    expect(defaultRowHeightPtFor(undefined)).toBe(15);
    expect(defaultRowHeightPtFor(0)).toBe(15);
  });

  it('lays a workbook out at its own Normal size', () => {
    expect(firstRowPt(10)).toBe(12.75);
    expect(firstRowPt(11)).toBe(15);
  });
});
