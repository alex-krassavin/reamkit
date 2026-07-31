// ECMA-376 §18.3.1.65 `<pageSetUpPr fitToPage>` beside §18.3.1.74 `<rowBreaks>`.
// Fit-to-page and a manual break are two answers to the same question and Excel
// takes the scaling one. Honouring both asks for the impossible — a break costs
// a page at every scale — so sheet-fit-breaks.xlsx, which says fit on one page
// and breaks after row 20, printed on two where the reference prints one.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';

function grid(rows: number): Array<Array<string>> {
  return Array.from({ length: rows }, (_, r) => [`Row ${String(r + 1)}`, String((r + 1) * 100)]);
}

/** How many pages the flow asks for: one, plus every forced break in the body. */
function forcedBreaks(bytes: Uint8Array): number {
  const { flow } = Ream.parse(bytes);
  let breaks = 0;
  for (const el of flow.body) {
    if (el.kind === 'paragraph' && el.paragraph.properties.pageBreakBefore) breaks++;
    if (el.kind !== 'table') continue;
    for (const row of el.table.rows) if (row.properties.pageBreakBefore) breaks++;
  }
  return breaks;
}

describe('fit-to-page and manual breaks (§18.3.1.65)', () => {
  it('drops a row break the sheet is being fitted past', () => {
    expect(
      forcedBreaks(
        buildXlsx({
          rows: grid(40),
          rowBreaks: [20],
          fitToPage: true,
          pageSetup: { fitToHeight: 1 },
        }),
      ),
    ).toBe(0);
  });

  it('keeps the break when nothing is being fitted', () => {
    expect(forcedBreaks(buildXlsx({ rows: grid(40), rowBreaks: [20] }))).toBe(1);
  });

  it('keeps the break when the axis is explicitly unconstrained', () => {
    // `fitToHeight="0"` is "as many pages down as it takes" (§18.3.1.63), so
    // the sheet is only being fitted ACROSS and the row break still divides it.
    expect(
      forcedBreaks(
        buildXlsx({
          rows: grid(40),
          rowBreaks: [20],
          fitToPage: true,
          pageSetup: { fitToWidth: 1, fitToHeight: 0 },
        }),
      ),
    ).toBe(1);
  });

  it('drops a column break the sheet is being fitted past', () => {
    expect(
      forcedBreaks(
        buildXlsx({
          rows: [['a', 'b', 'c', 'd']],
          columns: [{ min: 1, max: 4, widthChars: 24 }],
          colBreaks: [2],
          fitToPage: true,
          pageSetup: { fitToWidth: 1 },
        }),
      ),
    ).toBe(0);
  });
});
