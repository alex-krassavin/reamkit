// The regions a flowing reading takes one after the other: stretches read
// straight down, and bands where blocks stand side by side and each reads down
// on its own (an invoice's stack of labels beside the address it bills).

import { describe, expect, it } from 'vitest';

import type { TextRun } from '@/pdf-reader/content';
import { regionsOf } from '@/pdf-reader/regions';

/** A run of `text` at (x, y), each character half the size wide. */
function run(text: string, x: number, y: number, size = 9): TextRun {
  return {
    text,
    x,
    y,
    endX: x + text.length * size * 0.5,
    endY: y,
    fontSizePt: size,
    colorHex: '000000',
  } as TextRun;
}

/** A line of prose that runs from the left margin past both stacks. */
const ACROSS =
  'A line of text that runs all the way across the page, from one margin to the other, as prose does';

describe('the regions of a column', () => {
  it('reads two stacks with leading of their own as blocks side by side', () => {
    // bigboundingbox.pdf: "Invoice Date / Jun 3, 2013 / Invoice Number" beside
    // "Orange Demo Inc. / 23 Main Street / Central City", a point or three off
    // each other's lines. Read across, each line caught one of the other's.
    const runs = [
      run('Invoice Date', 360, 669.5),
      run('Jun 3, 2013', 360, 658.8),
      run('Invoice Number', 360, 640.9),
      run('INV-0046', 360, 630.1),
      run('Orange Demo Inc.', 472, 668.5),
      run('23 Main Street', 472, 655.5),
      run('Central City', 472, 642.5),
      run('MARINEVILLE', 472, 629.5),
    ];
    const regions = regionsOf(runs);
    expect(regions).toHaveLength(1);
    const band = regions[0];
    if (band?.kind !== 'side') throw new Error('a band of blocks');
    expect(band.cells.map((c) => c.runs.map((r) => r.text))).toEqual([
      ['Invoice Date', 'Jun 3, 2013', 'Invoice Number', 'INV-0046'],
      ['Orange Demo Inc.', '23 Main Street', 'Central City', 'MARINEVILLE'],
    ]);
  });

  it('leaves blocks that share their baselines one text, read across', () => {
    // An address beside the one it bills, line for line: a row-by-row setting.
    const runs = [
      run('548 Market Street', 30, 700),
      run('PMB 90375', 30, 686.5),
      run('Organization', 250, 700),
      run('111500', 250, 686.5),
    ];
    expect(regionsOf(runs).map((r) => r.kind)).toEqual(['flow']);
  });

  it('ends the band where a line crosses its gutter, and reads on from there', () => {
    const runs = [
      run('Invoice Date', 360, 669.5),
      run('Orange Demo Inc.', 472, 667.5),
      run('Jun 3, 2013', 360, 658.8),
      run('23 Main Street', 472, 655.5),
      run(ACROSS, 30, 640),
    ];
    const regions = regionsOf(runs);
    expect(regions.map((r) => r.kind)).toEqual(['side', 'flow']);
    expect(regions[1]?.kind === 'flow' && regions[1].runs.map((r) => r.text)).toEqual([ACROSS]);
  });
});
