import { describe, expect, it } from 'vitest';

import type { Item } from '@/core/line-breaker';
import { FORCED_BREAK, breakLines, greedyBreakLines } from '@/core/line-breaker';

const box = (width: number): Item => ({ type: 'box', width });
const glue = (width: number, stretch = 0, shrink = 0): Item => ({
  type: 'glue',
  width,
  stretch,
  shrink,
});
const forced = (): Item => ({ type: 'penalty', width: 0, penalty: FORCED_BREAK, flagged: false });

describe('greedyBreakLines — first-fit (E-PARITY FP3)', () => {
  it('packs each line to the last breakpoint that fits', () => {
    // Four 4-wide boxes joined by 1-wide spaces; target 10 fits two boxes + the
    // space (9) but not three. First-fit closes line 0 after the second box.
    const items: Array<Item> = [
      box(4),
      glue(1),
      box(4),
      glue(1),
      box(4),
      glue(1),
      box(4),
      glue(0, 1e6),
      forced(),
    ];
    // breaks point at the glue after box 2 (index 3) and the final sentinel (8):
    // line 0 = [0,3) "AAAA BBBB", line 1 = [4,8) "CCCC DDDD".
    expect(greedyBreakLines(items, 10)).toEqual([3, 8]);
  });

  it('puts an over-long word on its own line (emergency break)', () => {
    // A 15-wide box cannot fit a 10-wide line and has no earlier breakpoint —
    // it gets its own (overflowing) line rather than swallowing the next word.
    const items: Array<Item> = [box(15), glue(1), box(4), glue(0, 1e6), forced()];
    expect(greedyBreakLines(items, 10)).toEqual([1, 4]);
  });

  it('breaks off the overflow before an overflowing forced break', () => {
    // No final glue: the forced break itself overflows, so the last good break
    // (after box 0) closes line 0 and the wide word lands on the final line.
    const items: Array<Item> = [box(4), glue(1), box(9), forced()];
    expect(greedyBreakLines(items, 10)).toEqual([1, 3]);
  });

  it('returns no breaks for an empty stream', () => {
    expect(greedyBreakLines([], 10)).toEqual([]);
  });

  it('diverges from Knuth-Plass total-fit on a varied paragraph', () => {
    // Many varied-width boxes with stretchy glue: greedy packs each line to the
    // brim, while total-fit sacrifices some early fullness to avoid worse later
    // lines — so over a long paragraph the two choose different break points.
    const items: Array<Item> = [];
    for (let i = 0; i < 120; i++) {
      items.push(box(4 + ((i * 37) % 13)));
      items.push(glue(3, 6, 1));
    }
    items.push(glue(0, 1e6, 0), forced());
    const target = 60;
    expect(greedyBreakLines(items, target)).not.toEqual(breakLines(items, target).breaks);
  });
});
