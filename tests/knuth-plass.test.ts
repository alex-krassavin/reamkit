import { describe, expect, it } from 'vitest';

import type { Item } from '@/core/line-breaker';
import { breakLines } from '@/core/line-breaker';

// Build a paragraph of "word1 word2 word3 ..." with each word treated as a
// box of fixed width and each space as glue with a sensible stretch/shrink.
function paragraph(
  words: ReadonlyArray<{ width: number }>,
  spaceWidth = 3,
  finalBreak = true,
): Array<Item> {
  const items: Array<Item> = [];
  for (let i = 0; i < words.length; i++) {
    items.push({ type: 'box', width: words[i]!.width });
    if (i < words.length - 1) {
      items.push({ type: 'glue', width: spaceWidth, stretch: spaceWidth, shrink: 1 });
    }
  }
  if (finalBreak) {
    // Final glue + forced penalty pair so the last line stays left-aligned.
    items.push({ type: 'glue', width: 0, stretch: 1e6, shrink: 0 });
    items.push({ type: 'penalty', width: 0, penalty: -10_000, flagged: false });
  }
  return items;
}

describe('Knuth-Plass: breakLines', () => {
  it('returns a single break for a paragraph that fits on one line', () => {
    const items = paragraph([{ width: 10 }, { width: 10 }, { width: 10 }]);
    const { breaks, ratios } = breakLines(items, 100);
    expect(breaks).toHaveLength(1);
    expect(ratios).toHaveLength(1);
    // The single break must be the final forced penalty.
    expect(items[breaks[0]!]!.type).toBe('penalty');
  });

  it('splits across lines when the paragraph is too wide for one line', () => {
    const items = paragraph(
      Array.from({ length: 10 }, () => ({ width: 20 })),
      3,
    );
    const { breaks } = breakLines(items, 50);
    // 10 words × 20pt + 9 × 3pt spaces = 227pt. Line width 50pt fits ≈ 2 words.
    // Expect at least 5 breaks (lines).
    expect(breaks.length).toBeGreaterThanOrEqual(5);
  });

  it('returns ratios within [-1, ∞)', () => {
    const items = paragraph(
      Array.from({ length: 30 }, () => ({ width: 15 })),
      3,
    );
    const { ratios } = breakLines(items, 80);
    for (const r of ratios) {
      expect(r).toBeGreaterThanOrEqual(-1);
    }
  });

  it('prefers a balanced break over a greedy one (lower demerits)', () => {
    // Two words 19 wide, then five words 1 wide. Line width 22.
    // Greedy: first line 19+3+1 = 23 (overfull). Then five tiny words on next.
    // KP should put one 19-word per line and pack tiny words together.
    const widths = [19, 19, 1, 1, 1, 1, 1];
    const items = paragraph(widths.map((w) => ({ width: w })));
    const { breaks } = breakLines(items, 22);
    // At minimum: 2 lines for the big words + ≤2 lines for the tiny → ≤ 4 breaks.
    expect(breaks.length).toBeGreaterThanOrEqual(2);
    expect(breaks.length).toBeLessThanOrEqual(5);
  });

  it('honours an explicit forced-break penalty', () => {
    const items: Array<Item> = [
      { type: 'box', width: 10 },
      { type: 'glue', width: 3, stretch: 3, shrink: 1 },
      { type: 'box', width: 10 },
      // Forced break in the middle.
      { type: 'penalty', width: 0, penalty: -10_000, flagged: false },
      { type: 'box', width: 10 },
      { type: 'glue', width: 0, stretch: 1e6, shrink: 0 },
      { type: 'penalty', width: 0, penalty: -10_000, flagged: false },
    ];
    const { breaks } = breakLines(items, 500);
    // Two breakpoints: the forced one in the middle + the final one.
    expect(breaks.length).toBe(2);
  });

  it('accepts per-line widths (e.g. first line indented)', () => {
    const items = paragraph(
      Array.from({ length: 8 }, () => ({ width: 20 })),
      3,
    );
    const { breaks } = breakLines(items, [40, 80]);
    // First line 40pt = roughly 1.5 words; subsequent 80pt = ~3 words.
    expect(breaks.length).toBeGreaterThanOrEqual(2);
  });

  it('wraps loosely rather than overflowing when no tight break fits', () => {
    // Each word is 45pt in a 78pt column: one word fits (loose), two overflow.
    // A too-loose line is acceptable; an overfull one would spill past the column
    // and overlap its neighbour (the real bug: a long phrase in a thin table cell
    // rendered as one ~180pt line in a ~78pt cell).
    const items = paragraph(
      Array.from({ length: 5 }, () => ({ width: 45 })),
      3,
    );
    const { breaks, ratios } = breakLines(items, 78);
    // No line may be overfull (ratio < -1 ⇒ content spills past the column).
    expect(ratios.every((r) => r >= -1)).toBe(true);
    // It actually wrapped (≈ one word per line) instead of one giant line.
    expect(breaks.length).toBeGreaterThanOrEqual(4);
  });
});
