// Public BiDi API (Unicode UAX #9). The renderer uses three things:
//   - analyzeString: per-character embedding levels + paragraph base level
//   - reorderVisual: L2 reversal of a level array into visual order
//   - reverseGraphemes: code-point reversal used to lay RTL words out visually

import type { Direction } from '@/core/bidi/algorithm';
import { computeBidi, reorderVisual } from '@/core/bidi/algorithm';

export type { Direction, BidiResult } from '@/core/bidi/algorithm';
export type { BidiClass } from '@/core/bidi/char-types';
export { computeBidi, reorderVisual } from '@/core/bidi/algorithm';
export { bidiClass } from '@/core/bidi/char-types';

/** The result of {@link analyzeString}: per-code-point embedding levels for a string plus its paragraph base level. */
export interface StringBidi {
  /**
   * Embedding level per *code point* (not per UTF-16 unit). Index aligns with
   * `[...string]` iteration order.
   */
  readonly levels: ReadonlyArray<number>;
  /** The string's code points, in logical order (parallel to `levels`). */
  readonly codePoints: ReadonlyArray<number>;
  /** The resolved paragraph embedding level (0 = LTR, 1 = RTL). */
  readonly paragraphLevel: number;
}

/**
 * Run the UAX #9 algorithm over a string, returning per-code-point embedding
 * levels and the paragraph base level. Iterates by code point so surrogate pairs
 * stay intact and the result indices align with `[...text]`.
 *
 * @param text The paragraph text.
 * @param dir  The base direction; `'auto'` (the default) derives it from the
 *             first strong character (P2/P3).
 * @returns The {@link StringBidi} levels for `text`.
 */
export function analyzeString(text: string, dir: Direction = 'auto'): StringBidi {
  const codePoints: Array<number> = [];
  for (const ch of text) codePoints.push(ch.codePointAt(0)!);
  const { levels, paragraphLevel } = computeBidi(codePoints, dir);
  return { levels, codePoints, paragraphLevel };
}

/**
 * Whether a string contains any character that could trigger RTL reordering
 * (`R` or `AL` strong types, or Arabic/Hebrew presentation forms). Used as a fast
 * gate so pure-LTR paragraphs skip the BiDi machinery entirely.
 *
 * @param text The string to scan.
 * @returns `true` if any code point falls in the Hebrew/Arabic/Syriac/Thaana/NKo
 *          or Arabic presentation-form ranges.
 */
export function hasBidiCharacters(text: string): boolean {
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    // Hebrew, Arabic, Syriac, Thaana, NKo, Arabic presentation forms.
    if (
      (cp >= 0x0590 && cp <= 0x08ff) ||
      (cp >= 0xfb1d && cp <= 0xfdff) ||
      (cp >= 0xfe70 && cp <= 0xfeff)
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Reverse a string by code point (so surrogate pairs stay intact). Used to emit
 * an RTL run's glyphs in visual (right-to-left) order, since our glyph placement
 * advances left-to-right.
 *
 * @param text The string to reverse.
 * @returns `text` with its code points in reverse order.
 */
export function reverseByCodePoint(text: string): string {
  return [...text].reverse().join('');
}

/**
 * Given the embedding levels for a contiguous array of *tokens* (each token
 * carrying a single resolved level), return the visual order as a permutation of
 * token indices. A thin token-granularity wrapper over {@link reorderVisual} (L2).
 *
 * @param tokenLevels One resolved embedding level per token, in logical order.
 * @returns The visual order as a permutation where `result[k]` is the logical token index.
 */
export function reorderTokens(tokenLevels: ReadonlyArray<number>): Array<number> {
  return reorderVisual(tokenLevels);
}
export { segmentLevels } from '@/core/bidi/segments';
export type { BidiSegment } from '@/core/bidi/segments';
