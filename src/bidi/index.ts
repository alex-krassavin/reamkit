// Public BiDi API (Unicode UAX #9). The renderer uses three things:
//   - analyzeString: per-character embedding levels + paragraph base level
//   - reorderVisual: L2 reversal of a level array into visual order
//   - reverseGraphemes: code-point reversal used to lay RTL words out visually

import type { Direction } from '@/bidi/algorithm';
import { computeBidi, reorderVisual } from '@/bidi/algorithm';

export type { Direction, BidiResult } from '@/bidi/algorithm';
export type { BidiClass } from '@/bidi/char-types';
export { computeBidi, reorderVisual } from '@/bidi/algorithm';
export { bidiClass } from '@/bidi/char-types';

export interface StringBidi {
  // Embedding level per *code point* (not per UTF-16 unit). Index aligns with
  // [...string] iteration order.
  readonly levels: ReadonlyArray<number>;
  readonly codePoints: ReadonlyArray<number>;
  readonly paragraphLevel: number;
}

export function analyzeString(text: string, dir: Direction = 'auto'): StringBidi {
  const codePoints: Array<number> = [];
  for (const ch of text) codePoints.push(ch.codePointAt(0)!);
  const { levels, paragraphLevel } = computeBidi(codePoints, dir);
  return { levels, codePoints, paragraphLevel };
}

// Whether a string contains any character that could trigger RTL reordering
// (R or AL strong types, or Arabic/Hebrew presentation forms). Used as a fast
// gate so pure-LTR paragraphs skip the BiDi machinery entirely.
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

// Reverse a string by code point (so surrogate pairs stay intact). Used to
// emit an RTL run's glyphs in visual (right-to-left) order, since our glyph
// placement advances left-to-right.
export function reverseByCodePoint(text: string): string {
  return [...text].reverse().join('');
}

// Given the embedding levels for a contiguous array of *tokens* (each token
// carrying a single resolved level), return the visual order as a permutation
// of token indices.
export function reorderTokens(tokenLevels: ReadonlyArray<number>): Array<number> {
  return reorderVisual(tokenLevels);
}
