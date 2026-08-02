// Finding a piece of text in a PDF content stream, at the byte level.
//
// Glyphs are shown one of two ways. A run whose shaping leaves every glyph on
// its own advance is a single `<AAAABBBB> Tj`; one where pair kerning moved a
// glyph is a `[<AAAA> 41.02 <BBBB>] TJ`, the number being the nudge (ISO
// 32000-1 §9.4.3). Which of the two a given string gets depends on the font's
// kern table, so an assertion spelled `<hex> Tj` silently stops matching the
// day its string gains a kern pair. These build a pattern that accepts either.

import type { ParsedTtf } from '@/core/font/ttf-parser';

/** The 4-digit hex glyph ids for `text`, in order (Identity-H addressing). */
export function glyphHexes(parsed: ParsedTtf, text: string): Array<string> {
  return [...text].map((c) =>
    parsed.glyphForCodepoint(c.codePointAt(0)!).toString(16).padStart(4, '0').toUpperCase(),
  );
}

/**
 * A pattern matching `text` shown in a content stream, with or without the
 * kern splices shaping may have put between its glyphs.
 *
 * @param parsed The font the text is shown in.
 * @param text   The characters to look for.
 * @param flags  Extra regex flags (`'g'` to count occurrences).
 * @returns The pattern.
 */
export function showPattern(parsed: ParsedTtf, text: string, flags = ''): RegExp {
  const glyphs = glyphHexes(parsed, text);
  return new RegExp(`[<[]${glyphs.join(String.raw`(?:> -?[\d.]+ <)?`)}[>\\]]`, `u${flags}`);
}

/** How many times `text` is shown in `stream`. */
export function countShown(parsed: ParsedTtf, stream: string, text: string): number {
  return (stream.match(showPattern(parsed, text, 'g')) ?? []).length;
}
