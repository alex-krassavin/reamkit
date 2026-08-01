// ECMA-376 Part 1 §22.9.2.19 ST_Xstring — the `_xHHHH_` escape SpreadsheetML
// writes for a character XML itself cannot carry.
//
// XML 1.0 has no way to spell a control character: `&#13;` is legal but a
// carriage return inside `<t>` would be normalised away by any parser, so Excel
// writes `_x000D_` instead and every reader is expected to decode it. A cell
// holding four lines arrives as `Line 1_x000D_Line 2_x000D_…`, and drawn as it
// stands the escape is gibberish in the middle of the sentence.
//
// `_` itself is escaped as `_x005F_` when it would otherwise start a sequence,
// so decoding is a single left-to-right pass: never re-scan what a replacement
// produced, or the literal text `_x005F_x000D_` would decode twice into a
// carriage return the document never had.

const ESCAPE = /_x([0-9A-Fa-f]{4})_/g;

/**
 * Decode a SpreadsheetML `ST_Xstring` (§22.9.2.19): every `_xHHHH_` becomes the
 * character it names. Text carrying no escape is returned unchanged, so the
 * common path costs one failed match.
 *
 * @param text The raw `<t>` content.
 * @returns The text with its escapes resolved.
 */
export function decodeXstring(text: string): string {
  if (!text.includes('_x')) return text;
  return text.replace(ESCAPE, (whole, hex: string) => {
    const code = Number.parseInt(hex, 16);
    // A lone surrogate is not a character; leaving the escape as written is
    // closer to the truth than emitting an unpaired code unit.
    if (code >= 0xd800 && code <= 0xdfff) return whole;
    return String.fromCodePoint(code);
  });
}
