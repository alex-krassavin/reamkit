// CJK line breaking (Unicode UAX #14, common subset).
//
// CJK scripts run without spaces, so a whitespace-only tokenizer treats a whole
// ideographic run as one unbreakable box that overflows its container instead of
// wrapping. We give each "wide" character its own box and allow a line break
// between adjacent ones — except before a non-starter (closing punctuation, small
// kana) or after a non-ender (opening punctuation). This covers Han, kana and
// fullwidth forms; scripts that already use spaces (e.g. Hangul, Latin) wrap on
// their own and are deliberately left untouched.

/** True for a "wide" CJK code point that carries inter-character break opportunities. */
export function isCjkWide(cp: number): boolean {
  return (
    (cp >= 0x2e80 && cp <= 0x2eff) || // CJK Radicals Supplement
    (cp >= 0x3000 && cp <= 0x303f) || // CJK Symbols and Punctuation
    (cp >= 0x3040 && cp <= 0x30ff) || // Hiragana + Katakana
    (cp >= 0x3400 && cp <= 0x4dbf) || // CJK Unified Ideographs Extension A
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified Ideographs
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK Compatibility Ideographs
    (cp >= 0xff00 && cp <= 0xffef) || // Halfwidth and Fullwidth Forms
    (cp >= 0x20000 && cp <= 0x3ffff) // CJK Unified Ideographs Extension B and beyond
  );
}

// Non-starters — must not begin a line (UAX #14 CL/CP/EX/NS/IS): closing brackets,
// sentence/comma punctuation, small kana, iteration and prolonged-sound marks.
const NO_BREAK_BEFORE = new Set<number>(
  [
    ...'、。，．！？：；」』】）〕〉》｝〞〟｣‐・…‥ー〜々ゝゞ',
    ...'ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮ',
  ].map((c) => c.codePointAt(0)!),
);

// Non-enders — must not end a line before the break (UAX #14 OP): opening
// brackets / quotes.
const NO_BREAK_AFTER = new Set<number>([...'（｛「『【〔〈《〝｟'].map((c) => c.codePointAt(0)!));

/** True for a wide closing punctuation / non-starter that must not start a line. */
export function isCjkNoBreakBefore(cp: number): boolean {
  return NO_BREAK_BEFORE.has(cp);
}

/** True for a wide opening punctuation that must not end a line (break after it). */
export function isCjkNoBreakAfter(cp: number): boolean {
  return NO_BREAK_AFTER.has(cp);
}

/**
 * Split a whitespace-free segment so each wide CJK character becomes its own
 * piece, while non-CJK runs stay grouped. Returns the input as a single-element
 * array when it holds no CJK — so Latin / non-CJK tokenization stays byte-identical.
 *
 * @param seg A run of non-whitespace text.
 * @returns The pieces, in order.
 */
export function splitCjkSegment(seg: string): Array<string> {
  let hasWide = false;
  for (const ch of seg) {
    if (isCjkWide(ch.codePointAt(0)!)) {
      hasWide = true;
      break;
    }
  }
  if (!hasWide) return [seg];

  const out: Array<string> = [];
  let buf = '';
  for (const ch of seg) {
    if (isCjkWide(ch.codePointAt(0)!)) {
      if (buf) {
        out.push(buf);
        buf = '';
      }
      out.push(ch);
    } else {
      buf += ch;
    }
  }
  if (buf) out.push(buf);
  return out;
}

/**
 * Whether a line may break between code point `left` (end of the piece before the
 * boundary) and `right` (start of the piece after it). A break needs at least one
 * wide character on the boundary, and is suppressed before a non-starter or after
 * a non-ender.
 *
 * @param left  The last code point of the left piece.
 * @param right The first code point of the right piece.
 * @returns `true` when a break is allowed here.
 */
export function cjkBreakBetween(left: number, right: number): boolean {
  if (!isCjkWide(left) && !isCjkWide(right)) return false;
  if (isCjkNoBreakAfter(left)) return false;
  if (isCjkNoBreakBefore(right)) return false;
  return true;
}
