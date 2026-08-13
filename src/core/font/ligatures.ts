// Characters that stand for letters — and the letters they stand for.
//
// `ﬀ` is not a letter of any alphabet: it is one glyph for two f's, a decision
// of the FACE, which most faces make through their own `liga` table and never
// give a code point at all. `𝑝` is not a letter either: it is a lower-case p
// set in italic, given its own code point so that mathematics can say "the
// variable p" in plain text. No ordinary face carries either one.
//
// Two places need to know what they stand for:
//
//   - the page, when nothing on hand can draw the character — better the
//     letters than a box or a silence (bug1873345.pdf read "different" and we
//     set "di erent"; bug1529502.pdf reads "p ← trim(p)" and we set "□ ← □");
//   - `/ToUnicode`, for the ligatures only — a search for "flavor" does not
//     find "ﬂavor", while a `𝑝` a face DID draw is the character it says it is.

/**
 * The letters a character stands for, for a character nothing can draw.
 *
 * @param cp The code point.
 * @returns The letters, or `undefined` when the character is a letter itself.
 */
export function substituteLetters(cp: number): string | undefined {
  const ligature = LIGATURE_LETTERS.get(cp);
  if (ligature !== undefined) return ligature;
  // Mathematical Alphanumeric Symbols: thirteen styles of the Latin alphabet,
  // the Greek one and the digits, each member a compatibility form of the plain
  // character. The style is lost with the code point — the letter is not.
  if (cp >= 0x1d400 && cp <= 0x1d7ff) return decomposed(cp);
  // Letterlike Symbols, where a mapping to ONE letter is the same story (`ℎ` is
  // an italic h, `ℂ` a double-struck C) and a longer one is an abbreviation:
  // `℅` stands for "c/o" and `№` for "No", which are not shapes of a letter.
  if (cp >= 0x2100 && cp <= 0x214f) {
    const letters = decomposed(cp);
    return letters?.length === 1 ? letters : undefined;
  }
  return undefined;
}

/** Whether any code point of `text` is one {@link substituteLetters} answers for. */
export function hasSubstitutable(text: string): boolean {
  for (const ch of text) if (substituteLetters(ch.codePointAt(0)!) !== undefined) return true;
  return false;
}

/**
 * The letters a LIGATURE stands for, for the one place that asks only about
 * those: the `/ToUnicode` a drawn ligature glyph is stated by.
 *
 * @param cp The code point.
 * @returns The letters, or `undefined` when it is not one of these ligatures.
 */
export function lettersForLigature(cp: number): string | undefined {
  return LIGATURE_LETTERS.get(cp);
}

/** U+FB00..FB06 — `ﬅ` opens with a long s, which is an s. */
const LIGATURE_LETTERS: ReadonlyMap<number, string> = new Map([
  [0xfb00, 'ff'],
  [0xfb01, 'fi'],
  [0xfb02, 'fl'],
  [0xfb03, 'ffi'],
  [0xfb04, 'ffl'],
  [0xfb05, 'st'],
  [0xfb06, 'st'],
]);

/** What Unicode says the character is a form of, or `undefined` if nothing. */
function decomposed(cp: number): string | undefined {
  const ch = String.fromCodePoint(cp);
  const base = ch.normalize('NFKD');
  // The unassigned holes of the mathematical block decompose to themselves.
  return base === ch ? undefined : base;
}
