// The Latin typographic ligatures, and the letters each of them stands for.
//
// `ﬀ` is not a letter of any alphabet: it is one glyph for two f's, a decision
// of the FACE, which most faces make through their own `liga` table and never
// give a code point at all. Two places need to know that it is "ff":
//
//   - the page, when nothing on hand can draw the ligature — better the letters
//     than a character that goes out silently (bug1873345.pdf read "different"
//     and we set "di erent");
//   - `/ToUnicode`, because a search for "flavor" does not find "ﬂavor".

/** The letters a ligature stands for, or `undefined` when it is not one. */
export function lettersForLigature(cp: number): string | undefined {
  return LIGATURE_LETTERS.get(cp);
}

/** Whether any code point of `text` is one of these ligatures. */
export function hasLigature(text: string): boolean {
  for (const ch of text) if (LIGATURE_LETTERS.has(ch.codePointAt(0)!)) return true;
  return false;
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
