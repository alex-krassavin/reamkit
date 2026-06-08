// Arabic cursive joining (Unicode §9.2, "Arabic Cursive Joining"). Arabic
// letters take one of four contextual shapes — isolated, initial, medial,
// final — chosen from the joining types of the surrounding letters. This module
// is pure: it classifies code points and assigns each a form; the shaper then
// swaps in the matching glyph via the font's init/medi/fina GSUB lookups.
//
// Coverage: the standard Arabic block (U+0600–U+06FF) plus the tatweel and the
// ZWJ/ZWNJ controls — enough for ordinary Arabic text. Presentation-form blocks
// and the rarer extended letters fall back to dual-joining, which is correct for
// the vast majority of letters.

export type JoiningType = 'R' | 'L' | 'D' | 'C' | 'U' | 'T';
export type ArabicForm = 'isol' | 'init' | 'medi' | 'fina';

// Right-joining letters: they connect only to the preceding letter (so they can
// only ever be isolated or final). The canonical set from ArabicShaping.txt for
// the common Arabic letters (alef family, dal, dhal, reh, zain, waw, teh marbuta
// and a few extended forms).
const RIGHT_JOINING = new Set<number>([
  0x0622, 0x0623, 0x0624, 0x0625, 0x0627, 0x0629, 0x062f, 0x0630, 0x0631, 0x0632, 0x0648, 0x0671,
  0x0672, 0x0673, 0x0675, 0x0676, 0x0677, 0x0688, 0x0689, 0x068a, 0x068b, 0x068c, 0x068d, 0x068e,
  0x068f, 0x0690, 0x0691, 0x0692, 0x0693, 0x0694, 0x0695, 0x0696, 0x0697, 0x0698, 0x0699, 0x06c0,
  0x06c3, 0x06c4, 0x06c5, 0x06c6, 0x06c7, 0x06c8, 0x06c9, 0x06ca, 0x06cb, 0x06cd, 0x06cf, 0x06d2,
  0x06d3, 0x06d5, 0x06ee, 0x06ef,
]);

// Transparent (marks): skipped when looking at neighbours. The Arabic combining
// marks / harakat ranges plus the superscript alef.
function isTransparentMark(cp: number): boolean {
  return (
    (cp >= 0x0610 && cp <= 0x061a) ||
    (cp >= 0x064b && cp <= 0x065f) ||
    cp === 0x0670 ||
    (cp >= 0x06d6 && cp <= 0x06dc) ||
    (cp >= 0x06df && cp <= 0x06e4) ||
    cp === 0x06e7 ||
    cp === 0x06e8 ||
    (cp >= 0x06ea && cp <= 0x06ed)
  );
}

export function arabicJoiningType(cp: number): JoiningType {
  if (cp === 0x0640) return 'C'; // tatweel (kashida)
  if (cp === 0x200d) return 'C'; // ZWJ
  if (cp === 0x200c) return 'U'; // ZWNJ
  if (isTransparentMark(cp)) return 'T';
  if (RIGHT_JOINING.has(cp)) return 'R';
  // Remaining Arabic letters in the main block are dual-joining.
  if ((cp >= 0x0620 && cp <= 0x064a) || (cp >= 0x066e && cp <= 0x06d3) || cp === 0x06ff) {
    return 'D';
  }
  return 'U'; // non-Arabic / non-joining
}

const canJoinRight = (t: JoiningType): boolean => t === 'R' || t === 'D' || t === 'C'; // toward prev
const canJoinLeft = (t: JoiningType): boolean => t === 'L' || t === 'D' || t === 'C'; // toward next

// Assign a cursive form to each code point. Non-joining characters get 'isol'
// (their glyph is never in the positional maps, so this is a harmless default).
export function assignArabicForms(cps: ReadonlyArray<number>): Array<ArabicForm> {
  const types = cps.map(arabicJoiningType);
  const forms: Array<ArabicForm> = new Array(cps.length).fill('isol');

  // Nearest non-transparent neighbour on each side.
  const prevNonT: Array<number> = new Array(cps.length).fill(-1);
  const nextNonT: Array<number> = new Array(cps.length).fill(-1);
  let last = -1;
  for (let i = 0; i < cps.length; i++) {
    prevNonT[i] = last;
    if (types[i] !== 'T') last = i;
  }
  let nxt = -1;
  for (let i = cps.length - 1; i >= 0; i--) {
    nextNonT[i] = nxt;
    if (types[i] !== 'T') nxt = i;
  }

  for (let i = 0; i < cps.length; i++) {
    const t = types[i]!;
    if (t === 'U' || t === 'T') continue; // not shaped
    const p = prevNonT[i]!;
    const n = nextNonT[i]!;
    const joinsPrev = canJoinRight(t) && p >= 0 && canJoinLeft(types[p]!);
    const joinsNext = canJoinLeft(t) && n >= 0 && canJoinRight(types[n]!);
    forms[i] = joinsPrev && joinsNext ? 'medi' : joinsPrev ? 'fina' : joinsNext ? 'init' : 'isol';
  }
  return forms;
}
