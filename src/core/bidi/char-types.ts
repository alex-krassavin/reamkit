// Unicode UAX #9 — Bidirectional character types.
//
// We classify only the ranges that matter for the scripts we target (Latin,
// Hebrew, Arabic) plus the shared punctuation/format characters. Code points
// outside the listed ranges default to L (left-to-right), which is the
// correct default for the CJK / symbol / private-use blocks we do not yet
// shape specially.

export type BidiClass =
  // Strong
  | 'L'
  | 'R'
  | 'AL'
  // Weak
  | 'EN'
  | 'ES'
  | 'ET'
  | 'AN'
  | 'CS'
  | 'NSM'
  | 'BN'
  // Neutral
  | 'B'
  | 'S'
  | 'WS'
  | 'ON'
  // Explicit formatting
  | 'LRE'
  | 'LRO'
  | 'RLE'
  | 'RLO'
  | 'PDF'
  | 'LRI'
  | 'RLI'
  | 'FSI'
  | 'PDI';

interface Range {
  readonly lo: number;
  readonly hi: number;
  readonly cls: BidiClass;
}

// Explicit single code points (checked first — highest precedence).
const SINGLE: ReadonlyMap<number, BidiClass> = new Map<number, BidiClass>([
  [0x0009, 'S'], // TAB
  [0x000a, 'B'], // LF
  [0x000b, 'S'], // VT
  [0x000c, 'WS'], // FF
  [0x000d, 'B'], // CR
  [0x001c, 'B'],
  [0x001d, 'B'],
  [0x001e, 'B'],
  [0x001f, 'S'],
  [0x0020, 'WS'], // SPACE
  [0x0085, 'B'], // NEL
  [0x00a0, 'CS'], // NBSP
  [0x00ab, 'ON'],
  [0x00bb, 'ON'],
  [0x202a, 'LRE'],
  [0x202b, 'RLE'],
  [0x202c, 'PDF'],
  [0x202d, 'LRO'],
  [0x202e, 'RLO'],
  [0x2066, 'LRI'],
  [0x2067, 'RLI'],
  [0x2068, 'FSI'],
  [0x2069, 'PDI'],
  [0x061c, 'AL'], // ALM behaves as AL for our purposes
  [0x2028, 'WS'], // LINE SEPARATOR
  [0x2029, 'B'], // PARAGRAPH SEPARATOR
]);

// Ordered, non-overlapping ranges. First match wins.
const RANGES: ReadonlyArray<Range> = [
  // C0 controls not covered above → BN
  { lo: 0x0000, hi: 0x0008, cls: 'BN' },
  { lo: 0x000e, hi: 0x001b, cls: 'BN' },
  // ASCII punctuation / digits.
  { lo: 0x0021, hi: 0x0022, cls: 'ON' }, // ! "
  { lo: 0x0023, hi: 0x0025, cls: 'ET' }, // # $ %
  { lo: 0x0026, hi: 0x002a, cls: 'ON' }, // & ' ( ) *
  { lo: 0x002b, hi: 0x002b, cls: 'ES' }, // +
  { lo: 0x002c, hi: 0x002c, cls: 'CS' }, // ,
  { lo: 0x002d, hi: 0x002d, cls: 'ES' }, // -
  { lo: 0x002e, hi: 0x002f, cls: 'CS' }, // . /
  { lo: 0x0030, hi: 0x0039, cls: 'EN' }, // 0-9
  { lo: 0x003a, hi: 0x003a, cls: 'CS' }, // :
  { lo: 0x003b, hi: 0x0040, cls: 'ON' }, // ; < = > ? @
  { lo: 0x0041, hi: 0x005a, cls: 'L' }, // A-Z
  { lo: 0x005b, hi: 0x0060, cls: 'ON' }, // [ \ ] ^ _ `
  { lo: 0x0061, hi: 0x007a, cls: 'L' }, // a-z
  { lo: 0x007b, hi: 0x007e, cls: 'ON' }, // { | } ~
  // Latin-1 supplement letters → L (with a few symbol exceptions handled in SINGLE).
  { lo: 0x00a1, hi: 0x00a9, cls: 'ON' },
  { lo: 0x00aa, hi: 0x00aa, cls: 'L' },
  { lo: 0x00ac, hi: 0x00af, cls: 'ON' },
  { lo: 0x00b0, hi: 0x00b1, cls: 'ET' },
  { lo: 0x00b2, hi: 0x00b3, cls: 'EN' },
  { lo: 0x00b4, hi: 0x00b4, cls: 'ON' },
  { lo: 0x00b5, hi: 0x00b5, cls: 'L' },
  { lo: 0x00b6, hi: 0x00b9, cls: 'ON' },
  { lo: 0x00ba, hi: 0x00ba, cls: 'L' },
  { lo: 0x00bc, hi: 0x00bf, cls: 'ON' },
  { lo: 0x00c0, hi: 0x02b8, cls: 'L' }, // Latin extended / IPA
  // Combining diacritical marks → NSM
  { lo: 0x0300, hi: 0x036f, cls: 'NSM' },
  // Greek / Cyrillic → L
  { lo: 0x0370, hi: 0x0589, cls: 'L' },
  // Hebrew block.
  { lo: 0x0591, hi: 0x05bd, cls: 'NSM' }, // Hebrew points
  { lo: 0x05be, hi: 0x05be, cls: 'R' },
  { lo: 0x05bf, hi: 0x05bf, cls: 'NSM' },
  { lo: 0x05c0, hi: 0x05c0, cls: 'R' },
  { lo: 0x05c1, hi: 0x05c2, cls: 'NSM' },
  { lo: 0x05c3, hi: 0x05c3, cls: 'R' },
  { lo: 0x05c4, hi: 0x05c5, cls: 'NSM' },
  { lo: 0x05c6, hi: 0x05c6, cls: 'R' },
  { lo: 0x05c7, hi: 0x05c7, cls: 'NSM' },
  { lo: 0x05d0, hi: 0x05ea, cls: 'R' }, // Hebrew letters
  { lo: 0x05ef, hi: 0x05f4, cls: 'R' },
  // Arabic block.
  { lo: 0x0600, hi: 0x0605, cls: 'AN' }, // Arabic number signs
  { lo: 0x0606, hi: 0x0608, cls: 'AL' },
  { lo: 0x0609, hi: 0x060a, cls: 'ET' },
  { lo: 0x060b, hi: 0x060b, cls: 'AL' },
  { lo: 0x060c, hi: 0x060c, cls: 'CS' }, // Arabic comma
  { lo: 0x060d, hi: 0x060d, cls: 'AL' },
  { lo: 0x060e, hi: 0x060f, cls: 'ON' },
  { lo: 0x0610, hi: 0x061a, cls: 'NSM' },
  { lo: 0x061b, hi: 0x061b, cls: 'AL' },
  { lo: 0x061d, hi: 0x064a, cls: 'AL' }, // Arabic letters
  { lo: 0x064b, hi: 0x065f, cls: 'NSM' }, // Arabic marks
  { lo: 0x0660, hi: 0x0669, cls: 'AN' }, // Arabic-Indic digits
  { lo: 0x066a, hi: 0x066a, cls: 'ET' }, // %
  { lo: 0x066b, hi: 0x066c, cls: 'AN' }, // decimal/thousands sep
  { lo: 0x066d, hi: 0x066d, cls: 'AL' },
  { lo: 0x066e, hi: 0x066f, cls: 'AL' },
  { lo: 0x0670, hi: 0x0670, cls: 'NSM' },
  { lo: 0x0671, hi: 0x06d5, cls: 'AL' },
  { lo: 0x06d6, hi: 0x06dc, cls: 'NSM' },
  { lo: 0x06dd, hi: 0x06dd, cls: 'AN' },
  { lo: 0x06de, hi: 0x06e4, cls: 'NSM' },
  { lo: 0x06e5, hi: 0x06e6, cls: 'AL' },
  { lo: 0x06e7, hi: 0x06e8, cls: 'NSM' },
  { lo: 0x06e9, hi: 0x06e9, cls: 'ON' },
  { lo: 0x06ea, hi: 0x06ed, cls: 'NSM' },
  { lo: 0x06ee, hi: 0x06ef, cls: 'AL' },
  { lo: 0x06f0, hi: 0x06f9, cls: 'EN' }, // extended Arabic-Indic digits (EN)
  { lo: 0x06fa, hi: 0x06ff, cls: 'AL' },
  // Syriac / Arabic supplement → AL
  { lo: 0x0700, hi: 0x074f, cls: 'AL' },
  { lo: 0x0750, hi: 0x077f, cls: 'AL' }, // Arabic supplement
  // Thaana → AL-ish; treat as R-class via AL
  { lo: 0x0780, hi: 0x07bf, cls: 'AL' },
  // NKo → R
  { lo: 0x07c0, hi: 0x07ff, cls: 'R' },
  { lo: 0x0800, hi: 0x089f, cls: 'L' },
  { lo: 0x08a0, hi: 0x08ff, cls: 'AL' }, // Arabic extended-A
  // Arabic presentation forms.
  { lo: 0xfb1d, hi: 0xfb4f, cls: 'R' }, // Hebrew presentation forms
  { lo: 0xfb50, hi: 0xfdcf, cls: 'AL' }, // Arabic presentation forms-A
  { lo: 0xfdf0, hi: 0xfdff, cls: 'AL' },
  { lo: 0xfe70, hi: 0xfeff, cls: 'AL' }, // Arabic presentation forms-B
  // General punctuation neutrals.
  { lo: 0x2010, hi: 0x2027, cls: 'ON' },
  { lo: 0x2030, hi: 0x205e, cls: 'ON' },
];

export function bidiClass(cp: number): BidiClass {
  const single = SINGLE.get(cp);
  if (single !== undefined) return single;
  // Binary search would be ideal; the table is small enough for a linear scan
  // and char classification is not the hot path (line breaking dominates).
  for (const r of RANGES) {
    if (cp >= r.lo && cp <= r.hi) return r.cls;
  }
  return 'L';
}
