// Generate `src/core/font/indic-table.ts`: the character properties the Indic
// shaper (`src/core/font/indic-shaper.ts`) works from.
//
// A character's part in a syllable of the scripts of India is in Unicode's
// Indic_Syllabic_Category (consonant, vowel sign, virama, nukta…) and
// Indic_Positional_Category (a vowel sign drawn left of its consonant, above
// it, below it…). The shaper uses them as HarfBuzz uses them: mapped to the
// categories its syllable grammar is written in, a Ra marked as the letter a
// reph is made of, and a vowel sign's side mapped to the place it is sorted
// into, which differs script by script. That mapping is HarfBuzz's
// gen-indic-table.py and set_indic_properties, restated here. The
// normalization step needs the canonical decompositions and combining classes
// of the same blocks.
//
// This script does not query Unicode itself: point it at the dump that
// `scripts/indic-table/icu-dump.c` writes from ICU (its header says how).
//
//   npx tsx scripts/build-indic-table.ts /path/to/indic.tsv

import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

interface Row {
  readonly insc: string;
  readonly inpc: string;
  readonly ccc: number;
  readonly decomposition: ReadonlyArray<number>;
  readonly excluded: boolean;
}

const CATEGORY: Readonly<Record<string, number>> = {
  X: 0,
  C: 1,
  V: 2,
  N: 3,
  H: 4,
  ZWNJ: 5,
  ZWJ: 6,
  M: 7,
  SM: 8,
  A: 10,
  PLACEHOLDER: 11,
  DOTTEDCIRCLE: 12,
  RS: 13,
  MPst: 14,
  Repha: 15,
  Ra: 16,
  CM: 17,
  Symbol: 18,
  CS: 19,
};

const POSITION: Readonly<Record<string, number>> = {
  START: 0,
  RA_TO_BECOME_REPH: 1,
  PRE_M: 2,
  PRE_C: 3,
  BASE_C: 4,
  AFTER_MAIN: 5,
  ABOVE_C: 6,
  BEFORE_SUB: 7,
  BELOW_C: 8,
  AFTER_SUB: 9,
  BEFORE_POST: 10,
  POST_C: 11,
  AFTER_POST: 12,
  SMVD: 13,
  END: 14,
};

/** Indic_Syllabic_Category → HarfBuzz's category. */
const CATEGORY_OF: Readonly<Record<string, string>> = {
  Other: 'X',
  Avagraha: 'Symbol',
  Bindu: 'SM',
  Brahmi_Joining_Number: 'PLACEHOLDER',
  Cantillation_Mark: 'A',
  Consonant: 'C',
  Consonant_Dead: 'C',
  Consonant_Final: 'CM',
  Consonant_Head_Letter: 'C',
  Consonant_Initial_Postfixed: 'C',
  Consonant_Killer: 'M',
  Consonant_Medial: 'CM',
  Consonant_Placeholder: 'PLACEHOLDER',
  Consonant_Preceding_Repha: 'Repha',
  Consonant_Prefixed: 'X',
  Consonant_Subjoined: 'CM',
  Consonant_Succeeding_Repha: 'CM',
  Consonant_With_Stacker: 'CS',
  Gemination_Mark: 'SM',
  Invisible_Stacker: 'H',
  Joiner: 'ZWJ',
  Modifying_Letter: 'X',
  Non_Joiner: 'ZWNJ',
  Nukta: 'N',
  Number: 'PLACEHOLDER',
  Number_Joiner: 'PLACEHOLDER',
  Pure_Killer: 'M',
  Register_Shifter: 'RS',
  Syllable_Modifier: 'SM',
  Tone_Letter: 'X',
  Tone_Mark: 'N',
  Virama: 'H',
  Visarga: 'SM',
  Vowel: 'V',
  Vowel_Dependent: 'M',
  Vowel_Independent: 'V',
};

/** Indic_Positional_Category → HarfBuzz's position; a split sign takes its last part's. */
const POSITION_OF: Readonly<Record<string, string>> = {
  Not_Applicable: 'END',
  Left: 'PRE_C',
  Top: 'ABOVE_C',
  Bottom: 'BELOW_C',
  Right: 'POST_C',
  Bottom_And_Right: 'POST_C',
  Left_And_Right: 'POST_C',
  Top_And_Bottom: 'BELOW_C',
  Top_And_Bottom_And_Left: 'BELOW_C',
  Top_And_Bottom_And_Right: 'POST_C',
  Top_And_Left: 'ABOVE_C',
  Top_And_Left_And_Right: 'POST_C',
  Top_And_Right: 'POST_C',
  Overstruck: 'AFTER_MAIN',
  Visual_Order_Left: 'PRE_M',
};

/** The letters a reph is made of: each script's Ra. */
const RA = new Set([
  0x0930, 0x09b0, 0x09f0, 0x0a30, 0x0ab0, 0x0b30, 0x0bb0, 0x0c30, 0x0cb0, 0x0d30,
]);

/** Where HarfBuzz departs from the Unicode category, each with its reason there. */
const CATEGORY_OVERRIDES = new Map<number, string>([
  // Devanagari's grave and acute accents act like the bindus.
  [0x0953, 'SM'],
  [0x0954, 'SM'],
  [0x0a40, 'MPst'], // Gurmukhi II may follow a bindi or a tippi.
  [0x0a51, 'M'],
  [0x0a72, 'C'],
  [0x0a73, 'C'],
  [0x0afb, 'N'],
  [0x0b55, 'N'],
  [0x0980, 'PLACEHOLDER'],
  [0x09fc, 'PLACEHOLDER'],
  [0x0c80, 'PLACEHOLDER'],
  [0x1ced, 'A'],
  [0x1cf5, 'CS'],
  [0x1cf6, 'CS'],
  [0x2010, 'PLACEHOLDER'],
  [0x2011, 'PLACEHOLDER'],
  [0x25cc, 'DOTTEDCIRCLE'],
  ...range(0x1ce2, 0x1ce8).map((cp) => [cp, 'A'] as const),
  ...[...range(0xa8f2, 0xa8f7), ...range(0x1ce9, 0x1cec), ...range(0x1cee, 0x1cf1)].map(
    (cp) => [cp, 'Symbol'] as const,
  ),
]);

const POSITION_OVERRIDES = new Map<number, string>([
  [0x0a51, 'BELOW_C'],
  [0x0b01, 'BEFORE_SUB'], // Oriya's candrabindu is before-sub in the specification.
]);

const BLOCKS: ReadonlyArray<readonly [string, number]> = [
  ['Devanagari', 0x0900],
  ['Bengali', 0x0980],
  ['Gurmukhi', 0x0a00],
  ['Gujarati', 0x0a80],
  ['Oriya', 0x0b00],
  ['Tamil', 0x0b80],
  ['Telugu', 0x0c00],
  ['Kannada', 0x0c80],
  ['Malayalam', 0x0d00],
];

const blockOf = (cp: number): string =>
  BLOCKS.find(([, lo]) => cp >= lo && cp < lo + 0x80)?.[0] ?? 'Other';

/** Where a vowel sign drawn at `side` is sorted into, script by script (HarfBuzz's matra_pos_*). */
function matraPosition(cp: number, side: string): string {
  const block = blockOf(cp);
  if (side === 'PRE_C') return 'PRE_M';
  if (side === 'POST_C') {
    if (block === 'Devanagari') return 'AFTER_SUB';
    if (['Bengali', 'Gurmukhi', 'Gujarati', 'Oriya', 'Tamil', 'Malayalam'].includes(block)) {
      return 'AFTER_POST';
    }
    if (block === 'Telugu') return cp <= 0x0c42 ? 'BEFORE_SUB' : 'AFTER_SUB';
    if (block === 'Kannada') return cp < 0x0cc3 || cp > 0x0cd6 ? 'BEFORE_SUB' : 'AFTER_SUB';
    return 'AFTER_SUB';
  }
  if (side === 'ABOVE_C') {
    if (block === 'Gurmukhi') return 'AFTER_POST';
    if (block === 'Oriya') return 'AFTER_MAIN';
    if (block === 'Telugu' || block === 'Kannada') return 'BEFORE_SUB';
    return 'AFTER_SUB';
  }
  if (side === 'BELOW_C') {
    if (['Gurmukhi', 'Gujarati', 'Tamil', 'Malayalam'].includes(block)) return 'AFTER_POST';
    if (block === 'Telugu' || block === 'Kannada') return 'BEFORE_SUB';
    return 'AFTER_SUB';
  }
  return side;
}

const CONSONANTS = new Set(['C', 'CS', 'Ra', 'CM', 'V', 'PLACEHOLDER', 'DOTTEDCIRCLE']);

function properties(cp: number, row: Row): readonly [number, number] {
  let category = CATEGORY_OVERRIDES.get(cp) ?? CATEGORY_OF[row.insc] ?? 'X';
  let position = POSITION_OF[row.inpc] ?? 'END';
  if (CONSONANTS.has(category)) {
    position = 'BASE_C';
    if (RA.has(cp)) category = 'Ra';
  } else if (category === 'M' || category === 'MPst') {
    position = matraPosition(cp, position);
  } else if (category === 'SM' || category === 'A' || category === 'Symbol') {
    position = 'SMVD';
  }
  position = POSITION_OVERRIDES.get(cp) ?? position;
  return [CATEGORY[category]!, POSITION[position]!];
}

function range(lo: number, hi: number): Array<number> {
  return Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
}

const hex = (cp: number): string => `0x${cp.toString(16).padStart(4, '0')}`;
const DIGITS = '0123456789abcdefghijklmnopqrstuv';

function main(): void {
  const path = process.argv[2];
  if (!path) {
    process.stderr.write('usage: tsx scripts/build-indic-table.ts <indic.tsv>\n');
    process.exit(2);
  }
  const text = readFileSync(path, 'utf8');
  const version = /^# (ICU .*)$/mu.exec(text)?.[1] ?? 'ICU';
  const rows = new Map<number, Row>();
  for (const line of text.split('\n')) {
    if (line === '' || line.startsWith('#')) continue;
    const [cp, insc, inpc, , ccc, dec, excl] = line.split('\t');
    rows.set(parseInt(cp!, 16), {
      insc: insc!,
      inpc: inpc!,
      ccc: Number(ccc),
      decomposition: dec ? dec.split(' ').map((d) => parseInt(d, 16)) : [],
      excluded: excl === '1',
    });
  }
  const encode = (lo: number, hi: number): string =>
    range(lo, hi)
      .map((cp) => {
        const row = rows.get(cp);
        const [c, p] = row ? properties(cp, row) : [0, POSITION.END!];
        return DIGITS[c]! + DIGITS[p]!;
      })
      .join('');
  // A string constant, on its own line once `export const NAME = …;` would
  // run past Prettier's 100 columns.
  const constant = (name: string, value: string): string => {
    const one = `export const ${name} = '${value}';`;
    return one.length <= 100 ? one : `export const ${name} =\n  '${value}';`;
  };

  const inBlocks = (cp: number): boolean =>
    (cp >= 0x0900 && cp <= 0x0d7f) ||
    (cp >= 0x1cd0 && cp <= 0x1cff) ||
    (cp >= 0xa8e0 && cp <= 0xa8ff);
  // Outside the blocks: the characters HarfBuzz's table gives a part — the
  // digits, the dashes and the no-break space a vowel sign may be shown on,
  // the join controls, the dotted circle, and the superscripts the Vedic
  // texts use as tone marks.
  const singles = [...rows.keys()]
    .filter((cp) => cp < 0x0900 || (cp >= 0x2000 && cp < 0x2090) || cp === 0x25cc)
    .filter((cp) => rows.get(cp)!.insc !== 'Other' || CATEGORY_OVERRIDES.has(cp))
    .sort((a, b) => a - b);
  const decompositions = [...rows.entries()]
    .filter(([cp, r]) => cp >= 0x0900 && cp <= 0x0d7f && r.decomposition.length === 2)
    .sort(([a], [b]) => a - b);
  const combining = [...rows.entries()]
    .filter(([cp, r]) => r.ccc !== 0 && inBlocks(cp))
    .sort(([a], [b]) => a - b);

  const body = [
    '// The character properties of the scripts of India, as the Indic shaper',
    "// (`indic-shaper.ts`) uses them: Unicode's Indic_Syllabic_Category and",
    "// Indic_Positional_Category mapped to HarfBuzz's categories and positions,",
    '// and the canonical decompositions and combining classes its normalization',
    '// needs.',
    '//',
    `// GENERATED by \`scripts/build-indic-table.ts\` from ${version}.`,
    '// Do not edit by hand: fix the script and run it again.',
    '',
    '/** Category and position of U+0900..U+0D7F, two base-32 digits each. */',
    'export const MAIN =',
    BLOCKS.map(([, lo]) => `  '${encode(lo, lo + 0x7f)}'`).join(' +\n') + ';',
    '/** …of U+1CD0..U+1CFF, the Vedic extensions. */',
    constant('VEDIC', encode(0x1cd0, 0x1cff)),
    '/** …of U+A8E0..U+A8FF, Devanagari extended. */',
    constant('EXTENDED', encode(0xa8e0, 0xa8ff)),
    '/** The code points outside those blocks that take part: [code point, category, position]. */',
    'export const SINGLES: ReadonlyArray<readonly [number, number, number]> = [',
    ...singles.map((cp) => {
      const [c, p] = properties(cp, rows.get(cp)!);
      return `  [${hex(cp)}, ${c}, ${p}],`;
    }),
    '];',
    '/** Canonical decompositions in the blocks: [composite, first, second, excluded from composition]. */',
    'export const DECOMPOSITIONS: ReadonlyArray<readonly [number, number, number, boolean]> = [',
    ...decompositions.map(
      ([cp, r]) =>
        `  [${hex(cp)}, ${hex(r.decomposition[0]!)}, ${hex(r.decomposition[1]!)}, ${String(r.excluded)}],`,
    ),
    '];',
    '/** Canonical combining classes other than 0. */',
    'export const COMBINING: ReadonlyArray<readonly [number, number]> = [',
    ...combining.map(([cp, r]) => `  [${hex(cp)}, ${r.ccc}],`),
    '];',
    '',
  ].join('\n');
  const out = resolve('src/core/font/indic-table.ts');
  writeFileSync(out, body);
  process.stdout.write(
    `${out}: ${String(singles.length)} singles, ${String(decompositions.length)} decompositions, ${String(Math.round(body.length / 1024))}KB\n`,
  );
}

main();
