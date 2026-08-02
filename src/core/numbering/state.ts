// Numbering counter state and marker formatting.
//
// ECMA-376 Part 1 §17.9.21 (lvl/lvlText) — the lvlText is a printf-style
// template with %N placeholders where N is 1-indexed level. Each %N is
// substituted with the counter of level N-1, formatted according to that
// level's numFmt.

import type {
  AbstractNumbering,
  Numbering,
  NumberingFormat,
  NumberingLevel,
  NumberingReference,
} from '@/core/document-model';

/**
 * Mutable list-counter state for one numbering scope (e.g. a body or a single
 * header/footer band). Tracks a counter per list per level and advances them as
 * paragraphs are visited in order.
 */
export class NumberingState {
  // numId → counters indexed by ilvl (0..8). An empty slot means "not yet
  // started" — a sentinel of 0 could not, since §17.9.25 lets a level START at
  // zero and the second item would then seed itself again.
  private readonly counters = new Map<string, Array<number | undefined>>();

  /**
   * Advance the counter for `ref` (resetting deeper levels) and format its
   * marker text. The first hit at a level seeds it from the level's `start`;
   * subsequent hits increment.
   *
   * @param numbering The parsed numbering definitions.
   * @param ref       The paragraph's `numId` + `ilvl` reference.
   * @returns The formatted marker (e.g. `"2."`, `"•"`), or `null` when the
   *          reference does not resolve to a known list level.
   */
  resolveMarker(numbering: Numbering, ref: NumberingReference): string | null {
    const instance = numbering.numInstances.get(ref.numId);
    if (!instance) return null;
    const abstractNum = numbering.abstractNums.get(instance.abstractNumId);
    if (!abstractNum) return null;
    const level = abstractNum.levels.get(ref.ilvl);
    if (!level) return null;

    let arr = this.counters.get(ref.numId);
    if (!arr) {
      arr = new Array<number | undefined>(9).fill(undefined);
      this.counters.set(ref.numId, arr);
    }

    // Deeper levels reset whenever a shallower level advances.
    for (let k = ref.ilvl + 1; k < arr.length; k++) arr[k] = undefined;

    // §17.9.28 — where THIS instance starts the level, if it says.
    const startAt = (i: number, fallback: number | undefined): number =>
      instance.startOverrides?.get(i) ?? fallback ?? 0;

    const current = arr[ref.ilvl];
    arr[ref.ilvl] = current === undefined ? startAt(ref.ilvl, level.start) : current + 1;

    // A level that has not been reached yet still numbers the levels ABOVE it
    // in a multi-level marker; an unstarted one counts as its own start.
    const counts = arr.map((n, i) => n ?? startAt(i, abstractNum.levels.get(i)?.start));
    return formatLevelMarker(abstractNum, level, counts);
  }
}

export function formatLevelMarker(
  abstractNum: AbstractNumbering,
  currentLevel: NumberingLevel,
  counters: ReadonlyArray<number>,
): string {
  if (currentLevel.format === 'bullet') {
    return normalizeBullet(currentLevel.lvlText);
  }
  return currentLevel.lvlText.replace(/%(\d)/g, (_match, n) => {
    const lvlIdx = Number(n) - 1;
    const level = abstractNum.levels.get(lvlIdx);
    const counter = counters[lvlIdx] ?? 0;
    // §17.9.10 — a LEGAL level prints the levels ABOVE it in decimal, whatever
    // format each asks for, and keeps its own: listWithLgl.docx numbers its
    // chapters in Roman and its sections "Sect 1.01", not "Sect I.01".
    const legal = currentLevel.isLegal === true && lvlIdx !== currentLevel.ilvl;
    const fmt = legal ? 'decimal' : (level?.format ?? 'decimal');
    return formatCounter(fmt, counter);
  });
}

function formatCounter(format: NumberingFormat, n: number): string {
  // §17.9.25 lets a level start at zero, and a digit format prints it. A
  // letter or a numeral has no zero, so those stay blank.
  if (n < 0) return '';
  switch (format) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      // §17.18.59 — a leading zero below ten (01, 02, … 10, 11).
      return n < 10 ? `0${String(n)}` : String(n);
    case 'decimalFullWidth':
      return digitByDigit(n, FULL_WIDTH_DIGITS);
    case 'ordinal':
      return `${String(n)}${ordinalSuffix(n)}`;
    case 'lowerLetter':
      return n === 0 ? '' : toLetters(n).toLowerCase();
    case 'upperLetter':
      return n === 0 ? '' : toLetters(n);
    case 'lowerRoman':
      return n === 0 ? '' : toRoman(n).toLowerCase();
    case 'upperRoman':
      return n === 0 ? '' : toRoman(n);
    // A cycle, and past its end both references number the rest in digits.
    case 'ideographTraditional':
      return n === 0 ? '' : n <= HEAVENLY_STEMS.length ? HEAVENLY_STEMS[n - 1]! : String(n);
    case 'ideographZodiac':
      return n === 0 ? '' : n <= EARTHLY_BRANCHES.length ? EARTHLY_BRANCHES[n - 1]! : String(n);
    case 'ideographDigital':
    case 'koreanDigital2':
      return digitByDigit(n, IDEOGRAPH_DIGITS);
    case 'ideographLegalTraditional':
      // The formal numerals keep the unit's own digit: ten is 壹拾, not 拾.
      return counting(n, LEGAL_DIGITS, LEGAL_UNITS, false);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
    case 'koreanCounting':
    case 'taiwaneseCounting':
    case 'taiwaneseCountingThousand':
      return counting(n, IDEOGRAPH_DIGITS, COUNTING_UNITS, true);
    case 'hebrew1':
      return n === 0 ? '' : toHebrewNumeral(n);
    case 'hebrew2':
      // The alphabet as a plain sequence, cycling once it runs out.
      return n === 0 ? '' : HEBREW_ALPHABET[(n - 1) % HEBREW_ALPHABET.length]!;
    case 'decimalEnclosedCircle':
      // U+2460 ① … U+2473 ⑳; past twenty Word prints the number plainly.
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x245f + n) : String(n);
    case 'bullet':
    case 'none':
    default:
      return '';
  }
}

// §17.18.59 `hebrew1` — the gematria numerals. Hundreds run ק ר ש ת and repeat
// ת past four hundred; 15 and 16 are written טו and טז rather than spelling out
// the divine name.
const HEBREW_ONES = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט'] as const;
const HEBREW_TENS = ['', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ'] as const;
const HEBREW_HUNDREDS = ['', 'ק', 'ר', 'ש', 'ת'] as const;
const HEBREW_ALPHABET = [
  'א',
  'ב',
  'ג',
  'ד',
  'ה',
  'ו',
  'ז',
  'ח',
  'ט',
  'י',
  'כ',
  'ל',
  'מ',
  'נ',
  'ס',
  'ע',
  'פ',
  'צ',
  'ק',
  'ר',
  'ש',
  'ת',
] as const;

function toHebrewNumeral(n: number): string {
  let rest = n;
  let out = '';
  while (rest >= 400) {
    out += 'ת';
    rest -= 400;
  }
  out += HEBREW_HUNDREDS[Math.floor(rest / 100)] ?? '';
  rest %= 100;
  if (rest === 15 || rest === 16) {
    out += rest === 15 ? 'טו' : 'טז';
  } else {
    out += HEBREW_TENS[Math.floor(rest / 10)] ?? '';
    out += HEBREW_ONES[rest % 10] ?? '';
  }
  // Word and LibreOffice both print the bare letters in a list marker — the
  // geresh and gershayim a Hebrew numeral carries in running text are not part
  // of the number here.
  return out;
}

// §17.18.59 ideograph number formats. `IDEOGRAPH_DIGITS[0]` is the zero the
// digit-by-digit formats print (LibreOffice writes 10 as 一零).
const IDEOGRAPH_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'] as const;
const COUNTING_UNITS = ['', '十', '百', '千'] as const;
const LEGAL_DIGITS = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'] as const;
const LEGAL_UNITS = ['', '拾', '佰', '仟'] as const;
const HEAVENLY_STEMS = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'] as const;
const EARTHLY_BRANCHES = [
  '子',
  '丑',
  '寅',
  '卯',
  '辰',
  '巳',
  '午',
  '未',
  '申',
  '酉',
  '戌',
  '亥',
] as const;
const FULL_WIDTH_DIGITS = ['０', '１', '２', '３', '４', '５', '６', '７', '８', '９'] as const;

/** Each decimal digit as its own glyph — no units, no elision (10 → 一零). */
function digitByDigit(n: number, digits: ReadonlyArray<string>): string {
  return String(n)
    .split('')
    .map((d) => digits[Number(d)] ?? d)
    .join('');
}

/**
 * The counting systems: digits carrying place units (十百千). `elideTen` drops
 * the leading 一 of the tens place, which is what the plain counting systems do
 * (11 → 十一) and the formal one does not (11 → 壹拾壹). Above the units the
 * table covers, plain digits — a list marker never gets that far.
 */
function counting(
  n: number,
  digits: ReadonlyArray<string>,
  units: ReadonlyArray<string>,
  elideTen: boolean,
): string {
  if (n >= 10 ** units.length) return String(n);
  const ds = String(n).split('').map(Number);
  let out = '';
  ds.forEach((d, i) => {
    const place = ds.length - 1 - i;
    if (d === 0) {
      // An interior zero is spoken once, and never at the end (100 → 一百).
      if (out.length > 0 && !out.endsWith(digits[0]!) && ds.slice(i).some((x) => x > 0)) {
        out += digits[0]!;
      }
      return;
    }
    const head = elideTen && d === 1 && place === 1 && i === 0 ? '' : digits[d]!;
    out += head + units[place]!;
  });
  return out;
}

/** `1st`, `2nd`, `3rd`, `4th` — the English ordinal suffixes (§17.18.59). */
function ordinalSuffix(n: number): string {
  const two = n % 100;
  if (two >= 11 && two <= 13) return 'th';
  switch (n % 10) {
    case 1:
      return 'st';
    case 2:
      return 'nd';
    case 3:
      return 'rd';
    default:
      return 'th';
  }
}

function toLetters(n: number): string {
  // Spreadsheet-style: 1=A, 26=Z, 27=AA, 28=AB, …
  let s = '';
  let v = n;
  while (v > 0) {
    v--;
    s = String.fromCharCode(65 + (v % 26)) + s;
    v = Math.floor(v / 26);
  }
  return s;
}

function toRoman(n: number): string {
  const map: ReadonlyArray<readonly [number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let s = '';
  let v = n;
  for (const [value, ch] of map) {
    while (v >= value) {
      s += ch;
      v -= value;
    }
  }
  return s;
}

// Word's default bullets use private-use codepoints from the Symbol font
// (e.g. U+F0B7). Our substitute text fonts have no Symbol glyph, so we
// substitute the Unicode bullet (U+2022) which every general-purpose font
// covers.
function normalizeBullet(lvlText: string): string {
  if (lvlText.length === 0) return '•';
  const cp = lvlText.codePointAt(0)!;
  if (cp >= 0xe000 && cp <= 0xf8ff) return '•';
  return lvlText;
}
