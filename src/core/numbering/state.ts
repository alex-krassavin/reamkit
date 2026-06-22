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
  // numId → counters indexed by ilvl (0..8). 0 means "not yet started".
  private readonly counters = new Map<string, Array<number>>();

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
      arr = new Array<number>(9).fill(0);
      this.counters.set(ref.numId, arr);
    }

    // Deeper levels reset whenever a shallower level advances.
    for (let k = ref.ilvl + 1; k < arr.length; k++) arr[k] = 0;

    if (arr[ref.ilvl] === 0) {
      arr[ref.ilvl] = level.start;
    } else {
      arr[ref.ilvl]! += 1;
    }

    return formatLevelMarker(abstractNum, level, arr);
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
    const fmt = level?.format ?? 'decimal';
    return formatCounter(fmt, counter);
  });
}

function formatCounter(format: NumberingFormat, n: number): string {
  if (n <= 0) return '';
  switch (format) {
    case 'decimal':
      return String(n);
    case 'lowerLetter':
      return toLetters(n).toLowerCase();
    case 'upperLetter':
      return toLetters(n);
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperRoman':
      return toRoman(n);
    case 'bullet':
    case 'none':
    default:
      return '';
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
