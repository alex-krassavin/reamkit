import { describe, expect, it } from 'vitest';

import { analyzeString, hasBidiCharacters, reorderVisual, reverseByCodePoint } from '@/bidi';

// Hebrew sample letters (strong R).
const ALEF = 'א'; // א
const BET = 'ב'; // ב
const GIMEL = 'ג'; // ג
const HEB = ALEF + BET + GIMEL; // "אבג"

// Arabic sample letters (strong AL).
const ALIF = 'ا'; // ا
const BAA = 'ب'; // ب
const TAA = 'ت'; // ت
const ARABIC = ALIF + BAA + TAA;

describe('BiDi: paragraph level (P2/P3)', () => {
  it('LTR text gets paragraph level 0', () => {
    expect(analyzeString('hello').paragraphLevel).toBe(0);
  });

  it('Hebrew-first text auto-detects RTL paragraph level 1', () => {
    expect(analyzeString(HEB).paragraphLevel).toBe(1);
  });

  it('Arabic-first text auto-detects RTL paragraph level 1', () => {
    expect(analyzeString(ARABIC).paragraphLevel).toBe(1);
  });

  it('explicit direction overrides auto-detection', () => {
    expect(analyzeString(HEB, 'ltr').paragraphLevel).toBe(0);
    expect(analyzeString('hello', 'rtl').paragraphLevel).toBe(1);
  });

  it('leading neutral then Hebrew still resolves RTL', () => {
    expect(analyzeString(`  ${HEB}`).paragraphLevel).toBe(1);
  });
});

describe('BiDi: embedding levels', () => {
  it('pure LTR → all level 0', () => {
    const { levels } = analyzeString('abc');
    expect(levels).toEqual([0, 0, 0]);
  });

  it('pure Hebrew (auto RTL) → all level 1', () => {
    const { levels } = analyzeString(HEB);
    expect(levels).toEqual([1, 1, 1]);
  });

  it('Latin embedded in an RTL paragraph gets an even level', () => {
    // "אבג abc" base RTL: Hebrew at 1, the Latin "abc" rises to level 2.
    const { levels, paragraphLevel } = analyzeString(`${HEB} abc`);
    expect(paragraphLevel).toBe(1);
    // Indices: 0,1,2 Hebrew (1); 3 space; 4,5,6 latin.
    expect(levels[0]).toBe(1);
    expect(levels[4]).toBe(2);
    expect(levels[5]).toBe(2);
    expect(levels[6]).toBe(2);
  });

  it('Hebrew embedded in an LTR paragraph gets an odd level', () => {
    // "abc אבג" base LTR: latin level 0, Hebrew level 1.
    const { levels, paragraphLevel } = analyzeString(`abc ${HEB}`);
    expect(paragraphLevel).toBe(0);
    expect(levels[0]).toBe(0);
    expect(levels[4]).toBe(1);
    expect(levels[6]).toBe(1);
  });

  it('European digits after Latin stay level 0 (W7)', () => {
    const { levels } = analyzeString('a1');
    expect(levels).toEqual([0, 0]);
  });

  it('European digits in RTL context get level 2 (I1)', () => {
    // Hebrew then digits: digits become EN at level 2 inside RTL.
    const { levels } = analyzeString(`${ALEF}12`);
    expect(levels[0]).toBe(1); // Hebrew
    expect(levels[1]).toBe(2); // digit
    expect(levels[2]).toBe(2);
  });
});

describe('BiDi: paired brackets (N0)', () => {
  it('resolves a bracket pair from the strong type inside it', () => {
    // "a(bא)" base RTL: '(' sits between Latin a and Latin b (same direction L),
    // so N1 alone would give it an even (LTR) level — but the pair contains a
    // strong R (Hebrew alef), so N0 sets both brackets to R → an odd level.
    expect(analyzeString(`a(b${ALEF})`, 'rtl').levels[1]).toBe(1); // '('
    expect(analyzeString(`a(b${ALEF})`, 'rtl').levels[4]).toBe(1); // ')'
  });

  it('keeps brackets at the opposite direction when only that appears inside', () => {
    // "a(b)" base RTL: only Latin inside and a Latin before → N0 keeps the
    // brackets L (even level). Confirms N0 reads the content, not a constant.
    expect(analyzeString('a(b)', 'rtl').levels[1]).toBe(2);
  });
});

describe('BiDi: L2 reordering', () => {
  it('identity for all-LTR levels', () => {
    expect(reorderVisual([0, 0, 0])).toEqual([0, 1, 2]);
  });

  it('full reversal for all-RTL levels', () => {
    expect(reorderVisual([1, 1, 1])).toEqual([2, 1, 0]);
  });

  it('reverses only the RTL span in mixed content', () => {
    // Levels for "abc <HEB>" base LTR: [0,0,0, 0(space), 1,1,1].
    // Visual: latin stays, Hebrew span reverses.
    const order = reorderVisual([0, 0, 0, 0, 1, 1, 1]);
    expect(order).toEqual([0, 1, 2, 3, 6, 5, 4]);
  });

  it('nested: digits inside RTL keep internal LTR order', () => {
    // "<ALEF>12" base RTL → levels [1,2,2]. The number "12" (level 2) must
    // stay in logical order while the whole line is RTL.
    // Visual order: reverse level≥1 → [12-pair-then-alef], then reverse
    // level≥2 inside flips the digits back. Net: digits read "12", alef right.
    const order = reorderVisual([1, 2, 2]);
    // Highest level 2 run [1,2] reversed → [2,1]; then level 1 run [0,2,1]
    // reversed → [1,2,0]. So visual index 0 = logical 1, etc.
    expect(order).toEqual([1, 2, 0]);
  });
});

describe('BiDi: helpers', () => {
  it('hasBidiCharacters detects Hebrew and Arabic, ignores Latin', () => {
    expect(hasBidiCharacters('plain ascii')).toBe(false);
    expect(hasBidiCharacters(`mixed ${ALEF}`)).toBe(true);
    expect(hasBidiCharacters(ARABIC)).toBe(true);
  });

  it('reverseByCodePoint reverses without splitting surrogate pairs', () => {
    expect(reverseByCodePoint('abc')).toBe('cba');
    // U+1F600 is a surrogate pair; it must survive reversal intact.
    expect(reverseByCodePoint('a\u{1F600}b')).toBe('b\u{1F600}a');
  });
});
