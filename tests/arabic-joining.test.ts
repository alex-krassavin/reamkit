import { describe, expect, it } from 'vitest';

import { shapeText } from '@/core/font';
import { arabicJoiningType, assignArabicForms } from '@/core/font/arabic-joining';

const cps = (s: string): Array<number> => [...s].map((c) => c.codePointAt(0)!);

describe('Arabic joining types', () => {
  it('classifies the core letters', () => {
    expect(arabicJoiningType(0x0628)).toBe('D'); // beh — dual
    expect(arabicJoiningType(0x062d)).toBe('D'); // hah — dual
    expect(arabicJoiningType(0x0627)).toBe('R'); // alef — right
    expect(arabicJoiningType(0x0631)).toBe('R'); // reh — right
    expect(arabicJoiningType(0x0640)).toBe('C'); // tatweel — join-causing
    expect(arabicJoiningType(0x064e)).toBe('T'); // fatha — transparent mark
    expect(arabicJoiningType(0x0041)).toBe('U'); // Latin A — non-joining
  });
});

describe('assignArabicForms', () => {
  it('shapes a three-letter word (beh-hah-reh) init/medi/fina', () => {
    expect(assignArabicForms(cps('بحر'))).toEqual(['init', 'medi', 'fina']);
  });

  it('shapes a dual+right pair as init/final', () => {
    // beh (D) + alef (R): alef only joins on its right, so beh is initial, alef final.
    expect(assignArabicForms(cps('با'))).toEqual(['init', 'fina']);
  });

  it('leaves an isolated letter and non-joining characters isolated', () => {
    expect(assignArabicForms(cps('ا'))).toEqual(['isol']); // lone alef
    expect(assignArabicForms(cps('Aب'))).toEqual(['isol', 'isol']); // Latin breaks the run
  });

  it('skips transparent marks when choosing forms', () => {
    // beh + fatha(mark) + reh → beh joins reh through the mark → init, (mark), fina.
    expect(assignArabicForms(cps('بَر'))).toEqual(['init', 'isol', 'fina']);
  });
});

describe('shapeText applies cursive forms via init/medi/fina maps', () => {
  it('substitutes the contextual glyph for each position', () => {
    // beh→gid 10 (init 11), alef→gid 20 (fina 21).
    const gidOf = (cp: number): number => (cp === 0x0628 ? 10 : cp === 0x0627 ? 20 : 0);
    const advances = new Array(40).fill(100);
    const joiningForms = {
      init: new Map([[10, 11]]),
      medi: new Map<number, number>(),
      fina: new Map([[20, 21]]),
    };
    const shaped = shapeText('با', gidOf, advances, new Map(), new Map(), joiningForms);
    expect(shaped.gids).toEqual([11, 21]); // beh→initial, alef→final
  });

  it('is a no-op when the font has no joining forms', () => {
    const gidOf = (cp: number): number => (cp === 0x0628 ? 10 : 20);
    const advances = new Array(40).fill(100);
    const shaped = shapeText('با', gidOf, advances, new Map(), new Map());
    expect(shaped.gids).toEqual([10, 20]);
  });
});
