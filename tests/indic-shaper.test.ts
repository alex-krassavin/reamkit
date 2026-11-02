// The scripts of India, shaped. Every expected glyph, cluster, advance and
// offset here is HarfBuzz's own: hb-shape 15.0.0 with
// --remove-default-ignorables, on the same fixture faces — Noto Sans of each
// script cut down to the characters below (see fixtures/fonts/NOTICE.md).
// Each case is [glyph, cluster, advance, x offset, y offset] per glyph, in
// font units, in the order the glyphs are drawn.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { parseTtf } from '@/core/font';
import { hasIndic, shapeIndic } from '@/core/font/indic-shaper';

type Case = readonly [string, ReadonlyArray<readonly [number, number, number, number, number]>];

const CASES: Readonly<Record<string, ReadonlyArray<Case>>> = {
  Devanagari: [
    // the short i before ha, a half na
    [
      'हिन्दी',
      [
        [105, 0, 259, 0, 0],
        [17, 0, 531, 0, 0],
        [26, 2, 309, 0, 0],
        [8, 4, 531, 0, 0],
        [5, 4, 259, 0, 0],
      ],
    ],
    // a reph over ma, last in the syllable
    [
      'धर्म',
      [
        [9, 0, 615, 0, 0],
        [12, 1, 598, 0, 0],
        [99, 1, 0, 0, 0],
      ],
    ],
    // ksha, a ligature
    [
      'क्षमा',
      [
        [18, 0, 717, 0, 0],
        [12, 3, 598, 0, 0],
        [3, 3, 259, 0, 0],
      ],
    ],
    // pa with the ra under it
    [
      'प्रधान',
      [
        [47, 0, 569, 0, 0],
        [9, 3, 615, 0, 0],
        [3, 3, 259, 0, 0],
        [10, 5, 555, 0, 0],
      ],
    ],
    // qa: ka and a nukta, one glyph
    [
      'क़लम',
      [
        [21, 0, 768, 0, 0],
        [14, 2, 678, 0, 0],
        [12, 3, 598, 0, 0],
      ],
    ],
    // ka, virama, ZWJ: a half ka
    [
      'क्\u200dष',
      [
        [24, 0, 550, 0, 0],
        [16, 3, 578, 0, 0],
      ],
    ],
    // ka, virama, ZWNJ: the virama stays
    [
      'क्\u200cष',
      [
        [7, 0, 768, 0, 0],
        [19, 0, 0, -221, 0],
        [16, 3, 578, 0, 0],
      ],
    ],
    // a vowel sign on nothing sits on a dotted circle
    [
      'ि',
      [
        [4, 0, 259, 0, 0],
        [146, 0, 510, 0, 0],
      ],
    ],
    // a, short e: written as one vowel that is not one
    [
      'अॆ',
      [
        [2, 0, 764, 0, 0],
        [146, 0, 510, 0, 0],
        [6, 0, 0, 6, 0],
      ],
    ],
    // shri
    [
      'श्री',
      [
        [50, 0, 708, 0, 0],
        [5, 0, 259, 0, 0],
      ],
    ],
    // a reph and a short i in one syllable
    [
      'र्कि',
      [
        [121, 0, 259, 0, 0],
        [7, 0, 768, 0, 0],
        [143, 0, 0, 0, 0],
      ],
    ],
    // a ZWJ before a letter: its cluster goes forward to the letter
    ['\u200dक', [[7, 0, 768, 0, 0]]],
    // qa as one code point, taken apart into ka and the nukta
    ['क़', [[21, 0, 768, 0, 0]]],
  ],
  Bengali: [
    // o, written once and drawn on both sides
    [
      'কোথায়',
      [
        [72, 0, 361, 0, 0],
        [6, 0, 807, 0, 0],
        [14, 0, 266, 0, 0],
        [7, 2, 645, 0, 0],
        [14, 2, 266, 0, 0],
        [9, 4, 626, 0, 0],
        [13, 4, 0, -396, 108],
      ],
    ],
    // the short i before ka
    [
      'কি',
      [
        [15, 0, 266, 0, 0],
        [6, 0, 807, 0, 0],
      ],
    ],
    // ya-phala
    [
      'ক্য',
      [
        [6, 0, 807, 0, 0],
        [23, 0, 270, 0, 0],
      ],
    ],
    // a reph, after the base
    [
      'র্ক',
      [
        [6, 0, 807, 0, 0],
        [22, 0, 0, -361, 0],
      ],
    ],
    // shiksha
    [
      'শিক্ষা',
      [
        [63, 0, 266, 0, 0],
        [11, 0, 681, 0, 0],
        [24, 2, 859, 0, 0],
        [14, 2, 266, 0, 0],
      ],
    ],
    // au, on both sides
    [
      'কৌ',
      [
        [72, 0, 361, 0, 0],
        [6, 0, 807, 0, 0],
        [20, 0, 266, 0, 0],
      ],
    ],
    // o with nothing to sit on: taken apart, on a dotted circle
    [
      'ো',
      [
        [72, 0, 361, 0, 0],
        [3, 0, 510, 0, 0],
        [14, 0, 266, 0, 0],
      ],
    ],
    // a ZWJ before two marks: they do not attach across it (mark keeps its joiners)
    [
      '\u200dৣঁ',
      [
        [3, 0, 510, 0, 0],
        [21, 0, 0, 0, 0],
        [5, 0, 0, 0, 0],
      ],
    ],
  ],
  Gurmukhi: [
    // sihari before sa, a tippi
    [
      'ਸਿੰਘ',
      [
        [9, 0, 259, 0, 0],
        [8, 0, 596, 0, 0],
        [13, 0, 0, 0, 0],
        [4, 3, 757, 0, 0],
      ],
    ],
    // pa with the ra under it
    [
      'ਪ੍ਰੇਮ',
      [
        [5, 0, 593, 0, 0],
        [14, 0, 0, 0, 0],
        [11, 0, 0, 0, 0],
        [6, 4, 596, 0, 0],
      ],
    ],
    // a bindi before ii: one syllable, no dotted circle
    [
      'ਕਂੀ',
      [
        [3, 0, 622, 0, 0],
        [2, 0, 0, -169, 0],
        [10, 0, 259, 0, 0],
      ],
    ],
    // sihari before ka
    [
      'ਕਿ',
      [
        [9, 0, 259, 0, 0],
        [3, 0, 622, 0, 0],
      ],
    ],
  ],
  Gujarati: [
    // the short i before ka
    [
      'કિંમત',
      [
        [64, 0, 265, 0, 0],
        [3, 0, 511, 0, 0],
        [79, 0, 0, 0, 0],
        [7, 3, 594, 0, 0],
        [4, 4, 572, 0, 0],
      ],
    ],
    // a reph over ma
    [
      'ધર્મ',
      [
        [5, 0, 606, 0, 0],
        [7, 1, 594, 0, 0],
        [13, 1, 0, 0, 0],
      ],
    ],
    // a half sa
    [
      'પુસ્તક',
      [
        [46, 0, 568, 0, 0],
        [11, 0, 0, 0, 0],
        [21, 2, 529, 0, 0],
        [4, 4, 572, 0, 0],
        [3, 5, 511, 0, 0],
      ],
    ],
  ],
  Oriya: [
    // o, on both sides
    [
      'କୋ',
      [
        [8, 0, 534, 0, 0],
        [70, 0, 752, 0, 0],
        [5, 0, 201, 0, 0],
      ],
    ],
    // a reph, after the main consonant
    [
      'ର୍କ',
      [
        [70, 0, 752, 0, 0],
        [16, 0, 0, -486, 0],
      ],
    ],
    // ksha
    ['କ୍ଷ', [[32, 0, 785, 0, 0]]],
    // the i sign
    [
      'କି',
      [
        [70, 0, 752, 0, 0],
        [6, 0, 0, -685, 0],
      ],
    ],
    // ai, on both sides
    [
      'କୈ',
      [
        [8, 0, 534, 0, 0],
        [70, 0, 752, 0, 0],
        [13, 0, 0, -651, 0],
      ],
    ],
  ],
  Tamil: [
    // o, on both sides
    [
      'கொ',
      [
        [9, 0, 901, 0, 0],
        [2, 0, 825, 0, 0],
        [7, 0, 640, 0, 0],
      ],
    ],
    // ai, before the consonant
    [
      'கை',
      [
        [10, 0, 1160, 0, 0],
        [2, 0, 825, 0, 0],
      ],
    ],
    // ksha
    ['க்ஷ', [[15, 0, 1851, 0, 0]]],
    // shri
    ['ஸ்ரீ', [[26, 0, 1514, 0, 0]]],
    // au, on both sides
    [
      'கௌ',
      [
        [9, 0, 901, 0, 0],
        [2, 0, 825, 0, 0],
        [14, 0, 1070, 0, 0],
      ],
    ],
    // a dotted circle kerns with the ka after a ZWJ: a pair steps over the join control
    [
      'ா\u200dக்',
      [
        [28, 0, 562, 0, 0],
        [7, 0, 595, 0, 0],
        [16, 2, 825, 0, 0],
      ],
    ],
  ],
  Telugu: [
    // ksha
    [
      'క్షమ',
      [
        [15, 0, 522, 0, 0],
        [3, 3, 1058, 0, 0],
      ],
    ],
    // a reph asked for with ZWJ
    [
      'ర్\u200dక',
      [
        [2, 0, 522, 0, 0],
        [61, 0, 565, 0, 0],
      ],
    ],
    // o
    ['కొ', [[43, 0, 772, 0, 0]]],
    // shri
    [
      'శ్రీ',
      [
        [28, 0, 550, 0, 0],
        [58, 0, 0, 0, 0],
      ],
    ],
    // ai, in two parts
    ['కై', [[37, 0, 522, 0, 0]]],
    // the ai length mark sorts before the virama (combining class 5, not 91)
    [
      'క్ౖ',
      [
        [16, 0, 522, 0, 0],
        [14, 0, 0, 0, 0],
      ],
    ],
  ],
  Kannada: [
    // a reph, after the post-base forms
    [
      'ಧರ್ಮ',
      [
        [8, 0, 765, 0, 0],
        [9, 1, 1156, 0, 0],
        [28, 1, 567, 0, 0],
      ],
    ],
    // o, in two parts
    [
      'ಕೊ',
      [
        [85, 0, 574, 0, 0],
        [16, 0, 746, 0, 0],
      ],
    ],
    // stri: two below-base forms
    [
      'ಸ್ತ್ರೀ',
      [
        [11, 0, 709, 0, 0],
        [14, 0, 728, 0, 0],
        [105, 0, 208, 0, 0],
      ],
    ],
    // Ra, virama, ZWJ: read as Ra, ZWJ, virama
    [
      'ರ್\u200dಕ',
      [
        [10, 0, 651, 0, 0],
        [29, 0, 175, 0, 0],
      ],
    ],
    // ai, in two parts
    [
      'ಕೈ',
      [
        [85, 0, 574, 0, 0],
        [25, 0, 183, 0, 0],
      ],
    ],
  ],
  Malayalam: [
    // o, on both sides
    [
      'കൊ',
      [
        [8, 0, 715, 0, 0],
        [3, 0, 1038, 0, 0],
        [7, 0, 504, 0, 0],
      ],
    ],
    // kra: the ra drawn before ka (pref)
    [
      'ക്ര',
      [
        [19, 0, 244, 0, 0],
        [3, 0, 1038, 0, 0],
      ],
    ],
    // the dot reph
    [
      'ൎക',
      [
        [3, 0, 1038, 0, 0],
        [12, 0, 0, -241, 154],
      ],
    ],
    // dharmam, with a chillu
    [
      'ധർമ്മം',
      [
        [4, 0, 911, 0, 0],
        [14, 1, 736, 0, 0],
        [21, 2, 1240, 0, 0],
        [2, 2, 453, 0, 0],
      ],
    ],
    // ai, before the consonant
    [
      'കൈ',
      [
        [9, 0, 1430, 0, 0],
        [3, 0, 1038, 0, 0],
      ],
    ],
  ],
};

const face = (script: string) =>
  parseTtf(new Uint8Array(readFileSync(`tests/fixtures/fonts/NotoSans${script}-shaping.ttf`)));

describe('a word of each script, shaped as HarfBuzz shapes it', () => {
  for (const [script, cases] of Object.entries(CASES)) {
    it(script, () => {
      const parsed = face(script);
      for (const [text, want] of cases) {
        const run = shapeIndic(parsed, text);
        const got = run
          ? run.gids.map((g, i) => [
              g,
              run.clusters[i],
              run.advances[i],
              run.xOffsets[i],
              run.yOffsets[i],
            ])
          : [];
        expect(got, text).toEqual(want);
      }
    });
  }
});

describe('what it leaves to the caller', () => {
  it('shapes nothing without a letter of these scripts', () => {
    const devanagari = face('Devanagari');
    expect(shapeIndic(devanagari, 'abc')).toBeUndefined();
    // The dandas and the Vedic accents belong to no one script.
    expect(shapeIndic(devanagari, '।')).toBeUndefined();
    expect(shapeIndic(devanagari, '॑')).toBeUndefined();
    expect(hasIndic('abc ।')).toBe(false);
    expect(hasIndic('a क')).toBe(true);
  });

  it('shapes nothing in a face with no layout for the script', () => {
    const roboto = parseTtf(
      new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
    );
    expect(shapeIndic(roboto, 'हिन्दी')).toBeUndefined();
  });

  it('sets a script the face lacks glyph by glyph, beside one it shapes', () => {
    const devanagari = face('Devanagari');
    const run = shapeIndic(devanagari, 'हिক')!;
    // हि shaped (the i sign first), then the Bengali ka as the face's notdef.
    expect(run.gids).toEqual([105, 17, 0]);
    expect(run.clusters).toEqual([0, 0, 2]);
  });
});
