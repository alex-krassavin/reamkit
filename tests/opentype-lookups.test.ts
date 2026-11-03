// The GSUB and GPOS lookups of a face, run as a shaping engine runs them.
//
// Every expected glyph, cluster and position here is HarfBuzz's own: hb-shape
// 15.0.0, run on the same fixture face — the glyphs with --output-format=json,
// the buffer before and after each feature with --trace. The faces are Noto
// Sans Devanagari, Bengali and Newa, cut down with hb-subset to the characters
// below and keeping their layout tables (see fixtures/fonts/NOTICE.md).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { OtlFeature, OtlGlyph } from '@/core/font/opentype-lookups';
import { parseTtf } from '@/core/font';
import {
  OtlFace,
  pickScript,
  planLookups,
  position,
  substitute,
  wouldSubstitute,
} from '@/core/font/opentype-lookups';

const face = (name: string): OtlFace =>
  OtlFace.of(parseTtf(new Uint8Array(readFileSync(`tests/fixtures/fonts/${name}-subset.ttf`))));
const DEVANAGARI = face('NotoSansDevanagari');
const BENGALI = face('NotoSansBengali');
const NEWA = face('NotoSansNewa');
const KANNADA = face('NotoSansKannada');

const DEV2 = ['dev2', 'deva'];

/** A buffer of glyph ids and clusters, as HarfBuzz's trace prints one. */
const buffer = (f: OtlFace, glyphs: ReadonlyArray<readonly [number, number]>): Array<OtlGlyph> =>
  glyphs.map(([gid, cluster]) => {
    const joiner =
      gid === f.parsed.glyphForCodepoint(0x200d)
        ? 'zwj'
        : gid === f.parsed.glyphForCodepoint(0x200c)
          ? 'zwnj'
          : undefined;
    return {
      gid,
      cluster,
      mask: 1,
      cls: f.glyphClass(gid),
      syllable: 0,
      ligId: 0,
      ligComp: 0,
      ...(joiner ? { joiner } : {}),
    };
  });

const indic = (tag: string): OtlFeature => ({
  tag,
  mask: 1,
  manualZwj: true,
  manualZwnj: true,
  perSyllable: true,
});

/** Run one GSUB feature of the Devanagari face over a buffer; the glyphs and clusters after. */
function runFeature(
  features: ReadonlyArray<OtlFeature>,
  glyphs: ReadonlyArray<readonly [number, number]>,
): Array<[number, number]> {
  const buf = buffer(DEVANAGARI, glyphs);
  substitute(DEVANAGARI, planLookups(DEVANAGARI.gsub, DEV2, features), buf);
  return buf.map((g) => [g.gid, g.cluster]);
}

describe('GSUB, feature by feature', () => {
  it('forms a conjunct with a ligature (akhn)', () => {
    // क्ष: ka, virama, ssa become one glyph; the vowel sign after keeps its own cluster.
    expect(
      runFeature(
        [indic('akhn')],
        [
          [7, 0],
          [25, 0],
          [19, 2],
          [13, 3],
          [2, 3],
        ],
      ),
    ).toEqual([
      [21, 0],
      [13, 3],
      [2, 3],
    ]);
  });

  it('merges the whole cluster of a ligature’s last letter into it, as HarfBuzz does', () => {
    // ज्ञ: the aa that was one cluster with ña is now one with the conjunct.
    expect(
      runFeature(
        [indic('akhn')],
        [
          [8, 0],
          [25, 0],
          [9, 2],
          [2, 2],
          [12, 4],
        ],
      ),
    ).toEqual([
      [22, 0],
      [2, 0],
      [12, 4],
    ]);
    // श्री under rkrf: sha with the ra under it, and the long ii with them.
    expect(
      runFeature(
        [indic('rkrf')],
        [
          [18, 0],
          [25, 0],
          [15, 2],
          [4, 2],
        ],
      ),
    ).toEqual([
      [65, 0],
      [4, 0],
    ]);
  });

  it('leaves a conjunct unformed where a join control stands in it', () => {
    // क्‍ष and क्‌ष: the font's own rules decide what a ZWJ or a ZWNJ does
    // between the virama and ssa, and akhn has none for them.
    const zwj = [
      [7, 0],
      [25, 0],
      [278, 0],
      [19, 3],
    ] as const;
    expect(runFeature([indic('akhn')], zwj)).toEqual(zwj);
    const zwnj = [
      [7, 0],
      [25, 0],
      [277, 2],
      [19, 3],
    ] as const;
    expect(runFeature([indic('akhn')], zwnj)).toEqual(zwnj);
  });

  it('joins a nukta to its consonant (nukt)', () => {
    expect(
      runFeature(
        [indic('nukt')],
        [
          [7, 0],
          [26, 0],
          [5, 0],
          [24, 0],
        ],
      ),
    ).toEqual([
      [27, 0],
      [5, 0],
      [24, 0],
    ]);
  });

  it('runs a chained context whose lookup changes the length of the buffer (cjct)', () => {
    // विद्यालय: द्य is formed inside a context rule, and the places after it move.
    expect(
      runFeature(
        [indic('cjct')],
        [
          [3, 0],
          [17, 0],
          [11, 2],
          [25, 2],
          [14, 4],
          [2, 4],
          [16, 6],
          [14, 7],
        ],
      ),
    ).toEqual([
      [3, 0],
      [17, 0],
      [110, 2],
      [2, 2],
      [16, 6],
      [14, 7],
    ]);
    expect(
      runFeature(
        [indic('cjct')],
        [
          [15, 0],
          [2, 0],
          [41, 2],
          [10, 4],
          [150, 4],
        ],
      ),
    ).toEqual([
      [15, 0],
      [2, 0],
      [120, 2],
      [150, 2],
    ]);
    expect(
      runFeature(
        [indic('cjct')],
        [
          [11, 0],
          [25, 0],
          [17, 2],
          [2, 2],
          [15, 4],
        ],
      ),
    ).toEqual([
      [112, 0],
      [2, 0],
      [15, 4],
    ]);
  });

  it('runs the presentation stage as one, every feature’s lookups in LookupList order', () => {
    const stage = [
      ...['pres', 'abvs', 'blws', 'psts', 'haln'].map(indic),
      ...['calt', 'clig', 'liga', 'rclt', 'rlig'].map((tag) => ({ tag, mask: 1 })),
    ];
    // हिन्दी: the short i before ha takes the width of what it stands over.
    expect(
      runFeature(stage, [
        [3, 0],
        [20, 0],
        [35, 2],
        [11, 4],
        [4, 4],
      ]),
    ).toEqual([
      [155, 0],
      [20, 0],
      [35, 2],
      [11, 4],
      [4, 4],
    ]);
    expect(
      runFeature(stage, [
        [3, 0],
        [17, 0],
        [110, 2],
        [2, 2],
        [16, 6],
        [14, 7],
      ]),
    ).toEqual([
      [156, 0],
      [17, 0],
      [110, 2],
      [2, 2],
      [16, 6],
      [14, 7],
    ]);
    // क़ुँ: the u sign under a nukta takes its lower form.
    expect(
      runFeature(stage, [
        [27, 0],
        [5, 0],
        [24, 0],
      ]),
    ).toEqual([
      [27, 0],
      [263, 0],
      [24, 0],
    ]);
  });
});

const GPOS_FEATURES = ['abvm', 'blwm', 'mark', 'mkmk', 'curs', 'dist', 'kern'].map((tag) => ({
  tag,
  mask: 1,
}));

/** HarfBuzz's glyphs positioned by our GPOS: [advance, x offset, y offset] each. */
function positioned(
  f: OtlFace,
  scripts: ReadonlyArray<string>,
  glyphs: ReadonlyArray<readonly [number, number]>,
  gpos = true,
): Array<[number, number, number]> {
  const steps = gpos ? planLookups(f.gpos, scripts, GPOS_FEATURES) : [];
  return position(f, steps, buffer(f, glyphs)).map((p) => [p.xAdvance, p.xOffset, p.yOffset]);
}

describe('GPOS', () => {
  it('sets a mark on its base by their anchors (MarkBasePos)', () => {
    // हूँ: the candrabindu over ha, past the uu below it.
    const hu = [
      [143, 0],
      [24, 0],
    ] as const;
    expect(positioned(DEVANAGARI, DEV2, hu)).toEqual([
      [531, 0, 0],
      [0, 13, 0],
    ]);
    // कुं and क़ुँ: the u sign and the mark over it, both on the consonant.
    expect(
      positioned(DEVANAGARI, DEV2, [
        [7, 0],
        [5, 0],
        [23, 0],
      ]),
    ).toEqual([
      [768, 0, 0],
      [0, -221, 0],
      [0, -221, 0],
    ]);
    expect(
      positioned(DEVANAGARI, DEV2, [
        [27, 0],
        [263, 0],
        [24, 0],
      ]),
    ).toEqual([
      [768, 0, 0],
      [0, -221, 0],
      [0, -208, 0],
    ]);
    // …which, unpositioned, stand where the pen left them.
    expect(positioned(DEVANAGARI, DEV2, hu, false)).not.toEqual(positioned(DEVANAGARI, DEV2, hu));
  });

  it('kerns a pair by class (PairPos)', () => {
    expect(
      positioned(DEVANAGARI, DEV2, [
        [155, 0],
        [20, 0],
        [35, 2],
        [11, 4],
        [4, 4],
      ]),
    ).toEqual([
      [259, 0, 0],
      [531, 0, 0],
      [309, 0, 0],
      [531, 0, 0],
      [259, 0, 0],
    ]);
    expect(
      positioned(DEVANAGARI, DEV2, [
        [15, 0],
        [2, 0],
        [120, 2],
        [150, 2],
      ]),
    ).toEqual([
      [409, 0, 0],
      [259, 0, 0],
      [584, 0, 0],
      [0, -74, 0],
    ]);
  });

  it('stacks a mark on a mark, and places by context (MarkMarkPos, ChainContextPos)', () => {
    const BNG2 = ['bng2', 'beng'];
    // সংস্কৃতি: a conjunct of three with the vocalic r under it.
    expect(
      positioned(BENGALI, BNG2, [
        [13, 0],
        [4, 0],
        [48, 2],
        [69, 4],
        [75, 4],
        [79, 4],
        [18, 4],
        [17, 6],
        [7, 6],
      ]),
    ).toEqual([
      [623, 0, 0],
      [438, 0, 0],
      [745, 0, 0],
      [0, -580, 125],
      [0, -263, 98],
      [0, -250, 0],
      [0, -295, -5],
      [266, 0, 0],
      [707, 0, 0],
    ]);
    expect(
      positioned(BENGALI, BNG2, [
        [14, 0],
        [76, 0],
        [78, 0],
        [8, 2],
        [10, 3],
        [15, 3],
      ]),
    ).toEqual([
      [674, 0, 0],
      [0, -275, 165],
      [0, -200, 0],
      [553, 0, 0],
      [626, 0, 0],
      [0, -396, 108],
    ]);
    expect(
      positioned(BENGALI, BNG2, [
        [24, 0],
        [74, 2],
        [78, 2],
        [18, 2],
      ]),
    ).toEqual([
      [910, 0, 0],
      [0, -365, 23],
      [0, -200, 0],
      [0, -445, 1],
    ]);
  });

  it('joins glyphs cursively, the exit of one on the entry of the next (CursivePos)', () => {
    const NEWA_TAGS = ['newa'];
    expect(
      positioned(NEWA, NEWA_TAGS, [
        [7, 0],
        [24, 0],
        [15, 0],
      ]),
    ).toEqual([
      [585, 0, 0],
      [395, -375, 0],
      [0, -139, 0],
    ]);
    expect(
      positioned(NEWA, NEWA_TAGS, [
        [7, 0],
        [3, 0],
      ]),
    ).toEqual([
      [585, 0, 0],
      [318, -333, 0],
    ]);
    expect(
      positioned(NEWA, NEWA_TAGS, [
        [26, 0],
        [12, 0],
        [11, 0],
        [38, 0],
      ]),
    ).toEqual([
      [455, 0, 0],
      [448, -223, 0],
      [0, 55, 0],
      [0, -80, 0],
    ]);
    // …and a joined glyph rises with the one it hangs on.
    expect(
      positioned(NEWA, NEWA_TAGS, [
        [16, 0],
        [21, 0],
        [8, 0],
      ]),
    ).toEqual([
      [386, 0, 0],
      [227, 98, 0],
      [442, 231, 2],
    ]);
  });

  it('lets the second glyph of a pair start the next pair when it has no value of its own', () => {
    // ಙೈ: the two parts of the ai sign follow the nga, and both pairs kern.
    expect(
      positioned(
        KANNADA,
        ['knd2', 'knda'],
        [
          [4, 0],
          [60, 0],
          [16, 0],
        ],
      ),
    ).toEqual([
      [736, 0, 0],
      [257, 0, 0],
      [183, 0, 0],
    ]);
  });
});

describe('the lists', () => {
  const gid = (cp: number): number => DEVANAGARI.parsed.glyphForCodepoint(cp);
  const RA = gid(0x0930);
  const KA = gid(0x0915);
  const VIRAMA = gid(0x094d);

  it('says whether a feature would substitute a sequence, as a shaper asks', () => {
    const rphf = planLookups(DEVANAGARI.gsub, DEV2, [indic('rphf')]);
    const half = planLookups(DEVANAGARI.gsub, DEV2, [indic('half')]);
    expect(wouldSubstitute(DEVANAGARI, rphf, [RA, VIRAMA], true)).toBe(true);
    expect(wouldSubstitute(DEVANAGARI, rphf, [KA, VIRAMA], true)).toBe(false);
    expect(wouldSubstitute(DEVANAGARI, half, [KA, VIRAMA], true)).toBe(true);
    expect(wouldSubstitute(DEVANAGARI, half, [KA], true)).toBe(false);
  });

  it('picks the first script the table has, else its default', () => {
    expect(pickScript(DEVANAGARI.gsub, ['knd2', 'dev2', 'deva'])).toBe('dev2');
    expect(pickScript(DEVANAGARI.gsub, ['zzzz'])).toBe('DFLT');
    expect(pickScript(undefined, DEV2)).toBeUndefined();
  });

  it('runs a lookup two features name once, in LookupList order, for both', () => {
    const one = planLookups(DEVANAGARI.gsub, DEV2, [indic('akhn')]);
    const both = planLookups(DEVANAGARI.gsub, DEV2, [
      { tag: 'akhn', mask: 1 },
      { tag: 'akhn', mask: 2 },
    ]);
    expect(both.map((s) => s.lookup)).toEqual(one.map((s) => s.lookup));
    expect(both.every((s) => s.mask === 3)).toBe(true);
    const many = planLookups(DEVANAGARI.gsub, DEV2, ['half', 'akhn', 'nukt'].map(indic));
    expect(many.map((s) => s.lookup)).toEqual([...many.map((s) => s.lookup)].sort((a, b) => a - b));
  });
});
