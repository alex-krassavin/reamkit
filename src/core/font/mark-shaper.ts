// Combining marks on Latin, Greek and Cyrillic letters, placed by the face.
//
// A letter and a combining mark written apart — и and U+0306 for й, and the
// stress mark U+0301 a Russian textbook or a Chuvash dictionary puts on a vowel
// that has no composed form — were set as two glyphs side by side: the mark at
// its own default place, which a face designs for some average letter, and
// nothing moved it onto THIS one. Word draws й written that way exactly as it
// draws the composed letter, and sets a stressed а's acute on the а.
//
// The layout composes what has a composed letter (see `drawableText`); this is
// for what is left — a mark with no composed form, or a face without one. The
// face's own GPOS places it: `mark` sets a mark on its base by their anchors,
// `mkmk` a mark on a mark, `kern` and the ligatures run as they do for any
// other text. HarfBuzz, which LibreOffice and the browsers shape with, gives the
// same glyphs the same offsets: Arimo's acute over а at (-284, -340).

import type { IndicRun } from '@/core/font/indic-shaper';
import type { OtlFeature, OtlGlyph } from '@/core/font/opentype-lookups';
import type { ParsedTtf } from '@/core/font/ttf-parser';
import { OtlFace, planLookups, position, substitute } from '@/core/font/opentype-lookups';

/**
 * The combining marks a Latin, Greek or Cyrillic letter takes: the Combining
 * Diacritical Marks and their Extended and Supplement blocks, the Cyrillic
 * combining marks, the marks for symbols and the half marks.
 */
const MARK_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x0300, 0x036f],
  [0x0483, 0x0489],
  [0x1ab0, 0x1aff],
  [0x1dc0, 0x1dff],
  [0x20d0, 0x20ff],
  [0xfe20, 0xfe2f],
];

/** OpenType GDEF GlyphClassDef — the class a mark glyph is given. */
const MARK_CLASS = 3;

/**
 * Whether `text` holds a combining mark of the blocks Latin, Greek and
 * Cyrillic letters take — the text {@link shapeMarks} is for.
 *
 * @param text The text.
 */
export function hasCombiningMark(text: string): boolean {
  for (let i = 0; i < text.length; i++) if (isCombiningMark(text.charCodeAt(i))) return true;
  return false;
}

/**
 * Whether a code point is one of those marks.
 *
 * @param cp The code point.
 */
export function isCombiningMark(cp: number): boolean {
  if (cp < 0x0300) return false;
  for (const [lo, hi] of MARK_RANGES) if (cp >= lo && cp <= hi) return true;
  return false;
}

/** The script tags a face's layout is looked up by, for the letters of `cps`. */
function scriptTags(cps: ReadonlyArray<number>): ReadonlyArray<string> {
  for (const cp of cps) {
    if (cp >= 0x0400 && cp <= 0x052f) return ['cyrl'];
    if ((cp >= 0x0370 && cp <= 0x03ff) || (cp >= 0x1f00 && cp <= 0x1fff)) return ['grek'];
    if ((cp >= 0x41 && cp <= 0x24f) || (cp >= 0x1e00 && cp <= 0x1eff)) return ['latn'];
  }
  return ['latn'];
}

const feature = (tag: string): OtlFeature => ({ tag, mask: 1 });

/**
 * Shape a run whose letters carry combining marks, in one face: the face's
 * substitutions, then its positioning, the marks set on their letters.
 *
 * @param parsed The face.
 * @param text   The text.
 * @param kern   Whether to apply the face's `kern` feature (the layout profile decides).
 * @returns The glyphs and where each stands, or `undefined` when the text holds
 *          no such mark or the face has no positioning to set one with — the
 *          caller's own shaping applies then.
 */
export function shapeMarks(parsed: ParsedTtf, text: string, kern = true): IndicRun | undefined {
  if (!hasCombiningMark(text)) return undefined;
  const face = OtlFace.of(parsed);
  if (!face.gpos) return undefined;
  const cps = [...text].map((ch) => ch.codePointAt(0)!);
  const tags = scriptTags(cps);
  // A mark is one cluster with the letter it follows: a reader takes them
  // together, and letter spacing does not come between them.
  let base = 0;
  const buf: Array<OtlGlyph> = cps.map((cp, i) => {
    const mark = i > 0 && isCombiningMark(cp);
    if (!mark) base = i;
    const gid = parsed.glyphForCodepoint(cp);
    // A face without GDEF classes still has marks: the character says so.
    const cls = face.glyphClass(gid) || (mark ? MARK_CLASS : 1);
    return { gid, cluster: base, mask: 1, cls, syllable: 0, ligId: 0, ligComp: 0, sources: [i] };
  });
  substitute(face, planLookups(face.gsub, tags, ['ccmp', 'liga', 'clig'].map(feature)), buf);
  const gpos = planLookups(
    face.gpos,
    tags,
    [...(kern ? ['kern'] : []), 'mark', 'mkmk'].map(feature),
  );
  const pos = position(face, gpos, buf);
  const run: IndicRun = {
    gids: [],
    clusters: [],
    sources: [],
    advances: [],
    xOffsets: [],
    yOffsets: [],
  };
  buf.forEach((g, i) => {
    const p = pos[i]!;
    run.gids.push(g.gid);
    run.clusters.push(g.cluster);
    run.sources.push(g.sources ?? [g.cluster]);
    // A mark takes no room on the line, whatever its own advance says: it is
    // drawn over its letter, and the next letter stands where it would have
    // stood without it (HarfBuzz zeroes a GDEF mark's advance the same way).
    run.advances.push(g.cls === MARK_CLASS ? 0 : p.xAdvance);
    run.xOffsets.push(p.xOffset);
    run.yOffsets.push(p.yOffset);
  });
  return run;
}
