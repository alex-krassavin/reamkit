// §9.10.2 — the punctuation a glyph IS, read off the shape the page draws.
//
// A font that states no character for a glyph (a `/ToUnicode` entry of U+0000,
// a subset with no `cmap`) leaves nothing to write, and the page still shows
// the mark. Stripe writes every piece of punctuation it sets that way: its
// invoices read "6VOBWUGP-0010", "Kazakhstan VAT: 86-1696045" and
// "Aug 11–Sep 11, 2026" and have no hyphen, colon or dash in their text at all.
//
// Letters are beyond a shape test; a handful of punctuation marks are not. A
// dash is one flat bar at the height of the lower-case letters' middle, and its
// length says which dash; a colon is two dots stacked on the baseline; a
// bracket is one curve reaching below the baseline and over the capitals. Only
// the marks whose shape says one thing are read — anything else stays unread.

import type { PathSeg } from './content';

/** A glyph's contours, and the box they fill, in the page's y-up space. */
export interface GlyphShape {
  readonly segs: ReadonlyArray<PathSeg>;
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
}

/**
 * The punctuation mark a glyph's shape is, where it can be only one.
 *
 * @param glyph    The glyph as the page draws it.
 * @param baseline The y of the line it stands on.
 * @param size     The size the line is set at, which the shape is measured in.
 * @returns The character, or `undefined` where the shape does not say.
 */
export function punctuationOf(
  glyph: GlyphShape,
  baseline: number,
  size: number,
): string | undefined {
  if (!(size > 0)) return undefined;
  const width = (glyph.maxX - glyph.minX) / size;
  const height = (glyph.maxY - glyph.minY) / size;
  const bottom = (glyph.minY - baseline) / size;
  const top = (glyph.maxY - baseline) / size;
  const contours = glyph.segs.filter((s) => s.op === 'move').length;

  // A dash: one flat bar, standing clear of the baseline at the middle of the
  // lower case. How long it is says which: a hyphen is a third of an em, an en
  // dash half of one, an em dash the whole.
  if (contours === 1 && height <= FLAT_EM && width >= height * 2.5) {
    if (bottom < 0.15 || top > 0.6) return undefined;
    if (width < HYPHEN_LONGEST_EM) return '-';
    return width < EN_DASH_LONGEST_EM ? '–' : '—';
  }
  // A colon: two dots, one over the other, the lower on (or just over) the
  // baseline; a semicolon's lower one hangs below it.
  if (contours === 2 && width <= DOT_WIDEST_EM && height >= 0.3 && height <= 0.8) {
    if (bottom > 0.15) return undefined;
    return bottom < -0.05 ? ';' : ':';
  }
  // A full stop and a comma: one small mark at the baseline, the comma's tail
  // reaching below it.
  if (contours === 1 && width <= DOT_WIDEST_EM && height <= 0.4 && top <= 0.3) {
    if (bottom >= -0.05 && bottom <= 0.05 && height <= DOT_WIDEST_EM) return '.';
    return bottom < -0.05 ? ',' : undefined;
  }
  // A parenthesis: one narrow curve from below the baseline to over the
  // capitals. Its ENDS turn back toward what it encloses, so the end at the
  // top stands on the side the text is on: right of the middle for "(".
  if (contours === 1 && width <= 0.45 && height >= 0.75 && bottom <= -0.05 && top >= 0.6) {
    const peak = highestPoint(glyph.segs);
    if (peak === undefined) return undefined;
    const middle = (glyph.minX + glyph.maxX) / 2;
    return peak > middle ? '(' : ')';
  }
  return undefined;
}

/** The x of the highest point a glyph's contours pass through. */
function highestPoint(segs: ReadonlyArray<PathSeg>): number | undefined {
  let best: { x: number; y: number } | undefined;
  for (const s of segs) {
    if (s.op === 'close') continue;
    if (best === undefined || s.y > best.y) best = { x: s.x, y: s.y };
  }
  return best?.x;
}

/** The tallest a bar may be, in ems, and still be a dash rather than a block. */
const FLAT_EM = 0.15;

/** The longest a hyphen is, in ems; past it the bar is a dash. */
const HYPHEN_LONGEST_EM = 0.42;

/** The longest an en dash is, in ems; past it, an em dash. */
const EN_DASH_LONGEST_EM = 0.7;

/** The widest a dot is, in ems. */
const DOT_WIDEST_EM = 0.22;
