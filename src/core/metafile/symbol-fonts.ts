// Wingdings, Webdings, Symbol — fonts whose letters are not letters.
//
// A metafile names a typeface and hands over a string; the caller finds a font
// and draws the characters. That breaks for the symbol fonts, where the byte
// `0x6E` is not "n" but a filled circle, and where the substitute font a
// document is rendered with has no such mapping at all — so an embedded
// diagram's bullets came out as rows of "n" and "g" (45541_Footer's feasibility
// table, where every cell is one).
//
// These fonts carry a (3,0) symbol cmap: the codes live at U+F020…U+F0FF and
// are the low byte plus 0xF000. What they DRAW is public knowledge, so the
// handful that documents actually use are translated to the Unicode character
// that means the same thing, which any text font can draw. What is left
// untranslated is drawn as it was, which is no worse than before.

import type { VectorPath } from '@/core/vector';
import { PathBuilder } from '@/core/vector';

/**
 * The geometry a symbol character DRAWS, for the handful that are plain shapes.
 *
 * Translating them to Unicode is not enough on its own: the substitute font a
 * document renders with has no `●` either, and a missing glyph is a box with
 * its code number in it — which is what 45541_Footer's table filled sixty-eight
 * cells with. A shape this simple is better drawn than spelled.
 *
 * Measured off the fonts themselves: both fill the em square, one unit wide,
 * from 0.2 em below the baseline to 0.8 em above it.
 */
export type SymbolGeometry = 'circle' | 'square' | 'diamond';

/** Half a circle's control-point distance for a cubic quarter-arc. */
const KAPPA = 0.5522847498307936;

/**
 * The shapes a symbol string draws, one per character, or `undefined` when any
 * character is not one of them — a string that mixes a pictogram with a circle
 * stays text, which is the safer half of the trade.
 *
 * @param text   The string as stored.
 * @param family The typeface it is drawn with.
 */
export function symbolGeometryOf(
  text: string,
  family: string | undefined,
): Array<SymbolGeometry> | undefined {
  const table = family === undefined ? undefined : GEOMETRY_TABLES.get(normalize(family));
  if (!table || text.length === 0) return undefined;
  const out: Array<SymbolGeometry> = [];
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    const key = code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code;
    const shape = table.get(key);
    if (shape === undefined) return undefined;
    out.push(shape);
  }
  return out;
}

/**
 * The outline of one symbol in a box, as the path segments a picture draws.
 *
 * @param shape  Which symbol.
 * @param left   The box's left edge, in the metafile's logical units.
 * @param top    Its top edge (y grows DOWN, as a metafile has it).
 * @param size   The em, which is both the width and the height of the box.
 * @returns The closed outline.
 */
export function symbolOutline(
  shape: SymbolGeometry,
  left: number,
  top: number,
  size: number,
): VectorPath {
  const b = new PathBuilder();
  const r = size / 2;
  const [cx, cy] = [left + r, top + r];
  if (shape === 'circle') {
    const k = r * KAPPA;
    return b
      .moveTo(cx, top)
      .cubicTo(cx + k, top, left + size, cy - k, left + size, cy)
      .cubicTo(left + size, cy + k, cx + k, top + size, cx, top + size)
      .cubicTo(cx - k, top + size, left, cy + k, left, cy)
      .cubicTo(left, cy - k, cx - k, top, cx, top)
      .close()
      .build();
  }
  if (shape === 'diamond') {
    return b
      .moveTo(cx, top)
      .lineTo(left + size, cy)
      .lineTo(cx, top + size)
      .lineTo(left, cy)
      .close()
      .build();
  }
  return b
    .moveTo(left, top)
    .lineTo(left + size, top)
    .lineTo(left + size, top + size)
    .lineTo(left, top + size)
    .close()
    .build();
}

/** Whether a typeface's characters are symbols rather than letters. */
export function isSymbolFont(family: string | undefined): boolean {
  if (family === undefined) return false;
  return SYMBOL_TABLES.has(normalize(family));
}

/**
 * Translate a string drawn in a symbol font to the Unicode that means the same.
 *
 * @param text   The string as the metafile stored it.
 * @param family The typeface it was drawn with.
 * @returns The translated string; characters with no counterpart are kept.
 */
export function fromSymbolFont(text: string, family: string | undefined): string {
  const table = family === undefined ? undefined : SYMBOL_TABLES.get(normalize(family));
  if (!table) return text;
  let out = '';
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    // A (3,0) cmap puts the same characters at U+F0xx; both spellings arrive.
    const key = code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code;
    out += table.get(key) ?? ch;
  }
  return out;
}

function normalize(family: string): string {
  return family.trim().toLowerCase();
}

// Wingdings — the bullets, boxes and marks Office writes for lists and forms.
const WINGDINGS: ReadonlyMap<number, string> = new Map([
  [0x6c, '●'], // l — black circle
  [0x6d, '❍'], // m — shadowed circle
  [0x6e, '■'], // n — black square
  [0x6f, '□'], // o — white square
  [0x70, '❑'], // p — shadowed square
  [0x71, '❒'], // q — shadowed square
  [0x72, '▪'], // r — small black square
  [0x73, '▪'], // s
  [0x74, '▫'], // t — small white square
  [0x75, '◆'], // u — black diamond
  [0x76, '❖'], // v — diamond minus
  [0x77, '◇'], // w — white diamond
  [0x78, '✗'], // x — ballot X
  [0x9f, '•'], // Ÿ — bullet
  [0xa7, '▪'], // § — small square
  [0xfc, '✔'], // ü — check mark
  [0xfd, '✘'], // ý — heavy ballot X
  [0xfe, '☑'], // þ — boxed check
  [0xff, '☒'], // ÿ — boxed X
]);

// Webdings — a different assignment entirely; `n` is the filled circle the
// corpus deck fills eight rows of cells with.
const WEBDINGS: ReadonlyMap<number, string> = new Map([
  [0x67, '■'], // g — black square
  [0x6c, '■'], // l — black square
  [0x6e, '●'], // n — black circle
  [0x6f, '○'], // o — white circle
  [0x70, '□'], // p — white square
  [0x72, '◆'], // r — black diamond
  [0x73, '▲'], // s — black up triangle
  [0x74, '▼'], // t — black down triangle
  [0xa1, '●'], // ¡
]);

// Symbol — a Greek/mathematical face; the letters ARE Greek.
const SYMBOL: ReadonlyMap<number, string> = new Map([
  [0x61, 'α'],
  [0x62, 'β'],
  [0x63, 'χ'],
  [0x64, 'δ'],
  [0x65, 'ε'],
  [0x66, 'φ'],
  [0x67, 'γ'],
  [0x68, 'η'],
  [0x6c, 'λ'],
  [0x6d, 'μ'],
  [0x70, 'π'],
  [0x72, 'ρ'],
  [0x73, 'σ'],
  [0x74, 'τ'],
  [0x77, 'ω'],
  [0xb7, '•'], // · — bullet
  [0xd7, '×'],
  [0xb1, '±'],
  [0xa3, '≤'],
  [0xb3, '≥'],
  [0xb9, '≠'],
  [0xbb, '≈'],
]);

// What the geometric ones actually draw, measured off the fonts: Webdings `n`
// is one contour of four quadrant arcs filling the em, `g` is four corners.
const WINGDINGS_GEOMETRY: ReadonlyMap<number, SymbolGeometry> = new Map([
  [0x6c, 'circle'],
  [0x6e, 'square'],
  [0x6f, 'square'],
  [0x72, 'square'],
  [0x73, 'square'],
  [0x75, 'diamond'],
  [0x77, 'diamond'],
]);

const WEBDINGS_GEOMETRY: ReadonlyMap<number, SymbolGeometry> = new Map([
  [0x67, 'square'],
  [0x6c, 'square'],
  [0x6e, 'circle'],
]);

const GEOMETRY_TABLES: ReadonlyMap<string, ReadonlyMap<number, SymbolGeometry>> = new Map([
  ['wingdings', WINGDINGS_GEOMETRY],
  ['wingdings 2', WINGDINGS_GEOMETRY],
  ['wingdings 3', WINGDINGS_GEOMETRY],
  ['webdings', WEBDINGS_GEOMETRY],
  ['monotype sorts', WINGDINGS_GEOMETRY],
  ['zapfdingbats', WINGDINGS_GEOMETRY],
]);

const SYMBOL_TABLES: ReadonlyMap<string, ReadonlyMap<number, string>> = new Map([
  ['wingdings', WINGDINGS],
  ['wingdings 2', WINGDINGS],
  ['wingdings 3', WINGDINGS],
  ['webdings', WEBDINGS],
  ['symbol', SYMBOL],
  ['monotype sorts', WINGDINGS],
  ['zapfdingbats', WINGDINGS],
]);
