// §20.1.4.2.24 — the table style a slide table wears.
//
// A slide table states almost nothing about how it looks: its `a:tblPr` names a
// style by GUID and switches on the PARTS of it that apply — a header row,
// banded rows, a first column — and the style itself lives in
// `ppt/tableStyles.xml` as a set of conditional formats. Read without it a
// table is bare text: table_test2's blue header, its banding and every one of
// its rules are the style's (and so are two conference decks in the corpus).
//
// Each part carries `a:tcStyle` (a fill and the four borders) and `a:tcTxStyle`
// (bold, italic and a colour). They compose in the order the spec lists them,
// the nearest last, and the cell's own `a:tcPr` beats all of them.

import type { Border, CellBorders, TableCell } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';

import { pt } from '@/core/ir';
import { placeholderColors, readColorMods, resolveColorNode } from '@/core/drawingml/colors';
import { poAttr, poChildren, poIntAttr, poIs, poTag, poText } from '@/core/po-helpers';

/** Which conditional parts of a style a table asks for (`a:tblPr` flags). */
export interface TableStyleFlags {
  readonly firstRow: boolean;
  readonly lastRow: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
  readonly bandRow: boolean;
  readonly bandCol: boolean;
}

/** What one part of a table style says about a cell. */
export interface TableStylePart {
  readonly shadingHex?: string;
  /**
   * §20.1.2.3.1 — how opaque that shading is, when the part asks for less than
   * all of it. A cell's fill is a layer over the table's background, not over
   * the page, so the two have to be composed rather than each flattened to
   * white on its own.
   */
  readonly shadingAlpha?: number;
  readonly borders?: CellBorders;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly colorHex?: string;
}

/** Where a cell sits, which decides the parts that reach it. */
export interface CellPosition {
  readonly row: number;
  readonly rowCount: number;
  readonly col: number;
  readonly colCount: number;
}

/** The `a:tblPr` flags, all off unless the table says otherwise. */
export function tableStyleFlags(tblPr: PoNode | undefined): TableStyleFlags {
  const on = (name: string): boolean =>
    tblPr !== undefined && (poAttr(tblPr, name) === '1' || poAttr(tblPr, name) === 'true');
  return {
    firstRow: on('firstRow'),
    lastRow: on('lastRow'),
    firstCol: on('firstCol'),
    lastCol: on('lastCol'),
    bandRow: on('bandRow'),
    bandCol: on('bandCol'),
  };
}

/** The GUID a table names, if it names one (`a:tblPr/a:tableStyleId`). */
export function tableStyleId(tblPr: PoNode | undefined): string | undefined {
  const id = tblPr ? poChildren(tblPr).find((c) => poIs(c, 'a:tableStyleId')) : undefined;
  return id ? poText(id).trim() || undefined : undefined;
}

/** What the deck's theme lends a style that points at it instead of spelling it out. */
export interface TableStyleTheme {
  /** §20.1.4.1.14 `a:fillStyleLst` — the fills an `a:fillRef` indexes. */
  readonly fills?: ReadonlyArray<PoNode>;
  /** §20.1.4.2.19 `a:lnStyleLst` — the widths an `a:lnRef` indexes, in points. */
  readonly lineWidths?: ReadonlyArray<number>;
}

/**
 * The style a cell wears, composed from every part that reaches it.
 *
 * @param style  The `a:tblStyle` node.
 * @param flags  Which conditional parts the table asks for.
 * @param at     Where the cell sits.
 * @param colors The deck's colour resolver.
 * @param theme  The theme's style lists, for the parts that point at them.
 * @returns The composed part, empty when the style says nothing.
 */
export function cellStyle(
  style: PoNode,
  flags: TableStyleFlags,
  at: CellPosition,
  colors: ColorResolver,
  theme?: TableStyleTheme,
): TableStylePart {
  // §20.1.4.2.24 lists them in this order and the later ones win: the whole
  // table, then the banding, then the edge rows, then the edge columns.
  const names: Array<string> = ['a:wholeTbl'];
  if (flags.bandCol) names.push(bandName(at.col, flags.firstCol, 'V'));
  if (flags.bandRow) names.push(bandName(at.row, flags.firstRow, 'H'));
  if (flags.firstCol && at.col === 0) names.push('a:firstCol');
  if (flags.lastCol && at.col === at.colCount - 1) names.push('a:lastCol');
  if (flags.firstRow && at.row === 0) names.push('a:firstRow');
  if (flags.lastRow && at.row === at.rowCount - 1) names.push('a:lastRow');

  // §20.1.4.2.25 `a:tblBg` — the fill the whole TABLE stands on, under every
  // conditional part. The cells tile the table, so it composes as the bottom
  // layer of each one rather than as a shape of its own: bnc480256's style
  // paints its band at 40 % over this, and without it every second row came out
  // white where LibreOffice draws blue.
  let out: TableStylePart = tableBackground(style, colors, theme);
  for (const name of names) {
    const part = poChildren(style).find((c) => poIs(c, name));
    if (!part) continue;
    const next = partStyle(part, colors, theme);
    out = { ...out, ...next, ...overFill(out.shadingHex, next) };
  }
  const { shadingAlpha: _drop, ...flat } = out;
  return flat;
}

/**
 * A part's fill laid OVER what is already under it.
 *
 * A colour resolves flattened against the page, because that is what a shape's
 * fill sits on; a table cell's sits on the table's background. Undoing the one
 * and redoing the other needs only the alpha: bnc480256's banding is `accent1`
 * at 40% over a background the theme's gradient makes pale, and taken against
 * white instead it came out as the flat accent with every second row white.
 */
function overFill(under: string | undefined, part: TableStylePart): TableStylePart {
  const a = part.shadingAlpha;
  if (a === undefined || a >= 1 || part.shadingHex === undefined || under === undefined) return {};
  const flat = rgb(part.shadingHex);
  const back = rgb(under);
  if (!flat || !back) return {};
  // `flat = c·a + 1·(1 − a)`, so over `back` it is `flat − (1 − a)·(1 − back)`.
  const mix = flat.map((v, i) => v - (1 - a) * (1 - (back[i] ?? 1)));
  return { shadingHex: toHex(mix) };
}

// §20.1.2.3.1 — the opacity a fill's colour asks for, 1 when it asks for none.
function alphaOf(fill: PoNode): number {
  const walk = (node: PoNode): number | undefined => {
    for (const child of poChildren(node)) {
      if (poIs(child, 'a:alpha')) {
        const v = poIntAttr(child, 'val');
        if (v !== undefined) return Math.max(0, Math.min(1, v / 100000));
      }
      const deeper = walk(child);
      if (deeper !== undefined) return deeper;
    }
    return undefined;
  };
  return walk(fill) ?? 1;
}

function rgb(h: string): Array<number> | undefined {
  const m = /^([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})$/u.exec(h);
  return m ? [1, 2, 3].map((i) => parseInt(m[i] ?? '0', 16) / 255) : undefined;
}

function toHex(v: ReadonlyArray<number>): string {
  return v
    .map((c) =>
      Math.round(Math.max(0, Math.min(1, c)) * 255)
        .toString(16)
        .padStart(2, '0')
        .toUpperCase(),
    )
    .join('');
}

// The table's background fill, as a cell shading. `a:fillRef` points into the
// theme's fill style list; a fill written out in place is read as it stands.
function tableBackground(
  style: PoNode,
  colors: ColorResolver,
  theme?: TableStyleTheme,
): TableStylePart {
  const bg = poChildren(style).find((c) => poIs(c, 'a:tblBg'));
  if (!bg) return {};
  const solid = poChildren(bg).find((c) => poIs(c, 'a:solidFill'));
  const own = solid ? colorOf(solid, colors) : undefined;
  const hex =
    own ??
    refFillColor(
      poChildren(bg).find((c) => poIs(c, 'a:fillRef')),
      colors,
      theme,
    );
  return hex ? { shadingHex: hex } : {};
}

/**
 * §20.1.4.1.16 `a:fillRef` — a fill named by its place in the theme's list,
 * with the colour to put wherever that fill says `phClr`.
 *
 * Only its COLOUR is taken: a slot may be a gradient or a picture, and a cell's
 * shading is one colour. The first solid colour the slot names, or the
 * reference's own, is the nearest true answer.
 */
function refFillColor(
  ref: PoNode | undefined,
  colors: ColorResolver,
  theme?: TableStyleTheme,
): string | undefined {
  if (!ref) return undefined;
  const own = colorOf(ref, colors);
  const idx = poIntAttr(ref, 'idx');
  const slot = idx !== undefined && idx > 0 ? theme?.fills?.[idx - 1] : undefined;
  if (!slot) return own;
  // A slot writes its colours as `phClr`, the placeholder the reference fills
  // in, and it carries its OWN transforms over that: the Office theme's second
  // and third slots are gradients of `tint`/`shade` and `satMod`, and read as
  // the bare reference colour a table background came out the flat
  // accent where both references draw a wash of it.
  const named = own === undefined ? colors : placeholderColors(colors, own);
  // The slot IS the fill, not a wrapper round one.
  const solid = poIs(slot, 'a:solidFill') ? slot : undefined;
  if (solid) return colorOf(solid, named) ?? own;
  // A gradient is many colours and a cell's shading is one. The last stop is
  // the one these slots build their body from — the first is the highlight at
  // the very edge — and it is what LibreOffice's own rendering of bnc480256
  // matches to the byte.
  const grad = poIs(slot, 'a:gradFill') ? slot : undefined;
  const list = grad ? poChildren(grad).find((c) => poIs(c, 'a:gsLst')) : undefined;
  const stops = list ? poChildren(list).filter((c) => poIs(c, 'a:gs')) : [];
  const last = stops[stops.length - 1];
  return (last ? colorOf(last, named) : undefined) ?? own;
}

// Bands count from the first row/column that is NOT an edge one, so a table
// with a header starts its first band on the row below it.
function bandName(index: number, edge: boolean, axis: 'H' | 'V'): string {
  const i = edge ? index - 1 : index;
  return i % 2 === 0 ? `a:band1${axis}` : `a:band2${axis}`;
}

function partStyle(part: PoNode, colors: ColorResolver, theme?: TableStyleTheme): TableStylePart {
  const tcStyle = poChildren(part).find((c) => poIs(c, 'a:tcStyle'));
  const txStyle = poChildren(part).find((c) => poIs(c, 'a:tcTxStyle'));
  const fill = tcStyle ? poChildren(tcStyle).find((c) => poIs(c, 'a:fill')) : undefined;
  const solid = fill ? poChildren(fill).find((c) => poIs(c, 'a:solidFill')) : undefined;
  const shadingHex = solid ? colorOf(solid, colors) : undefined;
  const shadingAlpha = solid ? alphaOf(solid) : undefined;
  const bdr = tcStyle ? poChildren(tcStyle).find((c) => poIs(c, 'a:tcBdr')) : undefined;
  const borders = bdr ? partBorders(bdr, colors, theme) : undefined;
  const on = (name: string): boolean => poAttr(txStyle, name) === 'on';
  const colorHex = txStyle ? colorOf(txStyle, colors) : undefined;
  return {
    ...(shadingHex ? { shadingHex } : {}),
    ...(shadingAlpha !== undefined && shadingAlpha < 1 ? { shadingAlpha } : {}),
    ...(borders ? { borders } : {}),
    ...(txStyle && on('b') ? { bold: true } : {}),
    ...(txStyle && on('i') ? { italic: true } : {}),
    ...(colorHex ? { colorHex } : {}),
  };
}

const SIDES: ReadonlyArray<readonly [string, keyof CellBorders]> = [
  ['a:left', 'left'],
  ['a:right', 'right'],
  ['a:top', 'top'],
  ['a:bottom', 'bottom'],
  ['a:insideH', 'insideH'],
  ['a:insideV', 'insideV'],
];

function partBorders(
  bdr: PoNode,
  colors: ColorResolver,
  theme?: TableStyleTheme,
): CellBorders | undefined {
  const out: { -readonly [K in keyof CellBorders]?: Border } = {};
  for (const [tag, side] of SIDES) {
    const holder = poChildren(bdr).find((c) => poIs(c, tag));
    const ln = holder ? poChildren(holder).find((c) => poIs(c, 'a:ln')) : undefined;
    if (ln) {
      out[side] = lineBorder(ln, colors);
      continue;
    }
    // §20.1.4.2.19 — a side may point at the theme's line style list instead of
    // spelling the rule out: the reference carries the colour, the slot the
    // width. Every rule of bnc480256's table is written this way, and read as
    // "no line" the table came out with none at all.
    const border = refBorder(
      holder ? poChildren(holder).find((c) => poIs(c, 'a:lnRef')) : undefined,
      colors,
      theme,
    );
    if (border) out[side] = border;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function refBorder(
  ref: PoNode | undefined,
  colors: ColorResolver,
  theme?: TableStyleTheme,
): Border | undefined {
  if (!ref || poAttr(ref, 'idx') === '0') return undefined;
  const colorHex = colorOf(ref, colors);
  if (colorHex === undefined) return undefined;
  const idx = poIntAttr(ref, 'idx');
  const width = idx !== undefined && idx > 0 ? theme?.lineWidths?.[idx - 1] : undefined;
  return { style: 'single', colorHex, width: pt(width ?? 0.75) };
}

/**
 * §20.1.2.1 `a:ln` → the rule it draws on one side of a cell.
 *
 * Shared with the cell's OWN `a:lnL`/`a:lnR`/`a:lnT`/`a:lnB`, which are the
 * same element under a different name.
 *
 * @param ln     The line node.
 * @param colors The deck's colour resolver.
 * @returns The border, `none` when the line is one that draws nothing.
 */
export function lineBorder(ln: PoNode, colors: ColorResolver): Border {
  // An `a:noFill` outline is a border that is REMOVED, not one that is drawn —
  // and so is a colour at zero alpha: tdf164936's left rule is `2670C9` with
  // `<a:alpha val="0"/>`, and the cell has three sides, not four.
  if (poChildren(ln).some((c) => poIs(c, 'a:noFill')) || isTransparent(ln)) {
    return { style: 'none' };
  }
  const colorHex = colorOf(ln, colors);
  const w = poIntAttr(ln, 'w');
  return {
    style: 'single',
    ...(colorHex ? { colorHex } : {}),
    // §20.1.2.1 — `@w` is EMU, and a border's width is points.
    ...(w !== undefined ? { width: pt(Math.max(0.25, w / 12700)) } : {}),
  };
}

function isTransparent(ln: PoNode): boolean {
  const fill = poChildren(ln).find((c) => poIs(c, 'a:solidFill'));
  const color = fill ? poChildren(fill).find((c) => poTag(c) !== undefined) : undefined;
  const alpha = color ? readColorMods(color).find((m) => m.kind === 'alpha') : undefined;
  return alpha !== undefined && alpha.val <= 0.001;
}

/** The same part with the style's fill dropped, for a cell that says `a:noFill`. */
export function withoutFill(part: TableStylePart): TableStylePart {
  const { shadingHex: _fill, ...rest } = part;
  return rest;
}

// The first colour anywhere under a node — a fill, an outline and a text style
// all wrap theirs the same way.
function colorOf(node: PoNode, colors: ColorResolver): string | undefined {
  for (const child of poChildren(node)) {
    const hex = resolveColorNode(child, colors);
    if (hex !== undefined) return hex;
    const deeper = poTag(child) !== undefined ? colorOf(child, colors) : undefined;
    if (deeper !== undefined) return deeper;
  }
  return undefined;
}

/** A cell with the style's fill, borders and run properties filled in under its own. */
export function withCellStyle(cell: TableCell, part: TableStylePart): TableCell {
  const runProps = {
    ...(part.bold === true ? { bold: true } : {}),
    ...(part.italic === true ? { italic: true } : {}),
    ...(part.colorHex !== undefined ? { colorHex: part.colorHex } : {}),
  };
  const content =
    Object.keys(runProps).length === 0
      ? cell.content
      : cell.content.map((block) =>
          block.kind === 'paragraph'
            ? {
                ...block,
                paragraph: {
                  ...block.paragraph,
                  runs: block.paragraph.runs.map((run) => ({
                    ...run,
                    properties: { ...runProps, ...run.properties },
                  })),
                },
              }
            : block,
        );
  // Borders compose SIDE BY SIDE: a cell that states only its bottom rule keeps
  // the style's other three, which a whole-object override would throw away.
  const borders =
    part.borders || cell.properties.borders
      ? { ...part.borders, ...cell.properties.borders }
      : undefined;
  return {
    ...cell,
    properties: {
      ...(part.shadingHex !== undefined ? { shading: { colorHex: part.shadingHex } } : {}),
      ...cell.properties,
      ...(borders ? { borders } : {}),
    },
    content,
  };
}
