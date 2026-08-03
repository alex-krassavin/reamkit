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
import { readColorMods, resolveColorNode } from '@/core/drawingml/colors';
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

/**
 * The style a cell wears, composed from every part that reaches it.
 *
 * @param style  The `a:tblStyle` node.
 * @param flags  Which conditional parts the table asks for.
 * @param at     Where the cell sits.
 * @param colors The deck's colour resolver.
 * @returns The composed part, empty when the style says nothing.
 */
export function cellStyle(
  style: PoNode,
  flags: TableStyleFlags,
  at: CellPosition,
  colors: ColorResolver,
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

  let out: TableStylePart = {};
  for (const name of names) {
    const part = poChildren(style).find((c) => poIs(c, name));
    if (part) out = { ...out, ...partStyle(part, colors) };
  }
  return out;
}

// Bands count from the first row/column that is NOT an edge one, so a table
// with a header starts its first band on the row below it.
function bandName(index: number, edge: boolean, axis: 'H' | 'V'): string {
  const i = edge ? index - 1 : index;
  return i % 2 === 0 ? `a:band1${axis}` : `a:band2${axis}`;
}

function partStyle(part: PoNode, colors: ColorResolver): TableStylePart {
  const tcStyle = poChildren(part).find((c) => poIs(c, 'a:tcStyle'));
  const txStyle = poChildren(part).find((c) => poIs(c, 'a:tcTxStyle'));
  const fill = tcStyle ? poChildren(tcStyle).find((c) => poIs(c, 'a:fill')) : undefined;
  const solid = fill ? poChildren(fill).find((c) => poIs(c, 'a:solidFill')) : undefined;
  const shadingHex = solid ? colorOf(solid, colors) : undefined;
  const bdr = tcStyle ? poChildren(tcStyle).find((c) => poIs(c, 'a:tcBdr')) : undefined;
  const borders = bdr ? partBorders(bdr, colors) : undefined;
  const on = (name: string): boolean => poAttr(txStyle, name) === 'on';
  const colorHex = txStyle ? colorOf(txStyle, colors) : undefined;
  return {
    ...(shadingHex ? { shadingHex } : {}),
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

function partBorders(bdr: PoNode, colors: ColorResolver): CellBorders | undefined {
  const out: { -readonly [K in keyof CellBorders]?: Border } = {};
  for (const [tag, side] of SIDES) {
    const holder = poChildren(bdr).find((c) => poIs(c, tag));
    const ln = holder ? poChildren(holder).find((c) => poIs(c, 'a:ln')) : undefined;
    if (ln) out[side] = lineBorder(ln, colors);
  }
  return Object.keys(out).length > 0 ? out : undefined;
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
