// A Word table's column widths, as Word lays them out when it opens the file.
//
// §17.4.49 `w:tblGrid` is the grid Word last laid the table out at; §17.4.72
// `w:tcW` is the width each cell asks for, and §17.4.63 `w:tblW` the table's.
// Word 16.114 trusts the cells, not the grid. Measured on our own probes
// (`.git/wip-tcw-grid/ptcw*`, Word's PDFs against Ream's), in compatibility
// modes 15, 14 and none alike:
//
// - An AUTOFIT table is laid out afresh every time: from the cells' widths,
//   their text and the table's width. The grid plays no part. A grid of
//   2000+4000 twips around the cells "A" and "B" drew two 17.5pt columns; a
//   grid of 100+100 under cells that ask for 1in and 3in drew 72+216pt. The
//   docx package writes exactly that: `w:gridCol w:w="100"` for every column
//   and the real widths in the cells.
// - A FIXED table takes the cells' widths, and the grid's only for a column no
//   cell gives one. It never looks at the text.
//
// Where the grid already IS what the cells declare, it is kept: it is Word's
// own measurement of the text, which a substituted face can only approximate,
// and 2_table_doc.docx (two 221pt cells on a 221pt grid) must not move.

import type { Table, TableCell } from '@/core/document-model';

/** What one cell's content needs across, its margins left out. */
export interface CellContentWidths {
  /** The widest thing in it that cannot break: a word, a picture, a nested table. */
  readonly minPt: number;
  /** Its widest paragraph set on one line. */
  readonly maxPt: number;
}

/** What {@link wordColumnWidths} needs to know besides the table. */
export interface WordColumnsContext {
  /** How many columns the rows use (at least the grid's). */
  readonly colCount: number;
  /**
   * The width a percentage is a share of, and the most an autofit table of
   * no stated width may take: the text column the table stands in.
   */
  readonly availablePt: number;
  /** A cell's left and right margins, resolved. */
  padding: (cell: TableCell) => { readonly left: number; readonly right: number };
  /** A cell's content widths. Asked only of an autofit table. */
  content: (cell: TableCell) => CellContentWidths;
}

/**
 * The room Word leaves for a cell's text however narrow the cell is asked to
 * be: 20 twips. With the default 5.4pt margins a table of no width, or a cell
 * of 50 twips, drew 11.8pt cells; with margins of 20, 50 and 0pt a side the
 * cells were 41, 101 and (on a 5pt grid) 5pt. Border widths of 0, 0.5 and 3pt
 * changed nothing.
 */
const MIN_TEXT_PT = 1;

/**
 * What Word gives a fixed table's column whose grid width is 0 and that no cell
 * gives a width: a quarter inch (Word read the cells back as 18pt, with
 * margins of 0 and 5.4pt alike).
 */
const ZERO_GRID_COLUMN_PT = 18;

/** How far a grid column may stand from what its cells declare and still be theirs. */
function agrees(gridPt: number, declaredPt: number): boolean {
  return Math.abs(gridPt - declaredPt) <= Math.max(1, 0.02 * declaredPt);
}

/** A cell's own width, as it states it. */
type Declared =
  | { readonly kind: 'pct'; readonly fraction: number }
  | { readonly kind: 'abs'; readonly widthPt: number };

function declaredWidth(cell: TableCell): Declared | undefined {
  const p = cell.properties;
  if (p.widthFraction !== undefined) {
    // A share of nothing, or less, is no width at all: Word drew `pct 0`
    // cells as it drew cells of no width.
    return p.widthFraction > 0 ? { kind: 'pct', fraction: p.widthFraction } : undefined;
  }
  // `auto` and `nil` say "no width"; zero, a negative number and junk (which
  // the reader leaves unset) were drawn as no width too.
  if (p.width === undefined || p.widthType === 'auto' || p.widthType === 'nil') return undefined;
  return p.width > 0 ? { kind: 'abs', widthPt: p.width } : undefined;
}

/** The table's own width, when it states one Word can use. */
function statedWidth(table: Table, availablePt: number): number | undefined {
  const p = table.properties;
  if (p.widthType === 'dxa' && p.widthPt !== undefined && p.widthPt > 0) return p.widthPt;
  if (p.widthType === 'pct' && p.widthFraction !== undefined && p.widthFraction > 0) {
    return p.widthFraction * availablePt;
  }
  // `auto`, `nil`, a zero or negative width: the table is as wide as its
  // columns make it (Word drew a `pct 0` and a `pct -50%` table so).
  return undefined;
}

/** One column as the rows claim it. */
interface Column {
  /** The largest share of the table a cell of this column alone asks for. */
  pct: number;
  /** The largest absolute width such a cell asks for. */
  abs: number;
  /** The narrowest the column may be. */
  min: number;
  /** What its content wants on one line (autofit). */
  max: number;
  /** Its cells' margins: what is left of it with no text at all. */
  pad: number;
  /** A cell of its own starts in it. */
  owned: boolean;
  /** A spanning cell covers it. */
  spanned: boolean;
}

/** A cell across several columns: it claims only their sum. */
interface Span {
  readonly start: number;
  readonly end: number;
  readonly declared: Declared | undefined;
  readonly min: number;
  readonly max: number;
}

/**
 * The column widths of a Word table, by Word's rules for its layout mode.
 *
 * @param table The table.
 * @param ctx   The column count, the width to measure against, and the cells'
 *              margins and content.
 * @returns The width of each column in points, or `undefined` where the
 *          table's grid is already what its cells declare and stands as it is.
 */
export function wordColumnWidths(table: Table, ctx: WordColumnsContext): Array<number> | undefined {
  const n = ctx.colCount;
  const fixed = table.properties.layout === 'fixed';
  const grid = (i: number): number => table.grid[i] ?? 0;

  // How far the rows that actually divide the table reach. A grid column past
  // that is Word's own bookkeeping, touched only by a caption row spanning the
  // lot: fdo38414.docx carries a twelfth of an inch of it, and giving it width
  // hung that one row out past every other.
  // A column a cell starts in is no bookkeeping, though: gridbefore.docx
  // starts its first row's one cell two columns in.
  let reach = 0;
  for (const row of table.rows) {
    let at = row.properties.gridBefore ?? 0;
    if (row.cells.length > 0 && at > 0) reach = Math.max(reach, at + 1);
    for (const cell of row.cells) {
      at += Math.max(1, cell.properties.colSpan ?? 1);
      if (row.cells.length >= 2) reach = Math.max(reach, at);
    }
  }
  const live = reach > 0 && reach < n ? reach : n;

  const cols: Array<Column> = Array.from({ length: n }, () => ({
    pct: 0,
    abs: 0,
    min: 0,
    max: 0,
    pad: 0,
    owned: false,
    spanned: false,
  }));
  const spans: Array<Span> = [];
  // Every cell, spanning or not, for checking the grid against.
  const claims: Array<Span> = [];
  for (const row of table.rows) {
    let at = row.properties.gridBefore ?? 0;
    for (const cell of row.cells) {
      const span = Math.max(1, cell.properties.colSpan ?? 1);
      if (at >= n) break;
      const pad = ctx.padding(cell);
      const floor = pad.left + pad.right + MIN_TEXT_PT;
      let min = floor;
      let max = floor;
      if (!fixed) {
        const w = ctx.content(cell);
        min = Math.max(floor, w.minPt + pad.left + pad.right);
        max = Math.max(min, w.maxPt + pad.left + pad.right);
      }
      const declared = declaredWidth(cell);
      const end = Math.min(at + span, live);
      if (end > at) claims.push({ start: at, end, declared, min, max });
      if (span === 1) {
        const col = cols[at]!;
        col.owned = true;
        if (declared?.kind === 'pct') col.pct = Math.max(col.pct, declared.fraction);
        if (declared?.kind === 'abs') col.abs = Math.max(col.abs, declared.widthPt);
        col.min = Math.max(col.min, min);
        col.max = Math.max(col.max, max);
        col.pad = Math.max(col.pad, pad.left + pad.right);
      } else {
        for (let i = at; i < end; i++) cols[i]!.spanned = true;
        if (end > at) spans.push({ start: at, end, declared, min, max });
      }
      at += span;
    }
  }

  if (!fixed && gridIsDeclared(claims, live, table)) return undefined;

  // A spanning cell whose columns all but one are sized sizes that one too:
  // 2in across two columns, one of them 1in, makes the other 1in.
  for (let changed = true; changed; ) {
    changed = false;
    for (const s of spans) {
      if (s.declared?.kind !== 'abs') continue;
      let known = 0;
      const open: Array<number> = [];
      let share = false; // a share is settled with the table's width, later
      for (let i = s.start; i < s.end; i++) {
        const c = cols[i]!;
        if (c.pct > 0) share = true;
        else if (c.abs > 0) known += c.abs;
        else open.push(i);
      }
      if (share || open.length !== 1 || s.declared.widthPt - known <= 0) continue;
      cols[open[0]!]!.abs = s.declared.widthPt - known;
      changed = true;
    }
  }

  // Past the grid's end a fixed table has no column to size, and a cell there
  // no width of its own: a .doc row's trailing cells beyond its last edge
  // draw nothing.
  if (fixed) {
    for (let i = table.grid.length; i < live; i++) {
      const c = cols[i]!;
      if (c.pct === 0 && c.abs === 0) c.min = 0;
    }
  }

  // What each column asks for. Percentages wait for the table's width.
  //   fixed:   the cells' width, else the grid's (a quarter inch for none);
  //   autofit: the cells' width, else the content's on one line. A width too
  //            narrow for the column's longest word grows to it (1in for 15
  //            M's drew 135.8pt); the text alone never shrinks it (a sentence
  //            in a 1in cell wraps at 1in).
  // A column no cell starts in or spans (a `w:gridBefore` gap) keeps its grid.
  const isPct = (i: number): boolean => cols[i]!.pct > 0;
  const isAuto = (i: number): boolean =>
    !fixed && !isPct(i) && cols[i]!.abs === 0 && (cols[i]!.owned || cols[i]!.spanned);
  const want: Array<number> = cols.map((c, i) => {
    if (i >= live || isPct(i)) return 0;
    if (c.abs > 0) return Math.max(c.abs, c.min);
    if (isAuto(i)) return c.max;
    if (fixed && i >= table.grid.length) return 0;
    const g = grid(i);
    if (!fixed) return g;
    return Math.max(g > 0 ? g : ZERO_GRID_COLUMN_PT, c.min);
  });

  // A spanning cell's own width is a floor under the columns it covers.
  const stated = statedWidth(table, ctx.availablePt);
  for (const s of spans) {
    if (s.declared?.kind !== 'abs' && !(s.declared?.kind === 'pct' && stated !== undefined)) {
      continue;
    }
    const widthPt =
      s.declared.kind === 'abs' ? s.declared.widthPt : s.declared.fraction * (stated ?? 0);
    floorSpan(want, s, widthPt, isPct);
  }
  if (!fixed) {
    for (const s of spans) fitSpanText(want, cols, s, isPct);
    for (const s of spans) spreadSpanLine(want, s, isAuto, isPct);
  }

  // The table's width. Stated, it is that (Word drew 1in+3in cells in a 6in
  // table at 108+324pt). Else it is what the columns add up to, the
  // percentages taking their share of it: 1in beside 50% made a 2in table, a
  // 25% cell beside a 200pt grid column a 266.7pt one. A table of nothing but
  // percentages is as narrow as its narrowest share allows (25%+75% drew
  // 11.8+35.4pt, 17.5+52.5pt with text). An autofit table of no width never
  // takes more than the text column (5in+5in drew 261.6pt each); a fixed one
  // may, and so may a stated width (10in drew 720pt).
  let sumPct = 0;
  let sumWant = 0;
  let sumMin = 0;
  for (let i = 0; i < live; i++) {
    sumMin += cols[i]!.min;
    if (isPct(i)) sumPct += cols[i]!.pct;
    else sumWant += want[i]!;
  }
  let tableW: number;
  if (stated !== undefined) {
    tableW = stated;
  } else {
    if (sumPct < 1) tableW = sumWant / (1 - sumPct);
    else tableW = sumWant > 0 ? ctx.availablePt : 0;
    for (let i = 0; i < live; i++) {
      if (isPct(i)) tableW = Math.max(tableW, cols[i]!.min / cols[i]!.pct);
    }
    if (!fixed) tableW = Math.min(tableW, ctx.availablePt);
  }
  // No column goes under its minimum, whatever the table asked for: two 12 M
  // words in a 2in autofit table made it 234.9pt. Except past the text column
  // in autofit: there the words break (two 400pt words in a table of no width
  // and in a 6in one drew 261.4pt each).
  const limit = fixed ? Infinity : Math.max(ctx.availablePt, stated ?? 0);
  tableW = Math.max(tableW, Math.min(sumMin, limit));

  const out = new Array<number>(n).fill(0);
  // Percentages first. Asking for more than the whole, each takes at most
  // what the table has left, left to right: 75% beside 75% drew 75% and 25%.
  let left = tableW;
  for (let i = 0; i < live; i++) {
    if (!isPct(i)) continue;
    const share = cols[i]!.pct * tableW;
    const w = Math.max(cols[i]!.min, sumPct > 1 ? Math.min(share, Math.max(0, left)) : share);
    out[i] = w;
    left -= w;
  }

  // The rest share what the percentages leave.
  const rest: Array<number> = [];
  for (let i = 0; i < live; i++) if (!isPct(i)) rest.push(i);
  for (const i of rest) out[i] = want[i]!;
  let have = rest.reduce((s, i) => s + out[i]!, 0);
  const room = Math.max(0, left);
  if (have < room) {
    // Growing: an autofit table gives the slack to its columns of no width,
    // in proportion to their text (1in beside an empty cell in a full-width
    // table left 1in; "A" beside a sentence drew 44pt and 479pt). With every
    // column sized, or in a fixed table, all of them grow in proportion
    // (1in+3in on a full line drew 25% and 75%).
    const autos = rest.filter(isAuto);
    grow(out, autos.length > 0 ? autos : rest, room - have);
  } else if (have > room) {
    // Shrinking takes from each column what it has above its minimum, in
    // proportion: 1in+3in in a 2in table drew 39.1+105.2pt fixed, 40.8+103.2pt
    // autofit, where the minimum is the text's longest word. An autofit table
    // takes from its columns of no width first (1in beside a long sentence in
    // a full-width table kept its 1in).
    let excess = have - room;
    const groups = fixed ? [rest] : [rest.filter(isAuto), rest.filter((i) => !isAuto(i))];
    for (const g of groups) excess = shrink(out, cols, g, excess);
  }
  have = 0;
  for (let i = 0; i < live; i++) have += out[i]!;
  // Columns at their minimum that still overrun the table take it from the
  // percentages: 1in beside 50%+50% left the 1in column its text's width and
  // the two shares 252.9pt each.
  const all = Array.from({ length: live }, (_, i) => i);
  if (have > tableW) shrink(out, cols, all.filter(isPct), have - tableW);
  // Still too wide, the words themselves break: every column gives up its text
  // in proportion, down to its margins. A 600pt word beside "B" drew 506.5pt
  // and 16.4pt: 600 and 6.7pt of text scaled into the 501.7pt left between
  // the margins.
  have = 0;
  for (let i = 0; i < live; i++) have += out[i]!;
  if (!fixed && have > tableW) shrink(out, cols, all, have - tableW, 'pad');
  return out;
}

/**
 * Whether the grid is what the cells declare: every cell that states a width
 * spans that much of the grid (a spanning one at least that much), and every
 * column is under one.
 * Then the grid is the width Word last laid the cells out at, and stands — and
 * it alone says how a span divides between columns no single cell sizes.
 * Not where a cell's longest word would not fit its part of the grid, which
 * Word would have widened (1in cells on a 1in grid around 15 M's drew
 * 135.8pt).
 */
function gridIsDeclared(claims: ReadonlyArray<Span>, live: number, table: Table): boolean {
  if (table.rows.length === 0) return true;
  let total = 0;
  for (const g of table.grid) total += g;
  if (total <= 0 || table.grid.length < live) return false;
  const covered = new Array<boolean>(live).fill(false);
  for (const c of claims) {
    let g = 0;
    for (let i = c.start; i < c.end; i++) g += table.grid[i]!;
    if (g < c.min && !agrees(g, c.min)) return false;
    if (!c.declared) continue;
    const want = c.declared.kind === 'abs' ? c.declared.widthPt : c.declared.fraction * total;
    // A spanning cell's width is a floor: Word gives it more where the rows
    // around it need it (2in across a 3in and a 1in column drew 4in).
    if (c.end - c.start > 1 ? g < want && !agrees(g, want) : !agrees(g, want)) return false;
    for (let i = c.start; i < c.end; i++) covered[i] = true;
  }
  return covered.every(Boolean);
}

/**
 * A spanning cell's text set on one line asks its columns for that much
 * together: what they fall short of goes to those no cell gives a width, in
 * proportion to their own text. A full-width caption over "Year" and three
 * longer headings left "Year" as narrow beside the others as it is without it
 * (bug59058.docx).
 */
function spreadSpanLine(
  want: Array<number>,
  s: Span,
  isAuto: (i: number) => boolean,
  isPct: (i: number) => boolean,
): void {
  let have = 0;
  const autos: Array<number> = [];
  for (let i = s.start; i < s.end; i++) {
    if (isPct(i)) return;
    have += want[i]!;
    if (isAuto(i)) autos.push(i);
  }
  const deficit = s.max - have;
  if (deficit <= 0 || autos.length === 0) return;
  const base = autos.reduce((sum, i) => sum + want[i]!, 0);
  for (const i of autos) {
    want[i] = want[i]! + (base > 0 ? (deficit * want[i]!) / base : deficit / autos.length);
  }
}

/**
 * A spanning cell's width as a floor under its columns: what they fall short
 * of is shared over all of them in proportion to their widths. Word drew a
 * full-width first row over "A. Each", "B. Every", a 117.5pt "C. Most" and
 * "D. A lot" by widening all four 2.048 times (the sized one to 240.6pt);
 * without that row the 117.5pt stood and the other three took the slack. A
 * cell that spans columns says nothing about any ONE of them; sharing it out
 * per column instead let the full-width caption row of fdo38414.docx overrule
 * every column above it.
 */
function floorSpan(
  want: Array<number>,
  s: Span,
  widthPt: number,
  isPct: (i: number) => boolean,
): void {
  let have = 0;
  for (let i = s.start; i < s.end; i++) {
    if (isPct(i)) return; // its share is settled with the table's width
    have += want[i]!;
  }
  const deficit = widthPt - have;
  if (deficit <= 0) return;
  for (let i = s.start; i < s.end; i++) {
    want[i] = have > 0 ? want[i]! * (widthPt / have) : want[i]! + deficit / (s.end - s.start);
  }
}

/**
 * A spanning cell whose longest word or picture its columns cannot hold
 * widens them to it, sharing it in proportion to each column's width and its
 * own text together. Word set a 400pt word across a 1in and a 3in column at
 * 113.8+296.9pt, across a 1in column and one holding "B" at 295.2+115.5pt;
 * the columns it widens hold that width from then on.
 */
function fitSpanText(
  want: Array<number>,
  cols: Array<Column>,
  s: Span,
  isPct: (i: number) => boolean,
): void {
  let have = 0;
  let weight = 0;
  for (let i = s.start; i < s.end; i++) {
    if (isPct(i)) return; // its share is settled with the table's width
    have += want[i]!;
    weight += want[i]! + cols[i]!.max;
  }
  if (s.min <= have) return;
  const n = s.end - s.start;
  for (let i = s.start; i < s.end; i++) {
    const w = weight > 0 ? (s.min * (want[i]! + cols[i]!.max)) / weight : s.min / n;
    want[i] = w;
    cols[i]!.min = Math.max(cols[i]!.min, w);
  }
}

/** Widen `which` by `extra` in proportion to their widths (evenly when all are 0). */
function grow(out: Array<number>, which: ReadonlyArray<number>, extra: number): void {
  if (which.length === 0 || extra <= 0) return;
  const base = which.reduce((s, i) => s + out[i]!, 0);
  for (const i of which) {
    out[i] = out[i]! + (base > 0 ? (extra * out[i]!) / base : extra / which.length);
  }
}

/**
 * Narrow `which` by up to `excess`, each by its share of what the group has
 * above its floor: its minimum, or (once the text must break) its margins.
 *
 * @returns What is still to take once the group is at its floors.
 */
function shrink(
  out: Array<number>,
  cols: ReadonlyArray<Column>,
  which: ReadonlyArray<number>,
  excess: number,
  floor: 'min' | 'pad' = 'min',
): number {
  if (excess <= 0) return 0;
  const base = (i: number): number => (floor === 'min' ? cols[i]!.min : cols[i]!.pad);
  let spare = 0;
  for (const i of which) spare += Math.max(0, out[i]! - base(i));
  if (spare <= 0) return excess;
  const take = Math.min(excess, spare);
  for (const i of which) {
    const above = Math.max(0, out[i]! - base(i));
    out[i] = out[i]! - (take * above) / spare;
  }
  return excess - take;
}
