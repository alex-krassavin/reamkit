// ECMA-376 Part 1 §17.4 — Table parser.
// Parses w:tbl in preserveOrder format into the typed Table model.

import type {
  Border,
  BorderStyle,
  CellBorders,
  CellMargins,
  CellMerge,
  CellProperties,
  FloatAnchor,
  RowConditionalFormat,
  RowProperties,
  RunProperties,
  Table,
  TableCell,
  TableLook,
  TableProperties,
  TableRow,
} from '@/core/document-model';

import type { PoNode } from '@/core/po-helpers';
import type { ThemeFonts } from '@/core/drawingml/theme-parser';
import type { Pt } from '@/core/ir';
import type { ParseContext } from '@/word/document-parser';
import { DEFAULT_PARSE_CONTEXT, parseBodyElements, sdtRunProperties } from '@/word/document-parser';
import { eighthPtToPt, pt, twipsToPt } from '@/core/ir';
import {
  poAttr,
  poChildren,
  poChildrenWith,
  poFirstChild,
  poIntAttr,
  poIs,
  poToggle,
  poVal,
} from '@/core/po-helpers';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

const BORDER_STYLES = new Set<BorderStyle>([
  'none',
  'single',
  'double',
  'thick',
  'dotted',
  'dashed',
  'dashSmallGap',
]);

// §17.18.2 — the patterns drawn as more than one parallel line. Each becomes a
// `double`: two rules of the stated total width, which is what they all are.
const MULTI_LINE_BORDERS = new Set([
  'triple',
  'doubleWave',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'outset',
  'inset',
]);

const WIDTH_TYPES = new Set<'auto' | 'dxa' | 'pct' | 'nil'>(['auto', 'dxa', 'pct', 'nil']);
const HEIGHT_RULES = new Set<'auto' | 'atLeast' | 'exact'>(['auto', 'atLeast', 'exact']);

// §17.4.58 `w:tblpPr` → the anchor a drawing would state. `w:horzAnchor` and
// `w:vertAnchor` name the frame; `w:tblpX`/`w:tblpY` an offset in twips, and
// `w:tblpXSpec`/`w:tblpYSpec` an alignment instead. The four `…FromText`
// attributes are the standoff the wrapped text keeps.
function parseTablePosition(tblpPr: PoNode): FloatAnchor {
  const horz = poAttr(tblpPr, 'horzAnchor');
  const vert = poAttr(tblpPr, 'vertAnchor');
  const hRel = horz === 'page' ? 'page' : horz === 'margin' ? 'margin' : 'column';
  const vRel = vert === 'page' ? 'page' : vert === 'margin' ? 'margin' : 'paragraph';
  const xSpec = poAttr(tblpPr, 'tblpXSpec');
  const align =
    xSpec === 'center'
      ? 'center'
      : xSpec === 'right'
        ? 'right'
        : xSpec === 'left'
          ? 'left'
          : undefined;
  const x = poIntAttr(tblpPr, 'tblpX');
  const y = poIntAttr(tblpPr, 'tblpY');
  const dist = (name: string): Pt => twipsToPt(poIntAttr(tblpPr, name) ?? 0);
  return {
    // A floating table always has text beside it — `w:tblpPr` IS the wrap.
    wrap: 'square',
    posH: { relativeFrom: hRel, ...(align ? { align } : { offsetPt: twipsToPt(x ?? 0) }) },
    posV: { relativeFrom: vRel, offsetPt: twipsToPt(y ?? 0) },
    wrapDist: {
      topPt: dist('topFromText'),
      bottomPt: dist('bottomFromText'),
      leftPt: dist('leftFromText'),
      rightPt: dist('rightFromText'),
    },
  };
}

/**
 * Parse a `w:tbl` (ECMA-376 Part 1 §17.4) in preserveOrder shape into the typed
 * {@link Table} model: its properties (style ref, look, width, layout, borders,
 * default cell margins), column grid and rows. Vertical merges are resolved in a
 * second pass so cells carry the resolved {@link CellMerge} role, not the raw
 * `w:vMerge` markers.
 *
 * @param tbl The `w:tbl` PoNode.
 * @param ctx The document-wide parse context (resolvers for colour, images, etc.).
 * @returns The parsed table.
 */
export function parseTable(tbl: PoNode, ctx: ParseContext = DEFAULT_PARSE_CONTEXT): Table {
  const properties = parseTableProperties(poFirstChild(tbl, 'w:tblPr'));
  const grid = parseTableGrid(poFirstChild(tbl, 'w:tblGrid'));
  // Two-phase: collect rows with their raw §17.4.85 vMerge markers, then
  // resolve the markers into CellMerge roles — the model carries the resolved
  // semantics, not the OOXML encoding.
  const draftRows: Array<{ properties: RowProperties; cells: Array<DraftCell> }> = [];
  for (const tr of poChildrenThroughSdt(tbl, 'w:tr')) {
    if (isDeletedRow(tr)) continue;
    draftRows.push(parseTableRow(tr, ctx));
  }
  const roles = resolveMergeRoles(draftRows.map((r) => r.cells));
  const rows: Array<TableRow> = draftRows.map((draft, rowIdx) => ({
    properties: draft.properties,
    cells: draft.cells.map((d, cellIdx) => {
      const merge = roles[rowIdx]![cellIdx];
      return merge ? { ...d.cell, properties: { ...d.cell.properties, merge } } : d.cell;
    }),
  }));
  // §17.4.1 `w:bidiVisual` — the table reads RIGHT TO LEFT: its first cell is
  // the RIGHTMOST one. Reversing the row here (and the grid with it) leaves
  // every span, border and width resolving in the order they are drawn, so
  // table-rtl.docx's B column stands where the reference puts it, first.
  if (poFirstChild(poFirstChild(tbl, 'w:tblPr'), 'w:bidiVisual')) {
    return {
      properties,
      grid: [...grid].reverse(),
      rows: rows.map((row) => ({ ...row, cells: [...row.cells].reverse() })),
    };
  }
  return { properties, grid, rows };
}

// §17.5.2 — a content control may wrap the ROWS of a table or the CELLS of a
// row: `w:sdt` is chrome, and its `w:sdtContent` holds the real thing. Read as
// a plain child list the wrapper hid them, and cell-sdt-redline.docx's one-cell
// table came out with no cells at all.
function poChildrenThroughSdt(node: PoNode | undefined, tag: string): Array<PoNode> {
  return childrenThroughSdt(node, tag).map((c) => c.node);
}

// …and the run properties the wrapper lends what it holds (§17.5.2.28), for
// the callers that carry them further down.
function childrenThroughSdt(
  node: PoNode | undefined,
  tag: string,
  themeFonts?: ThemeFonts,
): Array<{ node: PoNode; sdtRunProps?: RunProperties }> {
  const out: Array<{ node: PoNode; sdtRunProps?: RunProperties }> = [];
  for (const child of poChildren(node)) {
    if (poIs(child, tag)) out.push({ node: child });
    else if (poIs(child, 'w:sdt')) {
      const props = sdtRunProperties(child, themeFonts);
      for (const inner of poChildrenWith(poFirstChild(child, 'w:sdtContent'), tag)) {
        out.push({ node: inner, ...(props ? { sdtRunProps: props } : {}) });
      }
    }
  }
  return out;
}

interface DraftCell {
  readonly cell: TableCell;
  readonly vMerge?: 'restart' | 'continue';
}

// ECMA-376 §17.4.85 (vMerge). Walk each logical column top-down and tag every
// cell with its position in a vertical merge group:
//   start  — vMerge="restart" with at least one following "continue"
//   middle — vMerge="continue" with another "continue" right after
//   end    — vMerge="continue" terminating a group
//   undefined — not merged (standalone)
function resolveMergeRoles(
  rows: ReadonlyArray<ReadonlyArray<DraftCell>>,
): Array<Array<CellMerge | undefined>> {
  const out: Array<Array<CellMerge | undefined>> = rows.map((r) =>
    new Array<CellMerge | undefined>(r.length).fill(undefined),
  );
  const colSlots = new Map<
    number,
    Array<{ rowIdx: number; cellIdx: number; vMerge: 'restart' | 'continue' | undefined }>
  >();
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx]!;
    let colIdx = 0;
    for (let cellIdx = 0; cellIdx < row.length; cellIdx++) {
      const d = row[cellIdx]!;
      let arr = colSlots.get(colIdx);
      if (!arr) {
        arr = [];
        colSlots.set(colIdx, arr);
      }
      arr.push({ rowIdx, cellIdx, vMerge: d.vMerge });
      colIdx += Math.max(1, d.cell.properties.colSpan ?? 1);
    }
  }
  for (const slots of colSlots.values()) {
    for (let i = 0; i < slots.length; i++) {
      const cur = slots[i]!;
      const next = slots[i + 1];
      const nextIsContinue = !!next && next.vMerge === 'continue';
      let role: CellMerge | undefined;
      if (cur.vMerge === 'restart') {
        role = nextIsContinue ? 'start' : undefined;
      } else if (cur.vMerge === 'continue') {
        role = nextIsContinue ? 'middle' : 'end';
      }
      out[cur.rowIdx]![cur.cellIdx] = role;
    }
  }
  return out;
}

function parseTableProperties(tblPr: PoNode | undefined): TableProperties {
  if (!tblPr) return {};
  const out: Mutable<TableProperties> = {};

  // §17.7.6 w:tblStyle + §17.4.62 w:tblLook — raw references the reader's
  // resolveTableStyles transform consumes (round-trip material afterwards).
  const tblStyle = poFirstChild(tblPr, 'w:tblStyle');
  if (tblStyle) {
    const id = poVal(tblStyle);
    if (id) out.styleId = id;
  }
  const tblLook = poFirstChild(tblPr, 'w:tblLook');
  if (tblLook) {
    const look = parseTableLook(tblLook);
    if (look) out.look = look;
  }

  const tblW = poFirstChild(tblPr, 'w:tblW');
  if (tblW) {
    const w = poIntAttr(tblW, 'w');
    const type = poAttr(tblW, 'type');
    // tblW @w is twips for type=dxa but fiftieths of a percent for type=pct
    // (5000 = 100% of the content width) — store each in its own field.
    if (w !== undefined && type === 'pct') out.widthFraction = w / 5000;
    else if (w !== undefined) out.widthPt = twipsToPt(w);
    if (type && WIDTH_TYPES.has(type as 'auto' | 'dxa' | 'pct' | 'nil')) {
      out.widthType = type as 'auto' | 'dxa' | 'pct' | 'nil';
    }
  }

  // §17.4.65 `w:tblInd` — the table's own indent from the text margin. Read
  // nowhere, a table that declares one was drawn flush to the margin:
  // NumberedList.docx indents its procedure table through its table style.
  const tblInd = poFirstChild(tblPr, 'w:tblInd');
  if (tblInd) {
    const w = poIntAttr(tblInd, 'w');
    const type = poAttr(tblInd, 'type');
    if (w !== undefined && (type === undefined || type === 'dxa')) out.indentPt = twipsToPt(w);
  }

  // §17.4.58 `w:tblpPr` — the table floats: it is placed at an anchor of its
  // own and the text runs past it. Read nowhere, fdo77887.docx's floating form
  // sat in the flow, taller than the page it should have shared, and left the
  // first page blank.
  const tblpPr = poFirstChild(tblPr, 'w:tblpPr');
  if (tblpPr) out.float = parseTablePosition(tblpPr);

  const tblLayout = poFirstChild(tblPr, 'w:tblLayout');
  if (tblLayout) {
    const t = poAttr(tblLayout, 'type');
    if (t === 'fixed' || t === 'auto') out.layout = t;
  }

  // §17.4.27 `w:tblPr/w:jc` — where a table narrower than the text column
  // sits in it. Read nowhere, fdo66474.docx's right-aligned header table sat
  // against the left margin, 170pt from where Word and LibreOffice put it.
  const tblJc = poVal(poFirstChild(tblPr, 'w:jc'));
  if (tblJc === 'center') out.alignment = 'center';
  else if (tblJc === 'right' || tblJc === 'end') out.alignment = 'right';
  else if (tblJc === 'left' || tblJc === 'start') out.alignment = 'left';

  const borders = parseBorders(poFirstChild(tblPr, 'w:tblBorders'));
  if (borders) out.borders = borders;

  const margins = parseCellMargins(poFirstChild(tblPr, 'w:tblCellMar'));
  if (margins) out.defaultCellMargins = margins;

  return out;
}

// §17.4.62 — modern files carry explicit flag attributes; legacy files encode
// the same flags in a hex @w:val bitmask (0020 firstRow, 0040 lastRow,
// 0080 firstColumn, 0100 lastColumn, 0200 noHBand, 0400 noVBand).
function parseTableLook(node: PoNode): TableLook | undefined {
  const flag = (name: string, bit: number): boolean | undefined => {
    const attr = poAttr(node, name);
    if (attr !== undefined) return attr === '1' || attr === 'true';
    const valRaw = poAttr(node, 'val');
    if (valRaw === undefined) return undefined;
    const mask = parseInt(valRaw, 16);
    return Number.isFinite(mask) ? (mask & bit) !== 0 : undefined;
  };
  const firstRow = flag('firstRow', 0x0020);
  const lastRow = flag('lastRow', 0x0040);
  const firstColumn = flag('firstColumn', 0x0080);
  const lastColumn = flag('lastColumn', 0x0100);
  const noHBand = flag('noHBand', 0x0200);
  const noVBand = flag('noVBand', 0x0400);
  const out: TableLook = {
    ...(firstRow !== undefined ? { firstRow } : {}),
    ...(lastRow !== undefined ? { lastRow } : {}),
    ...(firstColumn !== undefined ? { firstColumn } : {}),
    ...(lastColumn !== undefined ? { lastColumn } : {}),
    ...(noHBand !== undefined ? { noHBand } : {}),
    ...(noVBand !== undefined ? { noVBand } : {}),
  };
  return Object.keys(out).length > 0 ? out : undefined;
}

function parseTableGrid(tblGrid: PoNode | undefined): Array<Pt> {
  if (!tblGrid) return [];
  const cols: Array<Pt> = [];
  for (const gridCol of poChildrenWith(tblGrid, 'w:gridCol')) {
    const w = poIntAttr(gridCol, 'w');
    cols.push(twipsToPt(w ?? 0));
  }
  return cols;
}

/**
 * §17.13.5.15 — whether a row was deleted, under the reader's accept-all
 * reading of tracked changes (the same one that drops `w:del` runs).
 *
 * Two spellings. `w:trPr/w:del` says it outright. The other is how a producer
 * writes a row whose CONTENT was taken away: every paragraph mark in it carries
 * `w:del` or `w:moveFrom`, which deletes the mark, and a row whose every mark
 * is gone is a row that is gone. TC-table-DnD-move.docx moves a table that way,
 * and accepting only the runs left a ghost of empty bordered rows behind where
 * Word and LibreOffice leave nothing.
 *
 * @param tr The `w:tr` element.
 * @returns Whether to drop the row.
 */
function isDeletedRow(tr: PoNode): boolean {
  const trPr = poFirstChild(tr, 'w:trPr');
  if (trPr && (poFirstChild(trPr, 'w:del') ?? poFirstChild(trPr, 'w:moveFrom'))) return true;
  const cells = poChildrenThroughSdt(tr, 'w:tc');
  if (cells.length === 0) return false;
  let marks = 0;
  for (const tc of cells) {
    for (const p of poChildrenWith(tc, 'w:p')) {
      const rPr = poFirstChild(poFirstChild(p, 'w:pPr'), 'w:rPr');
      if (!rPr || !(poFirstChild(rPr, 'w:del') ?? poFirstChild(rPr, 'w:moveFrom'))) return false;
      marks++;
    }
    // A cell holding a nested table (or nothing at all) has no mark to delete,
    // so the row cannot be read as gone from its marks.
    if (poChildrenWith(tc, 'w:tbl').length > 0) return false;
  }
  return marks > 0;
}

function parseTableRow(
  tr: PoNode,
  ctx: ParseContext,
): { properties: RowProperties; cells: Array<DraftCell> } {
  const properties = parseRowProperties(poFirstChild(tr, 'w:trPr'));
  const cells: Array<DraftCell> = [];
  for (const { node, sdtRunProps } of childrenThroughSdt(tr, 'w:tc', ctx.themeFonts)) {
    cells.push(parseTableCell(node, ctx, sdtRunProps));
  }
  return { properties, cells };
}

// §17.4.7 `w:cnfStyle` — the conditional formats of the table style this row
// takes. Word writes both the explicit attributes and the §17.18.8 bit string
// (bit 1 firstRow, bit 2 lastRow); either says the same thing. A row claims
// them whatever its position is: calendar2.docx marks its weekday row
// `firstRow="1"` so the style paints it blue like the month above it, and
// reading the position alone drew it black.
function parseRowConditional(cnf: PoNode | undefined): RowConditionalFormat | undefined {
  if (!cnf) return undefined;
  const bits = poAttr(cnf, 'val');
  const bit = (i: number): boolean => bits?.[i] === '1';
  const flag = (name: string): boolean | undefined => {
    const v = poAttr(cnf, name);
    return v === undefined ? undefined : v === '1' || v === 'true' || v === 'on';
  };
  const firstRow = flag('firstRow') ?? bit(0);
  const lastRow = flag('lastRow') ?? bit(1);
  if (!firstRow && !lastRow) return undefined;
  return { ...(firstRow ? { firstRow } : {}), ...(lastRow ? { lastRow } : {}) };
}

function parseRowProperties(trPr: PoNode | undefined): RowProperties {
  if (!trPr) return {};
  const out: Mutable<RowProperties> = {};
  const trHeight = poFirstChild(trPr, 'w:trHeight');
  if (trHeight) {
    const val = poIntAttr(trHeight, 'val');
    const rule = poAttr(trHeight, 'hRule');
    if (val !== undefined) out.height = twipsToPt(val);
    if (rule && HEIGHT_RULES.has(rule as 'auto' | 'atLeast' | 'exact')) {
      out.heightRule = rule as 'auto' | 'atLeast' | 'exact';
    } else if (val !== undefined) {
      // §17.4.81 — a height stated with no `hRule` is a MINIMUM. (The schema
      // default is `auto`, but a row that says how tall it is is not asking to
      // be measured; Word and LibreOffice both read it this way.) Ignored, the
      // tall rows of TestTableCellAlign.docx collapsed to one line each and
      // took their cells' vertical alignment down with them — nothing to be
      // bottom OF.
      out.heightRule = 'atLeast';
    }
  }
  if (poFirstChild(trPr, 'w:cantSplit'))
    out.cantSplit = poToggle(poFirstChild(trPr, 'w:cantSplit')) ?? true;
  if (poFirstChild(trPr, 'w:tblHeader')) {
    out.isHeader = poToggle(poFirstChild(trPr, 'w:tblHeader')) ?? true;
  }
  // §17.4.14 — the row's cells begin this many grid columns in. gridbefore.docx
  // puts its one cell in the third column of a three-column grid, and read as a
  // first cell it landed under the row below's.
  const before = poIntAttr(poFirstChild(trPr, 'w:gridBefore'), 'val');
  if (before !== undefined && before > 0) out.gridBefore = before;
  const cnf = parseRowConditional(poFirstChild(trPr, 'w:cnfStyle'));
  if (cnf) out.conditional = cnf;
  return out;
}

function parseTableCell(tc: PoNode, ctx: ParseContext, sdtRunProps?: RunProperties): DraftCell {
  const { properties, vMerge } = parseCellProperties(poFirstChild(tc, 'w:tcPr'));
  const content = parseBodyElements(poChildren(tc), ctx, undefined, sdtRunProps);
  return { cell: { properties, content }, ...(vMerge ? { vMerge } : {}) };
}

function parseCellProperties(tcPr: PoNode | undefined): {
  properties: CellProperties;
  vMerge?: 'restart' | 'continue';
} {
  if (!tcPr) return { properties: {} };
  const out: Mutable<CellProperties> = {};
  let rawVMerge: 'restart' | 'continue' | undefined;
  const tcW = poFirstChild(tcPr, 'w:tcW');
  if (tcW) {
    const w = poIntAttr(tcW, 'w');
    // §17.4.72 — `w` means what `w:type` says it means. Read as twips whatever
    // the type, fdo38414.docx's `w:type="pct"` widths (fiftieths of a percent)
    // became 50pt columns; the type is now kept so the layout can resolve a
    // percentage against the table and the writer can put it back as it was.
    const type = poAttr(tcW, 'type');
    if (type === 'pct' || type === 'dxa' || type === 'auto' || type === 'nil') {
      out.widthType = type;
    }
    if (w !== undefined) {
      if (type === 'pct') out.widthFraction = w / 5000;
      else if (type !== 'auto' && type !== 'nil') out.width = twipsToPt(w);
    }
  }
  const gridSpan = poFirstChild(tcPr, 'w:gridSpan');
  if (gridSpan) {
    const v = poIntAttr(gridSpan, 'val');
    if (v !== undefined) out.colSpan = v;
  }
  const vMerge = poFirstChild(tcPr, 'w:vMerge');
  if (vMerge) {
    const v = poVal(vMerge);
    rawVMerge = v === 'restart' ? 'restart' : 'continue';
  }
  const borders = parseBorders(poFirstChild(tcPr, 'w:tcBorders'));
  if (borders) out.borders = borders;
  const margins = parseCellMargins(poFirstChild(tcPr, 'w:tcMar'));
  if (margins) out.margins = margins;
  // §17.4.33 — w:shd cell shading. We honour a direct @w:fill hex (the common
  // case, e.g. a coloured header row); "auto" / theme fills are left unshaded.
  const shd = poFirstChild(tcPr, 'w:shd');
  if (shd) {
    const fill = poAttr(shd, 'fill');
    if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill)) {
      out.shading = { colorHex: fill.toUpperCase() };
    }
  }
  // §17.4.84 `w:vAlign` — where the cell's content sits in a row taller than
  // it. TestTableCellAlign.docx is two rows of exactly that and we drew all
  // four cells from the top. `both` (justified) spreads the lines out; with no
  // such mode the closest honest reading is the top it already sat at.
  const vAlign = poVal(poFirstChild(tcPr, 'w:vAlign'));
  if (vAlign === 'center' || vAlign === 'bottom' || vAlign === 'top') out.verticalAlign = vAlign;
  // §17.4.20 `w:hideMark` — the end-of-cell mark does not count towards the
  // row's height. hidemark.docx's second row holds nothing but the marks, and
  // measured with them it stood a full line tall where Word and LibreOffice
  // draw a hairline strip.
  if (poFirstChild(tcPr, 'w:hideMark')) out.hideMark = true;

  return { properties: out, ...(rawVMerge ? { vMerge: rawVMerge } : {}) };
}

/**
 * A `w:tcBorders` / `w:tblBorders` / §17.6.10 `w:pgBorders` element: the rule on
 * each edge. Every one of them is spelled the same way.
 *
 * @param node The borders element, or `undefined`.
 * @returns The edges that name a rule, or `undefined` when none do.
 */
export function parseBorders(node: PoNode | undefined): CellBorders | undefined {
  if (!node) return undefined;
  const out: Mutable<CellBorders> = {};
  const top = parseBorder(poFirstChild(node, 'w:top'));
  const right = parseBorder(poFirstChild(node, 'w:right') ?? poFirstChild(node, 'w:end'));
  const bottom = parseBorder(poFirstChild(node, 'w:bottom'));
  const left = parseBorder(poFirstChild(node, 'w:left') ?? poFirstChild(node, 'w:start'));
  const insideH = parseBorder(poFirstChild(node, 'w:insideH'));
  const insideV = parseBorder(poFirstChild(node, 'w:insideV'));
  if (top) out.top = top;
  if (right) out.right = right;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  if (insideH) out.insideH = insideH;
  if (insideV) out.insideV = insideV;
  return Object.keys(out).length > 0 ? out : undefined;
}

function parseBorder(node: PoNode | undefined): Border | undefined {
  if (!node) return undefined;
  const raw = poVal(node);
  if (!raw) return undefined;
  // §17.18.2 names some hundred and eighty patterns and we draw six. `nil` and
  // `none` are a rule that is explicitly ABSENT — recorded, so it overrides the
  // one a table or style would otherwise lend the edge; all_gaps_word.docx
  // spells its grid away that way and dropping the edge boxed every cell.
  // Of the rest, the ones drawn as SEVERAL parallel lines read as a double
  // rule — fdo76586.docx boxes its table in `thickThinLargeGap` and a single
  // line is visibly one rule where the reference draws two. Anything else is
  // far closer to a solid rule of the stated width than to nothing at all.
  const val =
    raw === 'nil'
      ? 'none'
      : BORDER_STYLES.has(raw as BorderStyle)
        ? raw
        : MULTI_LINE_BORDERS.has(raw)
          ? 'double'
          : 'single';
  const sz = poIntAttr(node, 'sz');
  const color = poAttr(node, 'color');
  const out: Mutable<Border> = { style: val as BorderStyle };
  // `w:sz` on a compound pattern is the width of ONE of its lines; the rule as
  // a whole is that line, a gap and another line. A `double` here is drawn to
  // the TOTAL, so the three of them are what it is given.
  if (sz !== undefined) out.width = eighthPtToPt(MULTI_LINE_BORDERS.has(raw) ? sz * 3 : sz);
  // §17.3.1.24 `w:space` — the gap a paragraph rule keeps from the text, in
  // POINTS (a cell border has no such attribute and leaves it undefined).
  const space = poIntAttr(node, 'space');
  if (space !== undefined) out.spacePt = pt(space);
  if (color && color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(color)) {
    out.colorHex = color.toUpperCase();
  }
  return out;
}

function parseCellMargins(node: PoNode | undefined): CellMargins | undefined {
  if (!node) return undefined;
  const out: Mutable<CellMargins> = {};
  const top = poIntAttr(poFirstChild(node, 'w:top'), 'w');
  const bottom = poIntAttr(poFirstChild(node, 'w:bottom'), 'w');
  const left = poIntAttr(poFirstChild(node, 'w:left') ?? poFirstChild(node, 'w:start'), 'w');
  const right = poIntAttr(poFirstChild(node, 'w:right') ?? poFirstChild(node, 'w:end'), 'w');
  // §17.4.42 `w:tcMar` measures an INSET, and `w:w` is a signed type only
  // because ST_TblWidth is shared with the widths. A negative one is not a
  // narrower cell but a nonsense inset, and both references read it as none:
  // negative-cell-margin-twips.docx insets two of its three cells by −6160 and
  // −8800 twips, and taken at face value it drew their text 300pt off the left
  // edge of the page — the two cells came out empty.
  if (top !== undefined) out.top = twipsToPt(Math.max(0, top));
  if (bottom !== undefined) out.bottom = twipsToPt(Math.max(0, bottom));
  if (left !== undefined) out.left = twipsToPt(Math.max(0, left));
  if (right !== undefined) out.right = twipsToPt(Math.max(0, right));
  return Object.keys(out).length > 0 ? out : undefined;
}
