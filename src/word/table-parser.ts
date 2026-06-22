// ECMA-376 Part 1 §17.4 — Table parser.
// Parses w:tbl in preserveOrder format into the typed Table model.

import type {
  Border,
  BorderStyle,
  CellBorders,
  CellMargins,
  CellMerge,
  CellProperties,
  RowProperties,
  Table,
  TableCell,
  TableLook,
  TableProperties,
  TableRow,
} from '@/core/document-model';

import type { PoNode } from '@/core/po-helpers';
import type { Pt } from '@/core/ir';
import type { ParseContext } from '@/word/document-parser';
import { DEFAULT_PARSE_CONTEXT, parseBodyElements } from '@/word/document-parser';
import { eighthPtToPt, twipsToPt } from '@/core/ir';
import {
  poAttr,
  poChildren,
  poChildrenWith,
  poFirstChild,
  poIntAttr,
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
]);

const WIDTH_TYPES = new Set<'auto' | 'dxa' | 'pct' | 'nil'>(['auto', 'dxa', 'pct', 'nil']);
const HEIGHT_RULES = new Set<'auto' | 'atLeast' | 'exact'>(['auto', 'atLeast', 'exact']);

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
  for (const tr of poChildrenWith(tbl, 'w:tr')) {
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
  return { properties, grid, rows };
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

  const tblLayout = poFirstChild(tblPr, 'w:tblLayout');
  if (tblLayout) {
    const t = poAttr(tblLayout, 'type');
    if (t === 'fixed' || t === 'auto') out.layout = t;
  }

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

function parseTableRow(
  tr: PoNode,
  ctx: ParseContext,
): { properties: RowProperties; cells: Array<DraftCell> } {
  const properties = parseRowProperties(poFirstChild(tr, 'w:trPr'));
  const cells: Array<DraftCell> = [];
  for (const tc of poChildrenWith(tr, 'w:tc')) {
    cells.push(parseTableCell(tc, ctx));
  }
  return { properties, cells };
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
    }
  }
  if (poFirstChild(trPr, 'w:cantSplit'))
    out.cantSplit = poToggle(poFirstChild(trPr, 'w:cantSplit')) ?? true;
  if (poFirstChild(trPr, 'w:tblHeader')) {
    out.isHeader = poToggle(poFirstChild(trPr, 'w:tblHeader')) ?? true;
  }
  return out;
}

function parseTableCell(tc: PoNode, ctx: ParseContext): DraftCell {
  const { properties, vMerge } = parseCellProperties(poFirstChild(tc, 'w:tcPr'));
  const content = parseBodyElements(poChildren(tc), ctx);
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
    if (w !== undefined) out.width = twipsToPt(w);
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
  return { properties: out, ...(rawVMerge ? { vMerge: rawVMerge } : {}) };
}

function parseBorders(node: PoNode | undefined): CellBorders | undefined {
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
  const val = poVal(node);
  if (!val || !BORDER_STYLES.has(val as BorderStyle)) return undefined;
  const sz = poIntAttr(node, 'sz');
  const color = poAttr(node, 'color');
  const out: Mutable<Border> = { style: val as BorderStyle };
  if (sz !== undefined) out.width = eighthPtToPt(sz);
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
  if (top !== undefined) out.top = twipsToPt(top);
  if (bottom !== undefined) out.bottom = twipsToPt(bottom);
  if (left !== undefined) out.left = twipsToPt(left);
  if (right !== undefined) out.right = twipsToPt(right);
  return Object.keys(out).length > 0 ? out : undefined;
}
