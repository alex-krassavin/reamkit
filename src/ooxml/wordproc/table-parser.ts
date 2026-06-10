// ECMA-376 Part 1 §17.4 — Table parser.
// Parses w:tbl in preserveOrder format into the typed Table model.

import type {
  Border,
  BorderStyle,
  CellBorders,
  CellMargins,
  CellProperties,
  RowProperties,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/document-model';

import type { PoNode } from '@/ooxml/wordproc/po-helpers';
import type { Pt } from '@/ir';
import { eighthPtToPt, twipsToPt } from '@/ir';
import { parseBodyElements } from '@/ooxml/wordproc/document-parser';
import {
  poAttr,
  poChildren,
  poChildrenWith,
  poFirstChild,
  poIntAttr,
  poToggle,
  poVal,
} from '@/ooxml/wordproc/po-helpers';

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

export function parseTable(tbl: PoNode): Table {
  const properties = parseTableProperties(poFirstChild(tbl, 'w:tblPr'));
  const grid = parseTableGrid(poFirstChild(tbl, 'w:tblGrid'));
  const rows: Array<TableRow> = [];
  for (const tr of poChildrenWith(tbl, 'w:tr')) {
    rows.push(parseTableRow(tr));
  }
  return { properties, grid, rows };
}

function parseTableProperties(tblPr: PoNode | undefined): TableProperties {
  if (!tblPr) return {};
  const out: Mutable<TableProperties> = {};

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

function parseTableGrid(tblGrid: PoNode | undefined): Array<Pt> {
  if (!tblGrid) return [];
  const cols: Array<Pt> = [];
  for (const gridCol of poChildrenWith(tblGrid, 'w:gridCol')) {
    const w = poIntAttr(gridCol, 'w');
    cols.push(twipsToPt(w ?? 0));
  }
  return cols;
}

function parseTableRow(tr: PoNode): TableRow {
  const properties = parseRowProperties(poFirstChild(tr, 'w:trPr'));
  const cells: Array<TableCell> = [];
  for (const tc of poChildrenWith(tr, 'w:tc')) {
    cells.push(parseTableCell(tc));
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

function parseTableCell(tc: PoNode): TableCell {
  const properties = parseCellProperties(poFirstChild(tc, 'w:tcPr'));
  const content = parseBodyElements(poChildren(tc));
  return { properties, content };
}

function parseCellProperties(tcPr: PoNode | undefined): CellProperties {
  if (!tcPr) return {};
  const out: Mutable<CellProperties> = {};
  const tcW = poFirstChild(tcPr, 'w:tcW');
  if (tcW) {
    const w = poIntAttr(tcW, 'w');
    if (w !== undefined) out.width = twipsToPt(w);
  }
  const gridSpan = poFirstChild(tcPr, 'w:gridSpan');
  if (gridSpan) {
    const v = poIntAttr(gridSpan, 'val');
    if (v !== undefined) out.gridSpan = v;
  }
  const vMerge = poFirstChild(tcPr, 'w:vMerge');
  if (vMerge) {
    const v = poVal(vMerge);
    if (v === 'restart') out.vMerge = 'restart';
    else out.vMerge = 'continue';
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
  return out;
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
