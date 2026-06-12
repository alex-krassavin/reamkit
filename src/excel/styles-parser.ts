// ECMA-376 Part 1 §18.8 — xl/styles.xml.
//
// xlsx separates style attributes (fonts, fills, borders, numFmts) from
// "cell formats" (cellXfs). A cell carries an s="N" attribute that indexes
// into cellXfs. Each cellXf references font/fill/border/numFmt indices and
// optionally bundles its own alignment.
//
// We extract only what the renderer currently uses:
//   - fonts: size, bold, italic, color, name
//   - fills: solid foreground color (gray125/etc. ignored)
//   - cellXfs: numFmtId, fontId, fillId, alignment, applyXxx flags
//   - numFmts: custom format codes keyed by numFmtId
// Borders are deferred — the table renderer uses table-level borders.

import { XMLParser } from 'fast-xml-parser';

import type {
  XlsxBorder,
  XlsxBorderEdge,
  XlsxBorderStyleName,
  XlsxCellAlignment,
  XlsxCellXf,
  XlsxFill,
  XlsxFont,
  XlsxHorizontalAlign,
  XlsxStyles,
  XlsxVerticalAlign,
} from '@/core/spreadsheet-model';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // Tolerate an explicit `x:` namespace prefix (<x:styleSheet>, <x:fonts>, …)
  // used by some producers — see workbook-parser.ts.
  removeNSPrefix: true,
});

// The font/fill/border/xf/styles model types now live in
// @/core/spreadsheet-model; this parser imports them above and produces them.

export const EMPTY_XLSX_STYLES: XlsxStyles = {
  numFmts: new Map(),
  fonts: [],
  fills: [],
  borders: [],
  cellXfs: [],
};

export function parseXlsxStyles(data: Uint8Array): XlsxStyles {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const root = asObject(tree['styleSheet']);
  if (!root) return EMPTY_XLSX_STYLES;

  return {
    numFmts: parseNumFmts(root),
    fonts: parseFonts(root),
    fills: parseFills(root),
    borders: parseBorders(root),
    cellXfs: parseCellXfs(root),
  };
}

const VALID_BORDER_STYLES: ReadonlySet<XlsxBorderStyleName> = new Set([
  'none',
  'thin',
  'medium',
  'thick',
  'hair',
  'dashed',
  'dotted',
  'double',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
]);

function parseBorders(root: Record<string, unknown>): Array<XlsxBorder> {
  const node = asObject(root['borders']);
  if (!node) return [];
  const out: Array<XlsxBorder> = [];
  for (const item of asArray(node['border'])) {
    const obj = asObject(item);
    if (!obj) {
      out.push({});
      continue;
    }
    const border: Mutable<XlsxBorder> = {};
    const top = parseBorderEdge(asObject(obj['top']));
    const right = parseBorderEdge(asObject(obj['right']));
    const bottom = parseBorderEdge(asObject(obj['bottom']));
    const left = parseBorderEdge(asObject(obj['left']));
    if (top) border.top = top;
    if (right) border.right = right;
    if (bottom) border.bottom = bottom;
    if (left) border.left = left;
    out.push(border);
  }
  return out;
}

function parseBorderEdge(node: Record<string, unknown> | undefined): XlsxBorderEdge | undefined {
  if (!node) return undefined;
  const styleRaw = strAttr(node, 'style');
  if (!styleRaw) return undefined;
  if (!VALID_BORDER_STYLES.has(styleRaw as XlsxBorderStyleName)) return undefined;
  const colorHex = colorOf(asObject(node['color']));
  const edge: Mutable<XlsxBorderEdge> = { style: styleRaw as XlsxBorderStyleName };
  if (colorHex) edge.colorHex = colorHex;
  return edge;
}

function parseNumFmts(root: Record<string, unknown>): Map<number, string> {
  const result = new Map<number, string>();
  const node = asObject(root['numFmts']);
  if (!node) return result;
  for (const item of asArray(node['numFmt'])) {
    const obj = asObject(item);
    if (!obj) continue;
    const id = numAttr(obj, 'numFmtId');
    const code = strAttr(obj, 'formatCode');
    if (id !== undefined && code !== undefined) result.set(id, code);
  }
  return result;
}

function parseFonts(root: Record<string, unknown>): Array<XlsxFont> {
  const node = asObject(root['fonts']);
  if (!node) return [];
  const out: Array<XlsxFont> = [];
  for (const item of asArray(node['font'])) {
    const obj = asObject(item);
    if (!obj) continue;
    const font: Mutable<XlsxFont> = {};
    const sz = childValAttr(obj, 'sz');
    if (sz !== undefined) {
      const n = Number(sz);
      if (Number.isFinite(n)) font.sizePt = n;
    }
    if (hasChild(obj, 'b')) font.bold = childToggle(obj, 'b');
    if (hasChild(obj, 'i')) font.italic = childToggle(obj, 'i');
    if (hasChild(obj, 'u')) font.underline = childToggle(obj, 'u');
    const colorRgb = colorOf(asObject(obj['color']));
    if (colorRgb) font.colorHex = colorRgb;
    const nameVal = childValAttr(obj, 'name');
    if (nameVal) font.name = nameVal;
    out.push(font);
  }
  return out;
}

function parseFills(root: Record<string, unknown>): Array<XlsxFill> {
  const node = asObject(root['fills']);
  if (!node) return [];
  const out: Array<XlsxFill> = [];
  for (const item of asArray(node['fill'])) {
    const obj = asObject(item);
    if (!obj) continue;
    const fill: Mutable<XlsxFill> = {};
    const pf = asObject(obj['patternFill']);
    if (pf) {
      const pt = strAttr(pf, 'patternType');
      if (pt) fill.patternType = pt;
      const fg = colorOf(asObject(pf['fgColor']));
      const bg = colorOf(asObject(pf['bgColor']));
      if (fg) fill.fgColorHex = fg;
      if (bg) fill.bgColorHex = bg;
    }
    out.push(fill);
  }
  return out;
}

function parseCellXfs(root: Record<string, unknown>): Array<XlsxCellXf> {
  const node = asObject(root['cellXfs']);
  if (!node) return [];
  const out: Array<XlsxCellXf> = [];
  for (const item of asArray(node['xf'])) {
    const obj = asObject(item);
    if (!obj) continue;
    const numFmtId = numAttr(obj, 'numFmtId') ?? 0;
    const fontId = numAttr(obj, 'fontId') ?? 0;
    const fillId = numAttr(obj, 'fillId') ?? 0;
    const borderId = numAttr(obj, 'borderId') ?? 0;
    const xf: Mutable<XlsxCellXf> = { numFmtId, fontId, fillId, borderId };
    const applyNumberFormat = boolAttr(obj, 'applyNumberFormat');
    const applyFont = boolAttr(obj, 'applyFont');
    const applyFill = boolAttr(obj, 'applyFill');
    const applyBorder = boolAttr(obj, 'applyBorder');
    const applyAlignment = boolAttr(obj, 'applyAlignment');
    if (applyNumberFormat !== undefined) xf.applyNumberFormat = applyNumberFormat;
    if (applyFont !== undefined) xf.applyFont = applyFont;
    if (applyFill !== undefined) xf.applyFill = applyFill;
    if (applyBorder !== undefined) xf.applyBorder = applyBorder;
    if (applyAlignment !== undefined) xf.applyAlignment = applyAlignment;
    const align = asObject(obj['alignment']);
    if (align) {
      const a: Mutable<XlsxCellAlignment> = {};
      const h = strAttr(align, 'horizontal');
      const v = strAttr(align, 'vertical');
      const wrap = boolAttr(align, 'wrapText');
      if (h) a.horizontal = h as XlsxHorizontalAlign;
      if (v) a.vertical = v as XlsxVerticalAlign;
      if (wrap !== undefined) a.wrapText = wrap;
      if (Object.keys(a).length > 0) xf.alignment = a;
    }
    out.push(xf);
  }
  return out;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

function asObject(v: unknown): Record<string, unknown> | undefined {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return undefined;
  return v as Record<string, unknown>;
}

function asArray(v: unknown): ReadonlyArray<unknown> {
  if (v === null || v === undefined) return [];
  return Array.isArray(v) ? v : [v];
}

function strAttr(obj: Record<string, unknown>, name: string): string | undefined {
  const v = obj[`@_${name}`];
  return typeof v === 'string' ? v : undefined;
}

function numAttr(obj: Record<string, unknown>, name: string): number | undefined {
  const v = strAttr(obj, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function boolAttr(obj: Record<string, unknown>, name: string): boolean | undefined {
  const v = strAttr(obj, name);
  if (v === undefined) return undefined;
  return v === '1' || v === 'true';
}

function childValAttr(obj: Record<string, unknown>, childName: string): string | undefined {
  const child = asObject(obj[childName]);
  if (!child) return undefined;
  return strAttr(child, 'val');
}

function hasChild(obj: Record<string, unknown>, childName: string): boolean {
  return childName in obj;
}

function childToggle(obj: Record<string, unknown>, childName: string): boolean {
  const child = obj[childName];
  if (child === '' || child === null || child === undefined) return true;
  if (typeof child !== 'object') return true;
  const val = strAttr(child as Record<string, unknown>, 'val');
  if (val === undefined) return true;
  return !(val === 'false' || val === '0');
}

function colorOf(node: Record<string, unknown> | undefined): string | undefined {
  if (!node) return undefined;
  const rgb = strAttr(node, 'rgb');
  if (rgb) {
    // Excel stores ARGB; strip leading alpha if 8 hex digits
    if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return rgb.substring(2).toUpperCase();
    if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return rgb.toUpperCase();
  }
  return undefined;
}
