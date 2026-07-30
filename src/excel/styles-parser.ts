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
  Dxf,
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
import { applyColorMods } from '@/core/drawingml/colors';

import { INDEXED_COLORS } from '@/core/indexed-colors';

export { INDEXED_COLORS };

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
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

/** The empty style table, returned for a workbook with no (or a malformed) `styles.xml`. */
export const EMPTY_XLSX_STYLES: XlsxStyles = {
  numFmts: new Map(),
  fonts: [],
  fills: [],
  borders: [],
  cellXfs: [],
};

/** A workbook theme's colour scheme, slot name → RRGGBB (see `parseTheme`). */
export type ThemePalette = ReadonlyMap<string, string>;

/**
 * Parse `xl/styles.xml` (§18.8) into the {@link XlsxStyles} table the renderer
 * consumes: custom number formats, fonts, fills, borders, the cell formats
 * (`cellXfs` — each indexing font/fill/border/numFmt + optional alignment) and the
 * differential formats (`dxfs`) conditional formatting applies. Returns
 * {@link EMPTY_XLSX_STYLES} when the root is absent or malformed.
 *
 * @param theme The workbook's theme palette, resolving `<color theme="N">`.
 *              Without it those colours are dropped, and a sheet styled from the
 *              theme loses every fill and font colour it has.
 */
export function parseXlsxStyles(data: Uint8Array, theme?: ThemePalette): XlsxStyles {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const root = asObject(tree['styleSheet']);
  if (!root) return EMPTY_XLSX_STYLES;

  const dxfs = parseDxfs(root, theme);
  return {
    numFmts: parseNumFmts(root),
    fonts: parseFonts(root, theme),
    fills: parseFills(root, theme),
    borders: parseBorders(root, theme),
    cellXfs: parseCellXfs(root),
    ...(dxfs.length > 0 ? { dxfs } : {}),
  };
}

// §18.8.10 <dxfs> — differential formats a conditional-format rule applies on
// match. We read the font (bold/italic/color) and fill (solid highlight); a dxf
// fill's solid colour conventionally rides <bgColor> (Excel quirk).
function parseDxfs(root: Record<string, unknown>, theme: ThemePalette | undefined): Array<Dxf> {
  const node = asObject(root['dxfs']);
  if (!node) return [];
  const out: Array<Dxf> = [];
  for (const item of asArray(node['dxf'])) {
    const obj = asObject(item);
    if (!obj) {
      out.push({});
      continue;
    }
    const dxf: Mutable<Dxf> = {};
    const fontObj = asObject(obj['font']);
    if (fontObj) {
      const font: Mutable<XlsxFont> = {};
      if (hasChild(fontObj, 'b')) font.bold = childToggle(fontObj, 'b');
      if (hasChild(fontObj, 'i')) font.italic = childToggle(fontObj, 'i');
      const colorHex = colorOf(asObject(fontObj['color']), theme);
      if (colorHex) font.colorHex = colorHex;
      if (Object.keys(font).length > 0) dxf.font = font;
    }
    const pf = asObject(asObject(obj['fill'])?.['patternFill']);
    if (pf) {
      const fill: Mutable<XlsxFill> = {};
      const pt = strAttr(pf, 'patternType');
      if (pt) fill.patternType = pt;
      const fg = colorOf(asObject(pf['fgColor']), theme);
      const bg = colorOf(asObject(pf['bgColor']), theme);
      if (fg) fill.fgColorHex = fg;
      if (bg) fill.bgColorHex = bg;
      if (Object.keys(fill).length > 0) dxf.fill = fill;
    }
    // §18.8.9 — a dxf may format with nothing but an edge. Reading only font
    // and fill made such a rule a no-op: tdf171828.xlsx rules the boundary
    // under every year with one, and six of its twelve dxfs are border-only.
    // §18.8.9 — a dxf may carry a number format, which changes what the cell
    // SAYS and not just how it looks.
    const numFmtObj = asObject(obj['numFmt']);
    const code = numFmtObj ? strAttr(numFmtObj, 'formatCode') : undefined;
    if (code !== undefined && code.length > 0) dxf.numberFormat = code;
    const borderObj = asObject(obj['border']);
    if (borderObj) {
      const border: Mutable<XlsxBorder> = {};
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const edge = parseBorderEdge(asObject(borderObj[side]), theme);
        if (edge) border[side] = edge;
      }
      if (Object.keys(border).length > 0) dxf.border = border;
    }
    out.push(dxf);
  }
  return out;
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

function parseBorders(
  root: Record<string, unknown>,
  theme: ThemePalette | undefined,
): Array<XlsxBorder> {
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
    const top = parseBorderEdge(asObject(obj['top']), theme);
    const right = parseBorderEdge(asObject(obj['right']), theme);
    const bottom = parseBorderEdge(asObject(obj['bottom']), theme);
    const left = parseBorderEdge(asObject(obj['left']), theme);
    if (top) border.top = top;
    if (right) border.right = right;
    if (bottom) border.bottom = bottom;
    if (left) border.left = left;
    // §18.8.4 diagonal stroke + which corners it spans (E-SHEET W6).
    const diagonal = parseBorderEdge(asObject(obj['diagonal']), theme);
    if (diagonal) {
      border.diagonal = diagonal;
      if (boolAttr(obj, 'diagonalUp')) border.diagonalUp = true;
      if (boolAttr(obj, 'diagonalDown')) border.diagonalDown = true;
    }
    out.push(border);
  }
  return out;
}

function parseBorderEdge(
  node: Record<string, unknown> | undefined,
  theme: ThemePalette | undefined,
): XlsxBorderEdge | undefined {
  if (!node) return undefined;
  const styleRaw = strAttr(node, 'style');
  if (!styleRaw) return undefined;
  if (!VALID_BORDER_STYLES.has(styleRaw as XlsxBorderStyleName)) return undefined;
  const colorHex = colorOf(asObject(node['color']), theme);
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

function parseFonts(
  root: Record<string, unknown>,
  theme: ThemePalette | undefined,
): Array<XlsxFont> {
  const node = asObject(root['fonts']);
  if (!node) return [];
  const out: Array<XlsxFont> = [];
  for (const item of asArray(node['font'])) {
    // An empty element (`<font/>`, `<fill/>`, `<xf/>`) parses to a string, not
    // an object — but it still OCCUPIES ITS INDEX. Skipping it shifts every
    // later entry down one, and the ids in cellXfs then point at the wrong
    // record: tdf122336.xlsx writes `<font/><font><b/></font>`, so its bold
    // font landed at 0, `fontId="1"` resolved to nothing, and the header row
    // came out in the body weight.
    const obj = asObject(item);
    if (!obj) {
      out.push({});
      continue;
    }
    const font: Mutable<XlsxFont> = {};
    const sz = childValAttr(obj, 'sz');
    if (sz !== undefined) {
      const n = Number(sz);
      if (Number.isFinite(n)) font.sizePt = n;
    }
    if (hasChild(obj, 'b')) font.bold = childToggle(obj, 'b');
    if (hasChild(obj, 'i')) font.italic = childToggle(obj, 'i');
    if (hasChild(obj, 'u')) font.underline = childToggle(obj, 'u');
    const colorRgb = colorOf(asObject(obj['color']), theme);
    if (colorRgb) font.colorHex = colorRgb;
    const nameVal = childValAttr(obj, 'name');
    if (nameVal) font.name = nameVal;
    out.push(font);
  }
  return out;
}

function parseFills(
  root: Record<string, unknown>,
  theme: ThemePalette | undefined,
): Array<XlsxFill> {
  const node = asObject(root['fills']);
  if (!node) return [];
  const out: Array<XlsxFill> = [];
  for (const item of asArray(node['fill'])) {
    const obj = asObject(item);
    if (!obj) {
      out.push({});
      continue;
    }
    const fill: Mutable<XlsxFill> = {};
    const pf = asObject(obj['patternFill']);
    if (pf) {
      const pt = strAttr(pf, 'patternType');
      if (pt) fill.patternType = pt;
      const fg = colorOf(asObject(pf['fgColor']), theme);
      const bg = colorOf(asObject(pf['bgColor']), theme);
      if (fg) fill.fgColorHex = fg;
      if (bg) fill.bgColorHex = bg;
    }
    // §18.8.24 <gradientFill> (E-SHEET W6): the print model has no gradient, so
    // summarise the stops to their average colour and carry it as a solid fill —
    // the cell gets its intended background; on write-back it round-trips as solid.
    if (!fill.patternType) {
      const avg = averageGradientColor(asObject(obj['gradientFill']), theme);
      if (avg) {
        fill.patternType = 'solid';
        fill.fgColorHex = avg;
      }
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
    if (!obj) {
      out.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
      continue;
    }
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
      const indent = numAttr(align, 'indent');
      if (indent !== undefined && indent > 0) a.indent = indent;
      const rotation = numAttr(align, 'textRotation');
      if (rotation !== undefined && rotation !== 0) a.textRotation = rotation;
      const shrink = boolAttr(align, 'shrinkToFit');
      if (shrink) a.shrinkToFit = true;
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

// The mean of a gradientFill's stop colours (E-SHEET W6) — a representative solid.
function averageGradientColor(
  gf: Record<string, unknown> | undefined,
  theme: ThemePalette | undefined,
): string | undefined {
  if (!gf) return undefined;
  let r = 0;
  let g = 0;
  let b = 0;
  let n = 0;
  for (const s of asArray(gf['stop'])) {
    const so = asObject(s);
    const hex = so ? colorOf(asObject(so['color']), theme) : undefined;
    if (!hex) continue;
    r += parseInt(hex.slice(0, 2), 16);
    g += parseInt(hex.slice(2, 4), 16);
    b += parseInt(hex.slice(4, 6), 16);
    n++;
  }
  if (n === 0) return undefined;
  const h = (x: number): string =>
    Math.round(x / n)
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `${h(r)}${h(g)}${h(b)}`;
}

function colorOf(
  node: Record<string, unknown> | undefined,
  theme?: ThemePalette,
): string | undefined {
  if (!node) return undefined;
  const rgb = strAttr(node, 'rgb');
  if (rgb) {
    // Excel stores ARGB; strip leading alpha if 8 hex digits
    if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return rgb.substring(2).toUpperCase();
    if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return rgb.toUpperCase();
  }
  const indexed = strAttr(node, 'indexed');
  if (indexed !== undefined) {
    const i = Number(indexed);
    if (Number.isInteger(i) && i >= 0 && i < INDEXED_COLORS.length) return INDEXED_COLORS[i];
  }
  const themeIdx = strAttr(node, 'theme');
  if (themeIdx !== undefined && theme) {
    const slot = THEME_SLOTS[Number(themeIdx)];
    const base = slot ? theme.get(slot) : undefined;
    if (base) return applyTint(base, Number(strAttr(node, 'tint') ?? '0'));
  }
  return undefined;
}

/**
 * ECMA-376 §18.8.3 `<color theme>` — the index into the workbook theme's colour
 * scheme. Excel's order is NOT the order the theme part declares them in: the
 * first two slots are swapped, so `theme="0"` is the light background and
 * `theme="1"` the dark text, while `theme1.xml` writes `dk1` before `lt1`.
 */
const THEME_SLOTS: ReadonlyArray<string> = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/**
 * §18.8.19 `tint` — lighten (positive) or darken (negative) a theme colour by
 * scaling its HSL luminance. Excel writes it on nearly every theme colour it
 * uses, so ignoring it is not much better than ignoring the colour: a `theme="2"
 * tint="-0.5"` band is half as light as its slot.
 */
function applyTint(hex: string, tint: number): string {
  if (!Number.isFinite(tint) || tint === 0) return hex;
  // The same luminance arithmetic DrawingML spells as lumMod/lumOff.
  return tint < 0
    ? applyColorMods(hex, [{ kind: 'lumMod', val: 1 + tint }])
    : applyColorMods(hex, [
        { kind: 'lumMod', val: 1 - tint },
        { kind: 'lumOff', val: tint },
      ]);
}
