// ECMA-376 Part 1 §17.7 — Styles.xml parser.
//
// Produces a StyleSheet of document defaults + named styles. Inheritance
// (basedOn) is recorded but not resolved here — that happens in the
// style-cascade module so callers can decide which paragraph/run a style
// applies to.

import { XMLParser } from 'fast-xml-parser';

import type {
  Border,
  BorderStyle,
  CellBorders,
  CellMargins,
  CellShading,
  Style,
  StyleSheet,
  StyleType,
  TableStyleCondition,
  TableStyleConditionType,
  TableStyleLayer,
} from '@/core/document-model';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';
import { eighthPtToPt, twipsToPt } from '@/core/ir';

import { parseParagraphProperties } from '@/word/paragraph-properties';
import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr, getVal } from '@/word/xml-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const STYLE_TYPES = new Set<StyleType>(['paragraph', 'character', 'table', 'numbering']);

/**
 * Parse `word/styles.xml` (ECMA-376 Part 1 §17.7) into a {@link StyleSheet}: the
 * document defaults (`w:docDefaults` run/paragraph properties) plus the named
 * styles, keyed by `styleId`. Table styles additionally carry their base layer,
 * conditional (`w:tblStylePr`) layers and banding sizes. `basedOn` inheritance is
 * recorded but not resolved — that happens in the style-cascade module.
 *
 * @param stylesXml The raw `styles.xml` bytes.
 * @returns The parsed stylesheet, or `EMPTY_STYLE_SHEET` when the root is absent.
 */
export function parseStyles(stylesXml: Uint8Array): StyleSheet {
  const xml = decoder.decode(stylesXml);
  const tree = parser.parse(xml) as Record<string, unknown>;

  const root = asElement(tree['w:styles']);
  if (!root) return EMPTY_STYLE_SHEET;

  const docDefaults = asElement(root['w:docDefaults']);
  let defaultRunProperties = {};
  let defaultParagraphProperties = {};
  if (docDefaults) {
    const rPrDef = asElement(docDefaults['w:rPrDefault']);
    if (rPrDef) defaultRunProperties = parseRunProperties(rPrDef['w:rPr']);
    const pPrDef = asElement(docDefaults['w:pPrDefault']);
    if (pPrDef) defaultParagraphProperties = parseParagraphProperties(pPrDef['w:pPr']);
  }

  const styles = new Map<string, Style>();
  for (const s of asArray(root['w:style'])) {
    const el = asElement(s);
    if (!el) continue;
    const id = getAttr(el, 'styleId');
    const typeAttr = getAttr(el, 'type');
    if (!id || !typeAttr) continue;
    if (!STYLE_TYPES.has(typeAttr as StyleType)) continue;
    const type = typeAttr as StyleType;

    const basedOnVal = getVal(el['w:basedOn']);
    const isDefault = getAttr(el, 'default') === '1';

    const style: Style = {
      id,
      type,
      ...(basedOnVal ? { basedOn: basedOnVal } : {}),
      isDefault,
      runProperties: parseRunProperties(el['w:rPr']),
      paragraphProperties: parseParagraphProperties(el['w:pPr']),
      ...(type === 'table' ? parseTableStyleParts(el) : {}),
    };
    styles.set(id, style);
  }

  return { defaultRunProperties, defaultParagraphProperties, styles };
}

// ---------------------------------------------------------------------------
// §17.7.6 table styles. This file works on fast-xml-parser's FLAT tree (the
// table parser's border/margin readers work on PoNodes — the same flat/po
// duality rPr/pPr already live with), so the §17.4 border/shading/margin
// grammar is read here against the flat shape.
// ---------------------------------------------------------------------------

const CONDITION_TYPES = new Set<TableStyleConditionType>([
  'wholeTable',
  'band1Vert',
  'band2Vert',
  'band1Horz',
  'band2Horz',
  'firstCol',
  'lastCol',
  'firstRow',
  'lastRow',
  'nwCell',
  'neCell',
  'swCell',
  'seCell',
]);

function parseTableStyleParts(el: Record<string, unknown>): Partial<Style> {
  const out: {
    tableLayer?: TableStyleLayer;
    tableConditions?: Array<TableStyleCondition>;
    rowBandSize?: number;
    colBandSize?: number;
  } = {};

  const base = parseTableStyleLayer(el);
  if (base) out.tableLayer = base;

  const tblPr = asElement(el['w:tblPr']);
  if (tblPr) {
    const rowBand = intVal(tblPr['w:tblStyleRowBandSize']);
    const colBand = intVal(tblPr['w:tblStyleColBandSize']);
    if (rowBand !== undefined && rowBand > 0) out.rowBandSize = rowBand;
    if (colBand !== undefined && colBand > 0) out.colBandSize = colBand;
  }

  const conditions: Array<TableStyleCondition> = [];
  for (const c of asArray(el['w:tblStylePr'])) {
    const cond = asElement(c);
    if (!cond) continue;
    const typeAttr = getAttr(cond, 'type');
    if (!typeAttr || !CONDITION_TYPES.has(typeAttr as TableStyleConditionType)) continue;
    const layer = parseTableStyleLayer(cond);
    if (layer) conditions.push({ type: typeAttr as TableStyleConditionType, layer });
  }
  if (conditions.length > 0) out.tableConditions = conditions;
  return out;
}

// One layer = the element's tblPr (table borders + cell margins) + tcPr
// (region cell borders + shading) + rPr/pPr. tblBorders wins over tcBorders
// when both appear (the base layer's table grid vs a region's own edges).
function parseTableStyleLayer(el: Record<string, unknown>): TableStyleLayer | undefined {
  const tblPr = asElement(el['w:tblPr']);
  const tcPr = asElement(el['w:tcPr']);

  const borders =
    parseFlatBorders(tblPr ? asElement(tblPr['w:tblBorders']) : undefined) ??
    parseFlatBorders(tcPr ? asElement(tcPr['w:tcBorders']) : undefined);
  const cellMargins = parseFlatCellMargins(tblPr ? asElement(tblPr['w:tblCellMar']) : undefined);
  const shading = parseFlatShading(tcPr ? tcPr['w:shd'] : undefined);
  const runProperties = parseRunProperties(el['w:rPr']);
  const paragraphProperties = parseParagraphProperties(el['w:pPr']);

  const layer: TableStyleLayer = {
    ...(borders ? { borders } : {}),
    ...(cellMargins ? { cellMargins } : {}),
    ...(shading ? { shading } : {}),
    ...(Object.keys(runProperties).length > 0 ? { runProperties } : {}),
    ...(Object.keys(paragraphProperties).length > 0 ? { paragraphProperties } : {}),
  };
  return Object.keys(layer).length > 0 ? layer : undefined;
}

const FLAT_BORDER_STYLES = new Set<BorderStyle>([
  'none',
  'single',
  'double',
  'thick',
  'dotted',
  'dashed',
]);

function parseFlatBorder(node: unknown): Border | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const val = getAttr(el, 'val');
  if (!val || !FLAT_BORDER_STYLES.has(val as BorderStyle)) return undefined;
  const szRaw = getAttr(el, 'sz');
  const sz = szRaw !== undefined ? Number(szRaw) : NaN;
  const color = getAttr(el, 'color');
  return {
    style: val as BorderStyle,
    ...(Number.isFinite(sz) ? { width: eighthPtToPt(sz) } : {}),
    ...(color && color !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(color)
      ? { colorHex: color.toUpperCase() }
      : {}),
  };
}

function parseFlatBorders(el: Record<string, unknown> | undefined): CellBorders | undefined {
  if (!el) return undefined;
  const top = parseFlatBorder(el['w:top']);
  const right = parseFlatBorder(el['w:right'] ?? el['w:end']);
  const bottom = parseFlatBorder(el['w:bottom']);
  const left = parseFlatBorder(el['w:left'] ?? el['w:start']);
  const insideH = parseFlatBorder(el['w:insideH']);
  const insideV = parseFlatBorder(el['w:insideV']);
  const out: CellBorders = {
    ...(top ? { top } : {}),
    ...(right ? { right } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
    ...(insideH ? { insideH } : {}),
    ...(insideV ? { insideV } : {}),
  };
  return Object.keys(out).length > 0 ? out : undefined;
}

function parseFlatCellMargins(el: Record<string, unknown> | undefined): CellMargins | undefined {
  if (!el) return undefined;
  const read = (node: unknown): number | undefined => {
    const m = asElement(node);
    if (!m) return undefined;
    const w = getAttr(m, 'w');
    const n = w !== undefined ? Number(w) : NaN;
    return Number.isFinite(n) ? n : undefined;
  };
  const top = read(el['w:top']);
  const bottom = read(el['w:bottom']);
  const left = read(el['w:left'] ?? el['w:start']);
  const right = read(el['w:right'] ?? el['w:end']);
  const out: CellMargins = {
    ...(top !== undefined ? { top: twipsToPt(top) } : {}),
    ...(bottom !== undefined ? { bottom: twipsToPt(bottom) } : {}),
    ...(left !== undefined ? { left: twipsToPt(left) } : {}),
    ...(right !== undefined ? { right: twipsToPt(right) } : {}),
  };
  return Object.keys(out).length > 0 ? out : undefined;
}

// §17.4.33 w:shd — direct @w:fill hex only (same policy as the table parser).
function parseFlatShading(node: unknown): CellShading | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const fill = getAttr(el, 'fill');
  if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(fill)) {
    return { colorHex: fill.toUpperCase() };
  }
  return undefined;
}

function intVal(node: unknown): number | undefined {
  const v = getVal(node);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}
