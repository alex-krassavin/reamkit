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
import type { ThemeFonts } from '@/core/drawingml/theme-parser';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';
import { eighthPtToPt, twipsToPt } from '@/core/ir';

import { parseParagraphProperties } from '@/word/paragraph-properties';
import { parseRunProperties } from '@/word/run-properties';
import { shadingFillHex } from '@/word/shading';
import { cellMarginTwips, readTableMeasure } from '@/word/table-width';
import { asArray, asElement, getAttr, getVal } from '@/word/xml-helpers';

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
export function parseStyles(stylesXml: Uint8Array, themeFonts?: ThemeFonts): StyleSheet {
  const xml = decoder.decode(stylesXml);
  const tree = parser.parse(xml) as Record<string, unknown>;

  const root = asElement(tree['w:styles']);
  if (!root) return EMPTY_STYLE_SHEET;

  const docDefaults = asElement(root['w:docDefaults']);
  let defaultRunProperties = {};
  let defaultParagraphProperties = {};
  if (docDefaults) {
    const rPrDef = asElement(docDefaults['w:rPrDefault']);
    if (rPrDef) defaultRunProperties = parseRunProperties(rPrDef['w:rPr'], themeFonts);
    const pPrDef = asElement(docDefaults['w:pPrDefault']);
    if (pPrDef)
      defaultParagraphProperties = parseParagraphProperties(pPrDef['w:pPr'], 0, themeFonts);
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
    const styleName = getVal(el['w:name']);
    const builtIn = type === 'paragraph' ? builtInHeadingLevel(styleName) : undefined;
    const paragraphProperties = parseParagraphProperties(el['w:pPr'], 0, themeFonts);

    const style: Style = {
      id,
      type,
      ...(basedOnVal ? { basedOn: basedOnVal } : {}),
      isDefault,
      ...(styleName !== undefined ? { name: styleName } : {}),
      ...(builtIn !== undefined ? { builtInOutlineLevel: builtIn } : {}),
      runProperties: parseRunProperties(el['w:rPr'], themeFonts),
      // The built-in heading's level is the style's own, so the styles based
      // on it inherit it as they inherit everything else.
      paragraphProperties:
        builtIn !== undefined
          ? { ...paragraphProperties, outlineLevel: builtIn }
          : paragraphProperties,
      ...(type === 'table' ? parseTableStyleParts(el, themeFonts) : {}),
    };
    styles.set(id, style);
  }

  return { defaultRunProperties, defaultParagraphProperties, styles };
}

/**
 * The outline level (0–8) Word's built-in heading style of this primary name
 * stands at, or `undefined` for any other style.
 *
 * Word knows its built-in styles by NAME — §17.7.4.9 `w:name`, written in
 * English ("heading 1") in every locale, while the id is free (a Russian Word
 * writes `w:styleId="1"`). Measured in Word 16 for Mac on a probe
 * (`outline level of paragraph format`, 2026-11-20): "heading 4" with no
 * `w:outlineLvl` of its own is level 4; "heading 5" under the id `X5` is level
 * 5; a style named "Fancy" under the id `Heading6` is body text.
 *
 * @param name The style's `w:name`, if it has one.
 * @returns The 0-based outline level, or `undefined` when the name is no
 *          built-in heading's.
 */
function builtInHeadingLevel(name: string | undefined): number | undefined {
  const m = name !== undefined ? /^heading ([1-9])$/i.exec(name.trim()) : null;
  return m ? Number(m[1]) - 1 : undefined;
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

function parseTableStyleParts(
  el: Record<string, unknown>,
  themeFonts?: ThemeFonts,
): Partial<Style> {
  const out: {
    tableLayer?: TableStyleLayer;
    tableConditions?: Array<TableStyleCondition>;
    rowBandSize?: number;
    colBandSize?: number;
  } = {};

  const base = parseTableStyleLayer(el, themeFonts);
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
    const layer = parseTableStyleLayer(cond, themeFonts, true);
    if (layer) conditions.push({ type: typeAttr as TableStyleConditionType, layer });
  }
  if (conditions.length > 0) out.tableConditions = conditions;
  return out;
}

// One layer = the element's tblPr (table borders + cell margins) + tcPr
// (region cell borders + shading + cell margins) + rPr/pPr. tblBorders wins
// over tcBorders when both appear (the base layer's table grid vs a region's
// own edges). `region` marks a `w:tblStylePr`.
function parseTableStyleLayer(
  el: Record<string, unknown>,
  themeFonts?: ThemeFonts,
  region = false,
): TableStyleLayer | undefined {
  const tblPr = asElement(el['w:tblPr']);
  const tcPr = asElement(el['w:tcPr']);

  const borders =
    parseFlatBorders(tblPr ? asElement(tblPr['w:tblBorders']) : undefined) ??
    parseFlatBorders(tcPr ? asElement(tcPr['w:tcBorders']) : undefined);
  const cellMargins = parseFlatCellMargins(
    tblPr ? asElement(tblPr['w:tblCellMar']) : undefined,
    true,
  );
  // §17.4.42 — a style's `w:tcMar` is honoured too, and as the CELL's own
  // margin rather than the table's default: Word 16.114 insets a style's
  // `w:tcPr/w:tcMar` of 1080 by 54pt even under a table that states its own
  // `w:tblCellMar` of 720 (style-cellmar probes t14, t20, t13, t16).
  const cellLevelMargins = parseFlatCellMargins(
    tcPr ? asElement(tcPr['w:tcMar']) : undefined,
    false,
  );
  // §17.4.65 `w:tblInd` — the table indent a whole-table layer declares, read
  // as the table's own is (table-parser.ts): only an absolute measure indents.
  const tblInd = tblPr ? asElement(tblPr['w:tblInd']) : undefined;
  const indType = tblInd ? getAttr(tblInd, 'type') : undefined;
  const ind = tblInd ? readTableMeasure(getAttr(tblInd, 'w'), indType) : undefined;
  const indentPt =
    ind?.kind === 'dxa' && (ind.spelled || indType === undefined || indType === 'dxa')
      ? twipsToPt(ind.twips)
      : undefined;
  const shading = parseFlatShading(tcPr ? tcPr['w:shd'] : undefined);
  const runProperties = parseRunProperties(el['w:rPr'], themeFonts);
  const paragraphProperties = parseParagraphProperties(el['w:pPr'], 0, themeFonts);

  const layer: TableStyleLayer = {
    ...(borders ? { borders } : {}),
    ...(cellMargins ? { cellMargins } : {}),
    ...(cellLevelMargins ? { cellLevelMargins } : {}),
    // A region that states a tblPr or tcPr at all, margins or not, takes the
    // place of its namesake's margins up the basedOn chain. A derived style's
    // firstRow with nothing but a `w:shd` in its tcPr lost the 54pt its
    // parent's firstRow `w:tcMar` gave, one with only an `w:rPr` kept them
    // (style-cellmar probes k03, q02, q03, k01, k04).
    ...(region && (tblPr || tcPr) ? { restatesMargins: true } : {}),
    ...(indentPt !== undefined ? { indentPt } : {}),
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
  'dashSmallGap',
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

function parseFlatCellMargins(
  el: Record<string, unknown> | undefined,
  tableDefault: boolean,
): CellMargins | undefined {
  if (!el) return undefined;
  // A style's `w:tblCellMar` is the table's default layer: a side given as a
  // share or as auto/nil is no inset (cellMarginTwips), nor is a negative one.
  // Its `w:tcMar` reads as a cell's own, which drops such a side instead.
  const read = (node: unknown): number | undefined => {
    const m = asElement(node);
    return m ? cellMarginTwips(getAttr(m, 'w'), getAttr(m, 'type'), tableDefault) : undefined;
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

// §17.4.33 w:shd — the colour the shading paints, its pattern blended in (the
// table parser's policy; see `shadingFillHex`).
function parseFlatShading(node: unknown): CellShading | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const hex = shadingFillHex(getVal(el), getAttr(el, 'color'), getAttr(el, 'fill'));
  return hex !== undefined ? { colorHex: hex } : undefined;
}

function intVal(node: unknown): number | undefined {
  const v = getVal(node);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}
