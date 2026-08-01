// ECMA-376 Part 1 §18.3.1.99 — worksheet.xml.
// Walk sheetData/row/c, producing a flat list of cells with absolute
// row/column addresses and resolved text values.

import { XMLParser } from 'fast-xml-parser';

import type {
  CellType,
  CfOperator,
  CfRule,
  CfRuleAboveAverage,
  CfRuleCellIs,
  CfRuleColorScale,
  CfRuleDataBar,
  CfRuleDupUnique,
  CfRuleExpression,
  CfRuleIconSet,
  CfRuleText,
  CfRuleTimePeriod,
  CfRuleTop10,
  Cfvo,
  CfvoType,
  ColumnStyle,
  ColumnWidth,
  ConditionalFormat,
  DataValidation,
  DataValidationType,
  FormControlRef,
  HeaderFooter,
  HyperlinkRef,
  MergedRange,
  OleObjectRef,
  ParsedSparkline,
  ParsedWorksheet,
  RowHeight,
  RowStyle,
  SheetPane,
  SparklineKind,
  TimePeriodKind,
  WorksheetCell,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPrintOptions,
} from '@/core/spreadsheet-model';
import type { ThemePalette, WorkbookColors } from '@/excel/styles-parser';
import { INDEXED_COLORS } from '@/core/indexed-colors';
import { parseDxf } from '@/excel/styles-parser';
import { resolveInternalEntities } from '@/core/opc/xml-entities';
import { parseCellRef } from '@/excel/cell-reference';
import { parseAreaRef } from '@/excel/defined-name-ref';
import { decodeXstring } from '@/excel/escaped-text';

type MutableMerge = {
  -readonly [K in keyof MergedRange]: MergedRange[K];
};

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Entities a DOCTYPE
  // declares the parser never registers at all, so they are resolved before it
  // sees the text — see resolveInternalEntities.
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // Tolerate an explicit `x:` namespace prefix (<x:worksheet>, <x:sheetData>,
  // <x:row>, <x:c>) used by some producers — see workbook-parser.ts.
  removeNSPrefix: true,
});

// The worksheet/style model types now live in @/core/spreadsheet-model (the
// SpreadsheetML sibling of document-model); this parser imports them above.

/**
 * Parse `worksheet.xml` (§18.3.1.99) into a {@link ParsedWorksheet}: a flat list
 * of cells with absolute row/column addresses and their raw stored values, plus
 * the per-sheet geometry and print model (columns, merges, row heights,
 * pageMargins/pageSetup, breaks, frozen pane, conditional formats, data
 * validations, hyperlinks, header/footer, form controls, OLE objects, sparklines
 * and table parts). `r=` is optional on `<row>`/`<c>`; absent indices are inferred
 * from document order. A malformed root or missing `sheetData` yields an empty grid.
 */
export function parseWorksheet(data: Uint8Array, theme?: ThemePalette): ParsedWorksheet {
  const xml = resolveInternalEntities(decoder.decode(data));
  const tree = parser.parse(xml) as Record<string, unknown>;
  // §18.3.1.99 `<chartsheet>` — a sheet that is nothing but a chart. It has no
  // sheetData, but it does carry the same `<pageMargins>`, `<pageSetup>` and
  // `<drawing r:id>` a worksheet does, and reading none of them cost
  // 47813.xlsx its whole "Chart" tab: a page the reference fills with a plot of
  // 1700 points came out as the sheet after it.
  const worksheet = tree['worksheet'] ?? tree['chartsheet'];
  const emptyExtras = () => ({
    columns: [] as ReadonlyArray<ColumnWidth>,
    merges: [] as ReadonlyArray<MergedRange>,
    rowHeights: [] as ReadonlyArray<RowHeight>,
  });
  if (!worksheet || typeof worksheet !== 'object') {
    return { cells: [], maxRow: -1, maxColumn: -1, ...emptyExtras() };
  }
  const wsObj = worksheet as Record<string, unknown>;
  const pageMargins = parsePageMargins(wsObj);
  const pageSetup = parsePageSetup(wsObj);
  const fitToPage = parseFitToPage(wsObj);
  const printOptions = parsePrintOptions(wsObj);
  const rowBreaks = parseBreaks(wsObj, 'rowBreaks');
  const colBreaks = parseBreaks(wsObj, 'colBreaks');
  const pane = parsePane(wsObj);
  const drawingNode = wsObj['drawing'];
  const drawingRelId =
    drawingNode && typeof drawingNode === 'object'
      ? strAttr(drawingNode as Record<string, unknown>, 'id')
      : undefined;
  // §18.3.1.36 `<legacyDrawing>` — the pre-DrawingML shape part. A sheet's form
  // controls can live there and nowhere else (see vml-drawing.ts).
  const legacyNode = wsObj['legacyDrawing'];
  const legacyDrawingRelId =
    legacyNode && typeof legacyNode === 'object'
      ? strAttr(legacyNode as Record<string, unknown>, 'id')
      : undefined;
  const conditionalFormats = [
    ...parseConditionalFormatting(wsObj),
    // The 2009 extension's rules resolve their own colours; the workbook's
    // indexed table is not in reach here, and a `<x14:dxf>` naming an indexed
    // colour a workbook has REPLACED is rarer than the default is right.
    ...parseX14ConditionalFormatting(wsObj, { ...(theme ? { theme } : {}), indexed: INDEXED_COLORS }),
  ];
  const dataValidations = parseDataValidations(wsObj);
  const hyperlinks = parseHyperlinks(wsObj);
  const headerFooter = parseHeaderFooter(wsObj);
  const formControls = parseFormControls(wsObj);
  const oleObjects = parseOleObjects(wsObj);
  const sparklines = parseSparklines(wsObj);
  const tablePartRelIds = parseTableParts(wsObj);
  const sheetFormat = parseSheetFormatPr(wsObj);
  const printModel = {
    ...sheetFormat,
    ...(pageMargins ? { pageMargins } : {}),
    ...(pageSetup ? { pageSetup } : {}),
    ...(fitToPage ? { fitToPage } : {}),
    ...(printOptions ? { printOptions } : {}),
    ...(rowBreaks.length > 0 ? { rowBreaks } : {}),
    ...(colBreaks.length > 0 ? { colBreaks } : {}),
    ...(pane ? { pane } : {}),
    ...(drawingRelId !== undefined ? { drawingRelId } : {}),
    ...(legacyDrawingRelId !== undefined ? { legacyDrawingRelId } : {}),
    ...(conditionalFormats.length > 0 ? { conditionalFormats } : {}),
    ...(dataValidations.length > 0 ? { dataValidations } : {}),
    ...(hyperlinks.length > 0 ? { hyperlinks } : {}),
    ...(headerFooter ? { headerFooter } : {}),
    ...(formControls.length > 0 ? { formControls } : {}),
    ...(oleObjects.length > 0 ? { oleObjects } : {}),
    ...(sparklines.length > 0 ? { sparklines } : {}),
    ...(tablePartRelIds.length > 0 ? { tablePartRelIds } : {}),
  };
  const sheetData = wsObj['sheetData'];
  if (!sheetData || typeof sheetData !== 'object') {
    return {
      cells: [],
      maxRow: -1,
      maxColumn: -1,
      columns: parseColumns(wsObj),
      merges: parseMerges(wsObj),
      rowHeights: [],
      ...printModel,
    };
  }
  const rowRaw = (sheetData as Record<string, unknown>)['row'];
  const rows = Array.isArray(rowRaw) ? rowRaw : rowRaw !== undefined ? [rowRaw] : [];

  const cells: Array<WorksheetCell> = [];
  const rowHeights: Array<RowHeight> = [];
  const rowStyles: Array<RowStyle> = [];
  let maxRow = -1;
  let maxColumn = -1;

  // ECMA-376 §18.3.1.4/§18.3.1.73 — r= is optional on <row>/<c>: an absent row
  // index is "the previous row + 1"; an absent cell ref is "the previous cell's
  // column + 1" in the current row. Track running positions so r-less producers
  // (e.g. 56278.xlsx) don't render empty.
  let currentRow = -1;
  for (const row of rows) {
    if (!row || typeof row !== 'object') continue;
    const obj = row as Record<string, unknown>;
    const explicitRow = parseRowIndex(obj);
    currentRow = explicitRow !== undefined ? explicitRow : currentRow + 1;
    const height = parseRowHeight(obj, currentRow);
    if (height) rowHeights.push(height);
    // §18.3.1.73 — `s` is the row's style, and `customFormat` is what says to
    // use it. A row formatted this way writes no `<c>` for the cells it paints.
    const rowStyle = Number(strAttr(obj, 's'));
    const customFormat = boolAttr(obj, 'customFormat');
    if (customFormat && Number.isFinite(rowStyle) && rowStyle > 0) {
      rowStyles.push({ row: currentRow, styleIndex: rowStyle });
    }
    const cellRaw = obj['c'];
    const rowCells = Array.isArray(cellRaw) ? cellRaw : cellRaw !== undefined ? [cellRaw] : [];
    let prevCol = -1;
    for (const c of rowCells) {
      const parsed = parseCell(c, currentRow, prevCol + 1);
      if (!parsed) continue;
      prevCol = parsed.column;
      cells.push(parsed);
      if (parsed.row > maxRow) maxRow = parsed.row;
      if (parsed.column > maxColumn) maxColumn = parsed.column;
    }
  }

  const colStyles = parseColumnStyles(wsObj);
  return {
    cells,
    maxRow,
    maxColumn,
    columns: parseColumns(wsObj),
    merges: parseMerges(wsObj),
    rowHeights,
    ...(colStyles.length > 0 ? { columnStyles: colStyles } : {}),
    ...(rowStyles.length > 0 ? { rowStyles } : {}),
    ...printModel,
  };
}

function parseRowIndex(obj: Record<string, unknown>): number | undefined {
  const raw = strAttr(obj, 'r');
  if (!raw) return undefined;
  const n = Number(raw);
  return Number.isInteger(n) && n > 0 ? n - 1 : undefined;
}

function parseRowHeight(
  obj: Record<string, unknown>,
  rowIndex: number | undefined,
): RowHeight | undefined {
  if (rowIndex === undefined) return undefined;
  // A hidden row usually carries no `ht`, so it has to be recorded on `hidden`
  // alone — returning early when `ht` is absent lost it entirely.
  const hidden = boolAttr(obj, 'hidden');
  const htRaw = strAttr(obj, 'ht');
  const heightPt = htRaw !== undefined ? Number(htRaw) : Number.NaN;
  if (!Number.isFinite(heightPt) && !hidden) return undefined;
  const customRaw = strAttr(obj, 'customHeight');
  const customHeight = customRaw === '1' || customRaw === 'true';
  return {
    row: rowIndex,
    heightPt: Number.isFinite(heightPt) ? heightPt : 0,
    customHeight,
    ...(hidden ? { hidden } : {}),
  };
}

function parsePageMargins(ws: Record<string, unknown>): XlsxPageMargins | undefined {
  const node = ws['pageMargins'];
  if (!node || typeof node !== 'object') return undefined;
  const obj = node as Record<string, unknown>;
  const left = parseNumericAttr(obj, 'left');
  const right = parseNumericAttr(obj, 'right');
  const top = parseNumericAttr(obj, 'top');
  const bottom = parseNumericAttr(obj, 'bottom');
  if (left === undefined || right === undefined || top === undefined || bottom === undefined) {
    return undefined;
  }
  const header = parseNumericAttr(obj, 'header');
  const footer = parseNumericAttr(obj, 'footer');
  return {
    leftInches: left,
    rightInches: right,
    topInches: top,
    bottomInches: bottom,
    ...(header !== undefined ? { headerInches: header } : {}),
    ...(footer !== undefined ? { footerInches: footer } : {}),
  };
}

function parsePageSetup(ws: Record<string, unknown>): XlsxPageSetup | undefined {
  const node = ws['pageSetup'];
  if (!node || typeof node !== 'object') return undefined;
  const obj = node as Record<string, unknown>;
  const paperSize = parseNumericAttr(obj, 'paperSize');
  const orientationRaw = strAttr(obj, 'orientation');
  const orientation: XlsxPageSetup['orientation'] | undefined =
    orientationRaw === 'portrait' || orientationRaw === 'landscape' || orientationRaw === 'default'
      ? orientationRaw
      : undefined;
  const scale = parseNumericAttr(obj, 'scale');
  const fitToWidth = parseNumericAttr(obj, 'fitToWidth');
  const fitToHeight = parseNumericAttr(obj, 'fitToHeight');
  const printerSettingsRelId = strAttr(obj, 'id');
  const commentsRaw = strAttr(obj, 'cellComments');
  const cellComments: XlsxPageSetup['cellComments'] | undefined =
    commentsRaw === 'none' || commentsRaw === 'asDisplayed' || commentsRaw === 'atEnd'
      ? commentsRaw
      : undefined;
  if (
    paperSize === undefined &&
    orientation === undefined &&
    scale === undefined &&
    fitToWidth === undefined &&
    fitToHeight === undefined &&
    printerSettingsRelId === undefined &&
    cellComments === undefined
  ) {
    return undefined;
  }
  return {
    ...(paperSize !== undefined ? { paperSize: Math.round(paperSize) } : {}),
    ...(orientation !== undefined ? { orientation } : {}),
    ...(scale !== undefined ? { scale } : {}),
    ...(fitToWidth !== undefined ? { fitToWidth: Math.round(fitToWidth) } : {}),
    ...(fitToHeight !== undefined ? { fitToHeight: Math.round(fitToHeight) } : {}),
    ...(printerSettingsRelId !== undefined ? { printerSettingsRelId } : {}),
    ...(cellComments !== undefined ? { cellComments } : {}),
  };
}

// ECMA-376 §18.3.1.82/§18.3.1.65 — <sheetPr><pageSetUpPr fitToPage="1"/>.
function parseFitToPage(ws: Record<string, unknown>): boolean {
  const sheetPr = ws['sheetPr'];
  if (!sheetPr || typeof sheetPr !== 'object') return false;
  const pr = (sheetPr as Record<string, unknown>)['pageSetUpPr'];
  if (!pr || typeof pr !== 'object') return false;
  const raw = strAttr(pr as Record<string, unknown>, 'fitToPage');
  return raw === '1' || raw === 'true';
}

// ECMA-376 §18.3.1.70 — <printOptions gridLines horizontalCentered ...>.
function parsePrintOptions(ws: Record<string, unknown>): XlsxPrintOptions | undefined {
  const node = ws['printOptions'];
  if (!node || typeof node !== 'object') return undefined;
  const obj = node as Record<string, unknown>;
  const flag = (key: string): boolean | undefined => {
    const raw = strAttr(obj, key);
    if (raw === undefined) return undefined;
    return raw === '1' || raw === 'true';
  };
  const gridLines = flag('gridLines');
  const headings = flag('headings');
  const horizontalCentered = flag('horizontalCentered');
  const verticalCentered = flag('verticalCentered');
  if (
    gridLines === undefined &&
    headings === undefined &&
    horizontalCentered === undefined &&
    verticalCentered === undefined
  ) {
    return undefined;
  }
  return {
    ...(gridLines !== undefined ? { gridLines } : {}),
    ...(headings !== undefined ? { headings } : {}),
    ...(horizontalCentered !== undefined ? { horizontalCentered } : {}),
    ...(verticalCentered !== undefined ? { verticalCentered } : {}),
  };
}

// ECMA-376 §18.3.1.66 — <sheetViews><sheetView><pane>. A frozen (or frozen-split)
// pane freezes `ySplit` leading rows and `xSplit` leading columns; a plain "split"
// pane (a resizable divider, no freeze) is ignored. Frozen panes do not affect
// print/PDF — they are carried for round-trip + HTML sticky panes (E-SHEET SE2).
function parsePane(ws: Record<string, unknown>): SheetPane | undefined {
  const views = ws['sheetViews'];
  if (!views || typeof views !== 'object') return undefined;
  const viewRaw = (views as Record<string, unknown>)['sheetView'];
  const view = Array.isArray(viewRaw) ? viewRaw[0] : viewRaw;
  if (!view || typeof view !== 'object') return undefined;
  const paneNode = (view as Record<string, unknown>)['pane'];
  if (!paneNode || typeof paneNode !== 'object') return undefined;
  const pane = paneNode as Record<string, unknown>;
  const state = strAttr(pane, 'state');
  if (state !== 'frozen' && state !== 'frozenSplit') return undefined;
  const count = (key: string): number => {
    const raw = strAttr(pane, key);
    const n = raw !== undefined ? Number(raw) : NaN;
    return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
  };
  const frozenCols = count('xSplit');
  const frozenRows = count('ySplit');
  if (frozenCols === 0 && frozenRows === 0) return undefined;
  return { frozenRows, frozenCols };
}

// ECMA-376 §18.3.1.74/§18.3.1.14 — <rowBreaks>/<colBreaks> with <brk id=".."/>.
// Returns the (verbatim) ids of breaks; an absent id is skipped.
function parseBreaks(ws: Record<string, unknown>, tag: 'rowBreaks' | 'colBreaks'): Array<number> {
  const node = ws[tag];
  if (!node || typeof node !== 'object') return [];
  const brkRaw = (node as Record<string, unknown>)['brk'];
  const items = Array.isArray(brkRaw) ? brkRaw : brkRaw !== undefined ? [brkRaw] : [];
  const out: Array<number> = [];
  for (const b of items) {
    if (!b || typeof b !== 'object') continue;
    const id = parseNumericAttr(b as Record<string, unknown>, 'id');
    if (id !== undefined && Number.isInteger(id) && id >= 0) out.push(id);
  }
  return out;
}

function parseNumericAttr(obj: Record<string, unknown>, key: string): number | undefined {
  const raw = strAttr(obj, key);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * ECMA-376 §18.3.1.81 `<sheetFormatPr>` — the sheet's default row height and
 * column width, which apply to every row/column that does not override them.
 *
 * Both were previously ignored, so a row without an explicit `ht` had no height
 * at all and ended up however tall its text wanted to be. A spreadsheet row has
 * a definite height; text metrics do not get a vote.
 */
function parseSheetFormatPr(ws: Record<string, unknown>): {
  defaultRowHeightPt?: number;
  defaultColWidthChars?: number;
  baseColWidthChars?: number;
} {
  const node = ws['sheetFormatPr'];
  if (!node || typeof node !== 'object') return {};
  const obj = node as Record<string, unknown>;
  const height = parseNumericAttr(obj, 'defaultRowHeight');
  const width = parseNumericAttr(obj, 'defaultColWidth');
  const base = parseNumericAttr(obj, 'baseColWidth');
  return {
    ...(height !== undefined && height > 0 ? { defaultRowHeightPt: height } : {}),
    ...(width !== undefined && width > 0 ? { defaultColWidthChars: width } : {}),
    ...(base !== undefined && base > 0 ? { baseColWidthChars: base } : {}),
  };
}

// §18.3.1.13 `<col style>` — the style every cell of the span takes when it has
// none of its own. Read separately from the width because the two are
// independent: a `<col min max style/>` with no `width` is a legal way to
// format a column, and the width reader below discards it.
function parseColumnStyles(ws: Record<string, unknown>): Array<ColumnStyle> {
  const colsNode = ws['cols'];
  if (!colsNode || typeof colsNode !== 'object') return [];
  const colRaw = (colsNode as Record<string, unknown>)['col'];
  const items = Array.isArray(colRaw) ? colRaw : colRaw !== undefined ? [colRaw] : [];
  const out: Array<ColumnStyle> = [];
  for (const c of items) {
    if (!c || typeof c !== 'object') continue;
    const obj = c as Record<string, unknown>;
    const min = Number(strAttr(obj, 'min'));
    const max = Number(strAttr(obj, 'max'));
    const styleIndex = Number(strAttr(obj, 'style'));
    if (!Number.isFinite(min) || !Number.isFinite(max)) continue;
    if (!Number.isFinite(styleIndex) || styleIndex <= 0) continue;
    out.push({ min, max, styleIndex });
  }
  return out;
}

function parseColumns(ws: Record<string, unknown>): Array<ColumnWidth> {
  const colsNode = ws['cols'];
  if (!colsNode || typeof colsNode !== 'object') return [];
  const colsObj = colsNode as Record<string, unknown>;
  const colRaw = colsObj['col'];
  const items = Array.isArray(colRaw) ? colRaw : colRaw !== undefined ? [colRaw] : [];
  const out: Array<ColumnWidth> = [];
  for (const c of items) {
    if (!c || typeof c !== 'object') continue;
    const obj = c as Record<string, unknown>;
    const min = Number(strAttr(obj, 'min'));
    const max = Number(strAttr(obj, 'max'));
    const width = Number(strAttr(obj, 'width'));
    if (Number.isFinite(min) && Number.isFinite(max) && Number.isFinite(width)) {
      const hidden = boolAttr(obj, 'hidden');
      out.push({ min, max, widthChars: width, ...(hidden ? { hidden } : {}) });
    }
  }
  return out;
}

function parseMerges(ws: Record<string, unknown>): Array<MergedRange> {
  const mergeNode = ws['mergeCells'];
  if (!mergeNode || typeof mergeNode !== 'object') return [];
  const mergeObj = mergeNode as Record<string, unknown>;
  const mergeRaw = mergeObj['mergeCell'];
  const items = Array.isArray(mergeRaw) ? mergeRaw : mergeRaw !== undefined ? [mergeRaw] : [];
  const out: Array<MergedRange> = [];
  for (const m of items) {
    if (!m || typeof m !== 'object') continue;
    const obj = m as Record<string, unknown>;
    const ref = strAttr(obj, 'ref');
    if (!ref) continue;
    const colonIdx = ref.indexOf(':');
    if (colonIdx < 0) continue;
    try {
      const start = parseCellRef(ref.substring(0, colonIdx));
      const end = parseCellRef(ref.substring(colonIdx + 1));
      const range: MutableMerge = {
        startColumn: Math.min(start.column, end.column),
        startRow: Math.min(start.row, end.row),
        endColumn: Math.max(start.column, end.column),
        endRow: Math.max(start.row, end.row),
      };
      out.push(range);
    } catch {
      // Ignore malformed merge refs.
    }
  }
  return out;
}

// §18.3.1.18 <conditionalFormatting sqref="…"> elements (one or more), each
// owning <cfRule>s. Reads the value/text-driven families: `cellIs`, `colorScale`,
// `dataBar`, `iconSet`, `top10`, `aboveAverage`, `duplicate/uniqueValues`, the
// text tests (`containsText`/`beginsWith`/…), plus `expression` (formula, run by
// the W9 engine) and `timePeriod` (a clock-relative date window). Only the 2010+
// extension variants (`dataBar2010`, `iconSet2010`, …) remain unparsed.
function parseConditionalFormatting(ws: Record<string, unknown>): Array<ConditionalFormat> {
  const raw = ws['conditionalFormatting'];
  const items = Array.isArray(raw) ? raw : raw !== undefined ? [raw] : [];
  const out: Array<ConditionalFormat> = [];
  for (const item of items) {
    if (!item || typeof item !== 'object') continue;
    const obj = item as Record<string, unknown>;
    const sqref = strAttr(obj, 'sqref');
    if (!sqref) continue;
    const ranges = parseSqref(sqref);
    if (ranges.length === 0) continue;
    const ruleRaw = obj['cfRule'];
    const ruleItems = Array.isArray(ruleRaw) ? ruleRaw : ruleRaw !== undefined ? [ruleRaw] : [];
    const rules: Array<CfRule> = [];
    for (const rn of ruleItems) {
      if (!rn || typeof rn !== 'object') continue;
      const rule = parseCfRule(rn as Record<string, unknown>);
      if (rule) rules.push(rule);
    }
    if (rules.length > 0) out.push({ ranges, rules });
  }
  return out;
}

// ISO/IEC 29500 could not express everything Excel 2010 wanted of a conditional
// format, so the 2009 extension carries the rest in <extLst>: the same rule
// families, plus the ones the 2006 schema has no room for. The shape differs in
// three places — the range is a child <xm:sqref> and not an attribute, the
// formulas are <xm:f> and not <formula>, and the format is written INSIDE the
// rule as <x14:dxf> rather than pointing at the workbook's table.
//
// Nineteen corpus workbooks put rules here and we read none of them:
// tdf122102.xlsx paints four cells yellow, green, red and grey, and we printed
// four plain ones.
/** The rule types the extension states as a formula rather than a literal. */
const TEXT_RULES: ReadonlySet<string> = new Set([
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
]);

function parseX14ConditionalFormatting(
  ws: Record<string, unknown>,
  colors: WorkbookColors,
): Array<ConditionalFormat> {
  const out: Array<ConditionalFormat> = [];
  for (const ext of toArray(asObjectNode(ws['extLst'])?.['ext'])) {
    const group = asObjectNode(asObjectNode(ext)?.['conditionalFormattings']);
    if (!group) continue;
    for (const cf of toArray(group['conditionalFormatting'])) {
      const obj = asObjectNode(cf);
      if (!obj) continue;
      // <xm:sqref> is the element's text, not an attribute.
      const sqref = typeof obj['sqref'] === 'string' ? obj['sqref'] : undefined;
      if (!sqref) continue;
      const ranges = parseSqref(sqref);
      if (ranges.length === 0) continue;
      const rules: Array<CfRule> = [];
      for (const rn of toArray(obj['cfRule'])) {
        const node = asObjectNode(rn);
        if (!node) continue;
        const rule = parseCfRule(asBaseCfRule(node));
        if (!rule) continue;
        // colorScale/dataBar/iconSet colour themselves; the rest take a dxf.
        const dxf = parseDxf(node['dxf'], colors);
        rules.push(
          'dxfId' in rule && Object.keys(dxf).length > 0 ? ({ ...rule, dxf }) : rule,
        );
      }
      if (rules.length > 0) out.push({ ranges, rules });
    }
  }
  return out;
}

/**
 * Rewrite a `<x14:cfRule>` into the shape {@link parseCfRule} reads, so one
 * reader serves both spellings of the same rule.
 */
function asBaseCfRule(node: Record<string, unknown>): Record<string, unknown> {
  const formulas = toArray(node['f']).filter((f) => typeof f === 'string' || typeof f === 'number');
  // One `<formula>` parses as a string and several as an array; the readers
  // below expect that shape, not an array of one.
  const out: Record<string, unknown> = {
    ...node,
    formula: formulas.length === 1 ? formulas[0] : formulas,
  };
  // Every value-driven family refuses a rule with no `dxfId`, and the extension
  // has none — it wrote the format inline instead. Stand one in; the parsed
  // `<x14:dxf>` beside it is what the renderer actually reads.
  if (out['@_dxfId'] === undefined && node['dxf'] !== undefined) out['@_dxfId'] = '0';
  // A text rule states its needle in an attribute — a LITERAL. The extension
  // exists because the needle is a reference (`$B$1`, `Munka1!$A$1`), and it
  // writes the whole test as its first formula: `NOT(ISERROR(SEARCH($B$1,A1)))`.
  // Read that instead of comparing against the reference's own spelling, which
  // no cell contains.
  if (TEXT_RULES.has(String(out['@_type'])) && formulas.length > 0) {
    out['@_type'] = 'expression';
    out['formula'] = formulas[0];
  }
  return out;
}

const DV_TYPES: ReadonlySet<string> = new Set<DataValidationType>([
  'none',
  'whole',
  'decimal',
  'list',
  'date',
  'time',
  'textLength',
  'custom',
]);

// §18.3.1.32 <dataValidations> / §18.3.1.33 <dataValidation> — per-range input
// constraints (E-SHEET SV1). Reads the main-namespace validations: the visual
// signal is `type="list"` (an in-cell dropdown), but every field rides through
// for a faithful read→write round-trip. x14 (cross-sheet list source)
// validations carried in <extLst> are a documented v1 omission.
function parseDataValidations(ws: Record<string, unknown>): Array<DataValidation> {
  const node = asObjectNode(ws['dataValidations']);
  if (!node) return [];
  const out: Array<DataValidation> = [];
  for (const dv of toArray(node['dataValidation'])) {
    const obj = asObjectNode(dv);
    if (!obj) continue;
    const sqref = strAttr(obj, 'sqref');
    if (!sqref) continue;
    const ranges = parseSqref(sqref);
    if (ranges.length === 0) continue;
    const typeRaw = strAttr(obj, 'type');
    const type: DataValidationType =
      typeRaw && DV_TYPES.has(typeRaw) ? (typeRaw as DataValidationType) : 'none';
    const operator = strAttr(obj, 'operator');
    const errorStyle = strAttr(obj, 'errorStyle');
    const formula1 = formulaText(obj['formula1']);
    const formula2 = formulaText(obj['formula2']);
    const promptTitle = strAttr(obj, 'promptTitle');
    const prompt = strAttr(obj, 'prompt');
    const errorTitle = strAttr(obj, 'errorTitle');
    const error = strAttr(obj, 'error');
    out.push({
      type,
      ranges,
      ...(operator !== undefined ? { operator } : {}),
      ...(boolAttr(obj, 'allowBlank') ? { allowBlank: true } : {}),
      ...(boolAttr(obj, 'showDropDown') ? { showDropDown: true } : {}),
      ...(boolAttr(obj, 'showInputMessage') ? { showInputMessage: true } : {}),
      ...(boolAttr(obj, 'showErrorMessage') ? { showErrorMessage: true } : {}),
      ...(errorStyle !== undefined ? { errorStyle } : {}),
      ...(formula1 !== undefined ? { formula1 } : {}),
      ...(formula2 !== undefined ? { formula2 } : {}),
      ...(promptTitle !== undefined ? { promptTitle } : {}),
      ...(prompt !== undefined ? { prompt } : {}),
      ...(errorTitle !== undefined ? { errorTitle } : {}),
      ...(error !== undefined ? { error } : {}),
    });
  }
  return out;
}

function boolAttr(obj: Record<string, unknown>, key: string): boolean {
  const raw = strAttr(obj, key);
  return raw === '1' || raw === 'true';
}

// §18.3.1.46 <headerFooter><oddHeader>/<oddFooter> — the sheet's print header and
// footer format strings (E-SHEET W4). The projection expands the &-codes; v1 reads
// the odd (= default) header/footer (even/first variants are a later refinement).
function parseHeaderFooter(ws: Record<string, unknown>): HeaderFooter | undefined {
  const node = asObjectNode(ws['headerFooter']);
  if (!node) return undefined;
  const oddHeader = formulaText(node['oddHeader']);
  const oddFooter = formulaText(node['oddFooter']);
  if (!oddHeader && !oddFooter) return undefined;
  return {
    ...(oddHeader ? { oddHeader } : {}),
    ...(oddFooter ? { oddFooter } : {}),
  };
}

// §18.3.1.47 <hyperlinks><hyperlink ref r:id location display tooltip> — raw cell
// hyperlinks (E-SHEET W3). removeNSPrefix turns r:id into id (mirrors the drawing
// relId); the reader resolves relId → an external URL. A hyperlink with neither a
// relId nor a location carries no target and is dropped.
function parseHyperlinks(ws: Record<string, unknown>): Array<HyperlinkRef> {
  const node = asObjectNode(ws['hyperlinks']);
  if (!node) return [];
  const out: Array<HyperlinkRef> = [];
  for (const h of toArray(node['hyperlink'])) {
    const obj = asObjectNode(h);
    if (!obj) continue;
    const ref = strAttr(obj, 'ref');
    if (!ref) continue;
    const relId = strAttr(obj, 'id');
    const location = strAttr(obj, 'location');
    if (relId === undefined && location === undefined) continue;
    const display = strAttr(obj, 'display');
    const tooltip = strAttr(obj, 'tooltip');
    out.push({
      ref,
      ...(relId !== undefined ? { relId } : {}),
      ...(location !== undefined ? { location } : {}),
      ...(display !== undefined ? { display } : {}),
      ...(tooltip !== undefined ? { tooltip } : {}),
    });
  }
  return out;
}

// sqref is whitespace-separated areas ("A1:A10 C1:C5"); each resolves to a box.
function parseSqref(sqref: string): Array<MergedRange> {
  const out: Array<MergedRange> = [];
  for (const token of sqref.split(/\s+/)) {
    if (!token) continue;
    const r = parseAreaRef(token);
    if (r) {
      out.push({
        startColumn: r.startColumn,
        startRow: r.startRow,
        endColumn: r.endColumn,
        endRow: r.endRow,
      });
    }
  }
  return out;
}

const CF_OPERATORS: ReadonlySet<string> = new Set<CfOperator>([
  'lessThan',
  'lessThanOrEqual',
  'equal',
  'notEqual',
  'greaterThanOrEqual',
  'greaterThan',
  'between',
  'notBetween',
]);

function parseCfRule(obj: Record<string, unknown>): CfRule | undefined {
  const priority = parseNumericAttr(obj, 'priority') ?? 0;
  const type = strAttr(obj, 'type');
  switch (type) {
    case 'cellIs':
      return parseCellIsRule(obj, priority);
    case 'colorScale':
      return parseColorScaleRule(obj, priority);
    case 'dataBar':
      return parseDataBarRule(obj, priority);
    case 'iconSet':
      return parseIconSetRule(obj, priority);
    case 'top10':
      return parseTop10Rule(obj, priority);
    case 'aboveAverage':
      return parseAboveAverageRule(obj, priority);
    case 'duplicateValues':
    case 'uniqueValues':
      return parseDupUniqueRule(type, obj, priority);
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return parseTextRule(type, obj, priority);
    case 'expression':
      return parseExpressionRule(obj, priority);
    case 'timePeriod':
      return parseTimePeriodRule(obj, priority);
    default:
      return undefined; // dataBar2010 / iconSet2010 / etc. — skipped
  }
}

const TIME_PERIODS: ReadonlySet<string> = new Set<TimePeriodKind>([
  'today',
  'yesterday',
  'tomorrow',
  'last7Days',
  'thisWeek',
  'lastWeek',
  'nextWeek',
  'thisMonth',
  'lastMonth',
  'nextMonth',
]);

// §18.3.1.10 type="expression" — the single <formula> is the predicate; the
// conditional-format layer evaluates it per cell against the cached grid values
// (E-SHEET W9). A rule without a formula is dropped.
function parseExpressionRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleExpression | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const formula = formulaText(obj['formula']);
  if (formula === undefined) return undefined;
  return { type: 'expression', priority, formula, dxfId };
}

// §18.3.1.10 type="timePeriod" — the timePeriod attribute names the clock-
// relative window; Excel also emits a helper <formula> which we keep verbatim
// for write-back. An unknown period is dropped (no faithful render).
function parseTimePeriodRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleTimePeriod | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const period = strAttr(obj, 'timePeriod');
  if (period === undefined || !TIME_PERIODS.has(period)) return undefined;
  const formula = formulaText(obj['formula']);
  return {
    type: 'timePeriod',
    priority,
    timePeriod: period as TimePeriodKind,
    dxfId,
    ...(formula !== undefined ? { formula } : {}),
  };
}

// §18.3.1.10 type="top10" — `rank` (default 10) top values take the dxf; `percent`
// reads rank as a percentage, `bottom` flips to the lowest values.
function parseTop10Rule(obj: Record<string, unknown>, priority: number): CfRuleTop10 | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const rank = parseNumericAttr(obj, 'rank') ?? 10;
  return {
    type: 'top10',
    priority,
    rank,
    percent: boolAttr(obj, 'percent'),
    bottom: boolAttr(obj, 'bottom'),
    dxfId,
  };
}

// §18.3.1.10 type="aboveAverage" — `aboveAverage` defaults to true; `equalAverage`
// makes it inclusive; `stdDev` (when present) shifts the threshold by N std-devs.
function parseAboveAverageRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleAboveAverage | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const aboveRaw = strAttr(obj, 'aboveAverage');
  const aboveAverage = aboveRaw === undefined ? true : aboveRaw === '1' || aboveRaw === 'true';
  const stdDev = parseNumericAttr(obj, 'stdDev');
  return {
    type: 'aboveAverage',
    priority,
    aboveAverage,
    equalAverage: boolAttr(obj, 'equalAverage'),
    ...(stdDev !== undefined ? { stdDev } : {}),
    dxfId,
  };
}

// §18.3.1.10 type="duplicateValues" | "uniqueValues" — attribute-only; the dxf
// paints repeated (or one-off) values across the range.
function parseDupUniqueRule(
  type: 'duplicateValues' | 'uniqueValues',
  obj: Record<string, unknown>,
  priority: number,
): CfRuleDupUnique | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  return { type, priority, dxfId };
}

// §18.3.1.10 the text tests — `text` is the needle; Excel's generated `<formula>`
// is preserved verbatim (for faithful write-back) but matched directly, not run.
function parseTextRule(
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith',
  obj: Record<string, unknown>,
  priority: number,
): CfRuleText | undefined {
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const text = strAttr(obj, 'text');
  if (text === undefined) return undefined;
  const formula = formulaText(obj['formula']);
  return { type, priority, text, dxfId, ...(formula !== undefined ? { formula } : {}) };
}

// §18.3.1.49 <iconSet iconSet="3TrafficLights1"> — N cfvo thresholds (N = 3/4/5)
// naming the per-value buckets; `reverse` flips icon order.
function parseIconSetRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleIconSet | undefined {
  const is = obj['iconSet'];
  if (!is || typeof is !== 'object') return undefined;
  const isObj = is as Record<string, unknown>;
  const cfvos = parseCfvos(isObj['cfvo']);
  if (cfvos.length < 3) return undefined;
  const iconSet = strAttr(isObj, 'iconSet') ?? '3TrafficLights1';
  const reverseRaw = strAttr(isObj, 'reverse');
  const reverse = reverseRaw === '1' || reverseRaw === 'true';
  return { type: 'iconSet', priority, iconSet, cfvos, ...(reverse ? { reverse } : {}) };
}

// §18.3.1.28 <dataBar> — 2 cfvo stops (lower/upper) + a fill <color>; optional
// minLength/maxLength percent bounds. Extra cfvos (axis variants) are ignored.
function parseDataBarRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleDataBar | undefined {
  const db = obj['dataBar'];
  if (!db || typeof db !== 'object') return undefined;
  const dbObj = db as Record<string, unknown>;
  const cfvos = parseCfvos(dbObj['cfvo']);
  if (cfvos.length < 2) return undefined;
  const colorHex = colorRgbHex(dbObj['color']);
  if (!colorHex) return undefined;
  const minLength = parseNumericAttr(dbObj, 'minLength');
  const maxLength = parseNumericAttr(dbObj, 'maxLength');
  // §18.3.1.28 — absent means true; only an explicit false hides the number.
  const showValueRaw = strAttr(dbObj, 'showValue');
  const showValue = showValueRaw === '0' || showValueRaw === 'false' ? false : undefined;
  return {
    type: 'dataBar',
    priority,
    cfvos,
    colorHex,
    ...(minLength !== undefined ? { minLength } : {}),
    ...(maxLength !== undefined ? { maxLength } : {}),
    ...(showValue === false ? { showValue } : {}),
  };
}

function parseCellIsRule(obj: Record<string, unknown>, priority: number): CfRuleCellIs | undefined {
  const operator = strAttr(obj, 'operator');
  if (!operator || !CF_OPERATORS.has(operator)) return undefined;
  const dxfId = parseNumericAttr(obj, 'dxfId');
  if (dxfId === undefined) return undefined;
  const fRaw = obj['formula'];
  const fItems = Array.isArray(fRaw) ? fRaw : fRaw !== undefined ? [fRaw] : [];
  const formulas: Array<string> = [];
  for (const f of fItems) {
    const text = formulaText(f);
    if (text !== undefined) formulas.push(text);
  }
  if (formulas.length === 0) return undefined;
  return { type: 'cellIs', priority, operator: operator as CfOperator, formulas, dxfId };
}

const CFVO_TYPES: ReadonlySet<string> = new Set<CfvoType>([
  'num',
  'percent',
  'max',
  'min',
  'percentile',
  'formula',
  'autoMin',
  'autoMax',
]);

// §18.3.1.16 <colorScale> — N <cfvo> stops paired with N <color>s (N = 2 or 3).
// A stop with an unknown type or a non-rgb colour aborts the rule (returns
// undefined) so a scale we cannot resolve faithfully is simply not applied.
function parseColorScaleRule(
  obj: Record<string, unknown>,
  priority: number,
): CfRuleColorScale | undefined {
  const cs = obj['colorScale'];
  if (!cs || typeof cs !== 'object') return undefined;
  const csObj = cs as Record<string, unknown>;
  const cfvos = parseCfvos(csObj['cfvo']);
  const colorsHex = parseScaleColors(csObj['color']);
  if (cfvos.length < 2 || cfvos.length !== colorsHex.length) return undefined;
  return { type: 'colorScale', priority, cfvos, colorsHex };
}

function parseCfvos(raw: unknown): Array<Cfvo> {
  const items = Array.isArray(raw) ? raw : raw !== undefined ? [raw] : [];
  const out: Array<Cfvo> = [];
  for (const it of items) {
    if (!it || typeof it !== 'object') continue;
    const o = it as Record<string, unknown>;
    const type = strAttr(o, 'type');
    if (!type || !CFVO_TYPES.has(type)) return []; // unknown stop → drop the rule
    const val = strAttr(o, 'val');
    out.push(val !== undefined ? { type: type as CfvoType, val } : { type: type as CfvoType });
  }
  return out;
}

function parseScaleColors(raw: unknown): Array<string> {
  const items = Array.isArray(raw) ? raw : raw !== undefined ? [raw] : [];
  const out: Array<string> = [];
  for (const it of items) {
    const hex = colorRgbHex(it);
    if (!hex) return []; // theme/indexed colour (no rgb) → drop the rule for v1
    out.push(hex);
  }
  return out;
}

// <color rgb="FFF8696B"> → "F8696B" (ARGB alpha stripped, upper-cased); matches
// styles-parser's convention. Returns undefined for theme/indexed/auto colours.
function colorRgbHex(node: unknown): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const rgb = strAttr(node as Record<string, unknown>, 'rgb');
  if (!rgb) return undefined;
  if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return rgb.substring(2).toUpperCase();
  if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return rgb.toUpperCase();
  return undefined;
}

// <formula>5</formula> — fast-xml-parser yields the string directly (text-only,
// no attributes) or a node carrying #text.
function formulaText(node: unknown): string | undefined {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  if (node && typeof node === 'object') {
    const t = (node as Record<string, unknown>)['#text'];
    if (typeof t === 'string') return t;
    if (typeof t === 'number') return String(t);
  }
  return undefined;
}

// Coerce a fast-xml-parser child (which collapses a single element to an object
// and repeats to an array) into an array.
function toArray(v: unknown): ReadonlyArray<unknown> {
  return Array.isArray(v) ? v : v === undefined || v === null ? [] : [v];
}

// x14 sparklines (E-SHEET SC2). Worksheet extLst → <x14:sparklineGroups> /
// <x14:sparklineGroup type=…> / <x14:sparkline> with <xm:f> (data range) +
// <xm:sqref> (host cell). removeNSPrefix:true strips the x14:/xm: prefixes, so
// the tags read plainly here. The group's <x14:colorSeries rgb=…> tints the
// series. An absent type is a line sparkline; "stacked" is win/loss.
function parseSparklines(ws: Record<string, unknown>): Array<ParsedSparkline> {
  const extLst = asObjectNode(ws['extLst']);
  if (!extLst) return [];
  const out: Array<ParsedSparkline> = [];
  for (const ext of toArray(extLst['ext'])) {
    const groupsNode = asObjectNode(asObjectNode(ext)?.['sparklineGroups']);
    if (!groupsNode) continue;
    for (const g of toArray(groupsNode['sparklineGroup'])) {
      const group = asObjectNode(g);
      if (!group) continue;
      const typeStr = strAttr(group, 'type');
      const kind: SparklineKind =
        typeStr === 'column' ? 'column' : typeStr === 'stacked' ? 'winLoss' : 'line';
      const colorHex = colorRgbHex(group['colorSeries']);
      const slNode = asObjectNode(group['sparklines']);
      if (!slNode) continue;
      for (const sl of toArray(slNode['sparkline'])) {
        const spark = asObjectNode(sl);
        if (!spark) continue;
        const dataRange = formulaText(spark['f']);
        const sqref = formulaText(spark['sqref']);
        if (dataRange && sqref) {
          out.push({ kind, dataRange, sqref, ...(colorHex ? { colorHex } : {}) });
        }
      }
    }
  }
  return out;
}

// §18.3.* form controls (E-SHEET W8). Declared either as a transitional
// <controls> child of the worksheet or, in modern files, under the x14 extLst —
// often wrapped in <mc:AlternateContent><mc:Choice Requires="x14">. removeNSPrefix
// strips the x14:/mc:/r: prefixes, so each <control name r:id> reads plainly; the
// relId resolves (in the reader) to the control's ctrlProp part (type + state).
// Deduped by relId so a Choice/Fallback or doubled declaration counts once.
function parseFormControls(ws: Record<string, unknown>): Array<FormControlRef> {
  const out: Array<FormControlRef> = [];
  const seen = new Set<string>();
  collectControls(asObjectNode(ws['controls']), out, seen);
  for (const ext of toArray(asObjectNode(ws['extLst'])?.['ext'])) {
    collectControls(asObjectNode(asObjectNode(ext)?.['controls']), out, seen);
  }
  return out;
}

function collectControls(
  node: Record<string, unknown> | undefined,
  out: Array<FormControlRef>,
  seen: Set<string>,
): void {
  if (!node) return;
  const direct = toArray(node['control']);
  const fromChoice = toArray(asObjectNode(node['AlternateContent'])?.['Choice']).flatMap((c) =>
    toArray(asObjectNode(c)?.['control']),
  );
  for (const c of [...direct, ...fromChoice]) {
    const obj = asObjectNode(c);
    if (!obj) continue;
    const relId = strAttr(obj, 'id');
    if (!relId || seen.has(relId)) continue;
    seen.add(relId);
    const name = strAttr(obj, 'name');
    const shapeId = strAttr(obj, 'shapeId');
    // §18.3.1.20 `<controlPr print>` — absent means print, so only an explicit
    // false is worth recording.
    const controlPr = asObjectNode(obj['controlPr']);
    const printAttr = controlPr ? strAttr(controlPr, 'print') : undefined;
    const print = printAttr === '0' || printAttr === 'false' ? false : undefined;
    out.push({
      relId,
      ...(name ? { name } : {}),
      ...(shapeId ? { shapeId } : {}),
      ...(print === false ? { print } : {}),
    });
  }
}

// §18.3.* <oleObjects><oleObject progId r:id> — embedded OLE / ActiveX controls
// (E-SHEET W10). Like form controls they are often wrapped in <mc:AlternateContent>
// <mc:Choice Requires="x14">; removeNSPrefix strips the prefixes so each reads as
// <oleObject progId id>. The relId resolves (in the reader) to the control's
// activeX part (the property bag). Deduped by relId.
function parseOleObjects(ws: Record<string, unknown>): Array<OleObjectRef> {
  const out: Array<OleObjectRef> = [];
  const seen = new Set<string>();
  collectOleObjects(asObjectNode(ws['oleObjects']), out, seen);
  for (const ext of toArray(asObjectNode(ws['extLst'])?.['ext'])) {
    collectOleObjects(asObjectNode(asObjectNode(ext)?.['oleObjects']), out, seen);
  }
  return out;
}

function collectOleObjects(
  node: Record<string, unknown> | undefined,
  out: Array<OleObjectRef>,
  seen: Set<string>,
): void {
  if (!node) return;
  const direct = toArray(node['oleObject']);
  // A workbook with TWO embedded objects writes two `mc:AlternateContent`
  // siblings, and the parser hands those back as an array — which the
  // single-node reader turned into nothing at all, losing both of
  // bug64512_embed.xlsx's attachments before anything could report them.
  const fromChoice = toArray(node['AlternateContent'])
    .flatMap((ac) => toArray(asObjectNode(ac)?.['Choice']))
    .flatMap((c) => toArray(asObjectNode(c)?.['oleObject']));
  for (const o of [...direct, ...fromChoice]) {
    const obj = asObjectNode(o);
    if (!obj) continue;
    const relId = strAttr(obj, 'id');
    if (!relId || seen.has(relId)) continue;
    seen.add(relId);
    const progId = strAttr(obj, 'progId');
    out.push({ relId, ...(progId ? { progId } : {}) });
  }
}

function asObjectNode(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' && !Array.isArray(v)
    ? (v as Record<string, unknown>)
    : undefined;
}

// §18.3.1.95 <tableParts><tablePart r:id="…"/> — the relationship ids of the
// sheet's table parts (E-SHEET SC3). removeNSPrefix turns r:id into id; the
// reader resolves each id to an xl/tables/tableN.xml part.
function parseTableParts(ws: Record<string, unknown>): Array<string> {
  const node = asObjectNode(ws['tableParts']);
  if (!node) return [];
  const out: Array<string> = [];
  for (const tp of toArray(node['tablePart'])) {
    const obj = asObjectNode(tp);
    const rid = obj ? strAttr(obj, 'id') : undefined;
    if (rid) out.push(rid);
  }
  return out;
}

function parseCell(c: unknown, fallbackRow: number, fallbackCol: number): WorksheetCell | null {
  if (!c || typeof c !== 'object') return null;
  const obj = c as Record<string, unknown>;
  const ref = strAttr(obj, 'r');
  // r= is optional (§18.3.1.4): without it the position is implied by order —
  // the current row and the column after the previous cell.
  //
  // A ref that is PRESENT but unreadable gets the same treatment: it carries no
  // position either, so document order is the only answer available, and it is
  // the one the spec already sanctions. Dropping the cell instead turned
  // tdf122336.xlsx — whose producer writes r="11_2" — into a blank page.
  //
  // Nor is such a ref worth guessing at. `11_2` reads as `<column>_<row>`, and
  // the file even corroborates the row half — every ref inside `<row r="2">`
  // ends in `_2` — but the columns it yields are wrong: that sheet's header row
  // is 11 labels in columns 1..11, its data row is 11 values, and reading the
  // refs as columns scatters them over 1, 4, 5, 7, 11, 13, 22, 29..32, filing
  // the start time under "Klantnaam" and stretching two rows across seven
  // pages. In document order every value lands under its own heading. The first
  // number is the producer's own field id, not a column.
  const implied = { column: fallbackCol, row: fallbackRow };
  let address: { column: number; row: number };
  if (ref) {
    try {
      address = parseCellRef(ref);
    } catch {
      address = implied;
    }
  } else {
    address = implied;
  }
  const typeStr = strAttr(obj, 't') ?? 'n';
  const type = validateCellType(typeStr);
  const styleStr = strAttr(obj, 's');
  const styleIndex = styleStr !== undefined ? Number(styleStr) : undefined;
  const v = obj['v'];
  const rawValue = textOf(v);
  // §18.3.1.4 `vm` — only an error cell can be standing in for a rich value, so
  // that is the only place the index is worth carrying.
  const vmStr = type === 'e' ? strAttr(obj, 'vm') : undefined;
  const vm = vmStr !== undefined ? Number(vmStr) : undefined;
  const base = {
    column: address.column,
    row: address.row,
    type,
    ...(vm !== undefined && Number.isInteger(vm) && vm > 0 ? { valueMetadataIndex: vm } : {}),
  } as const;
  if (type === 'inlineStr') {
    const is = obj['is'];
    // §18.3.1.4 — t="inlineStr" pairs with <is>, but producers exist that
    // declare the type and then write the text into <v> anyway
    // (duplicate-filename.xlsx). The text is right there; taking it beats
    // rendering a blank cell over a technicality.
    const fromIs = inlineStringText(is);
    const inlineText = fromIs !== '' ? fromIs : rawValue;
    return {
      ...base,
      rawValue: '',
      inlineText,
      ...(Number.isFinite(styleIndex) ? { styleIndex: styleIndex as number } : {}),
    };
  }
  return {
    ...base,
    rawValue,
    ...(Number.isFinite(styleIndex) ? { styleIndex: styleIndex as number } : {}),
  };
}

function validateCellType(t: string): CellType {
  if (
    t === 'n' ||
    t === 's' ||
    t === 'str' ||
    t === 'b' ||
    t === 'd' ||
    t === 'e' ||
    t === 'inlineStr'
  ) {
    return t;
  }
  return 'n';
}

function strAttr(obj: Record<string, unknown>, key: string): string | undefined {
  const v = obj[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}

function textOf(node: unknown): string {
  if (typeof node === 'string') return decodeXstring(node);
  if (typeof node === 'number') return String(node);
  if (!node || typeof node !== 'object') return '';
  const obj = node as Record<string, unknown>;
  const inner = obj['#text'];
  if (typeof inner === 'string') return decodeXstring(inner);
  if (typeof inner === 'number') return String(inner);
  return '';
}

// ECMA-376 / Excel limit: a cell holds at most 32 767 characters. Capping is
// spec-correct and a DoS guard against a crafted multi-MB inline string.
const MAX_CELL_CHARS = 32_767;

function inlineStringText(is: unknown): string {
  if (!is || typeof is !== 'object') return '';
  const obj = is as Record<string, unknown>;
  const t = obj['t'];
  const direct = textOf(t);
  if (direct) return direct.length > MAX_CELL_CHARS ? direct.slice(0, MAX_CELL_CHARS) : direct;
  // §18.4.8 `<is>` holds either a bare `<t>` or a sequence of `<r>` runs — and a
  // SINGLE run is one object, not an array of one. Reading only the array form
  // dropped every inline string a producer wrote as one formatted run:
  // 52348.xlsx labels its whole header row `<is><r><rPr/><t>Category</t></r></is>`
  // and every one of those cells came out empty.
  const raw = obj['r'];
  const runs: Array<unknown> = Array.isArray(raw) ? raw : raw !== undefined ? [raw] : [];
  const joined = runs
    .map((rr) => textOf((rr as Record<string, unknown> | undefined)?.['t']))
    .join('');
  return joined.length > MAX_CELL_CHARS ? joined.slice(0, MAX_CELL_CHARS) : joined;
}
