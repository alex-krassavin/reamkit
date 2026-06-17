// ECMA-376 Part 1 §18.3.1.99 — worksheet.xml.
// Walk sheetData/row/c, producing a flat list of cells with absolute
// row/column addresses and resolved text values.

import { XMLParser } from 'fast-xml-parser';

import type {
  CellType,
  CfOperator,
  CfRule,
  CfRuleCellIs,
  CfRuleColorScale,
  CfRuleDataBar,
  CfRuleIconSet,
  Cfvo,
  CfvoType,
  ColumnWidth,
  ConditionalFormat,
  DataValidation,
  DataValidationType,
  HeaderFooter,
  HyperlinkRef,
  MergedRange,
  ParsedSparkline,
  ParsedWorksheet,
  RowHeight,
  SheetPane,
  SparklineKind,
  WorksheetCell,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPrintOptions,
} from '@/core/spreadsheet-model';
import { parseCellRef } from '@/excel/cell-reference';
import { parseAreaRef } from '@/excel/defined-name-ref';

type MutableMerge = {
  -readonly [K in keyof MergedRange]: MergedRange[K];
};

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
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

export function parseWorksheet(data: Uint8Array): ParsedWorksheet {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const worksheet = tree['worksheet'];
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
  const conditionalFormats = parseConditionalFormatting(wsObj);
  const dataValidations = parseDataValidations(wsObj);
  const hyperlinks = parseHyperlinks(wsObj);
  const headerFooter = parseHeaderFooter(wsObj);
  const sparklines = parseSparklines(wsObj);
  const tablePartRelIds = parseTableParts(wsObj);
  const printModel = {
    ...(pageMargins ? { pageMargins } : {}),
    ...(pageSetup ? { pageSetup } : {}),
    ...(fitToPage ? { fitToPage } : {}),
    ...(printOptions ? { printOptions } : {}),
    ...(rowBreaks.length > 0 ? { rowBreaks } : {}),
    ...(colBreaks.length > 0 ? { colBreaks } : {}),
    ...(pane ? { pane } : {}),
    ...(drawingRelId !== undefined ? { drawingRelId } : {}),
    ...(conditionalFormats.length > 0 ? { conditionalFormats } : {}),
    ...(dataValidations.length > 0 ? { dataValidations } : {}),
    ...(hyperlinks.length > 0 ? { hyperlinks } : {}),
    ...(headerFooter ? { headerFooter } : {}),
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

  return {
    cells,
    maxRow,
    maxColumn,
    columns: parseColumns(wsObj),
    merges: parseMerges(wsObj),
    rowHeights,
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
  const htRaw = strAttr(obj, 'ht');
  if (htRaw === undefined) return undefined;
  const heightPt = Number(htRaw);
  if (!Number.isFinite(heightPt)) return undefined;
  const customRaw = strAttr(obj, 'customHeight');
  const customHeight = customRaw === '1' || customRaw === 'true';
  return { row: rowIndex, heightPt, customHeight };
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
  if (
    paperSize === undefined &&
    orientation === undefined &&
    scale === undefined &&
    fitToWidth === undefined &&
    fitToHeight === undefined
  ) {
    return undefined;
  }
  return {
    ...(paperSize !== undefined ? { paperSize: Math.round(paperSize) } : {}),
    ...(orientation !== undefined ? { orientation } : {}),
    ...(scale !== undefined ? { scale } : {}),
    ...(fitToWidth !== undefined ? { fitToWidth: Math.round(fitToWidth) } : {}),
    ...(fitToHeight !== undefined ? { fitToHeight: Math.round(fitToHeight) } : {}),
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
  const horizontalCentered = flag('horizontalCentered');
  const verticalCentered = flag('verticalCentered');
  if (
    gridLines === undefined &&
    horizontalCentered === undefined &&
    verticalCentered === undefined
  ) {
    return undefined;
  }
  return {
    ...(gridLines !== undefined ? { gridLines } : {}),
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
      out.push({ min, max, widthChars: width });
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
// owning <cfRule>s. Reads `cellIs` (SC1) and `colorScale` (SC1b) rules; dataBar
// and iconSet are skipped until a follow-up.
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
  switch (strAttr(obj, 'type')) {
    case 'cellIs':
      return parseCellIsRule(obj, priority);
    case 'colorScale':
      return parseColorScaleRule(obj, priority);
    case 'dataBar':
      return parseDataBarRule(obj, priority);
    case 'iconSet':
      return parseIconSetRule(obj, priority);
    default:
      return undefined; // dataBar2010/etc. — skipped
  }
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
  return {
    type: 'dataBar',
    priority,
    cfvos,
    colorHex,
    ...(minLength !== undefined ? { minLength } : {}),
    ...(maxLength !== undefined ? { maxLength } : {}),
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
  let address: { column: number; row: number };
  if (ref) {
    try {
      address = parseCellRef(ref);
    } catch {
      return null;
    }
  } else {
    address = { column: fallbackCol, row: fallbackRow };
  }
  const typeStr = strAttr(obj, 't') ?? 'n';
  const type = validateCellType(typeStr);
  const styleStr = strAttr(obj, 's');
  const styleIndex = styleStr !== undefined ? Number(styleStr) : undefined;
  const v = obj['v'];
  const rawValue = textOf(v);
  const base = { column: address.column, row: address.row, type } as const;
  if (type === 'inlineStr') {
    const is = obj['is'];
    const inlineText = inlineStringText(is);
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
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  if (!node || typeof node !== 'object') return '';
  const obj = node as Record<string, unknown>;
  const inner = obj['#text'];
  if (typeof inner === 'string') return inner;
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
  const r = obj['r'];
  if (Array.isArray(r)) {
    const joined = r
      .map((rr) => textOf((rr as Record<string, unknown> | undefined)?.['t']))
      .join('');
    return joined.length > MAX_CELL_CHARS ? joined.slice(0, MAX_CELL_CHARS) : joined;
  }
  return '';
}
