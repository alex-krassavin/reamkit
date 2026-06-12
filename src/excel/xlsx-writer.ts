// XLSX writer (E-SHEET SD1) — SheetDoc → .xlsx. The first writer that consumes
// the SpreadsheetML IR node directly (not the FlowDoc projection): a meaningful
// spreadsheet writer is only possible on the grid tree, because the cells carry
// their RAW value + style index, which the print-model projection deliberately
// resolves away. Re-emits workbook / worksheets / sharedStrings / styles through
// the core OPC writer, so the SheetDoc is a stable round-trip fixpoint (SD2).
//
// v1 writes cell values, shared strings, styles (numFmts/fonts/fills/borders/
// cellXfs/dxfs), merges, column widths and row heights. Conditional formats,
// sparklines and table parts are not yet written back (reported as losses);
// they round-trip through the reader but are dropped on write — SD3.

import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { Sheet, SheetDoc } from '@/core/ir/sheet';
import type { Loss } from '@/core/ir';
import type { OpcPart, Relationship } from '@/core/opc';
import type {
  CfRule,
  Cfvo,
  ConditionalFormat,
  DefinedName,
  ParsedSparkline,
  ParsedWorksheet,
  WorksheetCell,
  XlsxBorder,
  XlsxBorderEdge,
  XlsxCellXf,
  XlsxFill,
  XlsxFont,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxPrintOptions,
  XlsxStyles,
} from '@/core/spreadsheet-model';

import { FEATURES } from '@/core/ir';
import { buildOpcPackage } from '@/core/opc';

const encoder = new TextEncoder();

const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_WORKSHEET =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_SHARED_STRINGS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_SHARED_STRINGS =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function writeXlsx(sheet: SheetDoc): WriteResult {
  const losses: Array<Loss> = [];

  const worksheetParts: Array<OpcPart> = [];
  const sheetEntries: Array<{ name: string; sheetId: number; rid: string }> = [];
  const workbookRels: Array<Relationship> = [];

  sheet.sheets.forEach((s, i) => {
    const idx = i + 1;
    const rid = `rId${idx}`;
    worksheetParts.push({
      path: `xl/worksheets/sheet${idx}.xml`,
      data: encoder.encode(worksheetXml(s.grid)),
      contentType: CT_WORKSHEET,
    });
    workbookRels.push(rel(rid, REL_WORKSHEET, `worksheets/sheet${idx}.xml`));
    sheetEntries.push({ name: s.name, sheetId: idx, rid });
    recordSheetLosses(s, losses);
  });

  // Shared strings + styles come after the sheets in the relationship list.
  const sstRid = `rId${sheet.sheets.length + 1}`;
  const stylesRid = `rId${sheet.sheets.length + 2}`;
  workbookRels.push(rel(sstRid, REL_SHARED_STRINGS, 'sharedStrings.xml'));
  workbookRels.push(rel(stylesRid, REL_STYLES, 'styles.xml'));

  const bytes = buildOpcPackage({
    parts: [
      {
        path: 'xl/workbook.xml',
        data: encoder.encode(workbookXml(sheetEntries, sheet.definedNames, sheet.date1904)),
        contentType: CT_WORKBOOK,
      },
      ...worksheetParts,
      {
        path: 'xl/sharedStrings.xml',
        data: encoder.encode(sharedStringsXml(sheet.sharedStrings)),
        contentType: CT_SHARED_STRINGS,
      },
      {
        path: 'xl/styles.xml',
        data: encoder.encode(stylesXml(sheet.styles)),
        contentType: CT_STYLES,
      },
    ],
    rootRelationships: [rel('rId1', REL_OFFICE_DOCUMENT, 'xl/workbook.xml')],
    partRelationships: [{ sourcePart: 'xl/workbook.xml', relationships: workbookRels }],
  });

  return { bytes, losses };
}

export const xlsxWriter: DocumentWriter<SheetDoc> = {
  id: 'xlsx',
  consumes: 'sheet',
  supports: new Set([FEATURES.text, FEATURES.tables]),
  write: (doc) => writeXlsx(doc),
};

function rel(id: string, type: string, target: string): Relationship {
  return { id, type, target, targetMode: 'Internal' };
}

// Features in the SheetDoc not yet re-emitted. Conditional formatting and
// sparklines now write back (SD3b); table parts follow.
function recordSheetLosses(s: Sheet, losses: Array<Loss>): void {
  const g = s.grid;
  if (g.tables && g.tables.length > 0) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.tables,
      detail: `sheet "${s.name}": table parts not written back (SD3b tail)`,
    });
  }
}

// ── workbook.xml ───────────────────────────────────────────────────────────

function workbookXml(
  sheets: ReadonlyArray<{ name: string; sheetId: number; rid: string }>,
  definedNames: ReadonlyArray<DefinedName>,
  date1904: boolean,
): string {
  const sheetTags = sheets
    .map((s) => `<sheet name="${escapeAttr(s.name)}" sheetId="${s.sheetId}" r:id="${s.rid}"/>`)
    .join('');
  const definedNamesXml =
    definedNames.length > 0
      ? `<definedNames>${definedNames
          .map(
            (d) =>
              `<definedName name="${escapeAttr(d.name)}"${
                d.localSheetId !== undefined ? ` localSheetId="${d.localSheetId}"` : ''
              }>${escapeText(d.value)}</definedName>`,
          )
          .join('')}</definedNames>`
      : '';
  return (
    XML_DECL +
    `<workbook xmlns="${MAIN_NS}" xmlns:r="${R_NS}">` +
    (date1904 ? '<workbookPr date1904="1"/>' : '') +
    `<sheets>${sheetTags}</sheets>` +
    definedNamesXml +
    '</workbook>'
  );
}

// ── worksheet.xml ──────────────────────────────────────────────────────────

function worksheetXml(grid: ParsedWorksheet): string {
  const dimension =
    grid.maxRow >= 0 && grid.maxColumn >= 0
      ? `<dimension ref="A1:${cellRef(grid.maxRow, grid.maxColumn)}"/>`
      : '<dimension ref="A1"/>';

  const colsXml =
    grid.columns.length > 0
      ? `<cols>${grid.columns
          .map(
            (c) => `<col min="${c.min}" max="${c.max}" width="${c.widthChars}" customWidth="1"/>`,
          )
          .join('')}</cols>`
      : '';

  const heightByRow = new Map<number, { heightPt: number; customHeight: boolean }>();
  for (const rh of grid.rowHeights) {
    heightByRow.set(rh.row, { heightPt: rh.heightPt, customHeight: rh.customHeight });
  }

  // Group cells by row, ascending; cells within a row ascending by column.
  const cellsByRow = new Map<number, Array<WorksheetCell>>();
  for (const cell of grid.cells) {
    let row = cellsByRow.get(cell.row);
    if (!row) {
      row = [];
      cellsByRow.set(cell.row, row);
    }
    row.push(cell);
  }
  const rowIndices = new Set<number>([...cellsByRow.keys(), ...heightByRow.keys()]);
  const sortedRows = [...rowIndices].sort((a, b) => a - b);

  const rowsXml = sortedRows
    .map((r) => {
      const cells = (cellsByRow.get(r) ?? []).slice().sort((a, b) => a.column - b.column);
      const h = heightByRow.get(r);
      const heightAttrs = h
        ? ` ht="${h.heightPt}"${h.customHeight ? ' customHeight="1"' : ''}`
        : '';
      return `<row r="${r + 1}"${heightAttrs}>${cells.map(cellXml).join('')}</row>`;
    })
    .join('');

  const mergesXml =
    grid.merges.length > 0
      ? `<mergeCells count="${grid.merges.length}">${grid.merges
          .map(
            (m) =>
              `<mergeCell ref="${cellRef(m.startRow, m.startColumn)}:${cellRef(
                m.endRow,
                m.endColumn,
              )}"/>`,
          )
          .join('')}</mergeCells>`
      : '';

  // Worksheet child order follows the ECMA-376 §18.3.1.99 sequence so the output
  // opens in Excel and round-trips through the reader (E-SHEET SD3a).
  return (
    XML_DECL +
    `<worksheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}">` +
    (grid.fitToPage ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : '') +
    dimension +
    colsXml +
    `<sheetData>${rowsXml}</sheetData>` +
    mergesXml +
    conditionalFormattingXml(grid.conditionalFormats) +
    printOptionsXml(grid.printOptions) +
    pageMarginsXml(grid.pageMargins) +
    pageSetupXml(grid.pageSetup) +
    breaksXml('rowBreaks', grid.rowBreaks) +
    breaksXml('colBreaks', grid.colBreaks) +
    sparklineExtXml(grid.sparklines) +
    '</worksheet>'
  );
}

// §18.3.1.18 <conditionalFormatting> — one per range group; rules write back per
// type (E-SHEET SD3b). The dxfs the cellIs rules reference are emitted in
// styles.xml; the reader re-derives the colorScale/dataBar/iconSet visuals.
function conditionalFormattingXml(cfs: ReadonlyArray<ConditionalFormat> | undefined): string {
  if (!cfs || cfs.length === 0) return '';
  return cfs
    .map((cf) => {
      const sqref = cf.ranges.map(rangeRef).join(' ');
      return `<conditionalFormatting sqref="${sqref}">${cf.rules.map(cfRuleXml).join('')}</conditionalFormatting>`;
    })
    .join('');
}

function cfRuleXml(rule: CfRule): string {
  const p = ` priority="${rule.priority}"`;
  switch (rule.type) {
    case 'cellIs':
      return (
        `<cfRule type="cellIs"${p} operator="${rule.operator}" dxfId="${rule.dxfId}">` +
        rule.formulas.map((f) => `<formula>${escapeText(f)}</formula>`).join('') +
        '</cfRule>'
      );
    case 'colorScale':
      return (
        `<cfRule type="colorScale"${p}><colorScale>` +
        rule.cfvos.map(cfvoXml).join('') +
        rule.colorsHex.map((c) => `<color rgb="FF${c}"/>`).join('') +
        '</colorScale></cfRule>'
      );
    case 'dataBar':
      return (
        `<cfRule type="dataBar"${p}><dataBar` +
        (rule.minLength !== undefined ? ` minLength="${rule.minLength}"` : '') +
        (rule.maxLength !== undefined ? ` maxLength="${rule.maxLength}"` : '') +
        '>' +
        rule.cfvos.map(cfvoXml).join('') +
        `<color rgb="FF${rule.colorHex}"/></dataBar></cfRule>`
      );
    case 'iconSet':
      return (
        `<cfRule type="iconSet"${p}><iconSet iconSet="${rule.iconSet}"` +
        (rule.reverse ? ' reverse="1"' : '') +
        '>' +
        rule.cfvos.map(cfvoXml).join('') +
        '</iconSet></cfRule>'
      );
  }
}

function cfvoXml(cfvo: Cfvo): string {
  return `<cfvo type="${cfvo.type}"${cfvo.val !== undefined ? ` val="${escapeAttr(cfvo.val)}"` : ''}/>`;
}

// x14 sparklines back into the worksheet extLst (E-SHEET SD3b). Our model is flat
// (one kind/colour per sparkline), so each round-trips as its own group — the
// reader recovers the same flat list and the writer stays a byte-stable fixpoint.
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

function sparklineExtXml(sparklines: ReadonlyArray<ParsedSparkline> | undefined): string {
  if (!sparklines || sparklines.length === 0) return '';
  const groups = sparklines
    .map((sp) => {
      const type =
        sp.kind === 'column' ? ' type="column"' : sp.kind === 'winLoss' ? ' type="stacked"' : '';
      const color = sp.colorHex ? `<x14:colorSeries rgb="FF${sp.colorHex}"/>` : '';
      return (
        `<x14:sparklineGroup${type}>${color}<x14:sparklines>` +
        `<x14:sparkline><xm:f>${escapeText(sp.dataRange)}</xm:f>` +
        `<xm:sqref>${escapeText(sp.sqref)}</xm:sqref></x14:sparkline>` +
        '</x14:sparklines></x14:sparklineGroup>'
      );
    })
    .join('');
  return (
    `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}">` +
    `<x14:sparklineGroups xmlns:xm="${XM_NS}">${groups}</x14:sparklineGroups></ext></extLst>`
  );
}

// A range box → A1 reference (a single cell drops the redundant ":A1").
function rangeRef(r: {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}): string {
  const a = cellRef(r.startRow, r.startColumn);
  if (r.startRow === r.endRow && r.startColumn === r.endColumn) return a;
  return `${a}:${cellRef(r.endRow, r.endColumn)}`;
}

function printOptionsXml(po: XlsxPrintOptions | undefined): string {
  if (!po) return '';
  const flag = (name: string, v: boolean | undefined): string =>
    v !== undefined ? ` ${name}="${v ? '1' : '0'}"` : '';
  const attrs =
    flag('gridLines', po.gridLines) +
    flag('horizontalCentered', po.horizontalCentered) +
    flag('verticalCentered', po.verticalCentered);
  return attrs ? `<printOptions${attrs}/>` : '';
}

function pageMarginsXml(m: XlsxPageMargins | undefined): string {
  if (!m) return '';
  return (
    `<pageMargins left="${m.leftInches}" right="${m.rightInches}"` +
    ` top="${m.topInches}" bottom="${m.bottomInches}"` +
    (m.headerInches !== undefined ? ` header="${m.headerInches}"` : '') +
    (m.footerInches !== undefined ? ` footer="${m.footerInches}"` : '') +
    '/>'
  );
}

function pageSetupXml(s: XlsxPageSetup | undefined): string {
  if (!s) return '';
  const attr = (name: string, v: number | string | undefined): string =>
    v !== undefined ? ` ${name}="${v}"` : '';
  const attrs =
    attr('paperSize', s.paperSize) +
    attr('orientation', s.orientation) +
    attr('scale', s.scale) +
    attr('fitToWidth', s.fitToWidth) +
    attr('fitToHeight', s.fitToHeight);
  return attrs ? `<pageSetup${attrs}/>` : '';
}

function breaksXml(tag: string, breaks: ReadonlyArray<number> | undefined): string {
  if (!breaks || breaks.length === 0) return '';
  const brks = breaks.map((id) => `<brk id="${id}" man="1"/>`).join('');
  return `<${tag} count="${breaks.length}" manualBreakCount="${breaks.length}">${brks}</${tag}>`;
}

function cellXml(cell: WorksheetCell): string {
  const ref = cellRef(cell.row, cell.column);
  const sAttr = cell.styleIndex !== undefined ? ` s="${cell.styleIndex}"` : '';
  // `n` is the default type and is emitted without a t attribute (round-trips).
  const tAttr = cell.type === 'n' ? '' : ` t="${cell.type}"`;
  if (cell.type === 'inlineStr') {
    const text = cell.inlineText ?? '';
    return `<c r="${ref}"${sAttr}${tAttr}><is><t xml:space="preserve">${escapeText(text)}</t></is></c>`;
  }
  if (cell.rawValue === '') return `<c r="${ref}"${sAttr}${tAttr}/>`;
  return `<c r="${ref}"${sAttr}${tAttr}><v>${escapeText(cell.rawValue)}</v></c>`;
}

// ── sharedStrings.xml ──────────────────────────────────────────────────────

function sharedStringsXml(strings: ReadonlyArray<string>): string {
  const items = strings
    .map((s) => `<si><t xml:space="preserve">${escapeText(s)}</t></si>`)
    .join('');
  return (
    XML_DECL +
    `<sst xmlns="${MAIN_NS}" count="${strings.length}" uniqueCount="${strings.length}">` +
    items +
    '</sst>'
  );
}

// ── styles.xml ─────────────────────────────────────────────────────────────

function stylesXml(styles: XlsxStyles): string {
  const numFmts =
    styles.numFmts.size > 0
      ? `<numFmts count="${styles.numFmts.size}">${[...styles.numFmts]
          .map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`)
          .join('')}</numFmts>`
      : '';
  const fonts = `<fonts count="${styles.fonts.length}">${styles.fonts.map(fontXml).join('')}</fonts>`;
  const fills = `<fills count="${styles.fills.length}">${styles.fills.map(fillXml).join('')}</fills>`;
  const borders = `<borders count="${styles.borders.length}">${styles.borders
    .map(borderXml)
    .join('')}</borders>`;
  const cellStyleXfs =
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
  const cellXfs = `<cellXfs count="${styles.cellXfs.length}">${styles.cellXfs
    .map(cellXfXml)
    .join('')}</cellXfs>`;
  const dxfs =
    styles.dxfs && styles.dxfs.length > 0
      ? `<dxfs count="${styles.dxfs.length}">${styles.dxfs
          .map(
            (d) =>
              `<dxf>${d.font ? `<font>${fontInner(d.font)}</font>` : ''}${
                d.fill ? `<fill>${patternFill(d.fill)}</fill>` : ''
              }</dxf>`,
          )
          .join('')}</dxfs>`
      : '';
  return (
    XML_DECL +
    `<styleSheet xmlns="${MAIN_NS}">` +
    numFmts +
    fonts +
    fills +
    borders +
    cellStyleXfs +
    cellXfs +
    dxfs +
    '</styleSheet>'
  );
}

function fontXml(f: XlsxFont): string {
  return `<font>${fontInner(f)}</font>`;
}

function fontInner(f: XlsxFont): string {
  return (
    (f.sizePt !== undefined ? `<sz val="${f.sizePt}"/>` : '') +
    (f.bold ? '<b/>' : '') +
    (f.italic ? '<i/>' : '') +
    (f.underline ? '<u/>' : '') +
    (f.colorHex ? `<color rgb="FF${f.colorHex}"/>` : '') +
    (f.name ? `<name val="${escapeAttr(f.name)}"/>` : '')
  );
}

function fillXml(f: XlsxFill): string {
  return `<fill>${patternFill(f)}</fill>`;
}

function patternFill(f: XlsxFill): string {
  const inner =
    (f.fgColorHex ? `<fgColor rgb="FF${f.fgColorHex}"/>` : '') +
    (f.bgColorHex ? `<bgColor rgb="FF${f.bgColorHex}"/>` : '');
  return `<patternFill${f.patternType ? ` patternType="${f.patternType}"` : ''}>${inner}</patternFill>`;
}

function borderXml(b: XlsxBorder): string {
  return `<border>${edgeXml('left', b.left)}${edgeXml('right', b.right)}${edgeXml(
    'top',
    b.top,
  )}${edgeXml('bottom', b.bottom)}</border>`;
}

function edgeXml(name: string, edge: XlsxBorderEdge | undefined): string {
  if (!edge || !edge.style) return `<${name}/>`;
  const color = edge.colorHex ? `<color rgb="FF${edge.colorHex}"/>` : '';
  return `<${name} style="${edge.style}">${color}</${name}>`;
}

function cellXfXml(xf: XlsxCellXf): string {
  const apply =
    applyAttr('applyNumberFormat', xf.applyNumberFormat) +
    applyAttr('applyFont', xf.applyFont) +
    applyAttr('applyFill', xf.applyFill) +
    applyAttr('applyBorder', xf.applyBorder) +
    applyAttr('applyAlignment', xf.applyAlignment);
  const a = xf.alignment;
  const alignment = a
    ? `<alignment${a.horizontal ? ` horizontal="${a.horizontal}"` : ''}${
        a.vertical ? ` vertical="${a.vertical}"` : ''
      }${a.wrapText ? ' wrapText="1"' : ''}/>`
    : '';
  return (
    `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}"${apply}>` +
    alignment +
    '</xf>'
  );
}

function applyAttr(name: string, value: boolean | undefined): string {
  return value !== undefined ? ` ${name}="${value ? '1' : '0'}"` : '';
}

// ── helpers ────────────────────────────────────────────────────────────────

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// 0-based (row, col) → an A1 reference (column letters + 1-based row).
function cellRef(row: number, col: number): string {
  let n = col + 1;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row + 1}`;
}

function escapeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, '&quot;');
}
