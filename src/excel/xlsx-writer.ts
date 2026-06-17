// XLSX writer (E-SHEET SD1) — SheetDoc → .xlsx. The first writer that consumes
// the SpreadsheetML IR node directly (not the FlowDoc projection): a meaningful
// spreadsheet writer is only possible on the grid tree, because the cells carry
// their RAW value + style index, which the print-model projection deliberately
// resolves away. Re-emits workbook / worksheets / sharedStrings / styles through
// the core OPC writer, so the SheetDoc is a stable round-trip fixpoint (SD2).
//
// Writes the full grid surface: cell values, shared strings, styles (numFmts/
// fonts/fills/borders/cellXfs/dxfs), merges, column widths, row heights, the
// page model (margins/setup/print options/breaks), conditional formatting,
// sparklines (extLst), table parts (SD1 + SD3a/b) and embedded charts — a sheet's
// drawing + chart parts, re-serialized from the parsed Chart model (WT1).

import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { Sheet, SheetChartRef, SheetDoc } from '@/core/ir/sheet';
import type { Loss } from '@/core/ir';
import type { OpcPart, OpcPartRelationships, Relationship } from '@/core/opc';
import type {
  CfRule,
  Cfvo,
  ConditionalFormat,
  DataValidation,
  DefinedName,
  ExcelTable,
  ParsedSparkline,
  ParsedWorksheet,
  SheetPane,
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
import { chartSpaceXml } from '@/core/drawingml/chart-serializer';
import { buildOpcPackage } from '@/core/opc';

const encoder = new TextEncoder();
const EMU_PER_PT = 12700; // §20.1.2.1

const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_WORKSHEET =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_SHARED_STRINGS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_SHARED_STRINGS =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function writeXlsx(sheet: SheetDoc): WriteResult {
  const losses: Array<Loss> = [];

  const worksheetParts: Array<OpcPart> = [];
  const tableParts: Array<OpcPart> = [];
  const drawingParts: Array<OpcPart> = []; // drawingN.xml + chartN.xml (WT1)
  const worksheetRels: Array<OpcPartRelationships> = [];
  const drawingRels: Array<OpcPartRelationships> = [];
  const sheetEntries: Array<{ name: string; sheetId: number; rid: string }> = [];
  const workbookRels: Array<Relationship> = [];
  let tableCounter = 0;
  let drawingCounter = 0;
  let chartCounter = 0;

  sheet.sheets.forEach((s, i) => {
    const idx = i + 1;
    const rid = `rId${idx}`;
    const wsPath = `xl/worksheets/sheet${idx}.xml`;

    const wsRels: Array<Relationship> = [];

    // §18.5 table parts (E-SHEET SD3b): one tableN.xml per table + a worksheet
    // relationship; the worksheet's <tableParts> references them by rId.
    const tableRelIds: Array<string> = [];
    for (const t of s.grid.tables ?? []) {
      const tid = ++tableCounter;
      const tRid = `rIdTbl${tid}`;
      tableRelIds.push(tRid);
      tableParts.push({
        path: `xl/tables/table${tid}.xml`,
        data: encoder.encode(tableXml(t, tid)),
        contentType: CT_TABLE,
      });
      wsRels.push(rel(tRid, REL_TABLE, `../tables/table${tid}.xml`));
    }

    // §20.5 embedded charts (WT1): one drawingN.xml carrying the sheet's chart
    // frames, each referencing a chartN.xml part the chart-serializer emits.
    const drawingRelId = emitSheetCharts(s, sheet, idx, {
      drawingParts,
      drawingRels,
      wsRels,
      nextDrawing: () => ++drawingCounter,
      nextChart: () => ++chartCounter,
      losses,
    });

    if (wsRels.length > 0) worksheetRels.push({ sourcePart: wsPath, relationships: wsRels });

    worksheetParts.push({
      path: wsPath,
      data: encoder.encode(worksheetXml(s.grid, tableRelIds, drawingRelId)),
      contentType: CT_WORKSHEET,
    });
    workbookRels.push(rel(rid, REL_WORKSHEET, `worksheets/sheet${idx}.xml`));
    sheetEntries.push({ name: s.name, sheetId: idx, rid });
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
      ...tableParts,
      ...drawingParts,
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
    partRelationships: [
      { sourcePart: 'xl/workbook.xml', relationships: workbookRels },
      ...worksheetRels,
      ...drawingRels,
    ],
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

// ── embedded charts (WT1) ──────────────────────────────────────────────────

interface ChartEmitCtx {
  readonly drawingParts: Array<OpcPart>;
  readonly drawingRels: Array<OpcPartRelationships>;
  readonly wsRels: Array<Relationship>;
  readonly nextDrawing: () => number;
  readonly nextChart: () => number;
  readonly losses: Array<Loss>;
}

// Emit a sheet's chart frames: one drawingN.xml (the spreadsheetDrawing) plus a
// chartN.xml per frame (the serialized Chart). Returns the worksheet relationship
// id of the drawing, or undefined when the sheet has no resolvable charts.
function emitSheetCharts(
  s: Sheet,
  doc: SheetDoc,
  idx: number,
  ctx: ChartEmitCtx,
): string | undefined {
  const refs = s.charts ?? [];
  if (refs.length === 0 || !doc.chartData) return undefined;

  const did = ctx.nextDrawing();
  const drawingPath = `xl/drawings/drawing${did}.xml`;
  const anchors: Array<string> = [];
  const dRels: Array<Relationship> = [];
  refs.forEach((ref, ci) => {
    const chart = doc.chartData?.get(ref.chartPartPath);
    if (!chart) {
      ctx.losses.push({
        severity: 'dropped',
        feature: FEATURES.charts,
        detail: `sheet "${s.name}": a chart frame had no resolved data and was dropped`,
      });
      return;
    }
    const cid = ctx.nextChart();
    ctx.drawingParts.push({
      path: `xl/charts/chart${cid}.xml`,
      data: encoder.encode(chartSpaceXml(chart)),
      contentType: CT_CHART,
    });
    const cRel = `rIdChart${ci + 1}`;
    dRels.push(rel(cRel, REL_CHART, `../charts/chart${cid}.xml`));
    anchors.push(chartAnchorXml(ref, cRel, ci));
  });
  if (anchors.length === 0) return undefined;

  ctx.drawingParts.push({
    path: drawingPath,
    data: encoder.encode(
      `${XML_DECL}<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${anchors.join('')}</xdr:wsDr>`,
    ),
    contentType: CT_DRAWING,
  });
  ctx.drawingRels.push({ sourcePart: drawingPath, relationships: dRels });
  const dRid = `rIdDraw${idx}`;
  ctx.wsRels.push(rel(dRid, REL_DRAWING, `../drawings/drawing${did}.xml`));
  return dRid;
}

// A one-cell anchor placing one chart frame; the size comes from the parsed
// SheetChartRef (the original two-cell anchor was already resolved to points).
// Frames are stacked vertically so they do not overlap on re-open.
function chartAnchorXml(ref: SheetChartRef, chartRelId: string, index: number): string {
  const cx = Math.round(ref.widthPt * EMU_PER_PT);
  const cy = Math.round(ref.heightPt * EMU_PER_PT);
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${index * 16}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    '<xdr:graphicFrame macro="">' +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 2}" name="Chart ${index + 1}"/>` +
    '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${chartRelId}"/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame>' +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
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

function worksheetXml(
  grid: ParsedWorksheet,
  tableRelIds: ReadonlyArray<string> = [],
  drawingRelId?: string,
): string {
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
    sheetViewsXml(grid.pane) +
    colsXml +
    `<sheetData>${rowsXml}</sheetData>` +
    mergesXml +
    conditionalFormattingXml(grid.conditionalFormats) +
    dataValidationsXml(grid.dataValidations) +
    printOptionsXml(grid.printOptions) +
    pageMarginsXml(grid.pageMargins) +
    pageSetupXml(grid.pageSetup) +
    breaksXml('rowBreaks', grid.rowBreaks) +
    breaksXml('colBreaks', grid.colBreaks) +
    // §18.3.1.99 child order: <drawing> precedes <tableParts>, both before extLst.
    (drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : '') +
    (tableRelIds.length > 0
      ? `<tableParts count="${tableRelIds.length}">${tableRelIds
          .map((id) => `<tablePart r:id="${id}"/>`)
          .join('')}</tableParts>`
      : '') +
    sparklineExtXml(grid.sparklines) +
    '</worksheet>'
  );
}

// §18.3.1.66 <sheetViews><sheetView><pane> — re-emit a frozen pane so the freeze
// survives a round-trip. The reader reads only xSplit/ySplit/state; topLeftCell
// and activePane are written for Excel's benefit but ignored on re-read.
function sheetViewsXml(pane: SheetPane | undefined): string {
  if (!pane || (pane.frozenRows <= 0 && pane.frozenCols <= 0)) return '';
  const { frozenRows, frozenCols } = pane;
  const topLeftCell = cellRef(frozenRows, frozenCols);
  const activePane =
    frozenRows > 0 && frozenCols > 0 ? 'bottomRight' : frozenRows > 0 ? 'bottomLeft' : 'topRight';
  const attrs =
    (frozenCols > 0 ? ` xSplit="${frozenCols}"` : '') +
    (frozenRows > 0 ? ` ySplit="${frozenRows}"` : '') +
    ` topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"`;
  return `<sheetViews><sheetView workbookViewId="0"><pane${attrs}/></sheetView></sheetViews>`;
}

// §18.5.1.2 xl/tables/tableN.xml. The reader keeps no column names, so generic
// tableColumns are synthesized (it ignores them on re-read); the resolved
// header/band colours are NOT written — the reader re-derives them from the
// style name + theme (E-SHEET SD3b).
function tableXml(t: ExcelTable, id: number): string {
  const ref = rangeRef(t.ref);
  const name = t.name ?? `Table${id}`;
  const ncols = t.ref.endColumn - t.ref.startColumn + 1;
  const columns = Array.from(
    { length: ncols },
    (_, c) => `<tableColumn id="${c + 1}" name="Column${c + 1}"/>`,
  ).join('');
  const styleInfo = t.styleName
    ? `<tableStyleInfo name="${escapeAttr(t.styleName)}" showFirstColumn="${
        t.showFirstColumn ? 1 : 0
      }" showLastColumn="${t.showLastColumn ? 1 : 0}" showRowStripes="${
        t.showRowStripes ? 1 : 0
      }" showColumnStripes="${t.showColumnStripes ? 1 : 0}"/>`
    : '';
  return (
    XML_DECL +
    `<table xmlns="${MAIN_NS}" id="${id}" name="${escapeAttr(name)}" displayName="${escapeAttr(
      name,
    )}" ref="${ref}"${t.headerRowCount !== 1 ? ` headerRowCount="${t.headerRowCount}"` : ''} totalsRowShown="0">` +
    (t.autoFilter ? `<autoFilter ref="${ref}"/>` : '') +
    `<tableColumns count="${ncols}">${columns}</tableColumns>` +
    styleInfo +
    '</table>'
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
    case 'top10':
      return (
        `<cfRule type="top10"${p} dxfId="${rule.dxfId}" rank="${rule.rank}"` +
        (rule.percent ? ' percent="1"' : '') +
        (rule.bottom ? ' bottom="1"' : '') +
        '/>'
      );
    case 'aboveAverage':
      return (
        `<cfRule type="aboveAverage"${p} dxfId="${rule.dxfId}"` +
        // aboveAverage defaults to true; only the false case needs the attribute.
        (rule.aboveAverage ? '' : ' aboveAverage="0"') +
        (rule.equalAverage ? ' equalAverage="1"' : '') +
        (rule.stdDev !== undefined ? ` stdDev="${rule.stdDev}"` : '') +
        '/>'
      );
    case 'duplicateValues':
    case 'uniqueValues':
      return `<cfRule type="${rule.type}"${p} dxfId="${rule.dxfId}"/>`;
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return (
        `<cfRule type="${rule.type}"${p} operator="${rule.type}" dxfId="${rule.dxfId}"` +
        ` text="${escapeAttr(rule.text)}">` +
        (rule.formula !== undefined ? `<formula>${escapeText(rule.formula)}</formula>` : '') +
        '</cfRule>'
      );
  }
}

function cfvoXml(cfvo: Cfvo): string {
  return `<cfvo type="${cfvo.type}"${cfvo.val !== undefined ? ` val="${escapeAttr(cfvo.val)}"` : ''}/>`;
}

// §18.3.1.32 <dataValidations> — per-range input constraints written back so the
// SheetDoc stays a byte-stable fixpoint (E-SHEET SV1). Every field round-trips;
// `showDropDown` keeps ECMA's inverted sense (the attr is "1" to HIDE the in-cell
// dropdown). Sits between <conditionalFormatting> and <printOptions> per
// §18.3.1.99 child order.
function dataValidationsXml(dvs: ReadonlyArray<DataValidation> | undefined): string {
  if (!dvs || dvs.length === 0) return '';
  return `<dataValidations count="${dvs.length}">${dvs.map(dataValidationXml).join('')}</dataValidations>`;
}

function dataValidationXml(dv: DataValidation): string {
  const sqref = dv.ranges.map(rangeRef).join(' ');
  const attrs =
    (dv.type !== 'none' ? ` type="${dv.type}"` : '') +
    (dv.operator !== undefined ? ` operator="${escapeAttr(dv.operator)}"` : '') +
    (dv.allowBlank ? ' allowBlank="1"' : '') +
    (dv.showDropDown ? ' showDropDown="1"' : '') +
    (dv.showInputMessage ? ' showInputMessage="1"' : '') +
    (dv.showErrorMessage ? ' showErrorMessage="1"' : '') +
    (dv.errorStyle !== undefined ? ` errorStyle="${escapeAttr(dv.errorStyle)}"` : '') +
    (dv.promptTitle !== undefined ? ` promptTitle="${escapeAttr(dv.promptTitle)}"` : '') +
    (dv.prompt !== undefined ? ` prompt="${escapeAttr(dv.prompt)}"` : '') +
    (dv.errorTitle !== undefined ? ` errorTitle="${escapeAttr(dv.errorTitle)}"` : '') +
    (dv.error !== undefined ? ` error="${escapeAttr(dv.error)}"` : '') +
    ` sqref="${sqref}"`;
  const children =
    (dv.formula1 !== undefined ? `<formula1>${escapeText(dv.formula1)}</formula1>` : '') +
    (dv.formula2 !== undefined ? `<formula2>${escapeText(dv.formula2)}</formula2>` : '');
  return `<dataValidation${attrs}>${children}</dataValidation>`;
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
