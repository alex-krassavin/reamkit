import type {
  Alignment,
  BodyElement,
  Border,
  BorderStyle,
  CellBorders,
  CellProperties,
  CellShading,
  PageMargins,
  PageSize,
  ParagraphProperties,
  RowProperties,
  RunProperties,
  SectionProperties,
  Table,
  TableCell,
  TableRow,
} from '@/document-model';
import type { FontBytesByVariant } from '@/font';
import type { FetchLike } from '@/fonts';
import type { CoreProperties } from '@/opc';
import type {
  CellRange,
  DefinedName,
  MergedRange,
  ParsedWorksheet,
  WorksheetCell,
  XlsxBorder,
  XlsxBorderEdge,
  XlsxBorderStyleName,
  XlsxCellAlignment,
  XlsxCellXf,
  XlsxFont,
  XlsxHorizontalAlign,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxStyles,
} from '@/ooxml/spreadsheet';
import type { DocumentInfo, SignatureOptions, StyledRenderOptions } from '@/pdf';
import { eighthPtToPt, halfPtToPt, twipsToPt } from '@/ir';
import { FontRegistry } from '@/font';
import { fetchFontSet } from '@/fonts';
import { OpcPackage, parseCoreProperties } from '@/opc';
import {
  EMPTY_XLSX_STYLES,
  applyNumberFormat,
  parseAreaRef,
  parseSharedStrings,
  parseTitleRowRange,
  parseWorkbook,
  parseWorksheet,
  parseXlsxStyles,
} from '@/ooxml/spreadsheet';
import { EMPTY_STYLE_SHEET } from '@/ooxml/wordproc';
import { renderStyledPdf, signPdf } from '@/pdf';

const WORKBOOK_PART = 'xl/workbook.xml';
const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
const STYLES_PART = 'xl/styles.xml';
const CORE_PROPS_PART = 'docProps/core.xml';

// Excel "character width" → twips. Calibri 11pt default Maximum Digit Width
// is ~7 px ≈ 5.25 pt ≈ 105 twips. This is a coarse approximation but the
// auto-fit pass refines column widths against the actual cell text anyway.
const TWIPS_PER_EXCEL_CHAR = 105;

// Excel's default column width is 8.43 "characters" ≈ 64px ≈ 960 twips. Used for
// columns without an explicit <col width="..">.
const DEFAULT_COL_TWIPS = 960;

// Excel's default row height is ~15pt = 300 twips. Used (for the fitToHeight
// estimate) for rows without an explicit <row ht="..">.
const DEFAULT_ROW_TWIPS = 300;

// 1 point = 20 twips (Word/Excel unit conversion).
const TWIPS_PER_POINT = 20;

// 1 inch = 1440 twips.
const TWIPS_PER_INCH = 1440;

// ECMA-376 Part 1 §18.3.1.63 — pageSetup paperSize enumeration. Values map
// to twip-precision width × height (portrait). Landscape swaps the pair.
// Only the common sizes; anything else falls back to A4.
const PAPER_SIZES_TWIPS: ReadonlyMap<number, readonly [number, number]> = new Map([
  [1, [12240, 15840]], // Letter 8.5" × 11"
  [3, [15840, 24480]], // Tabloid 11" × 17"
  [5, [12240, 20160]], // Legal 8.5" × 14"
  [8, [16838, 23811]], // A3 297mm × 420mm
  [9, [11906, 16838]], // A4 210mm × 297mm
  [11, [8392, 11906]], // A5 148mm × 210mm
  [70, [5953, 8392]], // A6 105mm × 148mm
]);
const DEFAULT_PAPER_TWIPS: readonly [number, number] = [11906, 16838];

export interface ConvertXlsxOptions extends Omit<StyledRenderOptions, 'registry' | 'styles'> {
  readonly fontBytes?: Uint8Array;
  readonly fonts?: FontBytesByVariant;
  // Force a substitute font family for the auto-download path. Ignored when
  // `fonts`/`fontBytes` are supplied.
  readonly fontFamily?: string;
  // Injectable fetch for the auto-download path (defaults to global fetch).
  readonly fontFetch?: FetchLike;
  // For PDF/A-3 only: embed the input .xlsx as an associated source file
  // (/AFRelationship /Source). Ignored for other profiles.
  readonly embedSource?: boolean;
  // Digitally sign the output (ISO 32000 §12.8). Requires the async
  // convertXlsxToPdf (signing uses WebCrypto). Ignored by the sync converter.
  readonly signature?: SignatureOptions;
}

// Convert a .xlsx to PDF, downloading an open substitute font automatically
// when the caller does not supply one (no network when `fonts` is set).
export async function convertXlsxToPdf(
  xlsx: Uint8Array,
  options: ConvertXlsxOptions = {},
): Promise<Uint8Array> {
  // A signature (SignatureOptions) is also a SignaturePlaceholder, so it doubles
  // as the render-time placeholder; the crypto runs afterwards on the result.
  const { signature } = options;
  const renderOptions: ConvertXlsxOptions = signature
    ? { ...options, signaturePlaceholder: signature }
    : options;
  let pdf: Uint8Array;
  if (renderOptions.fonts ?? renderOptions.fontBytes) {
    pdf = convertXlsxToPdfSync(xlsx, renderOptions);
  } else {
    const fonts = await fetchFontSet({
      ...(options.fontFamily ? { family: options.fontFamily } : {}),
      ...(options.fontFetch ? { fetch: options.fontFetch } : {}),
    });
    pdf = convertXlsxToPdfSync(xlsx, { ...renderOptions, fonts });
  }
  return signature ? signPdf(pdf, signature) : pdf;
}

// Synchronous conversion. Requires the caller to supply `fonts`/`fontBytes`.
export function convertXlsxToPdfSync(xlsx: Uint8Array, options: ConvertXlsxOptions): Uint8Array {
  const fonts: FontBytesByVariant | undefined =
    options.fonts ?? (options.fontBytes ? { regular: options.fontBytes } : undefined);
  if (!fonts) {
    throw new Error('convertXlsxToPdfSync requires options.fonts or options.fontBytes');
  }
  const registry = FontRegistry.fromBytes(fonts);

  const pkg = OpcPackage.open(xlsx);
  const workbookData = pkg.getPart(WORKBOOK_PART);
  if (!workbookData) throw new Error('Not a valid xlsx: missing xl/workbook.xml');
  const { sheets, date1904, definedNames } = parseWorkbook(workbookData);
  if (sheets.length === 0) throw new Error('xlsx has no sheets');

  const sharedStringsData = pkg.getPart(SHARED_STRINGS_PART);
  const sharedStrings = sharedStringsData ? parseSharedStrings(sharedStringsData) : [];

  const stylesData = pkg.getPart(STYLES_PART);
  const styles = stylesData ? parseXlsxStyles(stylesData) : EMPTY_XLSX_STYLES;

  const workbookRels = pkg.getPartRelationships(WORKBOOK_PART);
  const body: Array<BodyElement> = [];

  // Page geometry comes from the first sheet's <pageSetup>/<pageMargins>.
  // The renderer only supports one section, so subsequent sheets share the
  // first sheet's geometry; multi-section support is M2/M4-grade work.
  let firstSheetSection: SectionProperties | undefined;

  for (let sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    const sheet = sheets[sheetIdx]!;
    const sheetRel = workbookRels.find((r) => r.id === sheet.relationshipId);
    if (!sheetRel) continue;
    const resolved = pkg.resolveRelatedPart(WORKBOOK_PART, sheetRel);
    if (!resolved) continue;
    const worksheet = parseWorksheet(resolved.data);
    if (sheetIdx === 0) {
      firstSheetSection = sectionFromWorksheet(worksheet);
    }

    // Each sheet after the first starts on its own PDF page. We do NOT print the
    // sheet name: LibreOffice Calc / Excel `--convert-to pdf` emit it nowhere in
    // the body (nor the default header) — a synthetic title is pure extra text
    // that diverges from the print golden. So emit an empty page-break-only
    // paragraph (no runs ⇒ no glyphs ⇒ no text) for sheets > 0.
    if (sheetIdx > 0) {
      body.push({
        kind: 'paragraph',
        paragraph: { properties: { pageBreakBefore: true }, runs: [] },
      });
    }

    const printArea = resolvePrintArea(definedNames, sheetIdx);
    const titleRows = resolvePrintTitleRows(definedNames, sheetIdx);
    const gridLines = worksheet.printOptions?.gridLines === true;
    body.push(
      ...worksheetToBody(worksheet, sharedStrings, styles, date1904, {
        ...(printArea ? { printArea } : {}),
        ...(titleRows ? { titleRows } : {}),
        gridLines,
      }),
    );
  }

  const coreData = pkg.getPart(CORE_PROPS_PART);
  const coreProps = coreData ? parseCoreProperties(coreData) : undefined;

  const {
    fontBytes: _ignoreA,
    fonts: _ignoreB,
    section: sectionOverride,
    info: callerInfo,
    embedSource,
    attachments: callerAttachments,
    signature: _ignoreSig,
    ...renderOptions
  } = options;
  void _ignoreSig;
  void _ignoreA;
  void _ignoreB;
  const section = sectionOverride ?? firstSheetSection;
  const info = mergeInfo(coreProps, callerInfo);
  // PDF/A-3 only: optionally embed the input .xlsx as an associated source file.
  const attachments = [...(callerAttachments ?? [])];
  if (embedSource && options.pdfA?.startsWith('PDF/A-3')) {
    attachments.push({
      name: 'source.xlsx',
      bytes: xlsx,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      relationship: 'Source',
      description: 'Source Excel workbook',
    });
  }
  return renderStyledPdf(body, {
    registry,
    styles: EMPTY_STYLE_SHEET,
    ...(section ? { section } : {}),
    ...(info ? { info } : {}),
    ...(attachments.length > 0 ? { attachments } : {}),
    ...renderOptions,
  });
}

function mergeInfo(
  core: CoreProperties | undefined,
  caller: DocumentInfo | undefined,
): DocumentInfo | undefined {
  if (!core && !caller) return undefined;
  return {
    ...(core?.title ? { title: core.title } : {}),
    ...(core?.creator ? { author: core.creator } : {}),
    ...(core?.subject ? { subject: core.subject } : {}),
    ...(core?.keywords ? { keywords: core.keywords } : {}),
    ...(core?.created ? { creationDate: core.created } : {}),
    ...(core?.modified ? { modificationDate: core.modified } : {}),
    ...caller,
  };
}

function sectionFromWorksheet(worksheet: ParsedWorksheet): SectionProperties | undefined {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  const margins = marginsFromXlsx(worksheet.pageMargins);
  if (!pageSize && !margins) return undefined;
  return {
    ...(pageSize ? { pageSize } : {}),
    ...(margins ? { margins } : {}),
    headers: [],
    footers: [],
  };
}

function pageSizeFromSetup(setup: XlsxPageSetup | undefined): PageSize | undefined {
  if (!setup) return undefined;
  const paper =
    setup.paperSize !== undefined
      ? (PAPER_SIZES_TWIPS.get(setup.paperSize) ?? DEFAULT_PAPER_TWIPS)
      : DEFAULT_PAPER_TWIPS;
  const [w, h] = setup.orientation === 'landscape' ? [paper[1], paper[0]] : paper;
  const orientation = setup.orientation === 'landscape' ? 'landscape' : 'portrait';
  // Only emit a PageSize when paperSize or a non-default orientation was set;
  // otherwise let the renderer apply its A4 default.
  if (setup.paperSize === undefined && setup.orientation !== 'landscape') return undefined;
  return { width: twipsToPt(w), height: twipsToPt(h), orientation };
}

function marginsFromXlsx(margins: XlsxPageMargins | undefined): PageMargins | undefined {
  if (!margins) return undefined;
  return {
    top: twipsToPt(Math.round(margins.topInches * TWIPS_PER_INCH)),
    right: twipsToPt(Math.round(margins.rightInches * TWIPS_PER_INCH)),
    bottom: twipsToPt(Math.round(margins.bottomInches * TWIPS_PER_INCH)),
    left: twipsToPt(Math.round(margins.leftInches * TWIPS_PER_INCH)),
    ...(margins.headerInches !== undefined
      ? { header: twipsToPt(Math.round(margins.headerInches * TWIPS_PER_INCH)) }
      : {}),
    ...(margins.footerInches !== undefined
      ? { footer: twipsToPt(Math.round(margins.footerInches * TWIPS_PER_INCH)) }
      : {}),
  };
}

const PRINT_AREA_NAME = '_xlnm.Print_Area';
const PRINT_TITLES_NAME = '_xlnm.Print_Titles';

// ECMA-376 §18.2.5 — resolve the sheet-scoped _xlnm.Print_Area defined name
// (localSheetId = the sheet's 0-based index) into a clipping range.
function resolvePrintArea(
  definedNames: ReadonlyArray<DefinedName>,
  sheetIdx: number,
): CellRange | undefined {
  for (const dn of definedNames) {
    if (dn.name === PRINT_AREA_NAME && dn.localSheetId === sheetIdx) {
      return parseAreaRef(dn.value);
    }
  }
  return undefined;
}

// ECMA-376 §18.2.5 — _xlnm.Print_Titles → the repeated row range (0-indexed).
function resolvePrintTitleRows(
  definedNames: ReadonlyArray<DefinedName>,
  sheetIdx: number,
): { readonly startRow: number; readonly endRow: number } | undefined {
  for (const dn of definedNames) {
    if (dn.name === PRINT_TITLES_NAME && dn.localSheetId === sheetIdx) {
      return parseTitleRowRange(dn.value);
    }
  }
  return undefined;
}

// ECMA-376 §18.3.1.63/§18.3.1.65 — print scaling. Excel either honors an
// explicit <pageSetup scale="N"> (percent) or, when <pageSetUpPr fitToPage="1">,
// shrinks the sheet so it fits `fitToWidth` pages across. We model both as a
// single uniform factor applied to fonts + row heights (the renderer's auto-fit
// then packs the now-smaller text without the aggressive wrapping it would do at
// full size). Shrink-only: enlarging interacts poorly with auto-fit and Excel's
// fit-to-page never enlarges. Floor at Excel's 10% minimum.
const MIN_PRINT_SCALE = 0.1;

function sheetContentWidthTwips(worksheet: ParsedWorksheet): number {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  const pageWidthTwips = pageSize ? Math.round(pageSize.width * 20) : DEFAULT_PAPER_TWIPS[0];
  const margins = marginsFromXlsx(worksheet.pageMargins);
  const left = margins ? Math.round(margins.left * 20) : TWIPS_PER_INCH;
  const right = margins ? Math.round(margins.right * 20) : TWIPS_PER_INCH;
  return Math.max(TWIPS_PER_INCH / 2, pageWidthTwips - left - right);
}

function sheetContentHeightTwips(worksheet: ParsedWorksheet): number {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  const pageHeightTwips = pageSize ? Math.round(pageSize.height * 20) : DEFAULT_PAPER_TWIPS[1];
  const margins = marginsFromXlsx(worksheet.pageMargins);
  const top = margins ? Math.round(margins.top * 20) : TWIPS_PER_INCH;
  const bottom = margins ? Math.round(margins.bottom * 20) : TWIPS_PER_INCH;
  return Math.max(TWIPS_PER_INCH / 2, pageHeightTwips - top - bottom);
}

function computePrintScale(
  worksheet: ParsedWorksheet,
  totalGridTwips: number,
  contentWidthTwips: number,
  totalGridHeightTwips: number,
  contentHeightTwips: number,
): number {
  const setup = worksheet.pageSetup;
  let s = 1;
  if (worksheet.fitToPage) {
    // fitToPage overrides any explicit scale. The binding (smaller) of the two
    // fit factors wins. fitToWidth defaults to 1 page (existing behavior);
    // fitToHeight constrains only when explicitly ≥1 (a width-only "fit all
    // columns" sheet has fitToHeight="0" and keeps flowing down as before).
    const fitW = setup?.fitToWidth ?? 1;
    let sW = 1;
    let sH = 1;
    if (fitW >= 1 && totalGridTwips > 0) sW = (contentWidthTwips * fitW) / totalGridTwips;
    const fitH = setup?.fitToHeight;
    if (fitH !== undefined && fitH >= 1 && totalGridHeightTwips > 0) {
      sH = (contentHeightTwips * fitH) / totalGridHeightTwips;
    }
    s = Math.min(sW, sH);
  } else if (setup?.scale !== undefined && setup.scale > 0) {
    s = setup.scale / 100;
  }
  return Math.min(1, Math.max(MIN_PRINT_SCALE, s));
}

// Excel's default cell font is 11pt (22 half-points); scale that when a run
// carries no explicit size so the whole sheet shrinks uniformly.
function scaleRunFont(props: RunProperties, scale: number): RunProperties {
  const hp = props.fontSizePt !== undefined ? Math.round(props.fontSizePt * 2) : 22;
  return { ...props, fontSizePt: halfPtToPt(Math.max(2, Math.round(hp * scale))) };
}

// Per-sheet budget on rendered characters — a DoS guard (see use site).
const MAX_SHEET_TEXT_CHARS = 1_000_000;

interface PrintModelOptions {
  // ECMA-376 §18.2.5 — _xlnm.Print_Area: render only this range (clipped to the
  // used range). Absent ⇒ the whole used range.
  readonly printArea?: CellRange;
  // ECMA-376 §18.3.1.70 — <printOptions gridLines="1">: when false (the Excel/
  // Calc default) NO synthetic cell gridlines are drawn — only borders that come
  // from cell styles. This is the dominant visual difference vs. a print preview.
  readonly gridLines: boolean;
  // ECMA-376 §18.2.5 — _xlnm.Print_Titles repeated row range (absolute, 0-indexed).
  // Rows in this range are flagged isHeader so the renderer repeats them at the
  // top of each continuation page.
  readonly titleRows?: { readonly startRow: number; readonly endRow: number };
}

function worksheetToBody(
  worksheet: ParsedWorksheet,
  sharedStrings: Array<string>,
  styles: XlsxStyles,
  date1904: boolean,
  print: PrintModelOptions,
): Array<BodyElement> {
  if (worksheet.maxRow < 0 || worksheet.maxColumn < 0) return [];

  // A worksheet can declare cells far beyond its real data — e.g. a style
  // applied to whole rows produces tens of thousands of EMPTY styled cells out
  // to column XFD (16384). Materializing a dense maxColumn-wide grid then
  // exhausts memory (CVLKRA-KYC: a 204 KB file blew >512 MB heap from a sheet
  // of 49 194 cells, only 48 of which carried a value). Bound the grid to the
  // "used range": the extent of cells that carry content (a value or inline
  // text) or are spanned by a merge. Empty styled cells outside it are dropped,
  // as LibreOffice/Excel clip to the used range anyway.
  let usedRow = -1;
  let usedCol = -1;
  for (const c of worksheet.cells) {
    if (c.rawValue !== '' || c.inlineText !== undefined) {
      if (c.row > usedRow) usedRow = c.row;
      if (c.column > usedCol) usedCol = c.column;
    }
  }
  for (const m of worksheet.merges) {
    if (m.endRow > usedRow) usedRow = m.endRow;
    if (m.endColumn > usedCol) usedCol = m.endColumn;
  }
  if (usedRow < 0 || usedCol < 0) return []; // nothing but empty styled cells

  // Print area (when defined) overrides the rendered window: Excel prints only
  // the _xlnm.Print_Area range. Clip it to the used range so a print area that
  // over-declares (e.g. to XFD) cannot re-introduce the dense-grid blow-up.
  let rowStart = 0;
  let colStart = 0;
  let rowEnd = usedRow;
  let colEnd = usedCol;
  if (print.printArea) {
    rowStart = Math.max(0, print.printArea.startRow);
    colStart = Math.max(0, print.printArea.startColumn);
    rowEnd = Math.min(usedRow, print.printArea.endRow);
    colEnd = Math.min(usedCol, print.printArea.endColumn);
  }
  if (rowEnd < rowStart || colEnd < colStart) return []; // empty print area

  // Defence-in-depth for untrusted input: even within the used range, cap the
  // materialized grid so a pathological sheet (real values seeded across the
  // 16384×1048576 cell space) cannot exhaust memory. A PDF table larger than
  // this is unreadable regardless.
  const MAX_GRID_COLS = 1024;
  const MAX_GRID_ROWS = 50_000;
  const rowCount = Math.min(rowEnd - rowStart + 1, MAX_GRID_ROWS);
  const colCount = Math.min(colEnd - colStart + 1, MAX_GRID_COLS);
  // Absolute index of the last in-window row/column (for merge clamping).
  const rowWindowEnd = rowStart + rowCount - 1;
  const colWindowEnd = colStart + colCount - 1;

  // Per-row height overrides keyed by LOCAL (window-relative) row index.
  const rowHeightMap = new Map<number, { heightTwips: number; heightRule: 'atLeast' }>();
  for (const h of worksheet.rowHeights) {
    const local = h.row - rowStart;
    if (local < 0 || local >= rowCount) continue;
    const heightTwips = Math.round(h.heightPt * TWIPS_PER_POINT);
    // customHeight="1" pins the row exactly; without it the height attr is
    // advisory. Use 'atLeast' in both cases so PDF wrapping (we have no
    // clipping) can still grow rows that need more vertical space than Excel
    // pinned them to — this avoids overlap at the cost of slight divergence
    // from Excel's exact-truncation behavior.
    rowHeightMap.set(local, { heightTwips, heightRule: 'atLeast' });
  }

  // Index cells by LOCAL (row - rowStart, col - colStart).
  const cellMatrix: Array<Array<WorksheetCell | undefined>> = Array.from(
    { length: rowCount },
    () => new Array<WorksheetCell | undefined>(colCount),
  );
  for (const cell of worksheet.cells) {
    const lr = cell.row - rowStart;
    const lc = cell.column - colStart;
    // Drop cells outside the materialized window (print area / trim / cap).
    if (lr >= 0 && lr < rowCount && lc >= 0 && lc < colCount) {
      cellMatrix[lr]![lc] = cell;
    }
  }

  // Index merges by ABSOLUTE (startRow, startCol) and mark non-origin positions
  // (also absolute). Only merges that intersect the window matter.
  const mergeOrigins = new Map<string, MergedRange>();
  const insideMerge = new Set<string>();
  for (const m of worksheet.merges) {
    if (m.endRow < rowStart || m.startRow > rowWindowEnd) continue;
    if (m.endColumn < colStart || m.startColumn > colWindowEnd) continue;
    mergeOrigins.set(key(m.startRow, m.startColumn), m);
    // Clamp the expansion to the materialized window so a pathological wide/tall
    // merge can't blow up the insideMerge set on untrusted input.
    const rEnd = Math.min(m.endRow, rowWindowEnd);
    const cEnd = Math.min(m.endColumn, colWindowEnd);
    for (let r = Math.max(m.startRow, rowStart); r <= rEnd; r++) {
      for (let c = Math.max(m.startColumn, colStart); c <= cEnd; c++) {
        if (!(r === m.startRow && c === m.startColumn)) insideMerge.add(key(r, c));
      }
    }
  }

  // Column widths (LOCAL index): prefer <col width="..">; fall back to Excel's
  // default column width (8.43 chars ≈ 960 twips).
  const columnWidths = new Array<number>(colCount).fill(DEFAULT_COL_TWIPS);
  for (const col of worksheet.columns) {
    const twips = Math.round(col.widthChars * TWIPS_PER_EXCEL_CHAR);
    for (let abs = col.min - 1; abs <= col.max - 1; abs++) {
      const i = abs - colStart;
      if (i >= 0 && i < colCount) columnWidths[i] = twips;
    }
  }

  // Print scaling (fit-to-page / explicit <pageSetup scale>) → uniform shrink of
  // fonts + row heights. `scaled` gates the change so unscaled sheets stay
  // byte-identical (1.0 ⇒ no-op).
  const totalGridTwips = columnWidths.reduce((sum, w) => sum + w, 0);
  // Total grid height for fitToHeight: sum the rendered rows' heights (custom
  // overrides, else Excel's ~15pt default). Wrapping can grow rows past this, so
  // it's an estimate — but shrinking fonts reduces wrapping toward it.
  let totalGridHeightTwips = 0;
  for (let r = 0; r < rowCount; r++) {
    totalGridHeightTwips += rowHeightMap.get(r)?.heightTwips ?? DEFAULT_ROW_TWIPS;
  }
  const printScale = computePrintScale(
    worksheet,
    totalGridTwips,
    sheetContentWidthTwips(worksheet),
    totalGridHeightTwips,
    sheetContentHeightTwips(worksheet),
  );
  const scaled = printScale < 0.999;

  // Manual <rowBreaks>: each brk id is the 0-based row that starts a new page →
  // force a page break before that (absolute) row.
  const breakRows = new Set(worksheet.rowBreaks ?? []);

  // DoS guard: bound the total rendered text per sheet. A crafted file can
  // reference a multi-MB string from thousands of cells (poc-shared-strings:
  // 12 000 cells × one ~1 MB string); shaping/line-breaking that per cell hangs
  // the renderer. Real sheets are far under this budget; once it is exhausted
  // the remaining cells render empty.
  let textBudget = MAX_SHEET_TEXT_CHARS;
  const rows: Array<TableRow> = [];
  for (let r = 0; r < rowCount; r++) {
    const absR = r + rowStart;
    const cells: Array<TableCell> = [];
    for (let c = 0; c < colCount; c++) {
      const absC = c + colStart;
      const merge = mergeOrigins.get(key(absR, absC));
      const insideNotOrigin = insideMerge.has(key(absR, absC));
      if (insideNotOrigin && !merge) {
        // Origin sits to our left in the same row OR above in same column.
        // Above-in-same-column case becomes a vMerge=continue cell so the
        // visual column count of this row stays equal to colCount.
        const verticalParent = verticalMergeParent(worksheet.merges, absR, absC);
        if (verticalParent && absC === verticalParent.startColumn) {
          cells.push(makeVerticalContinuation(verticalParent, colWindowEnd, absR, rowWindowEnd));
        }
        // else: this column is spanned horizontally by an earlier cell in
        // this row → omit entirely so gridSpan layout works.
        continue;
      }

      const ws = cellMatrix[r]?.[c];
      let text = ws ? resolveCellText(ws, sharedStrings, styles, date1904) : '';
      if (text.length > textBudget) text = text.slice(0, Math.max(0, textBudget));
      textBudget -= text.length;
      const xf = ws && ws.styleIndex !== undefined ? styles.cellXfs[ws.styleIndex] : undefined;
      const baseRunProps = xf ? runPropsFromXf(xf, styles) : {};
      const runProps = scaled ? scaleRunFont(baseRunProps, printScale) : baseRunProps;
      const alignment = xf ? alignmentFromXf(xf) : undefined;
      const shading = xf ? shadingFromXf(xf, styles) : undefined;
      const borders = xf ? bordersFromXf(xf, styles) : undefined;

      // Cell overflow (Excel/Calc print model): a non-wrapping cell's text
      // overflows into EMPTY neighbours to the right (left/general alignment) but
      // is CLIPPED where an occupied cell blocks it — Calc renders only the part
      // that fits, dropping the rest. We mirror that for string cells (we don't
      // model wrapText) so the rendered text matches Calc's (TextSim).
      if (
        text.length > 0 &&
        !merge &&
        ws &&
        (ws.type === 's' || ws.type === 'str' || ws.type === 'inlineStr') &&
        (alignment === undefined || alignment === 'left')
      ) {
        let availTwips = columnWidths[c]!;
        let cc = c + 1;
        while (cc < colCount && !cellHasContent(cellMatrix[r]?.[cc])) {
          availTwips += columnWidths[cc]!;
          cc++;
        }
        if (cc < colCount) {
          // an occupied cell stops the overflow → clip to the available width
          const charsFit = Math.max(1, Math.round(availTwips / TWIPS_PER_EXCEL_CHAR));
          if (text.length > charsFit) text = text.slice(0, charsFit);
        }
      }

      // Clamp a merge's horizontal span to the in-window columns so a merge
      // straddling the print-area edge cannot exceed the grid.
      const visibleEndCol = merge ? Math.min(merge.endColumn, colWindowEnd) : 0;
      const properties: CellProperties = {
        ...(merge && visibleEndCol > merge.startColumn
          ? { colSpan: visibleEndCol - merge.startColumn + 1 }
          : {}),
        ...(merge && Math.min(merge.endRow, rowWindowEnd) > merge.startRow
          ? { merge: 'start' as const }
          : {}),
        ...(shading ? { shading } : {}),
        ...(borders ? { borders } : {}),
      };

      const paragraphProps: ParagraphProperties = alignment ? { alignment } : {};
      cells.push({
        properties,
        content: [
          {
            kind: 'paragraph',
            paragraph: {
              properties: paragraphProps,
              runs: text.length > 0 ? [{ text, properties: runProps }] : [],
            },
          },
        ],
      });
    }
    const baseRowProps = rowHeightMap.get(r);
    const rowHeightTwips =
      scaled && baseRowProps !== undefined
        ? Math.max(1, Math.round(baseRowProps.heightTwips * printScale))
        : baseRowProps?.heightTwips;
    const isTitleRow =
      print.titleRows !== undefined &&
      absR >= print.titleRows.startRow &&
      absR <= print.titleRows.endRow;
    const rowProps = {
      ...(rowHeightTwips !== undefined
        ? { height: twipsToPt(rowHeightTwips), heightRule: 'atLeast' as const }
        : {}),
      ...(isTitleRow ? { isHeader: true } : {}),
      ...(breakRows.has(absR) ? { pageBreakBefore: true } : {}),
    };
    rows.push({ properties: rowProps, cells });
  }

  // Gridlines: Excel/Calc do NOT print cell gridlines unless <printOptions
  // gridLines="1"> is set. Default ⇒ no synthetic full grid; only borders that
  // come from cell styles are drawn. With gridLines on, lay a thin grid like a
  // print preview with "Gridlines" enabled.
  const thin: Border = { style: 'single', width: eighthPtToPt(4) };
  // <printOptions horizontalCentered="1"> centers the sheet within the print
  // margins.
  const centered = worksheet.printOptions?.horizontalCentered === true;
  const table: Table = {
    properties: {
      ...(print.gridLines
        ? {
            borders: {
              top: thin,
              bottom: thin,
              left: thin,
              right: thin,
              insideH: thin,
              insideV: thin,
            },
          }
        : {}),
      ...(centered ? { alignment: 'center' as const } : {}),
    },
    grid: columnWidths.map((w) => twipsToPt(w)),
    rows,
  };

  return [{ kind: 'table', table }];
}

function verticalMergeParent(
  merges: ReadonlyArray<MergedRange>,
  row: number,
  column: number,
): MergedRange | undefined {
  for (const m of merges) {
    if (
      column >= m.startColumn &&
      column <= m.endColumn &&
      row > m.startRow &&
      row <= m.endRow &&
      m.endRow > m.startRow
    ) {
      return m;
    }
  }
  return undefined;
}

function makeVerticalContinuation(
  merge: MergedRange,
  colWindowEnd: number,
  absR: number,
  rowWindowEnd: number,
): TableCell {
  const visibleEndCol = Math.min(merge.endColumn, colWindowEnd);
  const span = visibleEndCol - merge.startColumn + 1;
  const lastVisibleRow = Math.min(merge.endRow, rowWindowEnd);
  return {
    properties: {
      merge: absR < lastVisibleRow ? ('middle' as const) : ('end' as const),
      ...(span > 1 ? { colSpan: span } : {}),
    },
    content: [
      {
        kind: 'paragraph',
        paragraph: { properties: {}, runs: [] },
      },
    ],
  };
}

function resolveCellText(
  cell: WorksheetCell,
  sharedStrings: Array<string>,
  styles: XlsxStyles,
  date1904: boolean,
): string {
  if (cell.type === 'inlineStr') return cell.inlineText ?? '';
  if (cell.type === 's') {
    const idx = Number(cell.rawValue);
    if (Number.isInteger(idx) && idx >= 0 && idx < sharedStrings.length) {
      return sharedStrings[idx]!;
    }
    return '';
  }
  if (cell.type === 'b') return cell.rawValue === '1' ? 'TRUE' : 'FALSE';
  if (cell.type === 'str' || cell.type === 'e' || cell.type === 'd') return cell.rawValue;

  // numeric cell — apply numFmt if any
  const xf = cell.styleIndex !== undefined ? styles.cellXfs[cell.styleIndex] : undefined;
  const numFmtId = xf?.numFmtId ?? 0;
  return applyNumberFormat(cell.rawValue, numFmtId, styles.numFmts, date1904);
}

function runPropsFromXf(xf: XlsxCellXf, styles: XlsxStyles): RunProperties {
  if (!xf.applyFont && xf.fontId === 0) return {};
  const font: XlsxFont | undefined = styles.fonts[xf.fontId];
  if (!font) return {};
  const props: { -readonly [K in keyof RunProperties]: RunProperties[K] } = {};
  if (font.bold) props.bold = true;
  if (font.italic) props.italic = true;
  if (font.sizePt !== undefined) props.fontSizePt = halfPtToPt(Math.round(font.sizePt * 2));
  if (font.colorHex) props.colorHex = font.colorHex;
  return props;
}

function alignmentFromXf(xf: XlsxCellXf): Alignment | undefined {
  const align: XlsxCellAlignment | undefined = xf.alignment;
  if (!align) return undefined;
  return mapAlignment(align.horizontal);
}

function mapAlignment(h: XlsxHorizontalAlign | undefined): Alignment | undefined {
  if (!h) return undefined;
  if (h === 'left') return 'left';
  if (h === 'right') return 'right';
  if (h === 'center' || h === 'centerContinuous') return 'center';
  if (h === 'justify') return 'both';
  if (h === 'distributed') return 'distribute';
  return undefined;
}

function shadingFromXf(xf: XlsxCellXf, styles: XlsxStyles): CellShading | undefined {
  // Apply the fill when applyFill is explicitly true OR fillId > 1 (Excel
  // reserves fillId 0=none, 1=gray125 system fills — any user fill starts at 2).
  if (xf.applyFill === false) return undefined;
  if (xf.fillId === 0 || xf.fillId === 1) {
    if (!xf.applyFill) return undefined;
  }
  const fill = styles.fills[xf.fillId];
  if (!fill || fill.patternType !== 'solid' || !fill.fgColorHex) return undefined;
  return { colorHex: fill.fgColorHex };
}

function bordersFromXf(xf: XlsxCellXf, styles: XlsxStyles): CellBorders | undefined {
  if (xf.applyBorder === false) return undefined;
  if (xf.borderId === 0 && !xf.applyBorder) return undefined;
  const border: XlsxBorder | undefined = styles.borders[xf.borderId];
  if (!border) return undefined;
  const out: { -readonly [K in keyof CellBorders]: CellBorders[K] } = {};
  const top = mapBorderEdge(border.top);
  const right = mapBorderEdge(border.right);
  const bottom = mapBorderEdge(border.bottom);
  const left = mapBorderEdge(border.left);
  if (top) out.top = top;
  if (right) out.right = right;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  return Object.keys(out).length > 0 ? out : undefined;
}

function mapBorderEdge(edge: XlsxBorderEdge | undefined): Border | undefined {
  if (!edge || !edge.style || edge.style === 'none') return undefined;
  const { style, sizeEighthPt } = mapBorderStyle(edge.style);
  return {
    style,
    width: eighthPtToPt(sizeEighthPt),
    ...(edge.colorHex ? { colorHex: edge.colorHex } : {}),
  };
}

function mapBorderStyle(style: XlsxBorderStyleName): { style: BorderStyle; sizeEighthPt: number } {
  switch (style) {
    case 'hair':
      return { style: 'single', sizeEighthPt: 2 };
    case 'thin':
      return { style: 'single', sizeEighthPt: 4 };
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return { style: 'single', sizeEighthPt: 8 };
    case 'thick':
      return { style: 'thick', sizeEighthPt: 12 };
    case 'dashed':
    case 'dashDot':
    case 'dashDotDot':
    case 'slantDashDot':
      return { style: 'dashed', sizeEighthPt: 4 };
    case 'dotted':
      return { style: 'dotted', sizeEighthPt: 4 };
    case 'double':
      return { style: 'double', sizeEighthPt: 4 };
    default:
      return { style: 'single', sizeEighthPt: 4 };
  }
}

function key(row: number, col: number): string {
  return `${row},${col}`;
}

// A cell "has content" (blocks overflow / counts toward the used range) when it
// carries a value or inline text — empty styled cells do not.
function cellHasContent(cell: WorksheetCell | undefined): boolean {
  return !!cell && (cell.rawValue !== '' || cell.inlineText !== undefined);
}
