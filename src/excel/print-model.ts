// Excel print model + grid→Flow projection (ECMA-376 §18.2/§18.3/§18.8).
//
// Everything that turns a parsed worksheet into Flow body elements lives here:
// page setup (paper size, orientation, margins, scale/fitToPage), print area
// and Print_Titles resolution, gridline policy, manual breaks, and the cell
// grid → Table projection with merge handling and style interpretation.
//
// This is reader-side knowledge: the xlsx reader composes these to build a
// FlowDoc. The PDF converter consumes the FlowDoc and knows nothing about
// worksheets (oop-design §5.1 — this move broke the reader↔converter cycle).

import type {
  Alignment,
  BodyElement,
  Border,
  BorderStyle,
  CellBorders,
  CellDataBar,
  CellIcon,
  CellProperties,
  CellShading,
  CellSparkline,
  PageMargins,
  PageSize,
  ParagraphProperties,
  Run,
  RunProperties,
  SectionProperties,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';
import type {
  CellRange,
  CellType,
  DataValidation,
  DefinedName,
  MergedRange,
  ParsedWorksheet,
  SheetRichRun,
  WorksheetCell,
  XlsxBorder,
  XlsxBorderEdge,
  XlsxBorderStyleName,
  XlsxCellXf,
  XlsxFont,
  XlsxHorizontalAlign,
  XlsxPageMargins,
  XlsxPageSetup,
  XlsxStyles,
} from '@/excel';
import type { CellConditionalFormatter, CfOverride } from '@/excel/conditional-format';
import type { SheetHyperlink, SheetSlicer } from '@/core/ir/sheet';
import type { Loss } from '@/core/ir/loss';
import { FEATURES } from '@/core/ir/features';
import { eighthPtToPt, halfPtToPt, pt, twipsToPt } from '@/core/ir';
import {
  applyNumberFormat,
  generalToWidth,
  numberFormatColorHex,
  parseAreaRef,
  parseTitleRowRange,
} from '@/excel';
import { bandedTables, computeColumnBands } from '@/excel/column-bands';
import { buildConditionalFormatter } from '@/excel/conditional-format';

/**
 * Excel "character width" → twips: the default font's Maximum Digit Width,
 * ~7 px at 96 DPI ≈ 5.25 pt ≈ 105 twips.
 */
export const TWIPS_PER_EXCEL_CHAR = 105;

/** The point size {@link TWIPS_PER_EXCEL_CHAR} was measured at (Excel's default). */
const DEFAULT_FONT_PT = 11;

/**
 * ECMA-376 §18.3.1.13 — a `<col width>` measures characters of text, and the
 * rendered column is that many Maximum Digit Widths PLUS a fixed 5-pixel
 * padding: `px = chars × MDW + 5`. 5 px at 96 DPI = 3.75 pt = 75 twips.
 *
 * Omitting it made every column 75 twips narrow, which compounds: on a sheet of
 * equal 12-character columns the third one landed ~37 pt left of where
 * LibreOffice puts it. That {@link DEFAULT_COL_TWIPS} below is 960 is the proof
 * the padding belongs — 8.43 characters only reaches Excel's documented 64 px
 * default with it (8.43 × 7 + 5 = 64.01 px = 48.01 pt = 960 twips), so the
 * default was derived from the full formula while the explicit path dropped it.
 */
export const COL_PADDING_TWIPS = 75;

/** One screen pixel at 96 DPI, in twips — the unit Excel's width formula works in. */
const TWIPS_PER_PIXEL = 15;

/**
 * §18.3.1.13 — a `<col width>` in twips.
 *
 * The padding above is the formula for a column at least one character wide.
 * Excel documents a SEPARATE one below that: `px = Trunc(width × MDW + 0.5)`,
 * with no padding at all. It matters because the padding is a constant: on a
 * form drawn over a fine grid — tdf118668.xlsx rules 168 columns of 0.855
 * characters — 3.75pt of padding nearly doubles a 4.5pt column, and the sheet
 * came out 1384pt wide against the reference's 754, so it paginated across two
 * pages where every reader prints one.
 *
 * @param chars     The declared width in characters.
 * @param charTwips The Maximum Digit Width, in twips.
 * @returns The rendered column width in twips.
 */
export function columnTwips(chars: number, charTwips: number): number {
  return columnTwipsOf(chars, charTwips);
}

/**
 * §18.3.1.81 — the DEFAULT column, in twips, from whichever of the two the
 * sheet declares.
 *
 * The two are not the same number. `defaultColWidth` "includes margin padding
 * and extra padding for gridlines"; `baseColWidth` is the bare character count
 * and explicitly excludes them, so deriving a default from it adds the padding
 * once to reach a defaultColWidth and once more to render it. 47668.xlsx says
 * `baseColWidth="10"` and caches its picture's extent at 9753600 EMU = 768pt
 * over 12 columns plus 48pt — 60pt, or 80px, a column. Padding it once gives
 * 75px and squeezed that picture by 6%. LibreOffice measures ~79px here, so
 * both references agree against us.
 *
 * @param worksheet     The sheet.
 * @param charTwips     The Maximum Digit Width, in twips.
 * @param fallbackChars The width to use when the sheet declares neither.
 * @returns The default column width in twips.
 */
function defaultColumnTwips(
  worksheet: ParsedWorksheet,
  charTwips: number,
  fallbackChars: number,
): number {
  if (worksheet.defaultColWidthChars !== undefined) {
    return columnTwipsOf(worksheet.defaultColWidthChars, charTwips);
  }
  if (worksheet.baseColWidthChars !== undefined) {
    return columnTwipsOf(worksheet.baseColWidthChars, charTwips) + COL_PADDING_TWIPS;
  }
  return columnTwipsOf(fallbackChars, charTwips);
}

function columnTwipsOf(chars: number, charTwips: number): number {
  if (chars >= 1) return Math.round(chars * charTwips + COL_PADDING_TWIPS);
  const px = Math.trunc((chars * charTwips) / TWIPS_PER_PIXEL + 0.5);
  return px * TWIPS_PER_PIXEL;
}

/**
 * Excel insets a cell's text by ~2 px each side (1.5 pt at 96 DPI). The layout
 * engine's default is a word processor's 108 twips (5.4 pt), which is nearly
 * four times as much.
 */
const EXCEL_CELL_INSET_PT = 1.5;

/**
 * Excel's default column width is 8.43 "characters" ≈ 64px ≈ 960 twips. Used for
 * columns without an explicit `<col width="..">`.
 */
export const DEFAULT_COL_TWIPS = 960;

/** The character count behind {@link DEFAULT_COL_TWIPS}, for a non-Calibri unit. */
const DEFAULT_COL_CHARS = 8.43;

/**
 * The column-width unit in twips: the Maximum Digit Width of the font the
 * document is rendered in (§18.3.1.13), or Excel's own 7 px when the caller has
 * not measured one.
 */
function digitTwips(digitWidthPt: number | undefined): number {
  if (digitWidthPt === undefined || !Number.isFinite(digitWidthPt) || digitWidthPt <= 0) {
    return TWIPS_PER_EXCEL_CHAR;
  }
  return digitWidthPt * TWIPS_PER_POINT;
}

/**
 * ECMA-376 §18.3.1.81 — the row height Excel uses when a sheet declares no
 * `<sheetFormatPr defaultRowHeight>`: 15pt, the line height of its default
 * theme font (Calibri 11). Independent of whatever font we end up rendering
 * with — the height belongs to the document, not to the typesetter.
 */
const EXCEL_DEFAULT_ROW_HEIGHT_PT = 15;

/**
 * Excel's default row height is ~15pt = 300 twips. Used (for the `fitToHeight`
 * estimate) for rows without an explicit `<row ht="..">`.
 */
export const DEFAULT_ROW_TWIPS = 300;

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

/**
 * Build the page section (paper size + margins) from a worksheet's `<pageSetup>`
 * / `<pageMargins>`.
 *
 * Margins are always set, to Excel's own defaults when the worksheet declares
 * none (§18.3.1.62) — the renderer's fallback is a word processor's inch, which
 * is not what a spreadsheet prints. The paper size is left unset when the
 * worksheet names none, because there the file genuinely holds no answer: Excel
 * picks by locale and printer, and the renderer's deterministic A4 is as good
 * as anything we could invent.
 */
export function sectionFromWorksheet(worksheet: ParsedWorksheet): SectionProperties {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  return {
    ...(pageSize ? { pageSize } : {}),
    // Always set: a worksheet that omits <pageMargins> still prints with
    // Excel's margins, not a word processor's.
    margins: marginsFromXlsx(worksheet.pageMargins),
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

/**
 * ECMA-376 §18.3.1.62 — the page margins Excel writes when the user has not
 * touched them, in inches. A worksheet may omit `<pageMargins>` entirely, and
 * falling through to the renderer's default (a word processor's 1 inch) put the
 * grid 0.3 inch — 21.6 pt — right of where Excel and LibreOffice print it, on
 * every such sheet.
 */
const EXCEL_DEFAULT_MARGINS: PageMargins = {
  left: twipsToPt(Math.round(0.7 * TWIPS_PER_INCH)),
  right: twipsToPt(Math.round(0.7 * TWIPS_PER_INCH)),
  top: twipsToPt(Math.round(0.75 * TWIPS_PER_INCH)),
  bottom: twipsToPt(Math.round(0.75 * TWIPS_PER_INCH)),
  header: twipsToPt(Math.round(0.3 * TWIPS_PER_INCH)),
  footer: twipsToPt(Math.round(0.3 * TWIPS_PER_INCH)),
};

function marginsFromXlsx(margins: XlsxPageMargins | undefined): PageMargins {
  if (!margins) return EXCEL_DEFAULT_MARGINS;
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

/**
 * ECMA-376 §18.2.5 — resolve the sheet-scoped `_xlnm.Print_Area` defined name
 * (`localSheetId` = the sheet's 0-based index) into a clipping range.
 */
export function resolvePrintArea(
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

/** ECMA-376 §18.2.5 — `_xlnm.Print_Titles` → the repeated row range (0-indexed). */
export function resolvePrintTitleRows(
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

/**
 * The width one page has for content: the sheet's paper (declared or the A4
 * default) less its left and right margins, in points. The same measure the
 * column bands are packed into — a drawing layer banded on any other width
 * would not line up with the grid beside it.
 */
export function printableWidthPt(worksheet: ParsedWorksheet): number {
  return sheetContentWidthTwips(worksheet) / TWIPS_PER_POINT;
}

function sheetContentWidthTwips(worksheet: ParsedWorksheet): number {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  const pageWidthTwips = pageSize ? Math.round(pageSize.width * 20) : DEFAULT_PAPER_TWIPS[0];
  const margins = marginsFromXlsx(worksheet.pageMargins);
  const left = Math.round(margins.left * 20);
  const right = Math.round(margins.right * 20);
  return Math.max(TWIPS_PER_INCH / 2, pageWidthTwips - left - right);
}

/**
 * The printable HEIGHT of one page of this sheet, in points — its paper height
 * less its own top and bottom margins. The vertical twin of
 * {@link printableWidthPt}.
 *
 * @param worksheet The parsed worksheet (paper + margins).
 * @returns The height a page of this sheet has for content, in points.
 */
export function printableHeightPt(worksheet: ParsedWorksheet): number {
  return sheetContentHeightTwips(worksheet) / TWIPS_PER_POINT;
}

function sheetContentHeightTwips(worksheet: ParsedWorksheet): number {
  const pageSize = pageSizeFromSetup(worksheet.pageSetup);
  const pageHeightTwips = pageSize ? Math.round(pageSize.height * 20) : DEFAULT_PAPER_TWIPS[1];
  const margins = marginsFromXlsx(worksheet.pageMargins);
  const top = Math.round(margins.top * 20);
  const bottom = Math.round(margins.bottom * 20);
  return Math.max(TWIPS_PER_INCH / 2, pageHeightTwips - top - bottom);
}

/**
 * The grid as the fit-to-page search sees it: the extents that have to be packed
 * into pages, the breaks that force one, and the band repeated at the top of each
 * continuation page.
 */
interface FitGeometry {
  /** Visible row heights in twips, in order (unscaled). */
  readonly rowTwips: ReadonlyArray<number>;
  /** Indices into `rowTwips` that start a new page (manual `<rowBreaks>`). */
  readonly rowBreaks: ReadonlySet<number>;
  /** Height of the `_xlnm.Print_Titles` rows, repeated on every page after the first. */
  readonly titleTwips: number;
  /** Visible column widths in twips, in order (unscaled). */
  readonly colTwips: ReadonlyArray<number>;
  /** Indices into `colTwips` that start a new band (manual `<colBreaks>`). */
  readonly colBreaks: ReadonlySet<number>;
}

/**
 * How many pages the grid ACTUALLY paginates into at a given scale — the same
 * greedy fill the layout performs, run over one axis.
 *
 * A row does not split, so the closed form `contentHeight × fitToHeight / total`
 * is a lower bound and not an answer: whatever height the last row that will not
 * fit leaves unused is a page's worth of slack the formula never accounts for,
 * and print titles make it worse by consuming their height again on every page.
 */
function pageCountAtScale(
  extents: ReadonlyArray<number>,
  forcedBreaks: ReadonlySet<number>,
  repeatTwips: number,
  limitTwips: number,
  scale: number,
): number {
  let pages = 1;
  let used = 0;
  for (let i = 0; i < extents.length; i++) {
    const extent = extents[i]! * scale;
    if (used > 0 && (forcedBreaks.has(i) || used + extent > limitTwips)) {
      pages++;
      used = repeatTwips * scale;
    }
    used += extent;
  }
  return pages;
}

/**
 * The largest whole percentage at or below `closed` that paginates into at most
 * `maxPages` pages. Whole percentages because that is the unit Excel searches in
 * and stores the answer in (`<pageSetup scale="51">`), so the two agree exactly.
 * Monotone — shrinking never adds a page — so a binary search is enough.
 */
function searchFitScale(
  closed: number,
  maxPages: number,
  pageCount: (scale: number) => number,
): number {
  const floor = Math.round(MIN_PRINT_SCALE * 100);
  let hi = Math.min(100, Math.floor(closed * 100));
  if (hi <= floor) return Math.max(MIN_PRINT_SCALE, closed);
  if (pageCount(hi / 100) <= maxPages) return closed;
  let lo = floor;
  // Shrink only when shrinking gets there. A manual page break costs a page at
  // every scale, so a sheet with three of them and `fitToHeight="1"` can never
  // satisfy the target — searching for one drove the scale into the floor and
  // rendered AverageTaxRates.xlsx at 1pt. Unreachable ⇒ leave the closed form.
  if (pageCount(lo / 100) > maxPages) return closed;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (pageCount(mid / 100) <= maxPages) lo = mid;
    else hi = mid - 1;
  }
  return lo / 100;
}

function computePrintScale(
  worksheet: ParsedWorksheet,
  totalGridTwips: number,
  contentWidthTwips: number,
  totalGridHeightTwips: number,
  contentHeightTwips: number,
  geometry: FitGeometry,
): number {
  const setup = worksheet.pageSetup;
  let s = 1;
  if (worksheet.fitToPage) {
    // fitToPage overrides any explicit scale. The binding (smaller) of the two
    // fit factors wins. BOTH default to 1 page (§18.3.1.63 CT_PageSetup): a
    // sheet that turns fit-to-page on and names neither is asking to fit on one
    // page each way. Treating an absent fitToHeight as "unconstrained" — the
    // meaning of an explicit `fitToHeight="0"` — left such a sheet at full size
    // and let it flow down: bnc762542.xlsx resolved to no shrink at all where
    // the reference shrinks it to about four fifths.
    const fitW = setup?.fitToWidth ?? 1;
    const fitH = setup?.fitToHeight ?? 1;
    let sW = 1;
    let sH = 1;
    if (fitW >= 1 && totalGridTwips > 0) sW = (contentWidthTwips * fitW) / totalGridTwips;
    if (fitH >= 1 && totalGridHeightTwips > 0) {
      sH = (contentHeightTwips * fitH) / totalGridHeightTwips;
    }
    // The formula divides an area; the printer packs whole rows and whole
    // columns. Where the two disagree the sheet spills onto one more page than
    // it was asked to take, so shrink until it really fits — see
    // {@link pageCountAtScale}. tdf58243.xlsx asks for 2×2 pages: the formula
    // gives 52 %, the print titles it ignores bring that to Excel's stored 51 %,
    // and the row boundaries it also ignores mean only 50 % fits.
    if (fitW >= 1 && geometry.colTwips.length > 0) {
      sW = searchFitScale(sW, fitW, (scale) =>
        pageCountAtScale(geometry.colTwips, geometry.colBreaks, 0, contentWidthTwips, scale),
      );
    }
    if (fitH >= 1 && geometry.rowTwips.length > 0) {
      sH = searchFitScale(sH, fitH, (scale) =>
        pageCountAtScale(
          geometry.rowTwips,
          geometry.rowBreaks,
          geometry.titleTwips,
          contentHeightTwips,
          scale,
        ),
      );
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

// The grey Excel and LibreOffice print cell gridlines in — light enough that a
// declared cell border still reads as the heavier line on the page.
const PRINT_GRIDLINE_HEX = 'C0C0C0';

// A synthetic id for a conditional format's own number format, which arrives as
// a code rather than through the workbook's <numFmts> table.
const CF_NUMBER_FORMAT_ID = 1_000_001;

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
  // E-SHEET SC2 tail (TC3) — sheet name → grid, for sparklines whose data range
  // is sheet-qualified (Sheet2!A1:C1), and (W9 tail) an `expression` CF rule that
  // references another sheet or a defined name. Absent ⇒ same-sheet resolution only.
  readonly sheetGrids?: ReadonlyMap<string, ParsedWorksheet>;
  // This sheet's name (so a self-qualified Sheet1!A1 and a sheet-scoped defined
  // name resolve) and the workbook's defined names (W9 expression CF tail).
  readonly sheetName?: string;
  readonly definedNames?: ReadonlyArray<DefinedName>;
  // E-SHEET W3 — cell hyperlinks resolved to external URLs; a covered cell's run
  // takes the URL as its href (PDF /Link + HTML <a>). Absent ⇒ no cell links.
  readonly hyperlinks?: ReadonlyArray<SheetHyperlink>;
  // E-SHEET W6 — per-index rich-text runs for shared strings, parallel to the
  // sharedStrings array; a t="s" cell whose index has runs emits multiple runs
  // with their own formatting. Absent ⇒ every cell is a single run.
  readonly sharedStringRuns?: ReadonlyArray<ReadonlyArray<SheetRichRun> | undefined>;
  // E-SHEET W9 — the injected reference date for conditional-format `timePeriod`
  // windows and TODAY()/NOW() in `expression` rules. Absent ⇒ those no-op.
  readonly now?: Date;
  // §18.3.1.13 measures a column in Maximum Digit Widths of the workbook's own
  // default font, so the unit is a property of the FONT — see
  // ProjectSheetOptions.digitWidthPt. Absent ⇒ Excel's own 7 px.
  readonly digitWidthPt?: number;
  // The print scale this sheet resolved to, reported back to the caller. A
  // drawing is anchored to the grid and shrinks with it, but it is emitted
  // beside the grid rather than inside it — and the factor is only known here,
  // where the grid's own totals are. Recomputing it outside would mean
  // duplicating the used-range logic and, sooner or later, disagreeing with it.
  readonly scaleSink?: { value: number };
  // The left edge of each column band, in points at the print scale, reported
  // back the same way. A drawing anchored in the second band is anchored past
  // the first band's width, and printed at that offset it falls off the page
  // the layout has reached — shape-macro-ext-ref.xlsx lost a whole chart and
  // the macro button beside it that way. One entry per band table emitted; a
  // sheet that does not band reports a single 0.
  readonly bandSink?: { lefts: Array<number> };
  // How far the sheet's drawings reach from its origin, in points. A drawing
  // anchored past the last cell still has to be printed, so fit-to-page has to
  // fit IT too — measuring only the range the cells occupy left a sheet whose
  // values sit in a handful of cells and whose callout runs down to row 78 with
  // nothing to shrink (bnc762542.xlsx), where the reference shrinks to about
  // four fifths.
  readonly drawingExtentPt?: { readonly widthPt: number; readonly heightPt: number };
  // Sink for projection losses. The defence-in-depth caps below are correct —
  // a pathological sheet must not exhaust memory — but a cap that fires without
  // saying so is precisely the silent wrongness LossReport exists to prevent.
  // Absent ⇒ the caps still apply, they just go unreported (internal callers
  // that already know); readXlsx always supplies one.
  readonly losses?: Array<Loss>;
}

/**
 * Project one worksheet's grid into Flow body elements — a single {@link Table}
 * (or, when the sheet is wider than the page, several column-banded tables). The
 * full print model lives here: the used-range/print-area window, merge handling,
 * the style cascade (fonts/fills/borders/alignment), print scaling, conditional
 * formatting, sparklines, table/pivot banding, overflow and rotation. Memory- and
 * text-budget-bounded against untrusted input.
 *
 * @param worksheet    The parsed grid + per-sheet geometry.
 * @param sharedStrings The workbook shared-string table (`t="s"` cells index it).
 * @param styles       The workbook style table (cellXfs + fonts/fills/borders).
 * @param date1904     The 1904 date system flag (serial-to-date epoch).
 * @param print        The print-model knobs (print area, gridlines, title rows, …).
 * @returns The body elements (one table block, or banded tables); empty for a
 *   blank or empty-print-area sheet.
 */
export function worksheetToBody(
  worksheet: ParsedWorksheet,
  sharedStrings: ReadonlyArray<string>,
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
  // §18.3.1.13's unit of column width is the Maximum Digit Width of the
  // WORKBOOK's default font. When the file names one — Calibri 11, Arial 10,
  // the two Excel itself defaults to — that width is Excel's documented 7 px,
  // and LibreOffice reproduces it because it substitutes metric-compatible
  // faces (Carlito, Liberation Sans). Only when the file names no font at all
  // does the reader's own default decide, and then the face we render with is
  // the honest answer: tdf122336.xlsx declares `<font/>`, LibreOffice laid its
  // columns out in Caladea, and its 40-character column takes a page to itself
  // where ours took half of one.
  const charTwipsUnit =
    styles.fonts[0]?.name === undefined ? digitTwips(print.digitWidthPt) : TWIPS_PER_EXCEL_CHAR;
  // …but only the COLUMNS are measured in it. `charWidthUnits` already returns
  // each character's real width expressed in Excel's own 7px unit, so measuring
  // text with the render font's digit as well applies the same ratio twice —
  // 123.54/105 on tdf122336.xlsx, which is the "sixth" the note above warns
  // about, and it cut "Uitvoeringsdatum" to "Uitvoeringsdatu" inside a column
  // that is 5% WIDER than the reference's and had 7pt to spare.
  const textTwipsUnit = TWIPS_PER_EXCEL_CHAR;

  let usedRow = -1;
  let usedCol = -1;
  const contentAt = new Set<string>();
  for (const c of worksheet.cells) {
    if (c.rawValue !== '' || c.inlineText !== undefined) {
      contentAt.add(key(c.row, c.column));
      if (c.row > usedRow) usedRow = c.row;
      if (c.column > usedCol) usedCol = c.column;
    }
  }
  // A merge extends the range only when its ORIGIN holds something. The span
  // has to be covered or the merge is clipped — but a merge follows content, it
  // does not create any, and treating every merge as content lets pure
  // formatting define the sheet. 57893-many-merges.xlsx is 50 000 rows of
  // merges over cells that hold nothing at all: once blank space started
  // paginating honestly, that made 1042 empty pages out of a document
  // LibreOffice prints as one. Empty styled cells are already excluded here for
  // the same reason.
  for (const m of worksheet.merges) {
    if (!contentAt.has(key(m.startRow, m.startColumn))) continue;
    if (m.endRow > usedRow) usedRow = m.endRow;
    if (m.endColumn > usedCol) usedCol = m.endColumn;
  }
  // A sparkline host cell (E-SHEET SC2) is usually empty and just past the data;
  // keep it in the used range so the cell that carries the mini chart survives.
  for (const sp of worksheet.sparklines ?? []) {
    const host = parseAreaRef(sp.sqref);
    if (!host) continue;
    if (host.startRow > usedRow) usedRow = host.startRow;
    if (host.startColumn > usedCol) usedCol = host.startColumn;
  }
  if (usedRow < 0 || usedCol < 0) {
    // Nothing but empty styled cells — no grid to render. The sheet may still
    // carry drawings, and fit-to-page still has to fit THEM, so resolve the
    // scale before leaving: bnc762542.xlsx keeps every value in a styled-empty
    // cell and prints nothing but its callout.
    if (print.scaleSink) {
      print.scaleSink.value = computePrintScale(
        worksheet,
        Math.round((print.drawingExtentPt?.widthPt ?? 0) * TWIPS_PER_POINT),
        sheetContentWidthTwips(worksheet),
        Math.round((print.drawingExtentPt?.heightPt ?? 0) * TWIPS_PER_POINT),
        sheetContentHeightTwips(worksheet),
        // No grid to paginate — the drawing extent alone decides, so the closed
        // form is the whole answer here.
        { rowTwips: [], rowBreaks: new Set(), titleTwips: 0, colTwips: [], colBreaks: new Set() },
      );
    }
    return [];
  }

  // Overflow does not stop at the used range either. A long string in the last
  // used column keeps running across the empty grid to its right, and Excel and
  // LibreOffice both draw it on one line — there is nothing there to block it.
  // Without columns to run into, the cell keeps its own width and the layout
  // wraps the text inside it: open-as-read-only.xlsx declares `<dimension
  // ref="A1"/>` and its single sentence came out as nine stacked lines.
  usedCol += overflowColumnsPastUsedRange(
    worksheet,
    usedCol,
    sharedStrings,
    styles,
    date1904,
    charTwipsUnit,
  );

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
  // 16384×1048576 cell space) cannot exhaust memory.
  //
  // MAX_GRID_CELLS is the one that does the work. Per-dimension caps are not
  // sufficient on their own — their product is what allocates, which is how
  // LibreOffice's too-many-cols-rows.xlsx (2.5 KB, five real cells, declared as
  // A1:XFE16777217) exhausted a 6 GB heap. Amplification needs no zip bomb; a
  // declared extent is enough. Rows give way first: columns carry the record
  // shape, rows are homogeneous repetitions of it.
  //
  // The column limit is SpreadsheetML's own (§18.3.1.13 — column XFD is the
  // last), i.e. no limit of ours at all. It used to be 1024, on the reasoning
  // that a wider PDF table is unreadable — but unreadable is the author's call,
  // not ours, and it silently truncated documents that are simply wide: POI's
  // 53105.xlsx is 2 rows across all 16 384 columns, and we printed 103 pages of
  // it where Excel and LibreOffice print 1639. Rendering it in full costs
  // 339 ms and 69 MB, so the cell budget was carrying the load anyway.
  const MAX_GRID_COLS = 16_384;
  // §18.3.1.73 — the last row SpreadsheetML defines; used only to explain a clip.
  const SHEET_MAX_ROWS = 1_048_576;
  const MAX_GRID_ROWS = 50_000;
  const MAX_GRID_CELLS = 1_000_000;
  const wantRows = rowEnd - rowStart + 1;
  const wantCols = colEnd - colStart + 1;
  let colCount = Math.min(wantCols, MAX_GRID_COLS);
  let rowCount = Math.max(
    1,
    Math.min(wantRows, MAX_GRID_ROWS, Math.floor(MAX_GRID_CELLS / colCount)),
  );
  // A clipped window ends wherever the cap fell, which on an amplified sheet is
  // in the middle of nothing: too-many-cols-rows.xlsx puts its five values at
  // the extremes of A1:XFE16777217, so the surviving 976×1024 window holds one
  // of them and 999 423 blank cells. Blank space still paginates — the sheet
  // projected 2657 near-empty pages — so trim the tail the cap left behind.
  //
  // Only when a cap actually fired. A window that reflects the used range or a
  // declared print area ends where the author said it does, trailing blanks
  // included, and Excel prints those.
  if (rowCount < wantRows) {
    rowCount = Math.max(1, lastContentRow(worksheet, rowStart, rowCount, colStart, colCount) + 1);
  }
  if (colCount < wantCols) {
    colCount = Math.max(
      1,
      lastContentColumn(worksheet, rowStart, rowCount, colStart, colCount) + 1,
    );
  }
  // Name the reason. A sheet can ask for more than SpreadsheetML HAS — a cell
  // at XFE is past column XFD, the last one the format defines, and no reader
  // can put it anywhere the file actually names. That is a different fact from
  // "we stopped early to stay inside memory", and reporting the second when the
  // first is true tells the reader to buy more RAM for a file that is malformed.
  // too-many-cols-rows.xlsx is both at once: A1:XFE16777217 exceeds the format
  // in each direction, and only then does the memory cap bite.
  const reason = (want: number, limit: number, unit: 'row' | 'column', last: string): string =>
    want > limit
      ? `the sheet declares cells past ${unit} ${last}, the last SpreadsheetML defines`
      : 'memory guard';
  if (rowCount < wantRows) {
    print.losses?.push({
      severity: 'dropped',
      feature: FEATURES.tables,
      detail:
        `grid clipped to the first ${rowCount} rows of ${wantRows} in the used range ` +
        `(${reason(rowStart + wantRows, SHEET_MAX_ROWS, 'row', String(SHEET_MAX_ROWS))})`,
      ...(print.sheetName ? { where: `sheet "${print.sheetName}"` } : {}),
    });
  }
  if (colCount < wantCols) {
    print.losses?.push({
      severity: 'dropped',
      feature: FEATURES.tables,
      detail:
        `grid clipped to the first ${colCount} columns of ${wantCols} in the used range ` +
        `(${reason(colStart + wantCols, MAX_GRID_COLS, 'column', 'XFD')})`,
      ...(print.sheetName ? { where: `sheet "${print.sheetName}"` } : {}),
    });
  }
  // Absolute index of the last in-window row/column (for merge clamping).
  const rowWindowEnd = rowStart + rowCount - 1;
  const colWindowEnd = colStart + colCount - 1;

  // Row heights keyed by LOCAL (window-relative) row index.
  //
  // EVERY row gets one. A spreadsheet row has a definite height — §18.3.1.81
  // `<sheetFormatPr defaultRowHeight>`, or Excel's 15pt for its default theme
  // font — and leaving rows without one let the text's own leading decide,
  // which made the row pitch a property of the rendering font rather than of
  // the document. The per-row `ht` overrides below take precedence.
  const defaultRowTwips = Math.round(
    (worksheet.defaultRowHeightPt ?? EXCEL_DEFAULT_ROW_HEIGHT_PT) * TWIPS_PER_POINT,
  );
  const rowHeightMap = new Map<number, { heightTwips: number; heightRule: 'atLeast' | 'exact' }>();
  for (let r = 0; r < rowCount; r++) {
    rowHeightMap.set(r, { heightTwips: defaultRowTwips, heightRule: 'atLeast' });
  }
  for (const h of worksheet.rowHeights) {
    const local = h.row - rowStart;
    if (local < 0 || local >= rowCount) continue;
    const heightTwips = Math.round(h.heightPt * TWIPS_PER_POINT);
    // §18.3.1.73 customHeight="1" — the author fixed this height, and Excel
    // does NOT grow such a row to fit its content; it clips. Growing it is not
    // a harmless safety margin: the extra height moves every row after it and
    // changes where the page breaks. simple-monthly-budget.xlsx pins all 23 of
    // its rows, their sum fits its page with 5pt to spare, and one row carrying
    // a 20pt "62%" grew past its pin and pushed the last row — and the chart
    // anchored beside it — onto a second page the reference does not have.
    // Without the flag the height is advisory and content may still grow it.
    rowHeightMap.set(local, { heightTwips, heightRule: h.customHeight ? 'exact' : 'atLeast' });
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
  // §18.3.1.81 `<sheetFormatPr defaultColWidth>` overrides Excel's 8.43-char
  // default for every column no `<col>` covers — the horizontal twin of
  // defaultRowHeight, and ignored for the same reason.
  // defaultColWidth wins; failing that Excel derives the default column from
  // baseColWidth by the same characters + 5px formula; failing both, 8.43.
  const defaultColTwips = defaultColumnTwips(worksheet, charTwipsUnit, DEFAULT_COL_CHARS);
  const columnWidths = new Array<number>(colCount).fill(defaultColTwips);
  // §18.3.1.13/§18.3.1.73 `hidden` — Excel and LibreOffice print neither a
  // hidden column nor a hidden row. Rendering them put a hidden currency column
  // and seven hidden rows into AverageTaxRates.xlsx that no other reader shows,
  // and broke the table's structure around them.
  const hiddenCols = new Set<number>();
  for (const col of worksheet.columns) {
    const twips = columnTwips(col.widthChars, charTwipsUnit);
    for (let abs = col.min - 1; abs <= col.max - 1; abs++) {
      const i = abs - colStart;
      if (i < 0 || i >= colCount) continue;
      columnWidths[i] = twips;
      if (col.hidden) hiddenCols.add(i);
    }
  }
  const hiddenRows = new Set<number>();
  for (const h of worksheet.rowHeights) {
    if (h.hidden) hiddenRows.add(h.row - rowStart);
  }
  // The grid the rest of the projection works with: local indices of the
  // columns that actually print, in order.
  const visibleCols: Array<number> = [];
  for (let c = 0; c < colCount; c++) if (!hiddenCols.has(c)) visibleCols.push(c);
  const visibleWidths = visibleCols.map((c) => columnWidths[c]!);

  // Print scaling (fit-to-page / explicit <pageSetup scale>) → uniform shrink of
  // fonts + row heights. `scaled` gates the change so unscaled sheets stay
  // byte-identical (1.0 ⇒ no-op).
  const totalGridTwips = visibleWidths.reduce((sum, w) => sum + w, 0);
  // Total grid height for fitToHeight: sum the rendered rows' heights (custom
  // overrides, else Excel's ~15pt default). Wrapping can grow rows past this, so
  // it's an estimate — but shrinking fonts reduces wrapping toward it.
  let totalGridHeightTwips = 0;
  // The same rows as an ordered list, plus the breaks and the repeated title
  // band, so fit-to-page can paginate them instead of dividing an area.
  const fitRowTwips: Array<number> = [];
  const fitRowBreaks = new Set<number>();
  const manualRowBreaks = new Set(worksheet.rowBreaks ?? []);
  let fitTitleTwips = 0;
  for (let r = 0; r < rowCount; r++) {
    if (hiddenRows.has(r)) continue;
    const twips = rowHeightMap.get(r)?.heightTwips ?? DEFAULT_ROW_TWIPS;
    totalGridHeightTwips += twips;
    const abs = r + rowStart;
    if (manualRowBreaks.has(abs)) fitRowBreaks.add(fitRowTwips.length);
    if (print.titleRows && abs >= print.titleRows.startRow && abs <= print.titleRows.endRow) {
      fitTitleTwips += twips;
    }
    fitRowTwips.push(twips);
  }
  const manualColBreaks = new Set(worksheet.colBreaks ?? []);
  const fitColBreaks = new Set<number>();
  visibleCols.forEach((c, i) => {
    if (manualColBreaks.has(c + colStart)) fitColBreaks.add(i);
  });
  // The extent that has to fit is the grid's OR the drawings', whichever
  // reaches further — see PrintContext.drawingExtentPt.
  const drawing = print.drawingExtentPt;
  const printScale = computePrintScale(
    worksheet,
    Math.max(totalGridTwips, Math.round((drawing?.widthPt ?? 0) * TWIPS_PER_POINT)),
    sheetContentWidthTwips(worksheet),
    Math.max(totalGridHeightTwips, Math.round((drawing?.heightPt ?? 0) * TWIPS_PER_POINT)),
    sheetContentHeightTwips(worksheet),
    {
      rowTwips: fitRowTwips,
      rowBreaks: fitRowBreaks,
      titleTwips: fitTitleTwips,
      colTwips: visibleWidths,
      colBreaks: fitColBreaks,
    },
  );
  const scaled = printScale < 0.999;
  if (print.scaleSink) print.scaleSink.value = printScale;

  // Manual <rowBreaks>: each brk id is the 0-based row that starts a new page →
  // force a page break before that (absolute) row.
  const breakRows = new Set(worksheet.rowBreaks ?? []);

  // DoS guard: bound the total rendered text per sheet. A crafted file can
  // reference a multi-MB string from thousands of cells (poc-shared-strings:
  // 12 000 cells × one ~1 MB string); shaping/line-breaking that per cell hangs
  // the renderer. Real sheets are far under this budget; once it is exhausted
  // the remaining cells render empty.
  let textBudget = MAX_SHEET_TEXT_CHARS;
  // Reported once, on the cell that first runs the budget dry — every later
  // cell would repeat the same fact.
  let textBudgetReported = false;

  // Cells sharing a cellXf reuse the SAME properties objects — not equal
  // copies. The style cascade memoizes by object identity, so on grid-shaped
  // sheets the resolved-property population collapses to one per distinct xf
  // instead of one per cell (POI bug62181: ~0.4M cells of a handful of xfs
  // OOMed a 512 MB heap on per-cell copies).
  const runPropsByXf = new Map<XlsxCellXf | undefined, RunProperties>();
  const cellRunProps = (xf: XlsxCellXf | undefined): RunProperties => {
    let props = runPropsByXf.get(xf);
    if (props === undefined) {
      const base = xf ? runPropsFromXf(xf, styles) : {};
      props = scaled ? scaleRunFont(base, printScale) : base;
      runPropsByXf.set(xf, props);
    }
    return props;
  };
  const paraPropsByAlignment = new Map<Alignment | undefined, ParagraphProperties>();
  const cellParaProps = (alignment: Alignment | undefined): ParagraphProperties => {
    let props = paraPropsByAlignment.get(alignment);
    if (props === undefined) {
      props = alignment ? { alignment } : {};
      paraPropsByAlignment.set(alignment, props);
    }
    return props;
  };

  // §18.3.1.18 conditional formatting (E-SHEET SC1): a per-cell fill/font
  // override evaluated from the sheet's rules. undefined when the sheet has none
  // — the cell loop then takes the byte-identical base-format path.
  const cfFormatter: CellConditionalFormatter | undefined = buildConditionalFormatter(
    worksheet.conditionalFormats,
    styles,
    worksheet.cells,
    // Resolve a cell's string value for the W5 text / duplicate-unique rules and
    // the W9 expression engine's references to text cells.
    (cell) => resolveCellText(cell, sharedStrings, styles, date1904),
    date1904,
    print.now,
    print.sheetGrids,
    print.sheetName,
    print.definedNames,
  );

  // Sparklines (E-SHEET SC2): host-cell (absolute key) → resolved value series.
  // Empty when the sheet has no sparklines, so the cell loop stays unchanged.
  const sparklineByCell = buildSparklineLookup(worksheet, print);

  // Excel tables (E-SHEET SC3): cell (absolute key) → banded/header fill + header
  // text colour. Empty when the sheet has no table parts. Applied below the
  // cell's own fill and below conditional formatting.
  const tableFormatByCell = buildTableFormatLookup(worksheet);

  // §18.3.1.33 data-validation `list` cells (E-SHEET SV1): the ranges whose cells
  // should paint an in-cell dropdown affordance. Empty (the dropdown block is
  // then skipped, byte-identical) unless the sheet has a shown list validation.
  const dropdownRanges = listDropdownRanges(worksheet.dataValidations);

  // The column bands the grid will be split into, in LOCAL column indices — a
  // band boundary is a page edge, and overflowing text stops at one. Computed
  // here rather than at the split below because the row loop needs it: text
  // clipped against the whole grid's width, then sliced into a band, is clipped
  // to twice the room it actually gets and wraps anyway (the Infos sheet of
  // tdf171828.xlsx wrapped to three lines inside a one-line row).
  const bandEndOfCol = new Map<number, number>();
  {
    const widthsForBands = scaledColumnWidths(visibleWidths, printScale, scaled);
    const total = widthsForBands.reduce((sum, w) => sum + w, 0);
    const breaks = new Set<number>();
    for (const brk of worksheet.colBreaks ?? []) {
      const local = brk - colStart;
      if (local > 0 && local < colCount) breaks.add(local);
    }
    const wide = worksheet.fitToPage ? (worksheet.pageSetup?.fitToWidth ?? 1) : 1;
    const width = sheetContentWidthTwips(worksheet);
    if (colCount > 1 && (!scaled || wide > 1) && (total > width || breaks.size > 0)) {
      for (const band of computeColumnBands(widthsForBands, width, breaks)) {
        for (let i = band.start; i <= band.end; i++) {
          bandEndOfCol.set(visibleCols[i] ?? i, visibleCols[band.end] ?? band.end);
        }
      }
    }
  }

  const rows: Array<TableRow> = [];
  // Where the first print-title row landed in `rows`. Not derivable from the
  // row index afterwards — hidden rows are skipped, so the two do not agree.
  let titleRowIndex = -1;
  for (let r = 0; r < rowCount; r++) {
    if (hiddenRows.has(r)) continue;
    const absR = r + rowStart;
    const cells: Array<TableCell> = [];
    // Columns swallowed by a cell overflowing rightwards over them (see the
    // overflow block below). Reset per row; filled left-to-right, so a column
    // is always marked before the loop reaches it.
    const overflowed = new Set<number>();
    // The local column each emitted cell starts at, and the right-aligned cells
    // that want space to their left — both consumed by absorbLeftwards below.
    const cellColumns: Array<number> = [];
    const leftOverflow: Array<LeftOverflow> = [];
    for (let c = 0; c < colCount; c++) {
      if (hiddenCols.has(c) || overflowed.has(c)) continue;
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
      // General is not a fixed format: a spreadsheet shows as many decimals as
      // the column has room for and ROUNDS to that, and falls back to
      // scientific notation when the integer part alone will not fit. Rendering
      // every stored digit and letting the cell clip it turns
      // 4.3900875881221957 into "4.390087" (Sparklines.xlsx) and 1161014163
      // into "1161014" (escape-unicode.xlsx) — the second reads as a number a
      // thousand times smaller, with nothing to say a digit was cut.
      // …but only for a cell that HOLDS a number. §18.3.1.4 gives `t` the
      // default "n", so `<c r="C49" s="3"/>` — a cell present only to carry a
      // style — arrives here typed numeric with an empty value, and `Number('')`
      // is 0, not NaN. Rounding that emptiness printed a `0` in every styled
      // blank: invalid_ext_data_validation.xlsx grew a whole page holding one.
      if (
        ws?.type === 'n' &&
        ws.rawValue.length > 0 &&
        (styles.cellXfs[ws.styleIndex ?? 0]?.numFmtId ?? 0) === 0
      ) {
        const unit = charTwips(styles.cellXfs[ws.styleIndex ?? 0], styles, textTwipsUnit);
        // Whether a rendering fits is measured the way everything else on this
        // page is: charWidthUnits, which reports the face we DRAW in.
        if (unit > 0) {
          const room = columnWidths[c]! / unit;
          text = generalToWidth(ws.rawValue, (candidate) => estimateChars(candidate) <= room);
        }
      }
      // The full (pre-truncation) text feeds conditional-format text/dup rules (W5).
      const cfText = text.length > 0 ? text : undefined;
      if (text.length > textBudget) {
        text = text.slice(0, Math.max(0, textBudget));
        if (!textBudgetReported) {
          textBudgetReported = true;
          print.losses?.push({
            severity: 'dropped',
            feature: FEATURES.text,
            detail: `per-sheet text budget of ${MAX_SHEET_TEXT_CHARS} characters exhausted; the remaining cells render empty`,
            ...(print.sheetName ? { where: `sheet "${print.sheetName}"` } : {}),
          });
        }
      }
      textBudget -= text.length;
      // §18.3.1.4: `s` is optional and defaults to 0 — a cell without it is
      // formatted by cellXfs[0], the workbook's Normal style, not by nothing at
      // all. Treating "absent" as "unstyled" gave every such cell the layout's
      // own defaults: simple-monthly-budget.xlsx writes all of its item labels
      // without `s`, so a 9pt slate-blue table came out 11pt black.
      const xf = ws ? styles.cellXfs[ws.styleIndex ?? 0] : undefined;
      let runProps = cellRunProps(xf);
      // §18.8.31: the section that applied may name a colour — `[Red]-#,##0.00`
      // is how every accounting format marks a negative. It belongs to the
      // format, not to the font, and conditional formatting still overrides it.
      if (ws?.type === 'n' && xf?.numFmtId) {
        const fmtColor = numberFormatColorHex(ws.rawValue, xf.numFmtId, styles.numFmts);
        if (fmtColor !== undefined && fmtColor !== runProps.colorHex) {
          runProps = { ...runProps, colorHex: fmtColor };
        }
      }
      const alignment = alignmentFromXf(xf, ws?.type);
      let shading = xf ? shadingFromXf(xf, styles) : undefined;
      // A table's banded/header fill + header text colour sit below the cell's
      // own fill (used only when the cell declares none) and below conditional
      // formatting (E-SHEET SC3).
      const tableFmt = tableFormatByCell.get(key(absR, absC));
      if (!shading && tableFmt?.shading) shading = tableFmt.shading;
      if (tableFmt?.fontColorHex) runProps = { ...runProps, colorHex: tableFmt.fontColorHex };
      let borders = xf ? bordersFromXf(xf, styles) : undefined;
      let dataBar: CellDataBar | undefined;
      let icon: CellIcon | undefined;
      const sparkline = sparklineByCell.get(key(absR, absC));

      // Conditional formatting (E-SHEET SC1/SC1b/SC1c): the applicable rules
      // override the cell's fill/font and may add an in-cell data bar. Only number
      // cells carry a comparable value; no formatter ⇒ block skipped (byte-identical).
      if (cfFormatter && ws) {
        // §18.3.1.4 types a value-less cell "n" by default, and `Number('')` is
        // 0 — which is exactly how Excel and Calc compare it: a rule written
        // `cellIs equal 0` colours the empty cells too, and 48539.xlsx paints
        // its whole Pass/Fail column that way. What a blank is NOT is a value
        // that can repeat, so it goes to the formatter flagged.
        const blank = ws.rawValue.length === 0 && ws.inlineText === undefined;
        const cfValue =
          ws.type === 'n' && Number.isFinite(Number(ws.rawValue)) ? Number(ws.rawValue) : undefined;
        const over = cfFormatter(absR, absC, cfValue, cfText, blank);
        if (over) {
          if (over.fillHex) shading = { colorHex: over.fillHex };
          if (over.dataBar) dataBar = over.dataBar;
          // §18.8.9 — a rule may change how the VALUE reads, not just how the
          // cell looks. Every dxf in new_cond_format_test.xlsx is a number
          // format and nothing else, so every one of its rules was a no-op.
          if (over.numberFormat !== undefined && ws.type === 'n' && ws.rawValue.length > 0) {
            text = applyNumberFormat(
              ws.rawValue,
              CF_NUMBER_FORMAT_ID,
              new Map([[CF_NUMBER_FORMAT_ID, over.numberFormat]]),
              date1904,
            );
          }
          // §18.3.1.28 "Show Bar Only": the bar IS the cell's rendering, and
          // the number would sit on top of its own gauge. simple-monthly-budget
          // printed 2336 across the bar the reference draws bare.
          if (over.hideValue) text = '';
          if (over.icon) icon = over.icon;
          // §18.8.9 — a rule's edges win over the cell's own, edge by edge; a
          // border-only dxf is a whole rule that does nothing else.
          if (over.border) borders = { ...borders, ...mapXlsxBorder(over.border) };
          runProps = applyCfOverride(runProps, over);
        }
      }

      // Columns this cell's text runs over, set by the overflow block below.
      let overflowSpan = 1;
      // The right rule the run takes over from the last cell it absorbs.
      let overflowRightRule: Border | undefined;
      // §18.8.1 wrapText (E-SHEET W6): a wrapped cell keeps its full text — the
      // table cell layout breaks it to the cell width and the row grows (atLeast).
      const wrapText = xf?.alignment?.wrapText === true;
      // §18.8.1 textRotation (E-SHEET W6): a rotated / vertical cell renders its
      // text stacked top-to-bottom (one glyph per line) — the faithful flowed
      // rendering for the dominant vertical-header case (textRotation 255 exactly,
      // ±90° in reading orientation). The row grows to fit, like Excel.
      const rotation = xf?.alignment?.textRotation;
      const rotated = rotation !== undefined && rotation !== 0 && text.length > 0;
      // A shrinkToFit cell keeps its full text too — the layout scales the font to
      // fit rather than clipping (Excel shrinks instead of overflowing).
      const shrinkToFit = xf?.alignment?.shrinkToFit === true && !wrapText && !rotated;
      // Cell overflow (Excel/Calc print model): a non-wrapping cell's text
      // overflows into EMPTY neighbours to the right (left/general alignment) but
      // is CLIPPED where an occupied cell blocks it — Calc renders only the part
      // that fits, dropping the rest. We mirror that for string cells (a wrapText,
      // rotated or shrinkToFit cell is exempt) so the rendered text matches Calc's.
      if (
        text.length > 0 &&
        !merge &&
        !wrapText &&
        !rotated &&
        !shrinkToFit &&
        ws &&
        (ws.type === 's' || ws.type === 'str' || ws.type === 'inlineStr') &&
        // A CENTRED cell spills as well — Excel and Calc run it out both ways.
        // Requiring left alignment kept tdf171828.xlsx's centred "unter
        // Berücksichtung der Sondertilgungen" inside its own 73pt column, where
        // it was cut to "unter Berücksic"; every other reader runs it across.
        (alignment === 'left' || alignment === 'center')
      ) {
        let availTwips = columnWidths[c]!;
        let cc = c + 1;
        // What blocks Excel's overflow is CONTENT — the text is drawn straight
        // across a filled or bordered neighbour, which keeps everything it
        // holds. We implement overflow as a column span, and a span erases what
        // it covers, so decoration is not free for us: it is safe to span only
        // where nothing would be lost. That is the common case in practice,
        // because a decorated empty neighbour usually belongs to the same band
        // as the cell overflowing into it — tdf171828.xlsx puts its labels in a
        // filled, top-ruled block and we clipped "Kreditsumme" to "Kre".
        // Anywhere the paint actually differs we still stop and clip, which is
        // visibly short but never visibly wrong.
        const neighbourIsFree = (col: number): boolean =>
          !cellHasContent(cellMatrix[r]?.[col]) &&
          // A cell inside a merge belongs to that merge, empty or not. Spanning
          // over its ORIGIN made the origin's own span disappear and the rest
          // of the merge paint nothing at all — tdf171828.xlsx lost the fill
          // and the thick rule over a whole column of its "mtl. Betrag" row.
          !insideMerge.has(key(absR, col + colStart)) &&
          !mergeOrigins.has(key(absR, col + colStart)) &&
          (!cellPaintsSomething(cellMatrix[r]?.[col], styles) ||
            spanPreservesPaint(cellMatrix[r]?.[col], styles, shading, borders)) &&
          !sparklineByCell.has(key(absR, col + colStart)) &&
          !(dropdownRanges.length > 0 && rangesCover(dropdownRanges, absR, col + colStart));
        const bandEnd = bandEndOfCol.get(c) ?? colCount - 1;
        // A hidden column is not on the page: it neither blocks the text nor
        // gives it any room, and — crucially — it is not one of the columns a
        // colSpan counts (see below).
        while (cc <= bandEnd && (hiddenCols.has(cc) || neighbourIsFree(cc))) {
          if (!hiddenCols.has(cc)) availTwips += columnWidths[cc]!;
          cc++;
        }
        // And one more when the only thing in the way is that cell's own right
        // rule — the closing edge of the band, which the span takes over. See
        // absorbableRightRule.
        if (
          cc <= bandEnd &&
          !cellHasContent(cellMatrix[r]?.[cc]) &&
          !sparklineByCell.has(key(absR, cc + colStart)) &&
          !(dropdownRanges.length > 0 && rangesCover(dropdownRanges, absR, cc + colStart))
        ) {
          const rule = absorbableRightRule(cellMatrix[r]?.[cc], styles, shading, borders);
          if (rule) {
            availTwips += columnWidths[cc]!;
            overflowRightRule = rule;
            cc++;
          }
        }
        // Cut to the width the cell will have. The estimate cannot be exact —
        // the font is not known here — so CellProperties.noWrap below makes the
        // layout drop whatever still spills onto a second line. The cut is what
        // handles the case the layout cannot: a single unbreakable word, which
        // has no line break to drop and would otherwise run past the cell.
        text = clipToWidth(text, availTwips, charTwips(xf, styles, textTwipsUnit));
        // Claim the empty neighbours the text runs over. Without this the cell
        // keeps its single column's width and the layout WRAPS the text inside
        // it — three stacked lines where Excel and LibreOffice draw one. The
        // grid still totals the same width, so nothing else on the row moves.
        if (cc > c + 1) {
          // In VISIBLE columns. A span is consumed against the emitted grid,
          // which skips hidden columns, so counting absolute ones overruns it
          // by however many are hidden inside the run — and an overrunning
          // span swallows the cell past its end. tdf100034.xlsx hides two
          // columns mid-row and lost the value in the one after them
          // ("Social Advocates for Youth", 100.00) on two of its four pages.
          overflowSpan = 0;
          for (let k = c; k < cc; k++) if (!hiddenCols.has(k)) overflowSpan++;
          for (let k = c + 1; k < cc; k++) overflowed.add(k);
        }
      }

      // §18.8.1 shrinkToFit (E-SHEET W6): instead of clipping, scale the font down
      // so the text fits its own column on one line. The grid auto-sizes columns to
      // content, so we shrink against the column's character capacity (the same
      // char model the clip uses) rather than a post-layout width.
      if (shrinkToFit && ws && text.length > 0) {
        const charsFit = Math.max(1, columnWidths[c]! / TWIPS_PER_EXCEL_CHAR);
        if (text.length > charsFit) {
          // `xf` is non-nullish here (shrinkToFit was derived from xf.alignment).
          const baseHalfPt = (styles.fonts[xf.fontId]?.sizePt ?? 11) * 2;
          const scaledHalfPt = Math.max(2, Math.round((baseHalfPt * charsFit) / text.length));
          runProps = { ...runProps, fontSizePt: halfPtToPt(scaledHalfPt) };
        }
      }

      // A data-validation `list` cell (E-SHEET SV1). The HTML writer paints an
      // affordance for it; the paginated layout deliberately does not — see
      // CellProperties.dropdown.
      const dropdown = dropdownRanges.length > 0 && rangesCover(dropdownRanges, absR, absC);

      // A cell covered by an external hyperlink (E-SHEET W3) → its run takes the URL.
      const href =
        print.hyperlinks && print.hyperlinks.length > 0
          ? hyperlinkUrlAt(print.hyperlinks, absR, absC)
          : undefined;

      // Clamp a merge's horizontal span to the in-window columns so a merge
      // straddling the print-area edge cannot exceed the grid, and count it in
      // VISIBLE columns — a hidden column inside the merge is not one of the
      // grid columns the span is consumed against.
      const visibleEndCol = merge ? Math.min(merge.endColumn, colWindowEnd) : 0;
      let mergeSpan = 0;
      if (merge && visibleEndCol > merge.startColumn) {
        for (let a = merge.startColumn; a <= visibleEndCol; a++) {
          if (!hiddenCols.has(a - colStart)) mergeSpan++;
        }
      }
      const properties: CellProperties = {
        ...(mergeSpan > 1
          ? { colSpan: mergeSpan }
          : overflowSpan > 1
            ? { colSpan: overflowSpan }
            : {}),
        ...(merge && Math.min(merge.endRow, rowWindowEnd) > merge.startRow
          ? { merge: 'start' as const }
          : {}),
        ...(shading ? { shading } : {}),
        ...(dataBar ? { dataBar } : {}),
        ...(icon ? { icon } : {}),
        ...(sparkline ? { sparkline } : {}),
        ...(borders || overflowRightRule
          ? { borders: { ...borders, ...(overflowRightRule ? { right: overflowRightRule } : {}) } }
          : {}),
        ...(dropdown ? { dropdown: true } : {}),
        // §18.8.1: without wrapText a cell's text is one line, cut at its box.
        // Rotated and shrink-to-fit cells have their own handling.
        ...(!wrapText && !rotated && !shrinkToFit && !merge ? { noWrap: true } : {}),
        // §18.8.1 — a number that does not fit its column is shown as a row of
        // `#`, never truncated, because a truncated number is a DIFFERENT
        // number: "4/30/201" reads as a date and is not the one in the cell.
        // column-style-autofilter.xlsx has 3196 such cells.
        //
        // Decided HERE, in Excel's own character unit, and not from the width
        // the text finally renders at. The render font is not the workbook's,
        // and a face a few percent wider would otherwise fill a column with `#`
        // that every other reader shows the value in — trading a value that is
        // slightly wrong for no value at all. Only under a format of the cell's
        // own: a General number narrows by dropping decimals, and text stays
        // text, because a clipped word is still recognisably that word.
        ...(!wrapText &&
        !rotated &&
        !shrinkToFit &&
        !merge &&
        ws?.type === 'n' &&
        (xf?.numFmtId ?? 0) !== 0 &&
        text.length > 0 &&
        tooWideToShow(text, charTwips(xf, styles, textTwipsUnit), columnWidths[c]!)
          ? { hashOnOverflow: true }
          : {}),
        // §18.8.1 `<alignment vertical>` — a spreadsheet cell sits at the BOTTOM
        // of its box by default, which any row taller than its text shows. We
        // drew every cell against the top: on a sheet of 28.35pt rows the text
        // floated a line-height above where Excel and LibreOffice put it.
        verticalAlign: verticalAlignOf(xf),
      };

      // §18.8.1 indent (E-SHEET W6): a left indent of N levels ≈ N×3 characters,
      // applied as the paragraph's left indent on top of the cell padding.
      const indentLevels = xf?.alignment?.indent ?? 0;
      const baseParaProps = cellParaProps(alignment);
      const paragraphProps =
        indentLevels > 0
          ? { ...baseParaProps, indentLeft: twipsToPt(indentLevels * 3 * TWIPS_PER_EXCEL_CHAR) }
          : baseParaProps;
      // A shared-string cell whose index carries rich runs (E-SHEET W6) emits one
      // document-model run per <r>, each layering its <rPr> over the cell font;
      // every other cell stays a single run.
      const richRuns =
        ws && ws.type === 's' && print.sharedStringRuns
          ? print.sharedStringRuns[Number(ws.rawValue)]
          : undefined;
      const cellRuns: Array<Run> =
        richRuns && richRuns.length > 0
          ? richRuns.map((rr) => ({
              text: rr.text,
              properties: richRunProps(runProps, rr),
              ...(href ? { href } : {}),
            }))
          : text.length > 0
            ? [{ text, properties: runProps, ...(href ? { href } : {}) }]
            : [];
      // A rotated cell stacks its glyphs vertically (one centred paragraph per
      // character); every other cell is a single paragraph of its runs.
      // An EMPTY cell contributes no content at all — not an empty paragraph.
      // A paragraph with no runs still lays out a line box the height of its
      // font, which puts a ~13pt floor under every row and makes a declared
      // 12.75pt row impossible. In a spreadsheet an empty cell draws nothing and
      // the row's own height governs; fills, borders and the cell marks are cell
      // PROPERTIES and paint without it.
      // A hard line break inside a cell (§18.4.12 — a literal LF, however the
      // producer spelled it) is a break, not a character: it is the only way a
      // wrapping cell holds two lines. Drawn as text it came out as a
      // missing-glyph box mid-word. A cell that does not wrap shows one line in
      // Excel, so there the break reads as a space.
      const content: Array<BodyElement> = rotated
        ? stackedVerticalContent(text, runProps, href)
        : cellRuns.length === 0
          ? []
          : splitCellLines(cellRuns, wrapText).map((runs) => ({
              kind: 'paragraph' as const,
              paragraph: { properties: paragraphProps, runs },
            }));
      // A right-aligned cell overflows the other way — leftwards — and the
      // widest label in a form column is routinely wider than the column it
      // sits in. Without it the text is drawn from the cell's own left edge and
      // runs over whatever is to the right: tdf171828.xlsx has "Tilgungsart"
      // lying across the value beside it. Recorded here and resolved after the
      // row is built, because the free space is to the LEFT of a cell the loop
      // has already emitted.
      if (
        text.length > 0 &&
        !merge &&
        !wrapText &&
        !rotated &&
        !shrinkToFit &&
        c > 0 &&
        alignment === 'right' &&
        estimateChars(text) * charTwips(xf, styles, textTwipsUnit) > columnWidths[c]!
      ) {
        leftOverflow.push({
          index: cells.length,
          column: c,
          needTwips: estimateChars(text) * charTwips(xf, styles, textTwipsUnit) - columnWidths[c]!,
          shading,
          borders,
        });
      }
      cellColumns.push(c);
      cells.push({ properties, content });
    }
    absorbLeftwards(cells, cellColumns, leftOverflow, styles, (col) => {
      if (insideMerge.has(key(absR, col)) || mergeOrigins.has(key(absR, col))) return undefined;
      if (sparklineByCell.has(key(absR, col + colStart))) return undefined;
      if (dropdownRanges.length > 0 && rangesCover(dropdownRanges, absR, col + colStart)) {
        return undefined;
      }
      if (cellHasContent(cellMatrix[r]?.[col])) return undefined;
      return { cell: cellMatrix[r]?.[col], widthTwips: columnWidths[col]! };
    });
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
        ? {
            height: twipsToPt(rowHeightTwips),
            heightRule: baseRowProps?.heightRule ?? ('atLeast' as const),
          }
        : {}),
      ...(isTitleRow ? { isHeader: true } : {}),
      ...(breakRows.has(absR) ? { pageBreakBefore: true } : {}),
    };
    if (isTitleRow && titleRowIndex < 0) titleRowIndex = rows.length;
    rows.push({ properties: rowProps, cells });
  }

  // Gridlines: Excel/Calc do NOT print cell gridlines unless <printOptions
  // gridLines="1"> is set. Default ⇒ no synthetic full grid; only borders that
  // come from cell styles are drawn. With gridLines on, lay a thin grid like a
  // print preview with "Gridlines" enabled.
  // A printed gridline is a light-grey hairline, not a rule. Left colourless it
  // painted BLACK at half a point, which is heavier than the cell borders the
  // sheet actually declares — on tdf100034.xlsx the thin bottom rules under its
  // total rows read as black against LibreOffice's grey grid and were swamped
  // by ours, five times the ink of the reference's gridline.
  const thin: Border = {
    style: 'single',
    width: eighthPtToPt(2),
    colorHex: PRINT_GRIDLINE_HEX,
  };
  // <printOptions horizontalCentered="1"> centers the sheet within the print
  // margins.
  const centered = worksheet.printOptions?.horizontalCentered === true;
  const tableProperties: TableProperties = {
    // A spreadsheet cell insets its text by about 2 px (1.5 pt at 96 DPI), not
    // by a word processor's 108 twips / 5.4 pt. The wider inset shifted every
    // left-aligned value right and every right-aligned one left, and it eats
    // into a column whose width the author fixed — Excel does not reflow a
    // narrow column to make room for padding.
    // Vertically there is no inset at all: in a spreadsheet the row height IS
    // the cell box, which is the whole point of a declared `ht`. Leaving the
    // word-processor default of 5.4pt top AND bottom put a 10.8pt floor under
    // every row, so a 12.75pt row could not render at 12.75pt and a sheet of
    // 173 empty rows needed three pages where one was asked for.
    defaultCellMargins: {
      left: pt(EXCEL_CELL_INSET_PT),
      right: pt(EXCEL_CELL_INSET_PT),
      top: pt(0),
      bottom: pt(0),
    },
    // §17.4.20 tblLayout — the layout engine's default is auto-fit, where the
    // grid is only a hint and column widths are derived from cell content. That
    // is right for WordprocessingML and wrong here: `<col width="..">` is what
    // the author set and what Excel and LibreOffice print. Auto-fitting it made
    // every column after the first land somewhere else — up to 80pt off the
    // reference render on a four-column sheet whose columns were all declared
    // the same width.
    layout: 'fixed',
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
  };

  // E-SHEET SE1 — when an unscaled sheet is wider than the printable page (or
  // carries a manual column break), paginate across columns into bands instead of
  // squeezing it onto one page width. A fit-to-ONE-page sheet keeps the uniform
  // shrink path (fit-to-page overrides manual breaks in Excel). But fitToWidth=N>1
  // means "fit into N pages across" (SE-T): scale the columns, then band the
  // SCALED widths across those N (or fewer) pages.
  const contentWidthTwips = sheetContentWidthTwips(worksheet);
  const colBreaksLocal = new Set<number>();
  for (const brk of worksheet.colBreaks ?? []) {
    const local = brk - colStart;
    if (local > 0 && local < colCount) colBreaksLocal.add(local);
  }
  const fitWide = worksheet.fitToPage ? (worksheet.pageSetup?.fitToWidth ?? 1) : 1;
  // Round DOWN so the scaled columns pack into the intended page count (rounding
  // up can spill the last column of a band onto an extra page).
  const bandWidths = scaledColumnWidths(visibleWidths, printScale, scaled);
  const bandTotal = bandWidths.reduce((sum, w) => sum + w, 0);
  if (
    colCount > 1 &&
    (!scaled || fitWide > 1) &&
    (bandTotal > contentWidthTwips || colBreaksLocal.size > 0)
  ) {
    const bands = computeColumnBands(bandWidths, contentWidthTwips, colBreaksLocal);
    if (bands.length > 1) {
      if (print.bandSink) {
        print.bandSink.lefts = bands.map((band) =>
          twipsToPt(bandWidths.slice(0, band.start).reduce((sum, w) => sum + w, 0)),
        );
      }
      return bandedTables(rows, bandWidths, bands, tableProperties, titleRowIndex);
    }
  }

  // A frozen pane becomes a sticky-pane hint for the HTML writer (E-SHEET SE3).
  // Only on the single-table path — sticky across column bands is meaningless.
  const frozen =
    worksheet.pane && (worksheet.pane.frozenRows > 0 || worksheet.pane.frozenCols > 0)
      ? { rows: worksheet.pane.frozenRows, cols: worksheet.pane.frozenCols }
      : undefined;
  // §18.2.5 `_xlnm.Print_Titles` repeats its rows at the top of every page,
  // wherever they sit. The layout repeats only the leading header rows of a
  // table — a mid-table header is not a repeating title in Word, where the flag
  // comes from — so a title row with content above it (tdf171828.xlsx puts its
  // column headings in row 17, under a form block) never repeated. Cutting the
  // sheet in two at that row makes it leading, and two abutting tables of the
  // same grid lay out exactly as one.
  if (titleRowIndex > 0 && titleRowIndex < rows.length) {
    const titleStart = titleRowIndex;
    const grid = bandWidths.map((w) => twipsToPt(w));
    return [
      {
        kind: 'table',
        table: { properties: tableProperties, grid, rows: rows.slice(0, titleStart) },
      },
      { kind: 'table', table: { properties: tableProperties, grid, rows: rows.slice(titleStart) } },
    ];
  }

  const table: Table = {
    properties: frozen ? { ...tableProperties, frozen } : tableProperties,
    // The print scale shrinks the whole sheet, columns included — `bandWidths`
    // above already scales them for the banded path, and the single-table path
    // has to agree. It emitted full-width columns for a scaled sheet, which was
    // invisible while the layout auto-fitted the grid away and became a clipped
    // page the moment the grid started to count: 49156.xlsx declares
    // `scale="47"`, so its 1100pt of columns has to come down to ~517pt.
    grid: bandWidths.map((w) => twipsToPt(w)),
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

/**
 * Resolve a cell's displayed text from its RAW stored value: shared-string and
 * inline-string lookup, boolean → `TRUE`/`FALSE`, and the number-format pass
 * (§18.8) for numeric cells. Error/string/date cells pass through verbatim.
 *
 * @param cell         The parsed cell.
 * @param sharedStrings The shared-string table for `t="s"` cells.
 * @param styles       The style table (for the cell's number format).
 * @param date1904     The 1904 date system flag.
 * @returns The resolved display string.
 */
export function resolveCellText(
  cell: WorksheetCell,
  sharedStrings: ReadonlyArray<string>,
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
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const numFmtId = xf?.numFmtId ?? 0;
  return applyNumberFormat(cell.rawValue, numFmtId, styles.numFmts, date1904);
}

function runPropsFromXf(xf: XlsxCellXf, styles: XlsxStyles): RunProperties {
  // Font 0 is the workbook's Normal style, and it is a font like any other —
  // skipping it on the reasoning that "the default needs no properties" only
  // holds if the layout's default happens to match. It usually does not: a
  // workbook written against Arial declares `<sz val="10"/>`, and every cell
  // that inherits it rendered at the layout's 11pt. Ten percent oversize is not
  // a rounding difference — it inflates every row it appears in, so the drift
  // accumulates down the page (45540_classic_Header.xlsx ends ~18pt low) and,
  // because a column's width is quoted in that font's digits, spreads the
  // columns rightwards too.
  const font: XlsxFont | undefined = styles.fonts[xf.fontId];
  if (!font) return {};
  const props: { -readonly [K in keyof RunProperties]: RunProperties[K] } = {};
  if (font.bold) props.bold = true;
  if (font.italic) props.italic = true;
  if (font.underline) props.underline = 'single';
  if (font.sizePt !== undefined) props.fontSizePt = halfPtToPt(Math.round(font.sizePt * 2));
  if (font.colorHex) props.colorHex = font.colorHex;
  return props;
}

// Stacked vertical cell text (E-SHEET W6): one centred paragraph per character so
// the glyphs run top-to-bottom and the row grows to fit (Excel textRotation 255 /
// ±90° vertical text). Combining marks would ideally stay with their base, but a
// per-code-point split is a faithful approximation for the label text this targets.
function stackedVerticalContent(
  text: string,
  runProps: RunProperties,
  href: string | undefined,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  for (const ch of text) {
    out.push({
      kind: 'paragraph',
      paragraph: {
        properties: { alignment: 'center' },
        runs: [{ text: ch, properties: runProps, ...(href ? { href } : {}) }],
      },
    });
  }
  return out;
}

// Layer a rich-text run's own <rPr> over the cell's base font (E-SHEET W6). The
// run properties from the shared string win for what they set — bold/italic/
// underline, colour, size and super/subscript — inheriting the rest from the cell.
function richRunProps(base: RunProperties, rr: SheetRichRun): RunProperties {
  const out: { -readonly [K in keyof RunProperties]: RunProperties[K] } = { ...base };
  // A run's <rPr> is the whole font, not a set of additions: a run that omits
  // <b> inside a bold cell is NOT bold. Treated as "bold if set, inherit
  // otherwise", tdf171828.xlsx printed "ohne Sondertilgung" all bold where the
  // file bolds only "ohne".
  if (rr.bold !== undefined) out.bold = rr.bold;
  if (rr.italic !== undefined) out.italic = rr.italic;
  if (rr.underline) out.underline = 'single';
  if (rr.colorHex) out.colorHex = rr.colorHex;
  if (rr.sizePt !== undefined) out.fontSizePt = halfPtToPt(Math.round(rr.sizePt * 2));
  if (rr.vertAlign) out.verticalAlign = rr.vertAlign;
  return out;
}

// A conditional-format override applied over the base run props (CF wins for
// the properties it sets — font colour, bold, italic). Size is left untouched.
function applyCfOverride(base: RunProperties, o: CfOverride): RunProperties {
  return {
    ...base,
    ...(o.fontColorHex ? { colorHex: o.fontColorHex } : {}),
    ...(o.bold !== undefined ? { bold: o.bold } : {}),
    ...(o.italic !== undefined ? { italic: o.italic } : {}),
  };
}

/**
 * ECMA-376 §18.8.1 — a cell's horizontal alignment, falling back to `general`.
 *
 * "General" is not "left": it is decided by the VALUE. Numbers, dates and times
 * go right, booleans and errors centre, text goes left. Treating an absent
 * `<alignment>` as no alignment at all left every number hugging the left edge
 * of its column, tens of points from where Excel and LibreOffice put it — on
 * the sheets where it matters most, since a column of figures is the common
 * case.
 */
function alignmentFromXf(xf: XlsxCellXf | undefined, type: CellType | undefined): Alignment {
  const explicit = xf?.alignment ? mapAlignment(xf.alignment.horizontal) : undefined;
  if (explicit) return explicit;
  return generalAlignment(type);
}

function generalAlignment(type: CellType | undefined): Alignment {
  switch (type) {
    case 'n':
    case 'd':
      return 'right';
    case 'b':
    case 'e':
      return 'center';
    default:
      // 's' | 'str' | 'inlineStr' | an empty cell.
      return 'left';
  }
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

// Approximate fg coverage (0..1) of each non-solid §18.18.55 ST_PatternType, used
// to blend the pattern foreground over its background into one representative
// solid (E-SHEET W6). Excel renders these as repeating dot/line patterns; the
// print model has no tiling, so a density-weighted blend is the faithful summary.
const PATTERN_DENSITY: Readonly<Record<string, number>> = {
  darkGray: 0.75,
  mediumGray: 0.5,
  lightGray: 0.25,
  gray125: 0.125,
  gray0625: 0.0625,
  darkHorizontal: 0.5,
  darkVertical: 0.5,
  darkDown: 0.5,
  darkUp: 0.5,
  darkGrid: 0.62,
  darkTrellis: 0.75,
  lightHorizontal: 0.25,
  lightVertical: 0.25,
  lightDown: 0.25,
  lightUp: 0.25,
  lightGrid: 0.37,
  lightTrellis: 0.5,
};

function shadingFromXf(xf: XlsxCellXf, styles: XlsxStyles): CellShading | undefined {
  // Apply the fill when applyFill is explicitly true OR fillId > 1 (Excel
  // reserves fillId 0=none, 1=gray125 system fills — any user fill starts at 2).
  if (xf.applyFill === false) return undefined;
  if (xf.fillId === 0 || xf.fillId === 1) {
    if (!xf.applyFill) return undefined;
  }
  const fill = styles.fills[xf.fillId];
  if (!fill) return undefined;
  const pattern = fill.patternType;
  if (!pattern || pattern === 'none') return undefined;
  if (pattern === 'solid') return fill.fgColorHex ? { colorHex: fill.fgColorHex } : undefined;
  // §18.8.20 a non-solid patternFill (E-SHEET W6) → a solid blend of fg over bg.
  // A gradientFill (no patternType) is summarised by the reader into fgColorHex.
  const density = PATTERN_DENSITY[pattern];
  if (density === undefined || !fill.fgColorHex) return undefined;
  const blended = blendHex(fill.fgColorHex, fill.bgColorHex ?? 'FFFFFF', density);
  return blended ? { colorHex: blended } : undefined;
}

// Blend two 6-hex colours: `fraction` of `fg` over `(1 - fraction)` of `bg`.
function blendHex(fg: string, bg: string, fraction: number): string | undefined {
  const a = hexToRgb(fg);
  const b = hexToRgb(bg);
  if (!a || !b) return undefined;
  const mix = (x: number, y: number): string =>
    Math.max(0, Math.min(255, Math.round(x * fraction + y * (1 - fraction))))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `${mix(a[0], b[0])}${mix(a[1], b[1])}${mix(a[2], b[2])}`;
}

function hexToRgb(hex: string): readonly [number, number, number] | undefined {
  const h = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9A-Fa-f]{6}$/.test(h)) return undefined;
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
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
  // §18.8.4 diagonal strokes (E-SHEET W6): the same edge style drives one or both
  // diagonals, selected by diagonalUp / diagonalDown.
  const diagonal = border.diagonal ? mapBorderEdge(border.diagonal) : undefined;
  if (diagonal) {
    if (border.diagonalDown) out.diagonalDown = diagonal;
    if (border.diagonalUp) out.diagonalUp = diagonal;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** An {@link XlsxBorder}'s four edges as {@link CellBorders} — a dxf's, or a cell's. */
function mapXlsxBorder(border: XlsxBorder): CellBorders {
  const out: { -readonly [K in keyof CellBorders]: CellBorders[K] } = {};
  const top = mapBorderEdge(border.top);
  const right = mapBorderEdge(border.right);
  const bottom = mapBorderEdge(border.bottom);
  const left = mapBorderEdge(border.left);
  if (top) out.top = top;
  if (right) out.right = right;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  return out;
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

// §18.3.1.33 — the ranges of `list` data validations that should show an in-cell
// dropdown. ECMA's showDropDown is INVERTED ("1" HIDES the dropdown), so a list
// validation contributes its ranges unless the flag is set (E-SHEET SV1).
function listDropdownRanges(
  dvs: ReadonlyArray<DataValidation> | undefined,
): ReadonlyArray<MergedRange> {
  if (!dvs || dvs.length === 0) return [];
  const out: Array<MergedRange> = [];
  for (const dv of dvs) {
    if (dv.type !== 'list' || dv.showDropDown) continue;
    out.push(...dv.ranges);
  }
  return out;
}

function rangesCover(ranges: ReadonlyArray<MergedRange>, row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.startRow && row <= r.endRow && col >= r.startColumn && col <= r.endColumn) {
      return true;
    }
  }
  return false;
}

// §18.3.1.47 — the URL of the first hyperlink whose range covers the cell (W3).
function hyperlinkUrlAt(
  links: ReadonlyArray<SheetHyperlink>,
  row: number,
  col: number,
): string | undefined {
  for (const l of links) {
    const r = l.ref;
    if (row >= r.startRow && row <= r.endRow && col >= r.startColumn && col <= r.endColumn) {
      return l.url;
    }
  }
  return undefined;
}

// E-SHEET SC2 — resolve the sheet's sparklines to a host-cell → value-series map.
// The data range is resolved on THIS sheet (parseAreaRef drops any sheet
// qualifier); cross-sheet series are a documented v1 limitation. Empty map when
// the sheet has no sparklines, so the cell loop is untouched for everyone else.
function buildSparklineLookup(
  worksheet: ParsedWorksheet,
  print: PrintModelOptions,
): Map<string, CellSparkline> {
  const out = new Map<string, CellSparkline>();
  for (const sp of worksheet.sparklines ?? []) {
    const host = parseAreaRef(sp.sqref);
    const area = parseAreaRef(sp.dataRange);
    if (!host || !area) continue;
    const grid = resolveSeriesGrid(sp.dataRange, worksheet, print.sheetGrids);
    const values = collectSeriesValues(grid.cells, area, print);
    if (values.length === 0 || values.every((v) => v === null)) continue;
    out.set(key(host.startRow, host.startColumn), {
      kind: sp.kind,
      values,
      ...(sp.colorHex ? { colorHex: sp.colorHex } : {}),
    });
  }
  return out;
}

// A sheet-qualified data range (Sheet2!A1:C1, or 'My Sheet'!…) resolves against
// the named sheet's grid; an unqualified range (or an unknown sheet) stays on
// the current sheet (E-SHEET SC2 tail TC3).
function resolveSeriesGrid(
  dataRange: string,
  current: ParsedWorksheet,
  sheetGrids: ReadonlyMap<string, ParsedWorksheet> | undefined,
): ParsedWorksheet {
  if (!sheetGrids) return current;
  const firstToken = dataRange.split(',')[0] ?? dataRange;
  const bang = firstToken.lastIndexOf('!');
  if (bang < 0) return current;
  let name = firstToken.slice(0, bang).trim();
  if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
  return sheetGrids.get(name) ?? current;
}

// The numeric series inside an area, in reading order (row-major). Blanks and
// non-numeric cells are kept as gaps (null) so x-positions stay aligned. A range
// far larger than any real sparkline (e.g. a whole column) falls back to the
// compact populated series rather than enumerating millions of gaps.
const MAX_SPARKLINE_POINTS = 1000;

function collectSeriesValues(
  cells: ReadonlyArray<WorksheetCell>,
  area: CellRange,
  print: PrintModelOptions,
): Array<number | null> {
  const byKey = new Map<string, number>();
  for (const c of cells) {
    if (c.row < area.startRow || c.row > area.endRow) continue;
    if (c.column < area.startColumn || c.column > area.endColumn) continue;
    if (c.type !== 'n') continue;
    const v = Number(c.rawValue);
    if (Number.isFinite(v)) byKey.set(key(c.row, c.column), v);
  }
  const cellCount = (area.endRow - area.startRow + 1) * (area.endColumn - area.startColumn + 1);
  if (cellCount > MAX_SPARKLINE_POINTS) {
    // The populated values all survive; what is lost are the blank gaps that
    // held the x-positions apart, so the plotted spacing no longer matches.
    print.losses?.push({
      severity: 'degraded',
      feature: FEATURES.charts,
      detail: `sparkline range spans ${cellCount} cells (over ${MAX_SPARKLINE_POINTS}); plotted as the compact populated series, so blank gaps no longer hold their x-positions`,
      ...(print.sheetName ? { where: `sheet "${print.sheetName}"` } : {}),
    });
    const pts = [...byKey.entries()].map(([k, v]) => {
      const [r, col] = k.split(',').map(Number) as [number, number];
      return { r, col, v };
    });
    pts.sort((a, b) => a.r - b.r || a.col - b.col);
    return pts.map((p) => p.v);
  }
  const out: Array<number | null> = [];
  for (let r = area.startRow; r <= area.endRow; r++) {
    for (let col = area.startColumn; col <= area.endColumn; col++) {
      out.push(byKey.get(key(r, col)) ?? null);
    }
  }
  return out;
}

// E-SHEET SC3 — a table cell's resolved fill + (for header cells) font colour.
interface TableCellFormat {
  readonly shading?: CellShading;
  readonly fontColorHex?: string;
}

// A pivot rowItem @t that marks a total row — 'grand' (grand total) or any
// subtotal-function name; absent / 'data' / 'blank' are ordinary data rows.
function isPivotTotal(t: string | undefined): boolean {
  return t !== undefined && t !== 'data' && t !== 'blank';
}

// Cell (absolute key) → table/pivot format for header rows and banded data rows.
// The header rows take the header fill + text colour (white on a Medium/Dark
// accent); with showRowStripes, the 2nd/4th/… data row takes the band colour
// (band1 stays unfilled, like Excel). Excel tables count headers as
// headerRowCount; pivots as firstDataRow (the offset to the first data row).
// Bounded by real table/pivot sizes (E-SHEET SC3, E-PIVOT PV2).
function buildTableFormatLookup(worksheet: ParsedWorksheet): Map<string, TableCellFormat> {
  const out = new Map<string, TableCellFormat>();
  const band = (
    ref: MergedRange,
    headerRows: number,
    style: {
      headerHex?: string;
      bandHex?: string;
      headerTextHex?: string;
      showRowStripes: boolean;
    },
    // A pivot total/subtotal data row (by 0-based offset): emphasised like the
    // header rather than striped (E-PIVOT PV3). Tables pass none.
    isTotalRow?: (dataOffset: number) => boolean,
  ): void => {
    const firstDataRow = ref.startRow + headerRows;
    for (let r = ref.startRow; r <= ref.endRow; r++) {
      let colorHex: string | undefined;
      let fontColorHex: string | undefined;
      if (r < firstDataRow) {
        colorHex = style.headerHex;
        fontColorHex = style.headerTextHex;
      } else if (isTotalRow?.(r - firstDataRow)) {
        colorHex = style.headerHex;
        fontColorHex = style.headerTextHex;
      } else if (style.showRowStripes && style.bandHex) {
        colorHex = (r - firstDataRow) % 2 === 1 ? style.bandHex : undefined;
      }
      if (!colorHex && !fontColorHex) continue;
      const fmt: TableCellFormat = {
        ...(colorHex ? { shading: { colorHex } } : {}),
        ...(fontColorHex ? { fontColorHex } : {}),
      };
      for (let c = ref.startColumn; c <= ref.endColumn; c++) out.set(key(r, c), fmt);
    }
  };
  for (const t of worksheet.tables ?? []) band(t.ref, t.headerRowCount, t);
  for (const p of worksheet.pivotTables ?? [])
    band(p.ref, p.firstDataRow, p, (off) => isPivotTotal(p.rowItemTypes?.[off]));
  // Overlay grand-total / subtotal COLUMNS with the header emphasis (E-PIVOT
  // PV4), overriding whatever the row pass banded in that column.
  for (const p of worksheet.pivotTables ?? []) {
    if (!p.colItemTypes || p.headerHex === undefined) continue;
    const firstDataRow = p.ref.startRow + p.firstDataRow;
    const firstDataCol = p.ref.startColumn + p.firstDataCol;
    const totalFmt: TableCellFormat = {
      shading: { colorHex: p.headerHex },
      ...(p.headerTextHex ? { fontColorHex: p.headerTextHex } : {}),
    };
    for (let i = 0; i < p.colItemTypes.length; i++) {
      if (!isPivotTotal(p.colItemTypes[i])) continue;
      const c = firstDataCol + i;
      if (c > p.ref.endColumn) continue;
      for (let r = firstDataRow; r <= p.ref.endRow; r++) out.set(key(r, c), totalFmt);
    }
  }
  return out;
}

// A cell "has content" (blocks overflow / counts toward the used range) when it
// carries a value or inline text — empty styled cells do not.
function cellHasContent(cell: WorksheetCell | undefined): boolean {
  return !!cell && (cell.rawValue !== '' || cell.inlineText !== undefined);
}

/**
 * The window-local index of the last row inside `[rowStart, rowStart+rowCount)`
 * that carries anything — a value, or a merge reaching into the window. `-1`
 * when the window is entirely blank.
 */
function lastContentRow(
  worksheet: ParsedWorksheet,
  rowStart: number,
  rowCount: number,
  colStart: number,
  colCount: number,
): number {
  const rowEnd = rowStart + rowCount - 1;
  const colEnd = colStart + colCount - 1;
  let last = -1;
  for (const c of worksheet.cells) {
    if (!cellHasContent(c)) continue;
    if (c.row < rowStart || c.row > rowEnd || c.column < colStart || c.column > colEnd) continue;
    if (c.row - rowStart > last) last = c.row - rowStart;
  }
  for (const m of worksheet.merges) {
    if (m.startColumn > colEnd || m.endColumn < colStart) continue;
    const reach = Math.min(m.endRow, rowEnd) - rowStart;
    if (reach > last) last = reach;
  }
  return last;
}

/** The column twin of {@link lastContentRow}. */
function lastContentColumn(
  worksheet: ParsedWorksheet,
  rowStart: number,
  rowCount: number,
  colStart: number,
  colCount: number,
): number {
  const rowEnd = rowStart + rowCount - 1;
  const colEnd = colStart + colCount - 1;
  let last = -1;
  for (const c of worksheet.cells) {
    if (!cellHasContent(c)) continue;
    if (c.row < rowStart || c.row > rowEnd || c.column < colStart || c.column > colEnd) continue;
    if (c.column - colStart > last) last = c.column - colStart;
  }
  for (const m of worksheet.merges) {
    if (m.startRow > rowEnd || m.endRow < rowStart) continue;
    const reach = Math.min(m.endColumn, colEnd) - colStart;
    if (reach > last) last = reach;
  }
  return last;
}

/**
 * Whether an EMPTY cell still draws something of its own — a fill or a border.
 *
 * Such a cell cannot be swallowed by a neighbour's overflowing text: the span
 * that gives the text its width would take the paint with it.
 */
function cellPaintsSomething(cell: WorksheetCell | undefined, styles: XlsxStyles): boolean {
  if (!cell || cell.styleIndex === undefined) return false;
  const xf = styles.cellXfs[cell.styleIndex];
  if (!xf) return false;
  return shadingFromXf(xf, styles) !== undefined || bordersFromXf(xf, styles) !== undefined;
}

/**
 * How many empty columns past the used range the last column's text needs to
 * run into, bounded by the printable width (Excel stops at the page edge too).
 *
 * Only the last used column can want them — anywhere else the grid already has
 * neighbours. Zero for the overwhelming majority of sheets, which keeps their
 * projection exactly as it was.
 */
function overflowColumnsPastUsedRange(
  worksheet: ParsedWorksheet,
  usedCol: number,
  sharedStrings: ReadonlyArray<string>,
  styles: XlsxStyles,
  date1904: boolean,
  charTwipsUnit: number,
): number {
  const defaultTwips = defaultColumnTwips(worksheet, charTwipsUnit, DEFAULT_COL_CHARS);
  // The columns past the used range are not necessarily default-width: a `<col>`
  // range routinely covers far more columns than hold anything. Sizing the
  // budget by the default instead of by what the column will actually be made
  // the grid wider than the page, and a grid wider than the page is split into
  // bands — one extra, blank page for a document LibreOffice prints as one.
  const widthOf = (abs: number): number => {
    for (const col of worksheet.columns) {
      if (abs < col.min - 1 || abs > col.max - 1) continue;
      return col.hidden ? 0 : columnTwips(col.widthChars, charTwipsUnit);
    }
    return defaultTwips;
  };

  // The last cell of each row is the one that can overflow freely: nothing to
  // its right blocks it, so its text runs on until the page edge. Any earlier
  // cell is stopped by the next occupied one and clips inside the grid.
  const lastOfRow = new Map<number, WorksheetCell>();
  for (const cell of worksheet.cells) {
    if (!cellHasContent(cell)) continue;
    const prev = lastOfRow.get(cell.row);
    if (!prev || cell.column > prev.column) lastOfRow.set(cell.row, cell);
  }

  let needTwips = 0;
  for (const cell of lastOfRow.values()) {
    if (!(cell.type === 's' || cell.type === 'str' || cell.type === 'inlineStr')) continue;
    const xf = styles.cellXfs[cell.styleIndex ?? 0];
    const align = xf?.alignment;
    if (align?.wrapText || align?.shrinkToFit || align?.textRotation) continue;
    if (alignmentFromXf(xf, cell.type) !== 'left') continue;
    let room = 0;
    for (let abs = cell.column; abs <= usedCol; abs++) room += widthOf(abs);
    const text = resolveCellText(cell, sharedStrings, styles, date1904);
    needTwips = Math.max(
      needTwips,
      // Text is measured in Excel's own unit — see textTwipsUnit at the top of
      // worksheetToBody; charWidthUnits has already scaled into it.
      estimateChars(text) * charTwips(xf, styles, TWIPS_PER_EXCEL_CHAR) - room,
    );
  }
  if (needTwips <= 0) return 0;

  let gridTwips = 0;
  for (let abs = 0; abs <= usedCol; abs++) gridTwips += widthOf(abs);
  const limit = sheetContentWidthTwips(worksheet);

  // Bounded independently of the width budget: a run of `<col width="0.01">`
  // would otherwise take thousands of iterations to fill one page. No page can
  // show 256 columns of text, so this never binds on a real sheet.
  const MAX_OVERFLOW_COLS = 256;
  let extra = 0;
  while (needTwips > 0 && extra < MAX_OVERFLOW_COLS) {
    const w = widthOf(usedCol + extra + 1);
    if (w <= 0 || gridTwips + w > limit) break;
    gridTwips += w;
    needTwips -= w;
    extra++;
  }
  return extra;
}

/** Column widths as the bands see them: shrunk by the print scale, floor-rounded. */
function scaledColumnWidths(
  widths: ReadonlyArray<number>,
  printScale: number,
  scaled: boolean,
): Array<number> {
  return scaled ? widths.map((w) => Math.max(1, Math.floor(w * printScale))) : [...widths];
}

/**
 * Break a cell's runs at the hard line breaks inside them, or — for a cell that
 * does not wrap, and so shows one line — flatten those breaks to spaces.
 */
function splitCellLines(runs: ReadonlyArray<Run>, wrapText: boolean): Array<Array<Run>> {
  if (!runs.some((r) => /[\r\n]/.test(r.text))) return [[...runs]];
  if (!wrapText) return [runs.map((r) => ({ ...r, text: r.text.replace(/[\r\n]+/g, ' ') }))];
  const lines: Array<Array<Run>> = [[]];
  for (const run of runs) {
    const parts = run.text.split(/\r\n|\r|\n/);
    for (let i = 0; i < parts.length; i++) {
      if (i > 0) lines.push([]);
      const text = parts[i]!;
      if (text.length > 0) lines[lines.length - 1]!.push({ ...run, text });
    }
  }
  // Interior and trailing blanks are kept — "a\n\nb" is three lines in Excel
  // and the middle one is empty. Only a cell that is nothing BUT breaks
  // collapses to no content at all, the same as an empty cell.
  return lines.some((l) => l.length > 0) ? lines : [];
}

/** §18.8.1 `<alignment vertical>`, defaulting to Excel's bottom. */
function verticalAlignOf(xf: XlsxCellXf | undefined): 'top' | 'center' | 'bottom' {
  const v = xf?.alignment?.vertical;
  if (v === 'top') return 'top';
  if (v === 'center') return 'center';
  // `justify` and `distributed` spread the lines; with no such layout mode the
  // closest honest reading is the box's natural top.
  if (v === 'justify' || v === 'distributed') return 'top';
  return 'bottom';
}

/**
 * The width of one character of a cell's own font, in twips.
 *
 * {@link TWIPS_PER_EXCEL_CHAR} is the Maximum Digit Width of the DEFAULT 11pt
 * font, and column widths are measured in exactly those units — but the text
 * inside a cell is not necessarily written in it. Measuring a 7pt annotation by
 * the 11pt yardstick makes it read as half again as wide as it is, and the
 * overflow clip then cuts it in the middle of a word: tdf171828.xlsx lost the
 * "eff. Zins =" half of a label that fits its band perfectly well.
 */
function charTwips(xf: XlsxCellXf | undefined, styles: XlsxStyles, unit: number): number {
  const sizePt = (xf ? styles.fonts[xf.fontId]?.sizePt : undefined) ?? DEFAULT_FONT_PT;
  if (!Number.isFinite(sizePt) || sizePt <= 0) return unit;
  return (unit * sizePt) / DEFAULT_FONT_PT;
}

// How wide a character renders, in the column-width unit — Excel's Maximum
// Digit Width of 7 px, which is what TWIPS_PER_EXCEL_CHAR measures.
//
// Two corrections are folded into these numbers, and both matter. Text is
// proportional, so a run of spaces or of narrow letters is nowhere near as wide
// as the same count of digits — counting every character as one made a padded
// label read half again as wide as it is, and the overflow clip cut it mid-word.
// And the font that renders is not the font the column was measured against:
// Roboto's digit is 6.18 pt at 11 pt against Excel's 5.25 pt, so even an exact
// per-character model in the source font's own units would under-measure the
// page by a sixth. The factors below are measured off Roboto (Arial and
// Helvetica sit within a few percent of it) and already scaled into Excel units.
const NARROW_CHARS = new Set(' .,:;\'"!|il');
const SEMI_NARROW_CHARS = new Set('ftr()[]{}-/\\');
const WIDE_CHARS = new Set('MWmw@%');

function charWidthUnits(ch: string): number {
  if (NARROW_CHARS.has(ch)) return 0.53;
  if (SEMI_NARROW_CHARS.has(ch)) return 0.71;
  if (WIDE_CHARS.has(ch)) return 1.77;
  if (ch >= 'a' && ch <= 'z') return 1.12;
  if (ch >= 'A' && ch <= 'Z') return 1.36;
  return 1.18;
}

/**
 * Whether a number is too wide for its column to show it at all — the test
 * behind {@link CellProperties.hashOnOverflow}.
 *
 * Two corrections separate this from the clipping estimate, and both are about
 * how destructive the answer is. Filling a cell with `#` erases the value, so
 * the test has to be sure.
 *
 * `charWidthUnits` measures in the font we DRAW in, expressed in Excel's unit —
 * right for deciding what fits on the page, wrong for deciding what the
 * document itself considers too wide, since §18.3.1.13 defines that unit AS the
 * digit's width. Dividing by our own digit puts the text back in the document's
 * terms: a face 18 % wider than the workbook's would otherwise hash a column of
 * dates that every other reader shows (forum-mso-de-104083.xlsx).
 *
 * And then a tenth of headroom, because the estimate is an estimate. A value
 * that overruns by one percent is a value we should draw and let the cell clip;
 * only one that plainly cannot fit earns the hashes.
 */
function tooWideToShow(text: string, perCharTwips: number, columnTwips: number): boolean {
  const digit = charWidthUnits('0');
  if (!(digit > 0) || !(perCharTwips > 0)) return false;
  return (estimateChars(text) / digit) * perCharTwips > columnTwips * 1.1;
}

/** Width of `text` in column-width units, estimated character by character. */
function estimateChars(text: string): number {
  let n = 0;
  for (const ch of text) n += charWidthUnits(ch);
  return n;
}

/** The longest prefix of `text` that fits `availTwips`, by the same estimate. */
function clipToWidth(text: string, availTwips: number, perChar: number): string {
  if (perChar <= 0) return text;
  const budget = availTwips / perChar;
  let used = 0;
  let i = 0;
  for (const ch of text) {
    const w = charWidthUnits(ch);
    if (used + w > budget) break;
    used += w;
    i += ch.length;
  }
  return i >= text.length ? text : text.slice(0, Math.max(1, i));
}

/** A right-aligned cell that needs room to its left, recorded as its row is built. */
interface LeftOverflow {
  /** Index of the cell in the row's `cells` array. */
  readonly index: number;
  /** Its local column. */
  readonly column: number;
  /** How much wider than its own column the text is. */
  readonly needTwips: number;
  readonly shading: CellShading | undefined;
  readonly borders: CellBorders | undefined;
}

/** What a candidate column offers a leftward overflow, or undefined if it is not free. */
type LeftNeighbour = (
  column: number,
) => { cell: WorksheetCell | undefined; widthTwips: number } | undefined;

/**
 * Widen each right-aligned cell that overflows leftwards over the free columns
 * before it, deleting the cells those columns produced.
 *
 * The mirror of the rightward overflow the row loop does inline, but it cannot
 * be done inline: the space a right-aligned cell needs lies behind it, in cells
 * already emitted. Same freeness rule as rightwards — nothing in the column,
 * and nothing painted there that the span would erase.
 */
function absorbLeftwards(
  cells: Array<TableCell>,
  cellColumns: Array<number>,
  candidates: ReadonlyArray<LeftOverflow>,
  styles: XlsxStyles,
  neighbourAt: LeftNeighbour,
): void {
  // Right to left: absorbing shifts every later index, and a candidate never
  // reaches past one to its left (that one's own cell blocks it).
  for (let k = candidates.length - 1; k >= 0; k--) {
    const cand = candidates[k]!;
    let need = cand.needTwips;
    let first = cand.index;
    while (first > 0 && need > 0) {
      const prev = first - 1;
      const col = cellColumns[prev]!;
      // Only a single-column cell can be absorbed whole; a spanning one is
      // already carrying somebody's text.
      if ((cells[prev]!.properties.colSpan ?? 1) !== 1) break;
      if (cellColumns[first]! !== col + 1) break;
      const free = neighbourAt(col);
      if (!free) break;
      if (!spanPreservesPaint(free.cell, styles, cand.shading, cand.borders)) break;
      need -= free.widthTwips;
      first = prev;
    }
    if (first === cand.index) continue;
    const span = cand.index - first + 1;
    const cell = cells[cand.index]!;
    cells.splice(first, span, {
      ...cell,
      properties: { ...cell.properties, colSpan: (cell.properties.colSpan ?? 1) + span - 1 },
    });
    cellColumns.splice(first, span, cellColumns[first]!);
  }
}

/**
 * Whether spanning a cell's overflowing text across `neighbour` would leave the
 * page looking the same as drawing over it would.
 *
 * A span paints the ORIGIN's fill and the ORIGIN's outer borders across the
 * whole run, so it is lossless exactly when the neighbour's fill already
 * matches, the neighbour carries no vertical rule of its own to erase, and its
 * horizontal edges are the ones the origin will redraw anyway. A block of empty
 * cells continuing a filled, top-ruled band — which is what a decorated
 * neighbour almost always is — passes; a differently filled or boxed cell does
 * not, and the text is clipped short of it instead.
 */
function spanPreservesPaint(
  neighbour: WorksheetCell | undefined,
  styles: XlsxStyles,
  shading: CellShading | undefined,
  borders: CellBorders | undefined,
): boolean {
  if (!neighbour || neighbour.styleIndex === undefined) return true;
  const xf = styles.cellXfs[neighbour.styleIndex];
  if (!xf) return true;
  if (shadingFromXf(xf, styles)?.colorHex !== shading?.colorHex) return false;
  const nb = bordersFromXf(xf, styles);
  if (!nb) return true;
  if (nb.left || nb.right || nb.insideV || nb.diagonalUp || nb.diagonalDown) return false;
  return sameBorder(nb.top, borders?.top) && sameBorder(nb.bottom, borders?.bottom);
}

function sameBorder(a: Border | undefined, b: Border | undefined): boolean {
  if (a === undefined || b === undefined) return a === b;
  return a.style === b.style && a.width === b.width && a.colorHex === b.colorHex;
}

/**
 * The right rule a neighbour carries when that rule is the ONLY thing standing
 * between it and the overflowing text — in which case the span can simply adopt
 * it as its own right border and nothing is lost.
 *
 * {@link spanPreservesPaint} refuses any vertical rule, which is right for a
 * rule INSIDE the run (a span erases it) but wrong for the one that closes it:
 * the last cell's right edge is the boundary of the band the origin sits in,
 * and the span's own right border lands in exactly the same place. Refusing it
 * cut five labels short on every page of tdf171828.xlsx — "ohne Sondertilgung"
 * as "ohne Sondertil" — where the reference runs each of them to that very
 * edge.
 */
function absorbableRightRule(
  neighbour: WorksheetCell | undefined,
  styles: XlsxStyles,
  shading: CellShading | undefined,
  borders: CellBorders | undefined,
): Border | undefined {
  if (!neighbour || neighbour.styleIndex === undefined) return undefined;
  const xf = styles.cellXfs[neighbour.styleIndex];
  if (!xf) return undefined;
  if (shadingFromXf(xf, styles)?.colorHex !== shading?.colorHex) return undefined;
  const nb = bordersFromXf(xf, styles);
  if (!nb?.right) return undefined;
  if (nb.left || nb.insideV || nb.diagonalUp || nb.diagonalDown) return undefined;
  if (!sameBorder(nb.top, borders?.top) || !sameBorder(nb.bottom, borders?.bottom)) {
    return undefined;
  }
  return nb.right;
}

// E-SHEET SV2 — a slicer panel projected as a styled mini-table emitted after the
// grid (like chart frames). A caption header spans the button columns; each item
// is a button cell — the slicer accent fill + white text when selected, a light
// band when not. The thin box + inside rules read as the slicer's button grid.
const SLICER_WIDTH_PT = 108; // ≈ 1.5in, Excel's default slicer width
const SLICER_ROW_PT = 16;
const SLICER_UNSELECTED_HEX = 'F2F2F2';

/**
 * Project a slicer panel (E-SHEET SV2) into a styled mini-{@link Table} emitted
 * after the grid: a caption header spanning the button columns, then one button
 * cell per item — the slicer accent fill + white text when selected, a light
 * band when not; the last row padded so every row keeps the column count.
 */
export function slicerTable(slicer: SheetSlicer): Table {
  const cols = Math.max(1, slicer.columnCount);
  const colWidthPt = SLICER_WIDTH_PT / cols;
  // A printed gridline is a light-grey hairline, not a rule. Left colourless it
  // painted BLACK at half a point, which is heavier than the cell borders the
  // sheet actually declares — on tdf100034.xlsx the thin bottom rules under its
  // total rows read as black against LibreOffice's grey grid and were swamped
  // by ours, five times the ink of the reference's gridline.
  const thin: Border = {
    style: 'single',
    width: eighthPtToPt(2),
    colorHex: PRINT_GRIDLINE_HEX,
  };
  const rowProps = { height: pt(SLICER_ROW_PT), heightRule: 'atLeast' as const };
  const rows: Array<TableRow> = [];

  // Caption header spanning all columns.
  rows.push({
    properties: rowProps,
    cells: [
      {
        properties: {
          ...(cols > 1 ? { colSpan: cols } : {}),
          ...(slicer.headerHex ? { shading: { colorHex: slicer.headerHex } } : {}),
        },
        content: [
          slicerParagraph(slicer.caption, {
            bold: true,
            ...(slicer.headerTextHex ? { colorHex: slicer.headerTextHex } : {}),
          }),
        ],
      },
    ],
  });

  // Button rows: items chunked across `cols` columns; the last row is padded with
  // empty cells so every row keeps the column count.
  for (let i = 0; i < slicer.items.length; i += cols) {
    const cells: Array<TableCell> = [];
    for (let c = 0; c < cols; c++) {
      const item = slicer.items[i + c];
      if (!item) {
        cells.push({ properties: {}, content: [slicerParagraph('', {})] });
        continue;
      }
      const fill = item.selected ? slicer.selectedHex : SLICER_UNSELECTED_HEX;
      const textHex = item.selected ? slicer.selectedTextHex : undefined;
      cells.push({
        properties: fill ? { shading: { colorHex: fill } } : {},
        content: [slicerParagraph(item.label, textHex ? { colorHex: textHex } : {})],
      });
    }
    rows.push({ properties: rowProps, cells });
  }

  return {
    properties: {
      borders: { top: thin, bottom: thin, left: thin, right: thin, insideH: thin, insideV: thin },
    },
    grid: Array.from({ length: cols }, () => pt(colWidthPt)),
    rows,
  };
}

function slicerParagraph(text: string, runProps: RunProperties): BodyElement {
  return {
    kind: 'paragraph',
    paragraph: {
      properties: {},
      runs: text.length > 0 ? [{ text, properties: runProps }] : [],
    },
  };
}
