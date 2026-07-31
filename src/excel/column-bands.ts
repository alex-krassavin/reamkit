// E-SHEET SE1 — wide-sheet column-band pagination. When a sheet is wider than the
// printable page (and no print scaling is fitting it onto the page), Excel does
// NOT squeeze it — it paginates ACROSS columns: band 1 holds the leftmost columns
// that fit, band 2 the next, and so on, each band printing all of its rows before
// the next ("down, then over"). Ream emits one FlowDoc table per band; the generic
// layout paginates each band vertically and a page break separates the bands.

import type {
  BodyElement,
  CellProperties,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';

import type { Pt } from '@/core/ir';
import { twipsToPt } from '@/core/ir';

/** §18.3.1.70 — the width of the row-number column, in twips. */
const HEADING_COL_TWIPS = 460;

/** The A1-style letters of an absolute column index (0 → A, 26 → AA). */
function columnLetters(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/**
 * Wrap a grid in its printed row and column headings (§18.3.1.70 `headings`):
 * the column letters across the top, the row numbers down the left, each in a
 * thin box, the way the sheet looks on screen. A banded sheet gets its own
 * letters per band, which is what both references print.
 *
 * @param rows       The band's rows.
 * @param widths     The band's column widths in twips.
 * @param colStart   The absolute index of the band's first column.
 * @param rowNumbers The absolute 1-based row number of each row.
 * @returns The rows and widths with the heading band added.
 */
export function withHeadingBand(
  rows: ReadonlyArray<TableRow>,
  widths: ReadonlyArray<number>,
  colStart: number,
  rowNumbers: ReadonlyArray<number>,
): { rows: Array<TableRow>; widths: Array<number> } {
  const border = { style: 'single' as const, width: 0.5 as Pt, colorHex: '808080' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const head = (text: string): TableCell => ({
    properties: { borders, verticalAlign: 'center' as const },
    content: [
      {
        kind: 'paragraph' as const,
        paragraph: { properties: { alignment: 'center' as const }, runs: [{ text, properties: {} }] },
      },
    ],
  });
  const letters: Array<TableCell> = [head('')];
  for (let i = 0; i < widths.length; i++) letters.push(head(columnLetters(colStart + i)));
  // The letters repeat at the top of every page, the way both references print
  // them — which is what a table's leading header row already does.
  const out: Array<TableRow> = [{ properties: { isHeader: true }, cells: letters }];
  rows.forEach((row, i) => {
    out.push({
      properties: row.properties,
      cells: [head(String(rowNumbers[i] ?? i + 1)), ...row.cells],
    });
  });
  return { rows: out, widths: [HEADING_COL_TWIPS, ...widths] };
}

/** A contiguous run of columns that prints together as one band (E-SHEET SE1). */
export interface ColumnBand {
  /** First local column index (inclusive). */
  readonly start: number;
  /** Last local column index (inclusive). */
  readonly end: number;
}

/**
 * Greedy fill: accumulate column widths until the next column would overflow the
 * content width, then start a new band. A manual column break (`colBreaks` holds
 * local indices that begin a new page) always starts a new band. Every band keeps
 * at least one column — a single column wider than the page stands alone (the
 * layout shrinks it, the one case where fitting is unavoidable).
 *
 * @param columnWidths     Per-column widths in twips, by local index.
 * @param contentWidthTwips The printable content width one band must fit within.
 * @param colBreaks        Local column indices that force a new band.
 * @returns The bands in left-to-right order.
 */
export function computeColumnBands(
  columnWidths: ReadonlyArray<number>,
  contentWidthTwips: number,
  colBreaks: ReadonlySet<number>,
): Array<ColumnBand> {
  const n = columnWidths.length;
  if (n === 0) return [];
  const bands: Array<ColumnBand> = [];
  let start = 0;
  let acc = 0;
  for (let c = 0; c < n; c++) {
    const w = columnWidths[c] ?? 0;
    const atBreak = c > start && colBreaks.has(c);
    const overflows = c > start && acc + w > contentWidthTwips;
    if (atBreak || overflows) {
      bands.push({ start, end: c - 1 });
      start = c;
      acc = 0;
    }
    acc += w;
  }
  bands.push({ start, end: n - 1 });
  return bands;
}

/**
 * Build one table {@link BodyElement} per band from the fully materialised rows +
 * column widths (twips, local index). Bands after the first force a page break
 * before their first row, so each band starts on a fresh page. A horizontal span
 * that crosses a band boundary is clipped to the band it starts in; later bands
 * render the overlapped columns blank, matching Excel's split-at-the-break
 * behaviour.
 *
 * @param rows         The full sheet rows (every band slices the same rows).
 * @param columnWidths Per-column widths in twips, by local index.
 * @param bands        The bands from {@link computeColumnBands}.
 * @param properties   The shared table properties applied to every band.
 * @param titleRowIndex The print-title row, or -1.
 * @param drawingReachTwips How far right the sheet's drawings reach (0 if none).
 * @param headings §18.3.1.70 — the printed row/column headings, when the sheet
 *   asks for them: the absolute index of the grid's first column and the sheet
 *   row number of each emitted row. Each band gets ITS OWN letters, which is
 *   what both references print across the top of a continuation page.
 * @returns One table body element per band, in band order.
 */
export function bandedTables(
  rows: ReadonlyArray<TableRow>,
  columnWidths: ReadonlyArray<number>,
  bands: ReadonlyArray<ColumnBand>,
  properties: TableProperties,
  titleRowIndex = -1,
  drawingReachTwips = 0,
  headings?: { readonly colStart: number; readonly rowNumbers: ReadonlyArray<number> },
): Array<BodyElement> {
  return bands.flatMap((band, bandIndex) => {
    const bandLeft = columnWidths.slice(0, band.start).reduce((sum, w) => sum + w, 0);
    const bandTwips = columnWidths.slice(band.start, band.end + 1);
    const grid = bandTwips.map((w) => twipsToPt(w));
    const bandRowNumbers = [...(headings?.rowNumbers ?? [])];
    const bandRows: Array<TableRow> = rows.map((row, rowIndex) => {
      const cells = sliceRowCells(row.cells, band);
      const rowProps =
        bandIndex > 0 && rowIndex === 0
          ? { ...row.properties, pageBreakBefore: true }
          : row.properties;
      return { properties: rowProps, cells };
    });
    // A band with nothing in it draws nothing — but its rows still consume
    // their height, and on a sheet whose content sits at opposite ends of a
    // 16 000-column span that is thousands of bands' worth of blank page.
    // too-many-cols-rows.xlsx made 45 pages that way, 43 of them carrying
    // nothing but the running header and footer. The first band is kept
    // regardless, so a sheet that really is empty still renders as one.
    //
    // A drawing draws too, and it is not in any row: 57362.xlsx anchors its
    // chart at column H over nine cell-less columns, and dropping that band
    // dropped the page the chart prints on. A band the drawings reach into is
    // never empty, whatever its cells say.
    const hostsDrawing = bandLeft < drawingReachTwips;
    if (bandIndex > 0 && !hostsDrawing && !bandRows.some(rowDrawsSomething)) return [];
    // Trailing rows that draw nothing IN THIS BAND still paginate, and a band
    // whose columns hold a couple of values near the top otherwise carries the
    // sheet's whole row count as blank pages: tdf171828.xlsx's second band has
    // three cells in row 30 and ran to twelve pages, eleven of them empty. The
    // used range already ends at the last row with content, so on the first
    // band — which spans it — this trims nothing.
    while (bandRows.length > 1 && !rowDrawsSomething(bandRows[bandRows.length - 1]!)) {
      bandRows.pop();
      bandRowNumbers.pop();
    }
    // Cut at the print-title row so it leads its table and repeats, keeping the
    // two halves inside their own band — the bands paginate one after the other
    // (`pageOrder="downThenOver"`), so they must not be interleaved.
    const split = titleRowIndex > 0 && titleRowIndex < bandRows.length;
    const parts: Array<{ rows: Array<TableRow>; numbers: Array<number> }> = split
      ? [
          { rows: bandRows.slice(0, titleRowIndex), numbers: bandRowNumbers.slice(0, titleRowIndex) },
          { rows: bandRows.slice(titleRowIndex), numbers: bandRowNumbers.slice(titleRowIndex) },
        ]
      : [{ rows: bandRows, numbers: bandRowNumbers }];
    return parts.map((part) => {
      const headed = headings
        ? withHeadingBand(part.rows, bandTwips, headings.colStart + band.start, part.numbers)
        : undefined;
      return {
        kind: 'table' as const,
        table: {
          properties,
          grid: headed ? headed.widths.map((w) => twipsToPt(w)) : grid,
          rows: headed ? headed.rows : part.rows,
        },
      };
    });
  });
}

/** Whether a row would put anything on the page — content, fill, border or mark. */
function rowDrawsSomething(row: TableRow): boolean {
  return row.cells.some((cell) => {
    if (cell.content.length > 0) return true;
    const p = cell.properties;
    return (
      p.shading !== undefined ||
      p.borders !== undefined ||
      p.dataBar !== undefined ||
      p.icon !== undefined ||
      p.sparkline !== undefined ||
      p.dropdown === true
    );
  });
}

// The cells of one row restricted to a band. Walk the visible-slot list tracking
// each cell's column span (a horizontal-merge origin carries colSpan; the columns
// it spans have no cell of their own) to recover each cell's start column.
function sliceRowCells(cells: ReadonlyArray<TableCell>, band: ColumnBand): Array<TableCell> {
  const out: Array<TableCell> = [];
  let col = 0;
  for (const cell of cells) {
    const span = cell.properties.colSpan ?? 1;
    const cellStart = col;
    const cellEnd = col + span - 1;
    col += span;
    if (cellEnd < band.start || cellStart > band.end) continue;
    const clampedSpan = Math.min(cellEnd, band.end) - Math.max(cellStart, band.start) + 1;
    if (cellStart >= band.start) {
      // Starts within the band — keep it, narrowing the span if it runs past the end.
      out.push(clampedSpan === span ? cell : withColSpan(cell, clampedSpan));
    } else {
      // Spans in from an earlier band — its content already printed there; blank.
      out.push(blankCell(clampedSpan));
    }
  }
  return out;
}

function withColSpan(cell: TableCell, span: number): TableCell {
  const props: { -readonly [K in keyof CellProperties]: CellProperties[K] } = {
    ...cell.properties,
  };
  if (span > 1) props.colSpan = span;
  else delete props.colSpan;
  return { ...cell, properties: props };
}

function blankCell(span: number): TableCell {
  return {
    properties: span > 1 ? { colSpan: span } : {},
    content: [],
  };
}
