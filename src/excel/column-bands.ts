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

import { twipsToPt } from '@/core/ir';

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
 * @returns One table body element per band, in band order.
 */
export function bandedTables(
  rows: ReadonlyArray<TableRow>,
  columnWidths: ReadonlyArray<number>,
  bands: ReadonlyArray<ColumnBand>,
  properties: TableProperties,
): Array<BodyElement> {
  return bands.map((band, bandIndex) => {
    const grid = columnWidths.slice(band.start, band.end + 1).map((w) => twipsToPt(w));
    const bandRows: Array<TableRow> = rows.map((row, rowIndex) => {
      const cells = sliceRowCells(row.cells, band);
      const rowProps =
        bandIndex > 0 && rowIndex === 0
          ? { ...row.properties, pageBreakBefore: true }
          : row.properties;
      return { properties: rowProps, cells };
    });
    return { kind: 'table', table: { properties, grid, rows: bandRows } };
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
    content: [{ kind: 'paragraph', paragraph: { properties: {}, runs: [] } }],
  };
}
