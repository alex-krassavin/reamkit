// BIFF chart reader (XLS-6) — a `.xls` embedded chart is a nested BOF…EOF
// substream of chart records. Unlike OOXML (which caches the plotted numbers in
// the chart part), a BIFF chart references worksheet ranges through its AI
// records and reads the values straight from the cells — the same cached-grid
// approach the formula engine uses. We pull the chart type, each series' values
// (from its values-AI range), the shared categories and the series names, and map
// them onto the Chart model the renderer already draws.

import type { Chart, ChartSeries, ChartType } from '@/core/document-model';
import type { WorksheetCell } from '@/core/spreadsheet-model';

interface Rec {
  readonly type: number;
  readonly data: Uint8Array;
}

interface Range {
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

interface SeriesAcc {
  values?: Range;
  categories?: Range;
  nameRef?: Range;
  nameText?: string;
}

// Chart-type group records (MS-XLS §2.4.*).
const REC_BAR = 0x1017;
const REC_LINE = 0x1018;
const REC_PIE = 0x1019;
const REC_AREA = 0x101a;
const REC_SCATTER = 0x101b;
const REC_SERIES = 0x1003;
const REC_AI = 0x1051;
const REC_SERIESTEXT = 0x100d;

const COLS = 16384;

// The chart substream's records + the host sheet's cells/strings → a Chart, or
// undefined when there are no series to plot.
export function parseBiffChart(
  records: ReadonlyArray<Rec>,
  cells: ReadonlyArray<WorksheetCell>,
  sharedStrings: ReadonlyArray<string>,
): Chart | undefined {
  let type: ChartType = 'unknown';
  let barDir: 'col' | 'bar' | undefined;
  const accs: Array<SeriesAcc> = [];
  let current: SeriesAcc | undefined;

  for (const rec of records) {
    switch (rec.type) {
      case REC_BAR:
        type = 'bar';
        barDir = (u16(rec.data, 4) & 0x01) !== 0 ? 'bar' : 'col';
        break;
      case REC_LINE:
        type = 'line';
        break;
      case REC_PIE:
        type = 'pie';
        break;
      case REC_AREA:
        type = 'area';
        break;
      case REC_SCATTER:
        type = 'scatter';
        break;
      case REC_SERIES:
        current = {};
        accs.push(current);
        break;
      case REC_AI: {
        if (!current) break;
        const id = rec.data[0] ?? 0;
        const cce = u16(rec.data, 6);
        const range = parseAreaPtg(rec.data.subarray(8, 8 + cce));
        if (range) {
          if (id === 1)
            current.values = range; // values
          else if (id === 2)
            current.categories = range; // category labels
          else if (id === 0) current.nameRef = range; // series name
        }
        break;
      }
      case REC_SERIESTEXT:
        if (current) current.nameText = parseSeriesText(rec.data);
        break;
    }
  }

  if (accs.length === 0) return undefined;

  const byKey = new Map<number, WorksheetCell>();
  for (const c of cells) byKey.set(c.row * COLS + c.column, c);
  const numAt = (r: number, c: number): number => {
    const cell = byKey.get(r * COLS + c);
    if (cell?.type !== 'n') return 0;
    const n = Number(cell.rawValue);
    return Number.isFinite(n) ? n : 0;
  };
  const textAt = (r: number, c: number): string => {
    const cell = byKey.get(r * COLS + c);
    if (!cell) return '';
    if (cell.type === 's') return sharedStrings[Number(cell.rawValue)] ?? '';
    if (cell.type === 'inlineStr') return cell.inlineText ?? '';
    return cell.rawValue;
  };

  let categories: ReadonlyArray<string> = [];
  const series: Array<ChartSeries> = [];
  for (const acc of accs) {
    const values = acc.values ? rangeCells(acc.values).map(([r, c]) => numAt(r, c)) : [];
    if (acc.categories && categories.length === 0) {
      categories = rangeCells(acc.categories).map(([r, c]) => textAt(r, c));
    }
    const name = acc.nameText ?? (acc.nameRef ? textAt(acc.nameRef.r0, acc.nameRef.c0) : undefined);
    series.push({ values, ...(name ? { name } : {}) });
  }
  if (series.every((s) => s.values.length === 0)) return undefined;

  return {
    type,
    categories,
    series,
    hasLegend: true,
    ...(barDir ? { barDir } : {}),
  };
}

// The (row, col) cells of a range, in vector order (row-major) — a single-column
// or single-row range yields its values in order.
function rangeCells(range: Range): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (let r = range.r0; r <= range.r1; r++) {
    for (let c = range.c0; c <= range.c1; c++) out.push([r, c]);
  }
  return out;
}

// Parse an AI formula's ptg into a single range. Handles ptgArea (0x05 class) and
// ptgArea3d (0x1B class, which carries a leading ixti sheet index). Anything else
// (a name, a multi-token formula) is unsupported → undefined.
function parseAreaPtg(rgce: Uint8Array): Range | undefined {
  if (rgce.length < 1) return undefined;
  const base = rgce[0]! & 0x1f;
  let off = 1;
  if (base === 0x1b)
    off = 3; // area3d: skip ixti (u16)
  else if (base !== 0x05) return undefined;
  if (off + 8 > rgce.length) return undefined;
  return {
    r0: u16(rgce, off),
    r1: u16(rgce, off + 2),
    c0: u16(rgce, off + 4) & 0x3fff,
    c1: u16(rgce, off + 6) & 0x3fff,
  };
}

// §2.4.290 SeriesText — an unused u16, then a ShortXLUnicodeString (cch, flags,
// characters).
function parseSeriesText(d: Uint8Array): string {
  if (d.length < 4) return '';
  const cch = d[2]!;
  const high = (d[3]! & 0x01) !== 0;
  let s = '';
  if (high) {
    for (let i = 0; i < cch && 4 + i * 2 + 1 < d.length; i++) {
      s += String.fromCharCode(d[4 + i * 2]! | (d[4 + i * 2 + 1]! << 8));
    }
  } else {
    for (let i = 0; i < cch && 4 + i < d.length; i++) s += String.fromCharCode(d[4 + i]!);
  }
  return s;
}

function u16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
}
