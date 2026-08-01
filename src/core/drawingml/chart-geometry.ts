// Pure chart geometry: a Chart + box (w×h points) → a ChartScene of rectangles,
// polylines, wedges and labels in a LOCAL y-up frame (origin bottom-left). No
// PDF or font dependency — text widths come through an injected measure fn, so
// this module is unit-testable in isolation. The renderer converts the scene
// to draw commands (rects/polylines/wedges via the vector layer, labels via the
// text pass).

import type { Chart, ChartLineStyle, ChartMarker, ChartSeries } from '@/core/document-model';

import { applyNumberFormat } from '@/core/number-format';

/**
 * An axis-aligned rectangle in the scene's local y-up frame: bars, scatter
 * point markers, legend swatches. Position is the bottom-left corner.
 */
export interface ChartRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly fillHex?: string;
  readonly strokeHex?: string;
  readonly strokeWidthPt?: number;
}
/** An open stroked polyline: line-chart series, gridlines and axis lines. */
export interface ChartPolyline {
  readonly points: ReadonlyArray<readonly [number, number]>;
  readonly strokeHex: string;
  readonly widthPt: number;
}
/** A closed, filled polygon (area-chart bands). Drawn before strokes/labels. */
export interface ChartPolygon {
  readonly points: ReadonlyArray<readonly [number, number]>;
  readonly fillHex: string;
  readonly strokeHex?: string;
  readonly widthPt?: number;
}
/**
 * A circular sector (pie/doughnut slice), centred at `(cx, cy)` with radius `r`.
 * `startRad`/`sweepRad` are radians in the y-up frame; sweeps are negative for
 * Excel's clockwise winding. The doughnut hole is a white wedge drawn last.
 */
export interface ChartWedge {
  readonly cx: number;
  readonly cy: number;
  readonly r: number;
  readonly startRad: number;
  readonly sweepRad: number;
  readonly fillHex: string;
  readonly strokeHex?: string;
}
/** How a {@link ChartLabel} sits horizontally relative to its anchor point. */
export type LabelAlign = 'left' | 'center' | 'right';
/** A text label (title, axis tick, category, data value, legend entry). */
export interface ChartLabel {
  readonly text: string;
  /** Anchor point; `align` says how text sits relative to it. */
  readonly x: number;
  /** Text baseline. */
  readonly y: number;
  readonly sizePt: number;
  readonly colorHex: string;
  readonly align: LabelAlign;
  /**
   * §21.2.2.216 `c:title/c:txPr/a:bodyPr@rot` — a value-axis title reads
   * bottom-to-top (`rot="-5400000"`, the default every reader applies). Degrees
   * counter-clockwise about the anchor; `align` then runs along the ROTATED
   * reading direction.
   */
  readonly rotationDeg?: number;
}
/**
 * The fully laid-out chart: rectangles, polylines, wedges and labels (plus
 * optional filled polygons) in a local y-up frame, origin bottom-left. The
 * renderer maps these to draw commands — rects/polylines/wedges/polygons via the
 * vector layer, labels via the text pass.
 */
export interface ChartScene {
  readonly rects: ReadonlyArray<ChartRect>;
  readonly polylines: ReadonlyArray<ChartPolyline>;
  readonly wedges: ReadonlyArray<ChartWedge>;
  readonly labels: ReadonlyArray<ChartLabel>;
  readonly polygons?: ReadonlyArray<ChartPolygon>;
  /** §21.2.2.198 chart-space fill + outline: drawn under everything else. */
  readonly background?: ChartRect;
  /**
   * Major gridlines, drawn UNDER the plotted data. Kept apart from the other
   * polylines because z-order is the whole point: gridlines over the bars strip
   * every one of them with the axis's own ruling, which is not what any
   * spreadsheet draws.
   */
  readonly gridlines?: ReadonlyArray<ChartPolyline>;
}

/**
 * Injected text-width measurer: the rendered advance width (points) of `text` at
 * `sizePt`. Keeps this module free of any font/PDF dependency, so it is
 * unit-testable in isolation.
 */
export type MeasureText = (text: string, sizePt: number) => number;

/** Font size (points) for axis ticks, category/data labels and legend text. */
export const CHART_LABEL_PT = 9;
/** Font size (points) for the chart title. */
export const CHART_TITLE_PT = 13;
const AXIS_COLOR = '595959';
const GRID_COLOR = 'D9D9D9';
const LABEL_COLOR = '595959';
const TITLE_COLOR = '404040';

/** The Office accent cycle (RRGGBB) for series without an explicit colour. */
export const SERIES_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

/**
 * Resolve a series colour: the series' own `colorHex` if set, else cycling
 * through `cycle` (the chart's theme accent cycle) or, failing that,
 * {@link SERIES_COLORS} by index.
 *
 * @param s     The series.
 * @param i     The series index, used to pick from the cycle.
 * @param cycle Optional per-chart colour cycle; falls back to {@link SERIES_COLORS}.
 * @returns An RRGGBB hex string.
 */
export const seriesColor = (s: ChartSeries, i: number, cycle?: ReadonlyArray<string>): string =>
  s.colorHex ??
  (cycle && cycle.length > 0 ? cycle[i % cycle.length]! : SERIES_COLORS[i % SERIES_COLORS.length]!);

// ─── value-axis "nice numbers" (Heckbert) ──────────────────────────────────
function niceNum(range: number, round: boolean): number {
  const exp = Math.floor(Math.log10(range));
  const f = range / 10 ** exp;
  let nf: number;
  if (round) nf = f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10;
  else nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * 10 ** exp;
}

/** A value-axis scale: the rounded `min`/`max` extent and the tick `step`. */
export interface Scale {
  readonly min: number;
  readonly max: number;
  readonly step: number;
}

/**
 * Compute a human-friendly axis {@link Scale} for the data range using
 * Heckbert's "nice numbers" algorithm: rounded endpoints and a 1/2/5·10ⁿ step
 * that yields about `maxTicks` ticks. A degenerate range (`dataMin === dataMax`)
 * is widened by 1 so the axis is non-empty.
 *
 * @param dataMin  The smallest data value to cover.
 * @param dataMax  The largest data value to cover.
 * @param maxTicks Target upper bound on tick count (default 6).
 * @returns The rounded min/max and tick step.
 */
export function niceScale(dataMin: number, dataMax: number, maxTicks = 6): Scale {
  const lo = Math.min(dataMin, dataMax);
  let hi = Math.max(dataMin, dataMax);
  if (lo === hi) {
    hi = lo + 1;
  }
  const range = niceNum(hi - lo, false);
  const step = niceNum(range / Math.max(1, maxTicks - 1), true);
  // Excel leaves the top datum room to breathe: it adds about 5% before
  // rounding up to the major unit, so a max that lands exactly on a step still
  // clears the ceiling. Without it 57362.xlsx's 12-value bar touched the plot
  // frame where both references leave a gap and label the axis to 14.
  const headroom = hi > 0 ? hi * 1.05 : hi;
  return { min: Math.floor(lo / step) * step, max: Math.ceil(headroom / step) * step, step };
}

/**
 * Format an axis tick value, choosing decimal places from the tick `step` so
 * `0.25`-spaced ticks read `0.25` while integer steps drop the fraction.
 *
 * @param v    The tick value.
 * @param step The tick spacing (from {@link niceScale}).
 * @returns The label text.
 */
export function formatTick(v: number, step: number): string {
  if (Number.isInteger(step) && Number.isInteger(v)) return String(v);
  const decimals = Math.max(0, -Math.floor(Math.log10(step)));
  return v.toFixed(decimals).replace(/\.?0+$/, (m) => (m.includes('.') ? '' : m));
}

const ticks = (s: Scale): Array<number> => {
  const out: Array<number> = [];
  // +1e-9 guards floating accumulation at the top tick.
  for (let v = s.min; v <= s.max + 1e-9; v += s.step) out.push(Math.abs(v) < 1e-9 ? 0 : v);
  return out;
};

/**
 * The value axis's {@link Scale}: the ends the author fixed where they fixed
 * them (§21.2.2.157 `c:scaling/c:min|c:max`), the data's own range where they
 * did not. A fixed end is exact — nice-rounding it would move a number the
 * author chose — so only the tick step comes from the rounding pass.
 *
 * @param chart   The chart (for its fixed ends).
 * @param dataMin The smallest value to cover.
 * @param dataMax The largest value to cover.
 * @param hPt     The chart's height — how many ticks fit is a question about
 *                the plot, not about the numbers.
 * @returns The axis min/max and tick step.
 */
function axisScale(chart: Chart, dataMin: number, dataMax: number, hPt: number): Scale {
  const min = chart.valAxisMin ?? dataMin;
  const max = chart.valAxisMax ?? dataMax;
  // How many ticks fit is a question about the PLOT, not about the numbers: a
  // tall axis carries more of them. A fixed six drew 0/100/200/300 down a plot
  // where both references fit 0/50/…/300.
  const rounded = niceScale(min, max, tickBudget(hPt));
  return {
    min: chart.valAxisMin ?? rounded.min,
    max: chart.valAxisMax ?? rounded.max,
    step: rounded.step,
  };
}

// ─── shared cartesian frame (scale, plot area, axes, gridlines, labels) ──────
interface CartesianFrame {
  readonly x0: number;
  readonly y0: number;
  readonly plotW: number;
  readonly plotH: number;
  readonly nCats: number;
  readonly slot: number; // category slot size along the category axis
  readonly horizontal: boolean;
  readonly zeroOffset: number; // value-0 distance from the value-axis min end
  valueOffset: (v: number) => number; // value → distance along the value axis
  /** §21.2.2.9 — the same for the SECONDARY value axis, when the chart has one. */
  readonly valueOffset2?: (v: number) => number;
  // mutable scene chrome the caller appends series geometry to:
  readonly rects: Array<ChartRect>;
  readonly polylines: Array<ChartPolyline>;
  /** Gridlines, kept apart so they draw UNDER the plotted data. */
  readonly gridlines: Array<ChartPolyline>;
  readonly labels: Array<ChartLabel>;
}

// Build the plot frame and emit its chrome (title, gridlines, tick + category
// labels, axis lines, legend). The value axis always spans 0 so bar/line
// baselines are meaningful. `horizontal` puts the value axis along x (bar
// charts); line/column keep it along y.
interface FrameOpts {
  readonly dataRange?: readonly [number, number]; // override value-axis extent (stacked totals)
  readonly formatValue?: (v: number) => string; // override tick label text (percent axis)
}

// ── shared axis chrome (C10) ─────────────────────────────────────────────────
// These helpers deduplicate the cartesian chrome between buildFrame and
// buildScatterScene. They PUSH into the caller's buffers — the call order in
// each builder is the z-order of the emitted PDF operators, so callers invoke
// them at exactly the points the inlined code used to occupy.

function pushChartTitle(labels: Array<ChartLabel>, chart: Chart, wPt: number, hPt: number): void {
  if (!chart.title) return;
  labels.push({
    text: chart.title,
    x: wPt / 2,
    y: hPt - 4 - CHART_TITLE_PT,
    sizePt: CHART_TITLE_PT,
    colorHex: TITLE_COLOR,
    align: 'center',
  });
}

function buildLegendBlock(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ReturnType<typeof layoutLegend> {
  // A series with no <c:tx> still needs a legend key when the chart declares a
  // legend — otherwise a chart whose series are all unnamed shows none at all,
  // and nothing on the page says which colour is which. Excel labels them
  // Series1, Series2… and so do we. chart_hyperlink.xlsx is two unnamed series
  // under a <c:legend legendPos="b">, and we drew no legend for it.
  const legendEntries: Array<LegendEntry> = chart.series.map((s, i) => ({
    name: legendSeriesName(s, i),
    colorHex: seriesColor(s, i, chart.seriesColorCycle),
    // The key stands for what the series LOOKS like: a series whose own line
    // draws nothing is a scatter of markers, so its key is a swatch and not a
    // rule (SimpleScatterChart.xlsx).
    ...(s.line?.none ? { marker: 'box' as const } : isLineLike(s) ? { marker: 'line' as const } : {}),
  }));
  // A line chart's key is a LINE, not a filled box — that is what the series
  // looks like on the plot, and what both references draw.
  const marker: LegendMarker = chart.type === 'line' || chart.type === 'scatter' ? 'line' : 'box';
  return layoutLegend(
    legendEntries,
    chart.hasLegend,
    chart.legendPos ?? 'b',
    wPt,
    hPt,
    measure,
    marker,
  );
}

// Gridlines + tick labels along one axis. 'y': horizontal lines with
// right-aligned labels in the left gutter; 'x': vertical lines with centered
// labels under the plot.
function pushGridTicks(
  gridlines: Array<ChartPolyline>,
  labels: Array<ChartLabel>,
  tickVals: ReadonlyArray<number>,
  fmt: (v: number) => string,
  axis: 'x' | 'y',
  at: (v: number) => number,
  x0: number,
  y0: number,
  plotW: number,
  plotH: number,
  grid: ChartLineStyle | undefined,
): void {
  for (const v of tickVals) {
    if (axis === 'x') {
      const gx = at(v);
      gridlines.push({
        points: [
          [gx, y0],
          [gx, y0 + plotH],
        ],
        strokeHex: grid?.colorHex ?? GRID_COLOR,
        widthPt: 0.75,
      });
      labels.push({
        text: fmt(v),
        x: gx,
        y: y0 - CHART_LABEL_PT,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'center',
      });
    } else {
      const gy = at(v);
      gridlines.push({
        points: [
          [x0, gy],
          [x0 + plotW, gy],
        ],
        strokeHex: grid?.colorHex ?? GRID_COLOR,
        widthPt: 0.75,
      });
      labels.push({
        text: fmt(v),
        x: x0 - 3,
        y: gy - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'right',
      });
    }
  }
}

/**
 * §21.2.2.196 — an axis draws the rule its `c:spPr/a:ln` asks for. `<a:noFill/>`
 * draws none at all, and both references honour it: 57362.xlsx gives its value
 * axis a 0.75pt #D9D9D9 hairline where we drew a 1pt #595959 one, and hides its
 * secondary axis's line entirely while keeping its labels.
 */
function axisStroke(style: ChartLineStyle | undefined): { hex: string; widthPt: number } | null {
  if (style?.none) return null;
  return { hex: style?.colorHex ?? AXIS_COLOR, widthPt: style?.widthPt ?? 1 };
}

function pushAxisLines(
  polylines: Array<ChartPolyline>,
  x0: number,
  y0: number,
  plotW: number,
  plotH: number,
  chart: Chart,
): void {
  const val = axisStroke(chart.valAxisLine);
  if (val) {
    polylines.push({
      points: [
        [x0, y0],
        [x0, y0 + plotH],
      ],
      strokeHex: val.hex,
      widthPt: val.widthPt,
    });
  }
  const cat = axisStroke(chart.catAxisLine);
  if (cat) {
    polylines.push({
      points: [
        [x0, y0],
        [x0 + plotW, y0],
      ],
      strokeHex: cat.hex,
      widthPt: cat.widthPt,
    });
  }
}

function buildFrame(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
  horizontal: boolean,
  opts: FrameOpts = {},
): CartesianFrame {
  const rects: Array<ChartRect> = [];
  const polylines: Array<ChartPolyline> = [];
  const gridlines: Array<ChartPolyline> = [];
  const labels: Array<ChartLabel> = [];

  const nCats = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1);
  // §21.2.2.9 — a series on the secondary axis is measured against ITS axis, so
  // it is no part of the primary's range and the primary is no part of its.
  const onSecondary = chart.series.filter((s) => s.secondaryAxis);
  const primary = onSecondary.length > 0 ? chart.series.filter((s) => !s.secondaryAxis) : chart.series;
  const allVals = primary.flatMap((s) => s.values.slice(0, nCats));
  const [dataMin, dataMax] = opts.dataRange ?? [Math.min(0, ...allVals), Math.max(0, ...allVals)];
  const scale = axisScale(chart, dataMin, dataMax, hPt);
  const fmtVal = opts.formatValue ?? ((v: number): string => formatTick(v, scale.step));
  const tickVals = ticks(scale);
  const vals2 = onSecondary.flatMap((s) => s.values.slice(0, nCats));
  // The author's own min/max pin the PRIMARY axis (§21.2.2.157 reads one axis);
  // the secondary takes the nice range of its own data.
  const scale2 =
    vals2.length > 0
      ? niceScale(
          Math.min(0, ...vals2),
          Math.max(0, ...vals2),
          Math.min(10, Math.max(4, Math.round(hPt / 24))),
        )
      : undefined;
  const tickVals2 = scale2 ? ticks(scale2) : [];

  let top = 4;
  if (chart.title) top += CHART_TITLE_PT * 1.6;
  const legend = buildLegendBlock(chart, wPt, hPt, measure);
  const tick2W =
    scale2 && !horizontal
      ? Math.max(0, ...tickVals2.map((v) => measure(formatTick(v, scale2.step), CHART_LABEL_PT))) +
        4 +
        (chart.secondaryValAxisTitle ? CHART_LABEL_PT * 1.5 : 0)
      : 0;
  const plotRight = wPt - 4 - legend.rightWidth - tick2W;

  const tickLabelW = Math.max(0, ...tickVals.map((v) => measure(fmtVal(v), CHART_LABEL_PT))) + 4;
  const catLabelH = CHART_LABEL_PT * 1.6;
  const catTitleH = chart.catAxisTitle ? CHART_LABEL_PT * 1.5 : 0;

  const valTitleW = chart.valAxisTitle ? CHART_LABEL_PT * 1.5 : 0;
  const x0 = 4 + valTitleW + tickLabelW;
  const y0 = 4 + legend.bottomHeight + catTitleH + catLabelH;
  const plotW = Math.max(1, plotRight - x0);
  const plotH = Math.max(1, hPt - top - y0);

  const valueOffset = (v: number): number =>
    ((v - scale.min) / (scale.max - scale.min)) * (horizontal ? plotW : plotH);
  const zeroOffset = valueOffset(0);

  pushChartTitle(labels, chart, wPt, hPt);
  // Axis titles. A value-axis title reads bottom-to-top, in the gutter outside
  // its own tick labels; the category-axis title sits centred below the
  // category labels.
  if (chart.valAxisTitle) {
    labels.push({
      text: chart.valAxisTitle,
      x: 4 + CHART_LABEL_PT * 0.9,
      y: y0 + plotH / 2,
      sizePt: CHART_LABEL_PT,
      colorHex: LABEL_COLOR,
      align: 'center',
      rotationDeg: 90,
    });
  }
  if (chart.catAxisTitle) {
    labels.push({
      text: chart.catAxisTitle,
      x: x0 + plotW / 2,
      y: legend.bottomHeight + 2,
      sizePt: CHART_LABEL_PT,
      colorHex: LABEL_COLOR,
      align: 'center',
    });
  }

  if (horizontal) {
    pushGridTicks(
      gridlines,
      labels,
      tickVals,
      fmtVal,
      'x',
      (v) => x0 + valueOffset(v),
      x0,
      y0,
      plotW,
      plotH,
      chart.gridLine,
    );
  } else {
    pushGridTicks(
      gridlines,
      labels,
      tickVals,
      fmtVal,
      'y',
      (v) => y0 + valueOffset(v),
      x0,
      y0,
      plotW,
      plotH,
      chart.gridLine,
    );
  }

  const valueOffset2 =
    scale2 && !horizontal
      ? (v: number): number => ((v - scale2.min) / (scale2.max - scale2.min)) * plotH
      : undefined;
  if (scale2 && valueOffset2) {
    for (const v of tickVals2) {
      labels.push({
        text: formatTick(v, scale2.step),
        x: x0 + plotW + 3,
        y: y0 + valueOffset2(v) - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'left',
      });
    }
    const sec = axisStroke(chart.secondaryValAxisLine);
    if (sec) {
      polylines.push({
        points: [
          [x0 + plotW, y0],
          [x0 + plotW, y0 + plotH],
        ],
        strokeHex: sec.hex,
        widthPt: sec.widthPt,
      });
    }
    if (chart.secondaryValAxisTitle) {
      labels.push({
        text: chart.secondaryValAxisTitle,
        x: wPt - 4,
        y: y0 + plotH / 2,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'center',
        rotationDeg: 90,
      });
    }
  }

  const slot = (horizontal ? plotH : plotW) / nCats;
  // Label every Nth category, where N is what it takes for them not to collide.
  // Excel and Calc both thin a crowded axis; drawing all of them turned
  // 47813.xlsx's 1700 points into a solid black bar under the plot. The step is
  // measured, not guessed: the widest label plus a gap, over the slot.
  const need = horizontal
    ? CHART_LABEL_PT * 1.4
    : Math.max(0, ...chart.categories.map((t) => measure(t, CHART_LABEL_PT))) + 4;
  const step = Math.max(1, Math.ceil(need / Math.max(slot, 0.01)));
  for (let c = 0; c < nCats; c += step) {
    // No <c:cat> means the categories are the point indices, which is what
    // Excel and Calc both label the axis with — an unlabelled category axis
    // leaves the bars standing on nothing.
    const cat = chart.categories[c] ?? String(c + 1);
    if (!cat) continue;
    const center = (horizontal ? y0 : x0) + c * slot + slot / 2;
    if (horizontal) {
      labels.push({
        text: cat,
        x: x0 - 3,
        y: center - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'right',
      });
    } else {
      labels.push({
        text: cat,
        x: center,
        y: y0 - CHART_LABEL_PT,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'center',
      });
    }
  }

  pushAxisLines(polylines, x0, y0, plotW, plotH, chart);
  legend.emit(rects, labels);

  return {
    x0,
    y0,
    plotW,
    plotH,
    nCats,
    slot,
    horizontal,
    zeroOffset,
    valueOffset,
    ...(valueOffset2 ? { valueOffset2 } : {}),
    rects,
    polylines,
    gridlines,
    labels,
  };
}

const pctLabel = (v: number): string => `${Math.round(v * 100)}%`;

// A datum's printed value (c:dLbls/showVal): integers as-is, else ≤2 decimals.
// A data label carries the axis's number format too — the figure on the bar and
// the figure on the axis are the same quantity, and Excel prints both as
// currency. Without it a budget chart labelled its bars 3750 beside a $3,750
// axis.
const fmtDataLabel = (chart: Chart, v: number): string =>
  chartValueFormatter(chart)?.(v) ??
  (Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100));

function catCount(chart: Chart): number {
  return Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1);
}

// Per-category stacked extents: max of summed positives, min of summed negatives.
function stackedTotals(chart: Chart, nCats: number): { min: number; max: number } {
  let max = 0;
  let min = 0;
  for (let c = 0; c < nCats; c++) {
    let pos = 0;
    let neg = 0;
    for (const s of chart.series) {
      const v = s.values[c] ?? 0;
      if (v >= 0) pos += v;
      else neg += v;
    }
    max = Math.max(max, pos);
    min = Math.min(min, neg);
  }
  return { min, max };
}

// The value-axis range + tick formatter for a grouping. percentStacked pins
// 0..100%; plain stacked spans the summed totals; clustered/standard lets the
// frame derive it from individual values.
function groupingFrameOpts(chart: Chart, nCats: number): FrameOpts {
  const g = chart.grouping ?? 'clustered';
  if (g === 'percentStacked') return { dataRange: [0, 1], formatValue: pctLabel };
  const format = chartValueFormatter(chart);
  if (g === 'stacked') {
    const t = stackedTotals(chart, nCats);
    return {
      dataRange: [Math.min(0, t.min), Math.max(0, t.max)],
      ...(format ? { formatValue: format } : {}),
    };
  }
  return format ? { formatValue: format } : {};
}

// The axis's own number format (§21.2.2.121), as a tick/label formatter. The
// code grammar is the cells' (§18.8.31), so `"$"#,##0` prints $1,000 on the
// axis exactly as it does in the cell the value came from.
const CHART_FORMAT_ID = 1_000_000;
const formatterCache = new WeakMap<Chart, ((v: number) => string) | null>();

function chartValueFormatter(chart: Chart): ((v: number) => string) | undefined {
  const hit = formatterCache.get(chart);
  if (hit !== undefined) return hit ?? undefined;
  const code = chart.numberFormat;
  const made =
    code === undefined
      ? null
      : (
          (formats) => (v: number) =>
            applyNumberFormat(String(v), CHART_FORMAT_ID, formats)
        )(new Map([[CHART_FORMAT_ID, code]]));
  formatterCache.set(chart, made);
  return made ?? undefined;
}

// ─── bar / column chart (clustered, stacked, percentStacked) ────────────────
/**
 * Lay out a bar/column {@link Chart} into a {@link ChartScene}. Honours
 * `chart.grouping` (clustered / stacked / percentStacked) and `chart.barDir`
 * (column vs horizontal bar), over the shared cartesian frame.
 *
 * @param chart   The bar/column chart.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and reserve axis gutters.
 * @returns The positioned scene primitives.
 */
export function buildBarScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  const g = chart.grouping ?? 'clustered';
  const stacked = g === 'stacked' || g === 'percentStacked';
  const percent = g === 'percentStacked';
  const horizontal = chart.barDir === 'bar';
  const nCats = catCount(chart);
  const f = buildFrame(chart, wPt, hPt, measure, horizontal, groupingFrameOpts(chart, nCats));
  // §21.2.2.145 — a combo's other groups plot on this same frame. Their series
  // are NOT bars: they must not take a slot in the cluster (57362.xlsx's two
  // series would each get half a slot and the bars come out at half width), and
  // they draw as their own type over the top.
  const barIdx: Array<number> = [];
  const lineIdx: Array<number> = [];
  chart.series.forEach((s, i) => (isLineLike(s) ? lineIdx : barIdx).push(i));

  if (stacked) {
    const groupPad = f.slot * 0.15;
    const barW = f.slot - 2 * groupPad;
    for (let c = 0; c < f.nCats; c++) {
      const along = (horizontal ? f.y0 : f.x0) + c * f.slot + groupPad;
      const denom = percent
        ? barIdx.reduce((a, i) => a + Math.abs(chart.series[i]!.values[c] ?? 0), 0) || 1
        : 1;
      let cumPos = 0;
      let cumNeg = 0;
      for (const s of barIdx) {
        const series = chart.series[s]!;
        const v = (series.values[c] ?? 0) / denom; // fraction when percent, else raw
        const base = v >= 0 ? cumPos : cumNeg;
        const top = base + v;
        const o0 = f.valueOffset(base);
        const o1 = f.valueOffset(top);
        const lo = Math.min(o0, o1);
        const span = Math.abs(o1 - o0);
        const color = pointColor(series, c) ?? seriesColor(series, s, chart.seriesColorCycle);
        if (horizontal) f.rects.push({ x: f.x0 + lo, y: along, w: span, h: barW, fillHex: color });
        else f.rects.push({ x: along, y: f.y0 + lo, w: barW, h: span, fillHex: color });
        if (chart.showValues && span > CHART_LABEL_PT) {
          const raw = series.values[c] ?? 0;
          if (horizontal)
            f.labels.push(
              centeredLabel(fmtDataLabel(chart, raw), f.x0 + lo + span / 2, along + barW * 0.3),
            );
          else
            f.labels.push(
              centeredLabel(fmtDataLabel(chart, raw), along + barW / 2, f.y0 + lo + span / 2 - 3),
            );
        }
        if (v >= 0) cumPos = top;
        else cumNeg = top;
      }
    }
    pushComboLines(chart, f, lineIdx);
    return {
      rects: f.rects,
      polylines: f.polylines,
      gridlines: f.gridlines,
      wedges: [],
      labels: f.labels,
    };
  }

  // clustered: series side by side within each category slot
  const nSer = Math.max(1, barIdx.length);
  // §21.2.2.75: the slot holds `nSer` bars plus a gap of `gapWidth` percent of
  // one bar. Guessing a flat 15% padding gave every bar 0.63 of its slot —
  // 57362.xlsx asks for 219 and got bars 2.6× the reference's.
  // …and the bar then fills that width. A further 10 % shaved off it was a
  // leftover from before the gap was read: dataValidationTableRange.xlsx asks
  // for `gapWidth="0"`, where both references draw bars that touch, and ours
  // still showed a white line between every pair.
  const gap = (chart.gapPercent ?? 150) / 100;
  const barW = f.slot / (nSer + gap);
  const groupPad = (f.slot - barW * nSer) / 2;
  for (let c = 0; c < f.nCats; c++) {
    const slotStart = (horizontal ? f.y0 : f.x0) + c * f.slot + groupPad;
    for (let b = 0; b < barIdx.length; b++) {
      const s = barIdx[b]!;
      const series = chart.series[s]!;
      const len = f.valueOffset(series.values[c] ?? 0) - f.zeroOffset; // signed from zero line
      const color = pointColor(series, c) ?? seriesColor(series, s, chart.seriesColorCycle);
      const along = slotStart + b * barW;
      if (horizontal) {
        const bx = f.x0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: bx, y: along, w: Math.abs(len), h: barW, fillHex: color });
      } else {
        const by = f.y0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: along, y: by, w: barW, h: Math.abs(len), fillHex: color });
      }
      if (chart.showValues) {
        const raw = series.values[c] ?? 0;
        const txt = fmtDataLabel(chart, raw);
        if (horizontal) {
          const end = f.x0 + f.zeroOffset + len;
          f.labels.push({
            text: txt,
            x: end + (len >= 0 ? 3 : -3),
            y: along + barW * 0.25,
            sizePt: CHART_LABEL_PT,
            colorHex: LABEL_COLOR,
            align: len >= 0 ? 'left' : 'right',
          });
        } else {
          const end = f.y0 + f.zeroOffset + len;
          f.labels.push(
            centeredLabel(txt, along + barW * 0.45, len >= 0 ? end + 2 : end - CHART_LABEL_PT),
          );
        }
      }
    }
  }
  pushComboLines(chart, f, lineIdx);
  return {
    rects: f.rects,
    polylines: f.polylines,
    gridlines: f.gridlines,
    wedges: [],
    labels: f.labels,
  };
}

/**
 * The name a series shows in the legend: its own `c:tx`, or Excel's positional
 * `SeriesN` when it has none. Exported because the SUBSET has to know it too —
 * a name invented at draw time is a name no glyph collector ever saw, and
 * 57362.xlsx drew its unnamed series as "eries1", the capital S appearing
 * nowhere else on the page and so nowhere in the font.
 *
 * @param series The series.
 * @param index  Its zero-based index in the chart.
 * @returns The legend text.
 */
export function legendSeriesName(series: ChartSeries, index: number): string {
  return series.name && series.name.length > 0 ? series.name : `Series${index + 1}`;
}

/** Whether a series plots as a line rather than as a bar of its own cluster. */
function isLineLike(series: ChartSeries): boolean {
  return series.type === 'line' || series.type === 'scatter';
}

/**
 * Draw a combo's line-group series over an already-built cartesian frame, at the
 * category slot centres the bars use (§21.2.2.145).
 */
function pushComboLines(
  chart: Chart,
  f: CartesianFrame,
  lineIdx: ReadonlyArray<number>,
): void {
  for (const s of lineIdx) {
    const series = chart.series[s]!;
    const color = seriesColor(series, s, chart.seriesColorCycle);
    const at = (series.secondaryAxis ? f.valueOffset2 : undefined) ?? f.valueOffset;
    const pts: Array<readonly [number, number]> = [];
    for (let c = 0; c < f.nCats; c++) {
      pts.push([f.x0 + c * f.slot + f.slot / 2, f.y0 + at(series.values[c] ?? 0)]);
    }
    if (pts.length >= 2) f.polylines.push({ points: pts, strokeHex: color, widthPt: 1.5 });
    else if (pts.length === 1) {
      const [px, py] = pts[0]!;
      f.rects.push({ x: px - 1.5, y: py - 1.5, w: 3, h: 3, fillHex: color });
    }
  }
}

function centeredLabel(text: string, x: number, y: number): ChartLabel {
  return { text, x, y, sizePt: CHART_LABEL_PT, colorHex: LABEL_COLOR, align: 'center' };
}

// ─── area chart (standard, stacked, percentStacked) ─────────────────────────
function areaBand(
  top: ReadonlyArray<number>,
  base: ReadonlyArray<number>,
  f: CartesianFrame,
  xAt: (c: number) => number,
  fillHex: string,
): ChartPolygon {
  const pts: Array<readonly [number, number]> = [];
  for (let c = 0; c < f.nCats; c++) pts.push([xAt(c), f.y0 + f.valueOffset(top[c] ?? 0)]);
  for (let c = f.nCats - 1; c >= 0; c--) pts.push([xAt(c), f.y0 + f.valueOffset(base[c] ?? 0)]);
  return { points: pts, fillHex, strokeHex: fillHex, widthPt: 1 };
}

/**
 * Lay out an area {@link Chart} into a {@link ChartScene}: each series becomes a
 * filled polygon down to the value baseline (stacked when `chart.grouping` is
 * stacked / percentStacked), over the shared cartesian frame.
 *
 * @param chart   The area chart.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and reserve axis gutters.
 * @returns The positioned scene primitives.
 */
export function buildAreaScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  const g = chart.grouping ?? 'standard';
  const stacked = g === 'stacked' || g === 'percentStacked';
  const percent = g === 'percentStacked';
  const nCats = catCount(chart);
  const f = buildFrame(chart, wPt, hPt, measure, false, groupingFrameOpts(chart, nCats));
  const xAt = (c: number): number => f.x0 + c * f.slot + f.slot / 2;
  const polygons: Array<ChartPolygon> = [];

  if (stacked) {
    const cum = new Array<number>(nCats).fill(0);
    for (let s = 0; s < chart.series.length; s++) {
      const series = chart.series[s]!;
      const base = cum.slice();
      const top = cum.map((b, c) => {
        const denom = percent
          ? chart.series.reduce((a, ss) => a + Math.abs(ss.values[c] ?? 0), 0) || 1
          : 1;
        return b + (series.values[c] ?? 0) / denom;
      });
      polygons.push(areaBand(top, base, f, xAt, seriesColor(series, s, chart.seriesColorCycle)));
      for (let c = 0; c < nCats; c++) cum[c] = top[c]!;
    }
  } else {
    // standard: filled to baseline, back-to-front so series 0 stays on top
    const base = new Array<number>(nCats).fill(0);
    for (let s = chart.series.length - 1; s >= 0; s--) {
      const series = chart.series[s]!;
      const top = Array.from({ length: nCats }, (_, c) => series.values[c] ?? 0);
      polygons.push(areaBand(top, base, f, xAt, seriesColor(series, s, chart.seriesColorCycle)));
    }
  }
  return {
    rects: f.rects,
    polylines: f.polylines,
    gridlines: f.gridlines,
    wedges: [],
    labels: f.labels,
    polygons,
  };
}

// ─── scatter chart (numeric X/Y) ────────────────────────────────────────────
/**
 * Lay out a scatter {@link Chart} into a {@link ChartScene}: numeric X/Y series
 * plotted as marker points over a frame with two value axes (X from each
 * series' `xValues`, Y from its `values`).
 *
 * @param chart   The scatter chart.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and reserve axis gutters.
 * @returns The positioned scene primitives.
 */
export function buildScatterScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  const rects: Array<ChartRect> = [];
  const polylines: Array<ChartPolyline> = [];
  const gridlines: Array<ChartPolyline> = [];
  const labels: Array<ChartLabel> = [];

  const xs: Array<number> = [];
  const ys: Array<number> = [];
  for (const s of chart.series) {
    for (let i = 0; i < s.values.length; i++) {
      xs.push(s.xValues?.[i] ?? i);
      ys.push(s.values[i] ?? 0);
    }
  }
  if (xs.length === 0) return { rects, polylines, gridlines, wedges: [], labels };

  // Neither scatter axis is forced through 0 — but Excel only RAISES the floor
  // off it for data that genuinely sits far from the origin (a 100…110 series),
  // and runs from zero otherwise. SimpleScatterChart.xlsx plots 0.5 and 1.5 and
  // both references start its axis at 0 where we started it at 0.4. How many
  // ticks fit is a question about the plot, exactly as it is for the frame
  // charts: the x labels sit side by side, so they need more room than the y.
  const xScale = niceScale(...autoRange(xs), tickBudget(wPt / 2));
  const yScale = niceScale(...autoRange(ys), tickBudget(hPt));
  const xTicks = ticks(xScale);
  const yTicks = ticks(yScale);

  let top = 4;
  if (chart.title) top += CHART_TITLE_PT * 1.6;
  const legend = buildLegendBlock(chart, wPt, hPt, measure);
  const tickLabelW =
    Math.max(0, ...yTicks.map((v) => measure(formatTick(v, yScale.step), CHART_LABEL_PT))) + 4;
  const x0 = 4 + tickLabelW;
  const y0 = 4 + legend.bottomHeight + CHART_LABEL_PT * 1.6;
  const plotW = Math.max(1, wPt - 4 - legend.rightWidth - x0);
  const plotH = Math.max(1, hPt - top - y0);
  const xAt = (v: number): number => x0 + ((v - xScale.min) / (xScale.max - xScale.min)) * plotW;
  const yAt = (v: number): number => y0 + ((v - yScale.min) / (yScale.max - yScale.min)) * plotH;

  pushChartTitle(labels, chart, wPt, hPt);
  pushGridTicks(
    gridlines,
    labels,
    yTicks,
    (v) => formatTick(v, yScale.step),
    'y',
    yAt,
    x0,
    y0,
    plotW,
    plotH,
    chart.gridLine,
  );
  pushGridTicks(
    gridlines,
    labels,
    xTicks,
    (v) => formatTick(v, xScale.step),
    'x',
    xAt,
    x0,
    y0,
    plotW,
    plotH,
    chart.gridLine,
  );
  pushAxisLines(polylines, x0, y0, plotW, plotH, chart);

  // §21.2.2.161 — the style says whether the points are joined, marked, or
  // both. The schema's default is `marker`; a chart Excel writes as
  // "scatter with straight lines and markers" says `lineMarker`, and drawing
  // only its points left chartTitle_withTitleFormula.xlsx as four dots where
  // both references draw the line through them. A smooth style is drawn
  // straight — the curve is a fit we do not compute.
  const style = chart.scatterStyle ?? 'marker';
  const joins = style === 'line' || style === 'lineMarker' || style.startsWith('smooth');
  const marks = style === 'marker' || style === 'lineMarker' || style === 'smoothMarker';
  const wedges: Array<ChartWedge> = [];
  for (let s = 0; s < chart.series.length; s++) {
    const series = chart.series[s]!;
    const color = seriesColor(series, s, chart.seriesColorCycle);
    const pts: Array<readonly [number, number]> = [];
    for (let i = 0; i < series.values.length; i++) {
      const px = xAt(series.xValues?.[i] ?? i);
      const py = yAt(series.values[i] ?? 0);
      pts.push([px, py]);
      if (marks) pushMarker(rects, wedges, series.marker, px, py, color);
    }
    // …unless the series itself says its line draws nothing. That is how Excel
    // writes "scatter with markers only" — the group keeps `lineMarker` and the
    // series' own `a:ln` is `<a:noFill/>`.
    if (joins && !series.line?.none && pts.length >= 2) {
      polylines.push({ points: pts, strokeHex: color, widthPt: series.line?.widthPt ?? 1.5 });
    }
  }
  legend.emit(rects, labels);
  return { rects, polylines, gridlines, wedges, labels };
}

/**
 * The data range an automatic axis covers: its own extent, except that data
 * starting within 5/6 of the top reads as data that belongs against a zero
 * baseline — Excel's own rule for when an automatic minimum stays at 0.
 */
function autoRange(vals: ReadonlyArray<number>): [number, number] {
  const lo = Math.min(...vals);
  const hi = Math.max(...vals);
  return [lo > 0 && lo < (5 / 6) * hi ? 0 : lo, hi];
}

/** Ticks that fit along an axis `extentPt` long (mirrors {@link axisScale}). */
const tickBudget = (extentPt: number): number =>
  Math.min(10, Math.max(4, Math.round(extentPt / 24)));

/** Side (points) of the square stamped for a series that names no symbol. */
const DEFAULT_MARKER_PT = 4;

/**
 * Stamp one data point's marker. §21.2.2.107: `none` draws nothing at all even
 * when the scatter style marks its points, a round symbol is a disc (a full
 * wedge — drawn last, so it sits over the line joining the points), and every
 * other symbol keeps the square we have always drawn.
 *
 * @param rects  Collects a square marker.
 * @param wedges Collects a round marker.
 * @param marker The series' `c:marker`, if it declared one.
 * @param x      Point centre, scene x.
 * @param y      Point centre, scene y.
 * @param color  The series colour.
 */
function pushMarker(
  rects: Array<ChartRect>,
  wedges: Array<ChartWedge>,
  marker: ChartMarker | undefined,
  x: number,
  y: number,
  color: string,
): void {
  if (marker?.symbol === 'none') return;
  const size = marker?.sizePt ?? DEFAULT_MARKER_PT;
  if (marker?.symbol === 'circle' || marker?.symbol === 'dot') {
    wedges.push({ cx: x, cy: y, r: size / 2, startRad: 0, sweepRad: -2 * Math.PI, fillHex: color });
    return;
  }
  rects.push({ x: x - size / 2, y: y - size / 2, w: size, h: size, fillHex: color });
}

// ─── line chart ───────────────────────────────────────────────────────────────
/**
 * Lay out a line {@link Chart} into a {@link ChartScene}: each series becomes a
 * stroked polyline across the category slots, over the shared cartesian frame.
 * Unlike bars/areas the value axis auto-mins (it need not include 0).
 *
 * @param chart   The line chart.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and reserve axis gutters.
 * @returns The positioned scene primitives.
 */
export function buildLineScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  // Line charts auto-min: the value axis need not include 0 when the data sits
  // far from it (unlike bars/areas, which need a meaningful baseline at 0).
  const allVals = chart.series.flatMap((s) => s.values);
  const range: readonly [number, number] =
    allVals.length > 0 ? [Math.min(...allVals), Math.max(...allVals)] : [0, 1];
  // …and the axis's own number format applies here exactly as it does to a bar
  // chart's: 123233_charts.xlsx labels every one of its four charts in currency
  // and only the line chart came out in bare digits.
  const lineFormat = chartValueFormatter(chart);
  const f = buildFrame(chart, wPt, hPt, measure, false, {
    dataRange: range,
    ...(lineFormat ? { formatValue: lineFormat } : {}),
  });
  // §21.2.2.106 — the group's `c:marker` switch says whether its series stamp
  // their points. WithChart.xlsx turns it on and we drew two bare lines.
  const wedges: Array<ChartWedge> = [];
  for (let s = 0; s < chart.series.length; s++) {
    const series = chart.series[s]!;
    const color = seriesColor(series, s, chart.seriesColorCycle);
    const pts: Array<readonly [number, number]> = [];
    for (let c = 0; c < f.nCats; c++) {
      const x = f.x0 + c * f.slot + f.slot / 2;
      const y = f.y0 + f.valueOffset(series.values[c] ?? 0);
      pts.push([x, y]);
      if (chart.lineMarkers) pushMarker(f.rects, wedges, series.marker, x, y, color);
      if (chart.showValues)
        f.labels.push(centeredLabel(fmtDataLabel(chart, series.values[c] ?? 0), x, y + 3));
    }
    if (pts.length >= 2) {
      f.polylines.push({ points: pts, strokeHex: color, widthPt: 1.5 });
    } else if (pts.length === 1) {
      // A single data point: a small marker so it is visible.
      const [px, py] = pts[0]!;
      f.rects.push({ x: px - 1.5, y: py - 1.5, w: 3, h: 3, fillHex: color });
    }
  }
  return {
    rects: f.rects,
    polylines: f.polylines,
    gridlines: f.gridlines,
    wedges,
    labels: f.labels,
  };
}

// ─── pie chart ──────────────────────────────────────────────────────────────
// A pie's slices cycle the same colours a bar chart's SERIES do — the chart's
// own cycle when it has one, which is the workbook theme's accents.
const sliceColor = (
  series: ChartSeries,
  i: number,
  cycle?: ReadonlyArray<string>,
): string => {
  const palette = cycle && cycle.length > 0 ? cycle : SERIES_COLORS;
  return pointColor(series, i) ?? palette[i % palette.length]!;
};

/**
 * Lay out a pie/doughnut {@link Chart} into a {@link ChartScene}: the first
 * series' values become proportional wedges (a centre hole for doughnut),
 * with a legend instead of axes.
 *
 * @param chart   The pie/doughnut chart.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and the legend.
 * @returns The positioned scene primitives.
 */
export function buildPieScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  const rects: Array<ChartRect> = [];
  const wedges: Array<ChartWedge> = [];
  const labels: Array<ChartLabel> = [];

  const series = chart.series[0];
  const values = series ? series.values.map((v) => Math.max(0, v)) : [];
  const total = values.reduce((a, b) => a + b, 0);
  if (!series || total <= 0) return { rects, polylines: [], wedges, labels };

  let top = 4;
  if (chart.title) top += CHART_TITLE_PT * 1.6;
  // Pie legend lists categories (each in its slice colour). A pie written
  // without `<c:cat>` has none, and its legend came out empty — the same case
  // the category axis already answers with the point indices, which is what
  // both references list: WithThreeCharts.xlsx's pie legend reads 1 to 6.
  const legendEntries: Array<LegendEntry> = values.map((_, i) => ({
    name: chart.categories[i] ?? String(i + 1),
    colorHex: sliceColor(series, i, chart.seriesColorCycle),
  }));
  const legend = layoutLegend(
    legendEntries,
    chart.hasLegend,
    chart.legendPos ?? 'r',
    wPt,
    hPt,
    measure,
  );

  const availW = Math.max(1, wPt - 8 - legend.rightWidth);
  const availH = Math.max(1, hPt - top - 4 - legend.bottomHeight);
  const r = Math.max(1, (Math.min(availW, availH) / 2) * 0.95);
  const cx = 4 + availW / 2;
  const cy = 4 + legend.bottomHeight + availH / 2;

  // A doughnut is a pie with a central hole; place its labels out on the ring.
  const holeR = chart.doughnut ? r * 0.5 : 0;
  const labelR = chart.doughnut ? (holeR + r) / 2 : r * 0.6;

  // Excel pies start at 12 o'clock and sweep clockwise (negative in y-up).
  let ang = Math.PI / 2;
  for (let i = 0; i < values.length; i++) {
    const v = values[i]!;
    if (v <= 0) continue;
    const sweep = -(v / total) * 2 * Math.PI;
    wedges.push({
      cx,
      cy,
      r,
      startRad: ang,
      sweepRad: sweep,
      fillHex: sliceColor(series, i, chart.seriesColorCycle),
      strokeHex: 'FFFFFF',
    });
    const mid = ang + sweep / 2;
    const pct = Math.round((v / total) * 100);
    // A label the author typed wins over the one we would compute — see
    // ChartSeries.pointLabels. It is drawn whatever the slice's size, because
    // the author put it there on purpose.
    const custom = series.pointLabels?.find((l) => l.idx === i)?.text;
    if (custom !== undefined || pct >= 5) {
      labels.push({
        text: custom ?? `${pct}%`,
        x: cx + Math.cos(mid) * labelR,
        y: cy + Math.sin(mid) * labelR - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: custom !== undefined || chart.doughnut ? LABEL_COLOR : 'FFFFFF',
        align: 'center',
      });
    }
    ang += sweep;
  }
  // Punch the hole: a white disc over the wedge centres (drawn after slices).
  if (holeR > 0) {
    wedges.push({ cx, cy, r: holeR, startRad: 0, sweepRad: -2 * Math.PI, fillHex: 'FFFFFF' });
  }

  if (chart.title) {
    labels.push({
      text: chart.title,
      x: wPt / 2,
      y: hPt - 4 - CHART_TITLE_PT,
      sizePt: CHART_TITLE_PT,
      colorHex: TITLE_COLOR,
      align: 'center',
    });
  }
  legend.emit(rects, labels);
  return { rects, polylines: [], wedges, labels };
}

function pointColor(series: ChartSeries, idx: number): string | undefined {
  return series.pointColors?.find((p) => p.idx === idx)?.colorHex;
}

// ─── legend ─────────────────────────────────────────────────────────────────
interface LegendEntry {
  readonly name: string;
  readonly colorHex: string;
  /** Per-entry key shape — a combo's line series keeps its line (§21.2.2.145). */
  readonly marker?: LegendMarker;
}
type LegendMarker = 'box' | 'line';

interface LegendLayout {
  readonly rightWidth: number;
  readonly bottomHeight: number;
  emit: (rects: Array<ChartRect>, labels: Array<ChartLabel>) => void;
}

// Generic legend over (name, colour) entries — series for bar/line, categories
// (slices) for pie. Reserves a right column or a bottom row.
function layoutLegend(
  entries: ReadonlyArray<LegendEntry>,
  hasLegend: boolean,
  pos: 'r' | 'l' | 't' | 'b',
  wPt: number,
  hPt: number,
  measure: MeasureText,
  marker: LegendMarker = 'box',
): LegendLayout {
  if (!hasLegend || entries.length === 0) {
    return { rightWidth: 0, bottomHeight: 0, emit: () => {} };
  }
  const sw = CHART_LABEL_PT; // key height reference
  const gap = 4;
  // Neither reader draws the key as a square: Excel and Calc both draw a WIDE,
  // flat swatch about twice the text height across — 57362.xlsx's key is 20 x
  // 4.5pt against a 9pt label. A line key is the same width, drawn as the
  // stroke it stands for.
  const keyW = sw * 2.2;
  const keyH = sw * 0.55;
  const key = (x: number, y: number, e: LegendEntry): ChartRect =>
    (e.marker ?? marker) === 'line'
      ? { x, y: y + sw / 2 - 0.75, w: keyW, h: 1.5, fillHex: e.colorHex }
      : { x, y: y + (sw - keyH) / 2, w: keyW, h: keyH, fillHex: e.colorHex };
  const entryW = (e: LegendEntry): number => keyW + 3 + measure(e.name, CHART_LABEL_PT) + gap * 2;

  if (pos === 'r' || pos === 'l') {
    const colW = Math.max(...entries.map(entryW));
    return {
      rightWidth: pos === 'r' ? colW : 0,
      bottomHeight: 0,
      emit: (rects, labels) => {
        const lx = pos === 'r' ? wPt - colW + gap : gap;
        let ly = hPt / 2 + (entries.length * (sw + 4)) / 2 - sw;
        for (const e of entries) {
          rects.push(key(lx, ly, e));
          labels.push({
            text: e.name,
            x: lx + keyW + 3,
            y: ly + 1,
            sizePt: CHART_LABEL_PT,
            colorHex: LABEL_COLOR,
            align: 'left',
          });
          ly -= sw + 4;
        }
      },
    };
  }
  const totalW = entries.reduce((acc, e) => acc + entryW(e), 0);
  return {
    rightWidth: 0,
    bottomHeight: sw + 6,
    emit: (rects, labels) => {
      let lx = (wPt - totalW) / 2 + gap;
      const ly = 2;
      for (const e of entries) {
        rects.push(key(lx, ly, e));
        labels.push({
          text: e.name,
          x: lx + keyW + 3,
          y: ly + 1,
          sizePt: CHART_LABEL_PT,
          colorHex: LABEL_COLOR,
          align: 'left',
        });
        lx += entryW(e);
      }
    },
  };
}

/**
 * Lay out any supported {@link Chart} into a {@link ChartScene}, dispatching by
 * `chart.type` to the per-type builders.
 *
 * @param chart   The chart to lay out.
 * @param wPt     Frame width in points.
 * @param hPt     Frame height in points.
 * @param measure Text measurer used to size labels and reserve gutters.
 * @returns The positioned scene, or `null` for an unrenderable type (the
 *          renderer then reserves the box with a light border).
 */
export function buildChartScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene | null {
  const plotted = chart.catAxisReversed ? withReversedCategories(chart) : chart;
  const scene = buildTypedScene(plotted, wPt, hPt, measure);
  return scene && withFrame(scene, chart, wPt, hPt);
}

/**
 * §21.2.2.134 `maxMin` on the category axis — the categories run the other way,
 * which for a horizontal bar chart puts the FIRST one at the top (how every
 * ranked list reads). Reversing the plotted order says exactly that and leaves
 * every builder's geometry alone; the per-point overrides move with their
 * points.
 *
 * @param chart The chart as the file declares it.
 * @returns The same chart with its categories, values and per-point overrides
 *          in reverse.
 */
function withReversedCategories(chart: Chart): Chart {
  const n = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1);
  const flip = (i: number): number => n - 1 - i;
  return {
    ...chart,
    // An empty category list is not a list of empty labels: the builders label
    // an unlabelled axis with the point index, and padding it would silence it.
    categories:
      chart.categories.length > 0
        ? Array.from({ length: n }, (_, i) => chart.categories[flip(i)] ?? '')
        : chart.categories,
    series: chart.series.map((s) => ({
      ...s,
      values: Array.from({ length: n }, (_, i) => s.values[flip(i)] ?? 0),
      ...(s.pointColors
        ? { pointColors: s.pointColors.map((p) => ({ ...p, idx: flip(p.idx) })) }
        : {}),
      ...(s.pointLabels
        ? { pointLabels: s.pointLabels.map((p) => ({ ...p, idx: flip(p.idx) })) }
        : {}),
    })),
  };
}

function buildTypedScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene | null {
  if (chart.type === 'bar') return buildBarScene(chart, wPt, hPt, measure);
  if (chart.type === 'line') return buildLineScene(chart, wPt, hPt, measure);
  if (chart.type === 'pie') return buildPieScene(chart, wPt, hPt, measure);
  if (chart.type === 'area') return buildAreaScene(chart, wPt, hPt, measure);
  if (chart.type === 'scatter') return buildScatterScene(chart, wPt, hPt, measure);
  return null;
}

/**
 * §21.2.2.198 — the chart-space frame, first in the scene so everything else
 * draws over it. A chart that declares neither fill nor outline is returned
 * untouched, so a scene without one is unchanged rect for rect.
 *
 * @param scene The typed scene.
 * @param chart The chart (for its frame fill/outline).
 * @param wPt   The chart's width in points.
 * @param hPt   Its height.
 * @returns The scene with the frame behind it.
 */
function withFrame(scene: ChartScene, chart: Chart, wPt: number, hPt: number): ChartScene {
  if (!chart.frameFillHex && !chart.frameLineHex) return scene;
  const frame: ChartRect = {
    x: 0,
    y: 0,
    w: wPt,
    h: hPt,
    ...(chart.frameFillHex ? { fillHex: chart.frameFillHex } : {}),
    ...(chart.frameLineHex ? { strokeHex: chart.frameLineHex, strokeWidthPt: 0.75 } : {}),
  };
  return { ...scene, background: frame };
}
