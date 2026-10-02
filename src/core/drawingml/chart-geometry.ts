// Pure chart geometry: a Chart + box (w×h points) → a ChartScene of rectangles,
// polylines, wedges and labels in a LOCAL y-up frame (origin bottom-left). No
// PDF or font dependency — text widths come through an injected measure fn, so
// this module is unit-testable in isolation. The renderer converts the scene
// to draw commands (rects/polylines/wedges via the vector layer, labels via the
// text pass).

import type {
  Chart,
  ChartLineStyle,
  ChartMarker,
  ChartSeries,
  ChartTextStyles,
  ShapeDash,
} from '@/core/document-model';

import { applyNumberFormat, generalRenderings } from '@/core/number-format';

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
  /** §20.1.10.49 — the outline's preset dash, when it names one. */
  readonly strokeDash?: ShapeDash;
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
export interface ChartLabel extends ChartFont {
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
   * §21.2.2.145 the PLOT rectangle's own fill + outline, drawn over the chart
   * frame and under the gridlines.
   */
  readonly plotBackground?: ChartRect;
  /**
   * Major gridlines, drawn UNDER the plotted data. Kept apart from the other
   * polylines because z-order is the whole point: gridlines over the bars strip
   * every one of them with the axis's own ruling, which is not what any
   * spreadsheet draws.
   */
  readonly gridlines?: ReadonlyArray<ChartPolyline>;
}

/**
 * The face a chart's text is set in beyond its size: the family the chart
 * names (absent — the document's own) and its weight and slant.
 */
export interface ChartFont {
  readonly family?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

/**
 * Injected text-width measurer: the rendered advance width (points) of `text` at
 * `sizePt`, in `font` when the chart names one. Keeps this module free of any
 * font/PDF dependency, so it is unit-testable in isolation.
 */
export type MeasureText = (text: string, sizePt: number, font?: ChartFont) => number;

/** Font size (points) for axis ticks, category/data labels and legend text. */
export const CHART_LABEL_PT = 9;
/** Font size (points) for the chart title. */
export const CHART_TITLE_PT = 13;
const AXIS_COLOR = '595959';
const GRID_COLOR = 'D9D9D9';
const LABEL_COLOR = '595959';
const TITLE_COLOR = '404040';

/** How one kind of a chart's text is set: its size and colour, and its face. */
interface TextFace extends ChartFont {
  readonly sizePt: number;
  readonly colorHex: string;
}

/**
 * The face one role of the chart's text is set in: the chart's own
 * ({@link Chart.text}) where the reader resolved it, and otherwise the sizes
 * and greys this module has always drawn in, in the document's face.
 */
function faceOf(chart: Chart, role: keyof ChartTextStyles): TextFace {
  const own = chart.text?.[role];
  const title = role === 'title';
  return {
    sizePt: own?.sizePt ?? (title ? CHART_TITLE_PT : CHART_LABEL_PT),
    colorHex: own?.colorHex ?? (title ? TITLE_COLOR : LABEL_COLOR),
    ...(own?.family ? { family: own.family } : {}),
    ...(own?.bold ? { bold: true } : {}),
    ...(own?.italic ? { italic: true } : {}),
  };
}

/** The width of `text` set in `face`. */
const widthIn = (measure: MeasureText, text: string, face: TextFace): number =>
  measure(text, face.sizePt, face);

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

/** A value-axis scale: the rounded `min`/`max` extent and the tick `step`. */
export interface Scale {
  readonly min: number;
  readonly max: number;
  readonly step: number;
}

/**
 * The scale Excel gives a value axis it is left to choose, measured against
 * Excel's own PDF of its column charts (2026-10-01 and -02):
 *
 *  - the ends: values all of one sign put the axis's near end at zero —
 *    unless their spread is under a sixth of the largest of them, when it
 *    stops half a spread short of the nearest (93…97 runs from 91) — and the
 *    far end takes 5% of its reach from that near end before it is rounded
 *    out to the step. 70…96 reaches 100.8 and runs 0…120 by 20, where 5% of
 *    the data's own spread stopped at 97.3 and ran it 0…100 by 10. Both ends
 *    of a range across zero take 5% of the spread;
 *  - the step: the smallest 1, 2 or 5 × 10ⁿ that leaves at most ten intervals
 *    (`maxIntervals`), whatever the chart's height — 3 750 runs 0…4 000 by 500
 *    on a 150pt chart as on a 400pt one. Only an axis too short to hold the
 *    labels asks for fewer, and the caller says so.
 *
 * Every case of the probes lands where Excel puts it: −1…1.2 on −1.5…1.5 by
 * 0.5, 0.3…4.7 on 0…5 by 0.5, 120…950 on 0…1 000 by 100, 0.012…0.047 on
 * 0…0.05 by 0.005, 1E-150…3E-150 on 0…3.5E-150. Heckbert's nice numbers with
 * a budget of ticks by height stepped the budget chart by 1 000 and started
 * 93…97 at zero.
 *
 * @param dataMin      The smallest value to cover.
 * @param dataMax      The largest value to cover.
 * @param maxIntervals At most this many steps between the ends (default 10).
 * @returns The rounded min/max and tick step.
 */
export function niceScale(dataMin: number, dataMax: number, maxIntervals = 10): Scale {
  const end = paddedEnds(dataMin, dataMax);
  const step = stepFor(end.lo, end.hi, maxIntervals);
  return { min: below(end.lo, step), max: above(end.hi, step), step };
}

/**
 * The ends {@link niceScale} rounds out to its step: the data's, padded as
 * Excel pads them. Never past the largest double, which the padding of values
 * out at ±1.7E+308 would run them to.
 */
function paddedEnds(dataMin: number, dataMax: number): { lo: number; hi: number } {
  let lo = Math.min(dataMin, dataMax);
  let hi = Math.max(dataMin, dataMax);
  // One value is a spread from zero to it.
  if (lo === hi) {
    if (hi > 0) lo = 0;
    else if (lo < 0) hi = 0;
    else hi = 1;
  }
  const spread = hi - lo;
  if (lo >= 0) {
    const from = spread < hi / 6 ? lo - spread / 2 : 0;
    return { lo: from, hi: finite(hi + (hi - from) * 0.05) };
  }
  if (hi <= 0) {
    const to = spread < -lo / 6 ? hi + spread / 2 : 0;
    return { lo: finite(lo - (to - lo) * 0.05), hi: to };
  }
  return { lo: finite(lo - spread * 0.05), hi: finite(hi + spread * 0.05) };
}

/** `v`, or the largest double of its sign where it ran past it. */
const finite = (v: number): number => Math.min(Number.MAX_VALUE, Math.max(-Number.MAX_VALUE, v));

/** `v` rounded down to a multiple of `step`, short of the largest double. */
const below = (v: number, step: number): number => finite(Math.floor(v / step + 1e-9) * step);

/** `v` rounded up to a multiple of `step`, short of the largest double. */
const above = (v: number, step: number): number => finite(Math.ceil(v / step - 1e-9) * step);

/**
 * The smallest 1, 2 or 5 × 10ⁿ that cuts `lo`…`hi` into at most
 * `maxIntervals` steps. The share of the span each interval takes is reckoned
 * so that a span past the largest double has one: −1.7E+308…1.7E+308 spans
 * Infinity, and the search for its step never ended.
 */
function stepFor(lo: number, hi: number, maxIntervals: number): number {
  const most = Math.max(1, maxIntervals);
  const span = hi - lo;
  const share = Math.abs(Number.isFinite(span) ? span / most : hi / most - lo / most);
  // Nothing to divide: an axis whose ends the author made one.
  if (!(share > 0 && share < Infinity)) return 1;
  // The power of ten at or under the share — and over zero, as 1E-324 is not.
  let base = Math.max(Number.MIN_VALUE, 10 ** Math.floor(Math.log10(share)));
  for (;;) {
    for (const m of [1, 2, 5]) {
      if (share <= m * base * (1 + 1e-9)) return Math.min(m * base, Number.MAX_VALUE);
    }
    base *= 10;
  }
}

/** How many steps lie between a scale's ends — in halves where the span overflows. */
function intervalsOf(s: Scale): number {
  const span = s.max - s.min;
  return Number.isFinite(span) ? span / s.step : (s.max / 2 - s.min / 2) / (s.step / 2);
}

/**
 * How far along `s` the value `v` stands: 0 at its min, 1 at its max —
 * reckoned in halves where a difference overflows, as it does between ends
 * out at ±1.7E+308.
 */
function fractionOf(v: number, s: Scale): number {
  const span = s.max - s.min;
  const from = v - s.min;
  return Number.isFinite(span) && Number.isFinite(from)
    ? from / span
    : (v / 2 - s.min / 2) / (s.max / 2 - s.min / 2);
}

/**
 * The most intervals Excel cuts a value axis into. A finer step than that is
 * widened to a 500th of the axis: its own PDF ran 0…26.25 by 0.0525 for steps
 * of 0.05, 0.01 and 0.001 alike, and kept 0.0525 itself (2026-10-02).
 */
const MOST_INTERVALS = 500;

/** The most characters Excel gives a value-axis label under General, its sign among them. */
const AXIS_LABEL_CHARS = 9;

/**
 * A value-axis tick as Excel labels it under General. An axis is not a cell:
 * its General gives every label nine characters, the sign among them, on an
 * axis 150pt or 500pt long, upright or lying alike — Excel's own PDF of 44
 * axes (2026-10-01 and -02). Within them it is the cells' rule
 * ({@link generalRenderings}): the number as it is where it fits (0.25 — we
 * wrote 0.3 on an axis stepped by 0.25 — 500000000, −50000000, 0.0000003),
 * else the finest of it rounded or in scientific notation (0.1234568,
 * −0.123457, 123456771, 1E+09, 1.235E+09, −1.23E+09, 1.235E-05, 5E-151).
 * Places counted from the step and written by `toFixed` threw past a hundred
 * of them, which a step of 1E-150 asked for.
 *
 * @param v The tick value.
 * @returns The label text.
 */
export function formatTick(v: number): string {
  return generalRenderings(v, AXIS_LABEL_CHARS, true).next().value ?? String(v);
}

/**
 * The values a scale labels: its min and every step up to its max. Counted,
 * not summed: the sum drifted, and its tolerance was absolute, so a scale of
 * 1E-150 ran on until it passed 1E-9 and asked for more ticks than an array
 * holds. Never more than Excel's most intervals and one past them — rounding
 * the ends out to a widened step can add that one.
 */
const ticks = (s: Scale): Array<number> => {
  const count = Math.min(Math.floor(intervalsOf(s) + 1e-9), MOST_INTERVALS + 1);
  const out: Array<number> = [];
  for (let i = 0; i <= count; i++) {
    const sum = s.min + i * s.step;
    // In halves where the product runs past the largest double — on an axis
    // out at ±1.7E+308 the fifth tick read Infinity.
    const v = Number.isFinite(sum) ? sum : 2 * (s.min / 2 + i * (s.step / 2));
    // A tick a rounding error off zero is zero, or General writes 1.776E-15.
    out.push(Math.abs(v) < s.step * 1e-9 ? 0 : v);
  }
  return out;
};

/**
 * The value axis's {@link Scale}: the ends the author fixed where they fixed
 * them (§21.2.2.157 `c:scaling/c:min|c:max`), the data's, padded as
 * {@link niceScale} pads them, where they did not. A fixed end is exact —
 * nice-rounding it would move a number the author chose.
 *
 * The step is the author's (§21.2.2.98 `c:majorUnit`), else the 1-2-5 step
 * of the span between the ends — with both fixed, of the span they fix:
 * −1E+300…1E+300 runs by 2E+299 in Excel, where padding the ends first
 * stepped it by 5E+299. The automatic ends round out to an author's step
 * from where the padding left them, not from where a step of our own would
 * have: 10…25 by 0.1 runs 0…26.3, as in Excel, not 0…30. And no step cuts
 * the axis finer than {@link MOST_INTERVALS}: Excel widens one that would to
 * a 500th of the span and rounds the automatic ends out to the widened step
 * (−10…25 by 0.01 runs −11.781…26.796 by 0.077). A step of 0.00001 asked for
 * three million labels and ran out of stack.
 *
 * @param chart      The chart (for its fixed ends and step).
 * @param dataMin    The smallest value to cover.
 * @param dataMax    The largest value to cover.
 * @param extentPt   The axis's length — how many ticks fit is a question
 *                   about the plot, not about the numbers.
 * @param horizontal Whether the axis lies along the foot.
 * @returns The axis min/max and tick step.
 */
function axisScale(
  chart: Chart,
  dataMin: number,
  dataMax: number,
  extentPt: number,
  horizontal = false,
): Scale {
  const fixedMin = chart.valAxisMin;
  const fixedMax = chart.valAxisMax;
  const padded = paddedEnds(fixedMin ?? dataMin, fixedMax ?? dataMax);
  const lo = fixedMin ?? padded.lo;
  const hi = fixedMax ?? padded.hi;
  const rounded = (step: number): Scale => ({
    min: fixedMin ?? below(lo, step),
    max: fixedMax ?? above(hi, step),
    step,
  });
  const scale = rounded(
    chart.valAxisMajorUnit ?? stepFor(lo, hi, intervalsThatFit(extentPt, horizontal)),
  );
  if (intervalsOf(scale) <= MOST_INTERVALS) return scale;
  // A 500th of the span between the ends as they stood before rounding: a
  // step of 1E-320 rounds them out past the largest double.
  const span = hi - lo;
  return rounded(
    Number.isFinite(span) ? span / MOST_INTERVALS : hi / MOST_INTERVALS - lo / MOST_INTERVALS,
  );
}

/**
 * §21.2.2.95 — the plot's own rectangle where its author sized it (an
 * `inner` manual layout), in the scene's y-up frame: the axes' labels and
 * titles stand outside it, where the chart sets them.
 */
function innerPlot(
  chart: Chart,
  wPt: number,
  hPt: number,
): { x0: number; y0: number; plotW: number; plotH: number } | undefined {
  const box = chart.plotBox;
  if (!box?.inner) return undefined;
  return {
    x0: box.x * wPt,
    y0: hPt - (box.y + box.h) * hPt,
    plotW: Math.max(1, box.w * wPt),
    plotH: Math.max(1, box.h * hPt),
  };
}

/** The ends and step an author fixed on one axis, each absent where not. */
interface AxisFix {
  readonly min: number | undefined;
  readonly max: number | undefined;
  readonly unit: number | undefined;
}

/**
 * A scatter axis's {@link Scale}: {@link niceScale}'s where its author fixed
 * nothing, and where they did, their ends exactly and their step, the ends
 * they left to the data rounded out to it — never cut into more than
 * {@link MOST_INTERVALS}, as a frame chart's value axis is not.
 */
function scatterScale(lo: number, hi: number, maxIntervals: number, fix: AxisFix): Scale {
  if (fix.min === undefined && fix.max === undefined && fix.unit === undefined) {
    return niceScale(lo, hi, maxIntervals);
  }
  const auto = niceScale(fix.min ?? lo, fix.max ?? hi, maxIntervals);
  const step =
    fix.unit ??
    (fix.min !== undefined && fix.max !== undefined
      ? stepFor(fix.min, fix.max, maxIntervals)
      : auto.step);
  const scale: Scale = {
    min: fix.min ?? below(auto.min, step),
    max: fix.max ?? above(auto.max, step),
    step,
  };
  if (intervalsOf(scale) <= MOST_INTERVALS) return scale;
  const span = scale.max - scale.min;
  return {
    ...scale,
    step: Number.isFinite(span)
      ? span / MOST_INTERVALS
      : scale.max / MOST_INTERVALS - scale.min / MOST_INTERVALS,
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
  /** §21.2.2.145 — the plot rectangle's own fill/rule, drawn under both. */
  readonly plotBackground?: ChartRect;
  readonly labels: Array<ChartLabel>;
}

// Build the plot frame and emit its chrome (title, gridlines, tick + category
// labels, axis lines, legend). The value axis always spans 0 so bar/line
// baselines are meaningful. `horizontal` puts the value axis along x (bar
// charts); line/column keep it along y.
interface FrameOpts {
  readonly dataRange?: readonly [number, number]; // override value-axis extent (stacked totals)
  // The range IS the axis — no room added, nothing rounded out (a 100% stack).
  readonly exactRange?: boolean;
  readonly formatValue?: (v: number) => string; // override tick label text (percent axis)
}

// ── shared axis chrome (C10) ─────────────────────────────────────────────────
// These helpers deduplicate the cartesian chrome between buildFrame and
// buildScatterScene. They PUSH into the caller's buffers — the call order in
// each builder is the z-order of the emitted PDF operators, so callers invoke
// them at exactly the points the inlined code used to occupy.

/**
 * The chart title's lines. A title longer than the chart is wide wraps at its
 * spaces, as Excel's and Calc's do, rather than running out past both sides
 * of the frame: dataValidationTableRange.xlsx's "Ranking of Washington
 * Counties on Days per Patient (ALOS) in 2015" crossed the cells beside its
 * chart. A line holds at most four fifths of the chart's width.
 *
 * @param chart   The chart.
 * @param wPt     The chart's width.
 * @param measure The text measurer.
 * @returns The lines, top first; none for a chart without a title.
 */
function titleLines(chart: Chart, wPt: number, measure: MeasureText): Array<string> {
  if (!chart.title) return [];
  const face = faceOf(chart, 'title');
  const room = wPt * 0.8;
  const lines: Array<string> = [];
  let line = '';
  for (const word of chart.title.split(/\s+/).filter((w) => w.length > 0)) {
    const longer = line ? `${line} ${word}` : word;
    if (line && widthIn(measure, longer, face) > room) {
      lines.push(line);
      line = word;
    } else {
      line = longer;
    }
  }
  if (line) lines.push(line);
  return lines;
}

/** The band a title of `lines` takes at the top of the chart. */
const titleHeight = (lines: ReadonlyArray<string>, chart: Chart): number =>
  lines.length === 0 ? 0 : faceOf(chart, 'title').sizePt * (1.6 + 1.2 * (lines.length - 1));

function pushChartTitle(
  labels: Array<ChartLabel>,
  lines: ReadonlyArray<string>,
  wPt: number,
  hPt: number,
  chart: Chart,
): void {
  const face = faceOf(chart, 'title');
  lines.forEach((text, i) => {
    labels.push({
      text,
      x: wPt / 2,
      y: hPt - 4 - face.sizePt * (1 + 1.2 * i),
      ...face,
      align: 'center',
    });
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
    ...(s.line?.none
      ? { marker: 'box' as const }
      : isLineLike(s)
        ? { marker: 'line' as const }
        : {}),
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
    faceOf(chart, 'legend'),
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
  face: TextFace,
  atEnd = false,
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
        y: atEnd ? y0 + plotH + face.sizePt * 0.4 : y0 - face.sizePt,
        ...face,
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
        x: atEnd ? x0 + plotW + 3 : x0 - 3,
        y: gy - face.sizePt / 3,
        ...face,
        align: atEnd ? 'left' : 'right',
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

/**
 * The two axis lines: the upright one (in the left axis's style) and the one
 * lying along (in the bottom axis's). The category axis lies where it crosses
 * the value axis — `cross.y` up a column chart, `cross.x` along a bar chart —
 * and the other stands at the plot's edge.
 */
function pushAxisLines(
  polylines: Array<ChartPolyline>,
  x0: number,
  y0: number,
  plotW: number,
  plotH: number,
  chart: Chart,
  cross: { readonly x?: number; readonly y?: number } = {},
): void {
  const val = axisStroke(chart.valAxisLine);
  const x = x0 + (cross.x ?? 0);
  const y = y0 + (cross.y ?? 0);
  if (val) {
    polylines.push({
      points: [
        [x, y0],
        [x, y0 + plotH],
      ],
      strokeHex: val.hex,
      widthPt: val.widthPt,
    });
  }
  const cat = axisStroke(chart.catAxisLine);
  if (cat) {
    polylines.push({
      points: [
        [x0, y],
        [x0 + plotW, y],
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

  const nCats = catCount(chart);
  // §21.2.2.9 — a series on the secondary axis is measured against ITS axis, so
  // it is no part of the primary's range and the primary is no part of its.
  const onSecondary = chart.series.filter((s) => s.secondaryAxis);
  const primary =
    onSecondary.length > 0 ? chart.series.filter((s) => !s.secondaryAxis) : chart.series;
  const allVals = primary.flatMap((s) => s.values.slice(0, nCats));
  // The data's own extent: whether the axis reaches down to zero is the
  // scale's decision (niceScale), not the data's.
  const [dataMin, dataMax] = opts.dataRange ?? extentOf(allVals) ?? [0, 1];
  const scale = opts.exactRange
    ? {
        min: dataMin,
        max: dataMax,
        step: niceScale(dataMin, dataMax, intervalsThatFit(horizontal ? wPt : hPt, horizontal))
          .step,
      }
    : axisScale(chart, dataMin, dataMax, horizontal ? wPt : hPt, horizontal);
  const fmtVal = opts.formatValue ?? formatTick;
  const tickVals = ticks(scale);
  const vals2 = onSecondary.flatMap((s) => s.values.slice(0, nCats));
  // The author's own min/max pin the PRIMARY axis (§21.2.2.157 reads one axis);
  // the secondary takes the nice range of its own data.
  const range2 = extentOf(vals2);
  const scale2 = range2 ? niceScale(range2[0], range2[1], intervalsThatFit(hPt, false)) : undefined;
  const tickVals2 = scale2 ? ticks(scale2) : [];

  // §21.2.2.33/§21.2.2.207 — where the category axis crosses the value axis
  // (at zero, unless the file says its minimum, its maximum or a value), and
  // so where its labels stand: beside it there (`nextTo`), or at the low or
  // high end of the value axis. A chart with values below zero has them on
  // its zero line, as Excel and Calc draw it — 47813.xlsx labels its sine and
  // cosine along the middle, where we put them under the plot.
  const crosses = chart.catAxisCrosses;
  const crossValue = Math.min(
    Math.max(
      crosses === 'min' ? scale.min : crosses === 'max' ? scale.max : (crosses ?? 0),
      scale.min,
    ),
    scale.max,
  );
  const labelsAt: 'start' | 'cross' | 'end' | 'none' =
    chart.catTickLabelPos === 'none'
      ? 'none'
      : chart.catTickLabelPos === 'low'
        ? 'start'
        : chart.catTickLabelPos === 'high'
          ? 'end'
          : crossValue <= scale.min
            ? 'start'
            : crossValue >= scale.max
              ? 'end'
              : 'cross';
  // No <c:cat> means the categories are the point indices, which is what
  // Excel and Calc both label the axis with — an unlabelled category axis
  // leaves the bars standing on nothing.
  const catText = (c: number): string => chart.categories[c] ?? String(c + 1);
  const catFace = faceOf(chart, 'catAxis');
  const valFace = faceOf(chart, 'valAxis');
  const val2Face = faceOf(chart, 'secondaryValAxis');
  let widestCat = 0;
  for (let c = 0; c < nCats; c++)
    widestCat = Math.max(widestCat, widthIn(measure, catText(c), catFace));
  const catBand = catFace.sizePt * 1.6;
  // §21.2.2.115 — the outer levels of a category axis labelled on several,
  // each a row of its own under the categories' labels (a column chart's).
  const groupLevels = horizontal ? [] : (chart.categoryGroups ?? []);

  // §21.2.2.33 — the value axis lies where it crosses the category axis: at
  // the first category, unless the file says the last — and the first is at
  // the far end of an axis that runs backwards. dataValidationTableRange.xlsx
  // ranks its counties top down and reads their values along the top, as
  // Excel and Calc draw it; we drew them along the foot.
  const valAtEnd =
    chart.catAxisReversed === true
      ? chart.valAxisCrosses !== 'max'
      : chart.valAxisCrosses === 'max';
  const title = titleLines(chart, wPt, measure);
  const top =
    4 +
    titleHeight(title, chart) +
    (!horizontal && labelsAt === 'end' ? catBand : 0) +
    (horizontal && valAtEnd ? catBand : 0);
  const legend = buildLegendBlock(chart, wPt, hPt, measure);
  const tick2W =
    scale2 && !horizontal
      ? Math.max(0, ...tickVals2.map((v) => widthIn(measure, formatTick(v), val2Face))) +
        4 +
        (chart.secondaryValAxisTitle ? faceOf(chart, 'secondaryValAxisTitle').sizePt * 1.5 : 0)
      : 0;
  // A bar chart's value labels stand centred under their ticks along the
  // foot, so the last one reaches half its width past the plot's end: room is
  // kept for it inside the frame, or "80" is cut in two by the frame's edge.
  const tickHalf = (v: number | undefined): number =>
    horizontal && v !== undefined ? widthIn(measure, fmtVal(v), valFace) / 2 : 0;
  const lastTickHalf = tickHalf(tickVals[tickVals.length - 1]);
  const catLabelW = Math.min(wPt * 0.4, widestCat + 6);
  const tickLabelW = Math.max(0, ...tickVals.map((v) => widthIn(measure, fmtVal(v), valFace))) + 4;
  const plotRight =
    wPt -
    4 -
    legend.rightWidth -
    tick2W -
    (horizontal && labelsAt === 'end' ? catLabelW : lastTickHalf) -
    (!horizontal && valAtEnd ? tickLabelW : 0);

  // The left of the plot holds the axis that stands upright, the foot of it
  // the one that lies along — the value axis in a column chart, the category
  // axis in a bar chart, each with its labels and its title. Sizing the left
  // by the value ticks whatever the direction ran a bar chart's category names
  // out past the frame (dataValidationTableRange.xlsx's "Grays Harbor" beside
  // ticks no wider than "80"). A name is given at most two fifths of the width.
  // The category labels take room only where they stand at the plot's edge;
  // on the zero line they stand inside it.
  const leftTitle = horizontal ? chart.catAxisTitle : chart.valAxisTitle;
  const footTitle = horizontal ? chart.valAxisTitle : chart.catAxisTitle;
  const leftTitleFace = faceOf(chart, horizontal ? 'catAxisTitle' : 'valAxisTitle');
  const footTitleFace = faceOf(chart, horizontal ? 'valAxisTitle' : 'catAxisTitle');
  const own = innerPlot(chart, wPt, hPt);
  const x0 =
    own?.x0 ??
    4 +
      (leftTitle ? leftTitleFace.sizePt * 1.5 : 0) +
      (horizontal
        ? labelsAt === 'start'
          ? catLabelW
          : tickHalf(tickVals[0])
        : valAtEnd
          ? 0
          : tickLabelW);
  const y0 =
    own?.y0 ??
    4 +
      legend.bottomHeight +
      (footTitle ? footTitleFace.sizePt * 1.5 : 0) +
      ((horizontal ? !valAtEnd : labelsAt === 'start') ? catBand : 0) +
      (!horizontal && labelsAt === 'start' ? groupLevels.length * catBand : 0);
  const plotW = own?.plotW ?? Math.max(1, plotRight - x0);
  const plotH = own?.plotH ?? Math.max(1, hPt - top - y0);

  const valueOffset = (v: number): number => fractionOf(v, scale) * (horizontal ? plotW : plotH);
  // Bars grow from where the category axis crosses: zero, or the end of the
  // axis nearest it — bars over 93…97 grow up from 91, the bottom of their axis.
  const zeroOffset = valueOffset(crossValue);

  // §21.2.2.145 — the plot rectangle's own fill and rule, drawn UNDER the
  // gridlines and the data. Chart_Plot_BorderLine_Style.docx rules its plot in
  // a heavy orange dash-dot and we drew nothing around it at all.
  const plotBackground: ChartRect | undefined =
    (chart.plotFillHex ?? chart.plotLine?.colorHex)
      ? {
          x: x0,
          y: y0,
          w: plotW,
          h: plotH,
          ...(chart.plotFillHex ? { fillHex: chart.plotFillHex } : {}),
          ...(chart.plotLine?.colorHex
            ? {
                strokeHex: chart.plotLine.colorHex,
                strokeWidthPt: chart.plotLine.widthPt ?? 0.75,
                ...(chart.plotLine.dash ? { strokeDash: chart.plotLine.dash } : {}),
              }
            : {}),
        }
      : undefined;

  pushChartTitle(labels, title, wPt, hPt, chart);
  // Axis titles. The upright axis's title reads bottom-to-top, in the gutter
  // outside its own labels; the lying one's sits centred below its labels.
  if (leftTitle) {
    labels.push({
      text: leftTitle,
      x: 4 + leftTitleFace.sizePt * 0.9,
      y: y0 + plotH / 2,
      ...leftTitleFace,
      align: 'center',
      rotationDeg: 90,
    });
  }
  if (footTitle) {
    labels.push({
      text: footTitle,
      x: x0 + plotW / 2,
      y: legend.bottomHeight + 2,
      ...footTitleFace,
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
      valFace,
      valAtEnd,
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
      valFace,
      valAtEnd,
    );
  }

  const valueOffset2 =
    scale2 && !horizontal ? (v: number): number => fractionOf(v, scale2) * plotH : undefined;
  if (scale2 && valueOffset2) {
    for (const v of tickVals2) {
      labels.push({
        text: formatTick(v),
        x: x0 + plotW + 3,
        y: y0 + valueOffset2(v) - val2Face.sizePt / 3,
        ...val2Face,
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
        ...faceOf(chart, 'secondaryValAxisTitle'),
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
  // The labels measured are the labels drawn: a chart with no <c:cat> is
  // labelled with its point indices, and measuring the categories it does not
  // have stepped 47813.xlsx's 716 points by five, its numbers run together.
  const need = horizontal ? catFace.sizePt * 1.4 : widestCat + 4;
  const step = Math.max(1, Math.ceil(need / Math.max(slot, 0.01)));
  // Where the labels stand across the axis: beside the plot's start, on the
  // crossing, or beside its end.
  const across =
    labelsAt === 'start' ? 0 : labelsAt === 'end' ? (horizontal ? plotW : plotH) : zeroOffset;
  // Every Nth from where the axis starts — its far end when it runs
  // backwards, so the first category is labelled whichever way it reads.
  const labelled: Array<number> = [];
  if (labelsAt !== 'none') {
    if (chart.catAxisReversed === true) for (let c = nCats - 1; c >= 0; c -= step) labelled.push(c);
    else for (let c = 0; c < nCats; c += step) labelled.push(c);
  }
  for (const c of labelled) {
    const cat = catText(c);
    if (!cat) continue;
    const center = (horizontal ? y0 : x0) + c * slot + slot / 2;
    if (horizontal) {
      labels.push({
        text: cat,
        x: labelsAt === 'end' ? x0 + across + 3 : x0 + across - 3,
        y: center - catFace.sizePt / 3,
        ...catFace,
        align: labelsAt === 'end' ? 'left' : 'right',
      });
    } else {
      labels.push({
        text: cat,
        x: center,
        y: labelsAt === 'end' ? y0 + across + catFace.sizePt * 0.4 : y0 + across - catFace.sizePt,
        ...catFace,
        align: 'center',
      });
    }
  }
  // Each outer level's groups, centred under the categories they span, a
  // rule between one group and the next running down through the rows.
  if (labelsAt !== 'none' && labelsAt !== 'end') {
    groupLevels.forEach((groups, level) => {
      const rowY = y0 + across - catFace.sizePt - (level + 1) * catBand;
      groups.forEach((group, i) => {
        const end = Math.min(nCats, groups[i + 1]?.start ?? nCats);
        if (end <= group.start) return;
        labels.push({
          text: group.label,
          x: x0 + ((group.start + end) / 2) * slot,
          y: rowY,
          ...catFace,
          align: 'center',
        });
        if (i > 0) {
          polylines.push({
            points: [
              [x0 + group.start * slot, y0 + across],
              [x0 + group.start * slot, rowY - catFace.sizePt * 0.4],
            ],
            strokeHex: GRID_COLOR,
            widthPt: 0.75,
          });
        }
      });
    });
  }

  pushAxisLines(
    polylines,
    x0,
    y0,
    plotW,
    plotH,
    chart,
    horizontal
      ? { x: zeroOffset, ...(valAtEnd ? { y: plotH } : {}) }
      : { y: zeroOffset, ...(valAtEnd ? { x: plotW } : {}) },
  );
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
    ...(plotBackground ? { plotBackground } : {}),
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

/**
 * How many category slots a chart has: its labels, or its longest series —
 * counted in a loop, as spreading a file's series into one call overflows
 * the stack.
 */
function catCount(chart: Chart): number {
  let n = Math.max(chart.categories.length, 1);
  for (const s of chart.series) n = Math.max(n, s.values.length);
  return n;
}

/**
 * The least and the greatest of `values`, or undefined where none is a
 * number. Read in a loop: `Math.min(...values)` throws on a series of 200 000
 * points, as a call takes no more arguments than the stack holds.
 */
function extentOf(values: Iterable<number>): readonly [number, number] | undefined {
  let lo = Infinity;
  let hi = -Infinity;
  for (const v of values) {
    if (v < lo) lo = v;
    if (v > hi) hi = v;
  }
  return lo <= hi ? [lo, hi] : undefined;
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
  if (g === 'percentStacked') return { dataRange: [0, 1], exactRange: true, formatValue: pctLabel };
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

/**
 * A tick formatter for a number format code — made per scene and not kept: a
 * cache by code would grow with every code every document ever named.
 */
function formatterOf(code: string | undefined): ((v: number) => string) | undefined {
  if (code === undefined) return undefined;
  const formats = new Map([[CHART_FORMAT_ID, code]]);
  return (v: number) => applyNumberFormat(String(v), CHART_FORMAT_ID, formats);
}

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

/**
 * Whole percentages that add up to a hundred, as Excel labels a pie's slices:
 * each share's floor, and the points still missing given to the slices with
 * the largest remainders — 10, 20 and 15 of 45 read 22%, 45% and 33% in
 * Excel's PDF, where rounding each one reads 22, 44 and 33.
 *
 * @param values The slices' values, a slice that is not positive drawing none.
 * @param total  Their sum.
 * @returns Each slice's whole percentage.
 */
function wholePercents(values: ReadonlyArray<number>, total: number): Array<number> {
  const exact = values.map((v) => (v > 0 && total > 0 ? (v / total) * 100 : 0));
  const out = exact.map((e) => Math.floor(e));
  let missing = 100 - out.reduce((a, b) => a + b, 0);
  const byRemainder = exact
    .map((e, i) => ({ i, rest: e - (out[i] ?? 0) }))
    .filter(({ i }) => (values[i] ?? 0) > 0)
    .sort((a, b) => b.rest - a.rest || a.i - b.i);
  for (const { i } of byRemainder) {
    if (missing <= 0) break;
    out[i] = (out[i] ?? 0) + 1;
    missing--;
  }
  return out;
}

/**
 * §21.2.2.49 — the text a slice's data label shows: the parts the chart's
 * switches turn on, in Excel's order — series name, category name, value,
 * percentage — joined by its separator, ", " where it states none. The
 * labels' own number format is the value's where the value is shown, and the
 * percentage's where only that is; a percentage is otherwise whole.
 * LIBRE_OFFICE-100610-0.pptx's five pies show the 4, 6, 3 and 6 they ask for,
 * 45540_classic_Footer.xlsx its functions' names over their shares, and a pie
 * whose labels show nothing — or that has none — shows nothing, as Excel's.
 *
 * @returns The label, or undefined when the chart labels nothing.
 */
function sliceLabel(
  chart: Chart,
  series: ChartSeries,
  i: number,
  v: number,
  total: number,
  percent: number,
): string | undefined {
  const dl = chart.dataLabels;
  if (!dl) return undefined;
  const own = formatterOf(dl.numberFormat);
  const parts: Array<string> = [];
  if (dl.showSerName) parts.push(series.name ?? legendSeriesName(series, 0));
  if (dl.showCatName) parts.push(chart.categories[i] ?? String(i + 1));
  if (dl.showVal) parts.push(own ? own(v) : fmtDataLabel(chart, v));
  if (dl.showPercent) {
    parts.push(own && !dl.showVal ? own(v / total) : `${String(percent)}%`);
  }
  return parts.length > 0 ? parts.join(dl.separator ?? ', ') : undefined;
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
/**
 * §21.2.2.198 `c:ser/c:spPr/a:ln` — a bar's own outline, as the stroke fields a
 * {@link ChartRect} carries. Chart_BorderLine_Style.docx outlines each of its
 * series in a colour and dash of its own, and drawing bars unstroked lost all
 * of it.
 *
 * @param series The series the bar belongs to.
 * @returns The stroke fields, or an empty object when the series has no rule.
 */
function barOutline(series: ChartSeries): Partial<ChartRect> {
  const ln = series.line;
  if (!ln || ln.none === true || !ln.colorHex) return {};
  return {
    strokeHex: ln.colorHex,
    strokeWidthPt: ln.widthPt ?? 0.75,
    ...(ln.dash && ln.dash !== 'solid' ? { strokeDash: ln.dash } : {}),
  };
}

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

  // §21.2.2.75 — the gap between slots is a percentage of one bar's width,
  // whatever the grouping.
  const gap = (chart.gapPercent ?? 150) / 100;
  if (stacked) {
    // A stack is one bar wide: the slot holds it and its gap. A flat 15% of
    // padding each side made every stack 0.7 of its slot, where Excel's own PDF
    // gives one of `gapWidth="150"` 0.4 (2026-10-02) — WithChartSheet.xlsx's
    // stood nearly twice as wide as Calc draws them.
    const barW = f.slot / (1 + gap);
    const groupPad = (f.slot - barW) / 2;
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
        const outline = barOutline(series);
        if (horizontal)
          f.rects.push({ x: f.x0 + lo, y: along, w: span, h: barW, fillHex: color, ...outline });
        else f.rects.push({ x: along, y: f.y0 + lo, w: barW, h: span, fillHex: color, ...outline });
        if (chart.showValues && span > faceOf(chart, 'dataLabels').sizePt) {
          const raw = series.values[c] ?? 0;
          if (horizontal)
            f.labels.push(
              centeredLabel(
                fmtDataLabel(chart, raw),
                f.x0 + lo + span / 2,
                along + barW * 0.3,
                faceOf(chart, 'dataLabels'),
              ),
            );
          else
            f.labels.push(
              centeredLabel(
                fmtDataLabel(chart, raw),
                along + barW / 2,
                f.y0 + lo + span / 2 - 3,
                faceOf(chart, 'dataLabels'),
              ),
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
      ...(f.plotBackground ? { plotBackground: f.plotBackground } : {}),
      ...(f.plotBackground ? { plotBackground: f.plotBackground } : {}),
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
  // §21.2.2.131 — and the bars of a cluster lie `overlap` percent of a width
  // over each other, or apart where it is negative: Excel's PDF of three
  // series at `overlap="-27"` steps them 1.27 widths apart, each a 5.73rd of
  // the slot at `gapWidth="219"` (2026-10-02).
  const overlap = (chart.overlapPercent ?? 0) / 100;
  const barW = f.slot / (nSer - (nSer - 1) * overlap + gap);
  const pitch = barW * (1 - overlap);
  const groupPad = (f.slot - barW - (nSer - 1) * pitch) / 2;
  for (let c = 0; c < f.nCats; c++) {
    const slotStart = (horizontal ? f.y0 : f.x0) + c * f.slot + groupPad;
    for (let b = 0; b < barIdx.length; b++) {
      const s = barIdx[b]!;
      const series = chart.series[s]!;
      const len = f.valueOffset(series.values[c] ?? 0) - f.zeroOffset; // signed from zero line
      const color = pointColor(series, c) ?? seriesColor(series, s, chart.seriesColorCycle);
      const along = slotStart + b * pitch;
      const outline = barOutline(series);
      if (horizontal) {
        const bx = f.x0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: bx, y: along, w: Math.abs(len), h: barW, fillHex: color, ...outline });
      } else {
        const by = f.y0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: along, y: by, w: barW, h: Math.abs(len), fillHex: color, ...outline });
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
            ...faceOf(chart, 'dataLabels'),
            align: len >= 0 ? 'left' : 'right',
          });
        } else {
          const end = f.y0 + f.zeroOffset + len;
          const face = faceOf(chart, 'dataLabels');
          f.labels.push(
            centeredLabel(txt, along + barW * 0.45, len >= 0 ? end + 2 : end - face.sizePt, face),
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
    ...(f.plotBackground ? { plotBackground: f.plotBackground } : {}),
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
function pushComboLines(chart: Chart, f: CartesianFrame, lineIdx: ReadonlyArray<number>): void {
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

function centeredLabel(text: string, x: number, y: number, face: TextFace): ChartLabel {
  return { text, x, y, ...face, align: 'center' };
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
    ...(f.plotBackground ? { plotBackground: f.plotBackground } : {}),
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
  const [xLo, xHi] = extentOf(xs) ?? [0, 1];
  const [yLo, yHi] = extentOf(ys) ?? [0, 1];
  // …and each axis takes the ends, step and number format its author fixed
  // (§21.2.2.157, §21.2.2.98, §21.2.2.121), the lying one its own.
  const xScale = scatterScale(xLo, xHi, intervalsThatFit(wPt, true), {
    min: chart.xAxisMin,
    max: chart.xAxisMax,
    unit: chart.xAxisMajorUnit,
  });
  const yScale = scatterScale(yLo, yHi, intervalsThatFit(hPt, false), {
    min: chart.valAxisMin,
    max: chart.valAxisMax,
    unit: chart.valAxisMajorUnit,
  });
  const xTicks = ticks(xScale);
  const yTicks = ticks(yScale);
  const fmtX = formatterOf(chart.xNumberFormat) ?? formatTick;
  const fmtY = formatterOf(chart.numberFormat) ?? formatTick;

  const title = titleLines(chart, wPt, measure);
  const top = 4 + titleHeight(title, chart);
  const legend = buildLegendBlock(chart, wPt, hPt, measure);
  const xFace = faceOf(chart, 'catAxis');
  const yFace = faceOf(chart, 'valAxis');
  // The axis titles, as a frame chart sets them: the upright axis's reading
  // bottom-to-top in the gutter outside its labels, the lying one's centred
  // under its own.
  const yTitleFace = faceOf(chart, 'valAxisTitle');
  const xTitleFace = faceOf(chart, 'catAxisTitle');
  const tickLabelW = Math.max(0, ...yTicks.map((v) => widthIn(measure, fmtY(v), yFace))) + 4;
  const own = innerPlot(chart, wPt, hPt);
  const x0 = own?.x0 ?? 4 + (chart.valAxisTitle ? yTitleFace.sizePt * 1.5 : 0) + tickLabelW;
  const y0 =
    own?.y0 ??
    4 +
      legend.bottomHeight +
      (chart.catAxisTitle ? xTitleFace.sizePt * 1.5 : 0) +
      xFace.sizePt * 1.6;
  const plotW = own?.plotW ?? Math.max(1, wPt - 4 - legend.rightWidth - x0);
  const plotH = own?.plotH ?? Math.max(1, hPt - top - y0);
  const xAt = (v: number): number => x0 + fractionOf(v, xScale) * plotW;
  const yAt = (v: number): number => y0 + fractionOf(v, yScale) * plotH;

  pushChartTitle(labels, title, wPt, hPt, chart);
  if (chart.valAxisTitle) {
    labels.push({
      text: chart.valAxisTitle,
      x: 4 + yTitleFace.sizePt * 0.9,
      y: y0 + plotH / 2,
      ...yTitleFace,
      align: 'center',
      rotationDeg: 90,
    });
  }
  if (chart.catAxisTitle) {
    labels.push({
      text: chart.catAxisTitle,
      x: x0 + plotW / 2,
      y: legend.bottomHeight + 2,
      ...xTitleFace,
      align: 'center',
    });
  }
  pushGridTicks(
    gridlines,
    labels,
    yTicks,
    fmtY,
    'y',
    yAt,
    x0,
    y0,
    plotW,
    plotH,
    chart.gridLine,
    yFace,
  );
  pushGridTicks(
    gridlines,
    labels,
    xTicks,
    fmtX,
    'x',
    xAt,
    x0,
    y0,
    plotW,
    plotH,
    chart.gridLine,
    xFace,
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
 * How many intervals an axis `extentPt` long has room to label: ten, Excel's
 * own most, unless its labels would run into each other first — one label
 * line apiece stacked up a vertical axis, a few digits' width apiece along a
 * horizontal one. Excel crams nine labels up a 150pt column chart rather than
 * step it more coarsely, so the room is all that limits it.
 */
const intervalsThatFit = (extentPt: number, horizontal: boolean): number =>
  Math.min(
    10,
    Math.max(2, Math.floor((extentPt * 0.75) / (CHART_LABEL_PT * (horizontal ? 3.5 : 1.4)))),
  );

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
  // The data's own extent: whether the axis reaches down to zero is the
  // scale's to decide, by Excel's rule (niceScale) — chartex.docx's 1.8…5
  // line chart runs 0…6, as both references draw it, and 93…97 from 91.
  const allVals = chart.series.flatMap((s) => s.values);
  const range = extentOf(allVals) ?? [0, 1];
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
        f.labels.push(
          centeredLabel(
            fmtDataLabel(chart, series.values[c] ?? 0),
            x,
            y + 3,
            faceOf(chart, 'dataLabels'),
          ),
        );
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
    ...(f.plotBackground ? { plotBackground: f.plotBackground } : {}),
    wedges,
    labels: f.labels,
  };
}

// ─── pie chart ──────────────────────────────────────────────────────────────
// A pie's slices cycle the same colours a bar chart's SERIES do — the chart's
// own cycle when it has one, which is the workbook theme's accents.
const sliceColor = (series: ChartSeries, i: number, cycle?: ReadonlyArray<string>): string => {
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

  const title = titleLines(chart, wPt, measure);
  const top = 4 + titleHeight(title, chart);
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
    faceOf(chart, 'legend'),
  );

  // §21.2.2.95 — the box the plot was sized to, where the author sized it,
  // and the room left by the title and the legend where they did not.
  const plotBox = chart.plotBox;
  const availW = plotBox ? plotBox.w * wPt : Math.max(1, wPt - 8 - legend.rightWidth);
  const availH = plotBox ? plotBox.h * hPt : Math.max(1, hPt - top - 4 - legend.bottomHeight);
  const cx = plotBox ? (plotBox.x + plotBox.w / 2) * wPt : 4 + availW / 2;
  const cy = plotBox
    ? hPt - (plotBox.y + plotBox.h / 2) * hPt
    : 4 + legend.bottomHeight + availH / 2;

  // §21.2.2.143 — a 3-D pie is a disc tilted to its view's elevation: an
  // ellipse `tilt` as tall as it is wide, standing `thickness` deep, its
  // front edge showing. A flat pie is the disc seen from straight above.
  const disc = chart.pie3D;
  const elevation = disc ? (Math.min(90, Math.max(0, disc.rotX)) * Math.PI) / 180 : Math.PI / 2;
  // Floored a little above edge-on, where the disc would vanish; the depth is
  // Excel's, read off its PDF: a quarter of the radius seen edge on, 0.17 of
  // it at 30°, shrinking with the square of the elevation's cosine.
  const tilt = Math.max(Math.sin(elevation), 0.02);
  const thickF = disc ? 0.25 * Math.cos(elevation) ** 2 : 0;
  // §21.2.2.62 — each slice stands out from the centre by its explosion, and
  // the pie shrinks so the farthest of them still fits.
  const explosionOfSlice = (i: number): number =>
    (atPoint(series.pointExplosions, i)?.percent ?? series.explosion ?? 0) / 100;
  let maxOut = 0;
  for (let i = 0; i < values.length; i++) {
    if ((values[i] ?? 0) > 0) maxOut = Math.max(maxOut, explosionOfSlice(i));
  }
  // In the author's box the pie fills it, its smaller side whole, as Excel's
  // PDF has it; in the room left over, a little inside it.
  const fullR = Math.max(
    1,
    Math.min(availW / 2 / (1 + maxOut), availH / (2 * tilt * (1 + maxOut) + thickF)) *
      (plotBox ? 1 : 0.95),
  );

  // Excel pies start at 12 o'clock and sweep clockwise (negative in y-up),
  // turned by the first slice's angle (§21.2.2.68) or the 3-D view's.
  const turn = ((disc ? disc.rotY : (chart.firstSliceAngle ?? 0)) * Math.PI) / 180;
  const sweeps = values.map((v) => (v > 0 ? -(v / total) * 2 * Math.PI : 0));
  const starts: Array<number> = [];
  let ang = Math.PI / 2 - turn;
  for (const sweep of sweeps) {
    starts.push(ang);
    ang += sweep;
  }
  const midOf = (i: number): number => (starts[i] ?? 0) + (sweeps[i] ?? 0) / 2;

  // The labels first: what they say, where they stand and how much room they
  // take — a label set outside the pie takes its room from the pie's radius.
  const boxes = pieLabelBoxes(chart, series, values, total, sweeps, starts, fullR, wPt, measure);
  // Each label set outside takes its room on its own side: a label beside the
  // pie its width, one over or under it its height — the radius is what the
  // tightest of them leaves. A label the author placed stands where they put
  // it, and takes none.
  let room = fullR;
  for (const b of plotBox ? [] : boxes) {
    if (b.where !== 'out' || b.placement) continue;
    const c = Math.abs(Math.cos(b.mid));
    const sn = Math.abs(Math.sin(b.mid)) * tilt;
    if (c > 0.05) room = Math.min(room, (availW / 2 - PIE_LABEL_GAP - b.width) / c);
    if (sn > 0.05) room = Math.min(room, (availH / 2 - PIE_LABEL_GAP - b.height) / sn);
  }

  // The disc for a radius: its top face's centre — raised by half its depth,
  // so the face and the edge under it are centred together — and each
  // slice's own centre, moved out by its explosion.
  const discOf = (rr: number) => {
    const thickness = thickF * rr;
    const cyTop = cy + thickness / 2;
    const centre = (i: number): readonly [number, number] => {
      const out = explosionOfSlice(i) * rr;
      const mid = midOf(i);
      return [cx + Math.cos(mid) * out, cyTop + Math.sin(mid) * out * tilt];
    };
    const point = (i: number, t: number, along: number): readonly [number, number] => {
      const [x0, y0] = centre(i);
      return [x0 + Math.cos(t) * along, y0 + Math.sin(t) * along * tilt];
    };
    return { thickness, cyTop, centre, point };
  };

  // Where each label stands for a pie of radius `pieR` (the chart's place
  // for it, or the author's), kept inside the chart. A label the author
  // placed is placed in the chart's frame — from where it stood round the
  // full pie — and stays there however far the pie then gives way.
  const positionOf = (box: PieLabelBox, pieR: number): { bx: number; by: number } => {
    const rr = box.placement ? fullR : pieR;
    const d = discOf(rr);
    const i = box.index;
    const cos = Math.cos(box.mid);
    const sin = Math.sin(box.mid);
    // Where the chart would set it: on the ring of a doughnut, inside a
    // slice at the depth its position asks for, or past the slice's end,
    // leaning away from the pie on its own side — under the disc's edge
    // where the slice is in front.
    let bx: number;
    let by: number;
    if (box.where === 'out') {
      const [px, py0] = d.point(i, box.mid, rr + PIE_LABEL_GAP / Math.max(tilt, 0.2));
      const py = sin < 0 ? py0 - d.thickness : py0;
      bx = px + (cos >= 0 ? box.width / 2 : -box.width / 2);
      by = py + (sin * box.height) / 2;
    } else {
      const depth = chart.doughnut
        ? (rr * 0.5 + rr) / 2
        : rr * (box.where === 'ctr' ? 0.5 : box.where === 'end' ? 0.75 : 0.6);
      [bx, by] = d.point(i, box.mid, depth);
    }
    // …or where the author dragged it (§21.2.2.95): its own corner as a
    // fraction of the chart, or that far from where the chart would set it.
    const moved = box.placement;
    if (moved?.edge) {
      if (moved.x !== undefined) bx = moved.x * wPt + box.width / 2;
      if (moved.y !== undefined) by = hPt - moved.y * hPt - box.height / 2;
    } else if (moved) {
      bx += (moved.x ?? 0) * wPt;
      by -= (moved.y ?? 0) * hPt;
    }
    return {
      bx: Math.min(Math.max(bx, box.width / 2 + 2), wPt - box.width / 2 - 2),
      by: Math.min(Math.max(by, box.height / 2 + 2), hPt - box.height / 2 - 2),
    };
  };
  // …and no label meant to stand outside the pie lies over it: the pie gives
  // way to the labels set round it. 45544.xlsx's names stand where its author
  // dragged them round Excel's 3-D pie; a label dragged INTO its slice is
  // meant there, and is not in the way. Distances are the disc's own, its
  // height counted back up by its tilt.
  const across = (x: number, y: number, cyTop: number): number =>
    Math.hypot(x - cx, (y - cyTop) / tilt);
  let r = Math.max(fullR * 0.4, room);
  for (const box of plotBox ? [] : boxes) {
    if (box.where !== 'out' || !box.placement) continue;
    const { bx, by } = positionOf(box, r);
    const cyTop = discOf(fullR).cyTop;
    if (across(bx, by, cyTop) < fullR * 0.9) continue;
    const nx = Math.min(Math.max(cx, bx - box.width / 2), bx + box.width / 2);
    const ny = Math.min(Math.max(cyTop, by - box.height / 2), by + box.height / 2);
    r = Math.max(fullR * 0.4, Math.min(r, across(nx, ny, cyTop) - PIE_LABEL_GAP / 2));
  }

  const d = discOf(r);
  // A doughnut is a pie with a central hole; place its labels out on the ring.
  const holeR = chart.doughnut ? r * 0.5 : 0;
  const polygons: Array<ChartPolygon> = [];
  if (disc) {
    // The edge first, back to front, then every slice's face over it: the
    // front of the disc shows its depth in each slice's colour, darkened, and
    // a slice standing out shows its cut sides.
    const sides: Array<ChartPolygon & { readonly depth: number }> = [];
    for (let i = 0; i < values.length; i++) {
      if ((values[i] ?? 0) <= 0) continue;
      const fill = shadeHex(sliceColor(series, i, chart.seriesColorCycle), 0.7);
      const a0 = starts[i]!;
      const a1 = a0 + sweeps[i]!;
      for (const [from, to] of frontArcs(a1, a0)) {
        const topArc = arcPoints(from, to, (t) => d.point(i, t, r));
        const bottom = topArc.map(([x, y]) => [x, y - d.thickness] as const).reverse();
        sides.push({
          points: [...topArc, ...bottom],
          fillHex: fill,
          strokeHex: fill,
          widthPt: 0.5,
          depth: Math.min(...topArc.map(([, y]) => y)),
        });
      }
      if (explosionOfSlice(i) > 0) {
        const [x0, y0] = d.centre(i);
        for (const t of [a0, a1]) {
          const [ex, ey] = d.point(i, t, r);
          sides.push({
            points: [
              [x0, y0],
              [ex, ey],
              [ex, ey - d.thickness],
              [x0, y0 - d.thickness],
            ],
            fillHex: fill,
            strokeHex: fill,
            widthPt: 0.5,
            depth: Math.min(y0, ey),
          });
        }
      }
    }
    sides.sort((a, b) => b.depth - a.depth);
    for (const { depth: _depth, ...side } of sides) polygons.push(side);
    for (let i = 0; i < values.length; i++) {
      if ((values[i] ?? 0) <= 0) continue;
      const a0 = starts[i]!;
      polygons.push({
        points: [d.centre(i), ...arcPoints(a0, a0 + sweeps[i]!, (t) => d.point(i, t, r))],
        fillHex: sliceColor(series, i, chart.seriesColorCycle),
        strokeHex: 'FFFFFF',
        widthPt: 0.75,
      });
    }
  } else {
    for (let i = 0; i < values.length; i++) {
      if ((values[i] ?? 0) <= 0) continue;
      const [x0, y0] = d.centre(i);
      wedges.push({
        cx: x0,
        cy: y0,
        r,
        startRad: starts[i]!,
        sweepRad: sweeps[i]!,
        fillHex: sliceColor(series, i, chart.seriesColorCycle),
        strokeHex: 'FFFFFF',
      });
    }
    // Punch the hole: a white disc over the wedge centres (drawn after slices).
    if (holeR > 0) {
      wedges.push({ cx, cy, r: holeR, startRad: 0, sweepRad: -2 * Math.PI, fillHex: 'FFFFFF' });
    }
  }

  const polylines: Array<ChartPolyline> = [];
  const face = faceOf(chart, 'dataLabels');
  const lineH = face.sizePt * 1.2;
  const placed = boxes.map((box) => ({ box, ...positionOf(box, r) }));
  // Labels set outside the pie that would run into each other are stacked
  // apart, top down, as Excel's best fit spreads the names of a run of thin
  // slices — and kept off the chart's foot. A label its author placed stays
  // where they put it, and the ones the chart sets give way to it.
  const outside = placed
    .filter((p) => p.box.where === 'out' || across(p.bx, p.by, d.cyTop) >= r)
    .sort((a, b) => b.by - a.by);
  const overlaps = (a: (typeof placed)[number], b: (typeof placed)[number]): boolean =>
    Math.abs(a.bx - b.bx) < (a.box.width + b.box.width) / 2 &&
    Math.abs(a.by - b.by) < (a.box.height + b.box.height) / 2 + 2;
  const fixed = outside.filter((p) => p.box.placement);
  const chartSet = outside.filter((p) => !p.box.placement);
  // Spreading compares each label with the ones before it, and a pie of a
  // hundred thousand slices — a file can ask for one — made that a hang:
  // past a few dozen labels a pie is a crowd no spreading reads, and they
  // stay where they stand.
  if (fixed.length + chartSet.length <= MOST_SPREAD_LABELS) {
    chartSet.forEach((here, k) => {
      // Every label its author placed is in the way wherever it stands, and so
      // is every one the chart set above this one; moving under one may land
      // on another, so the walk goes on until none is hit.
      const blocks = (other: (typeof placed)[number]): boolean => overlaps(other, here);
      for (let pass = 0; pass <= fixed.length + k; pass++) {
        const hit = fixed.find(blocks) ?? chartSet.slice(0, k).find(blocks);
        if (!hit) break;
        here.by = hit.by - (hit.box.height + here.box.height) / 2 - 2;
      }
    });
    const lowest = chartSet.reduce<(typeof placed)[number] | undefined>(
      (low, p) => (low === undefined || p.by < low.by ? p : low),
      undefined,
    );
    const under = lowest ? 2 - (lowest.by - lowest.box.height / 2) : 0;
    if (under > 0) for (const p of chartSet) p.by += under;
  }
  for (const { box, bx, by } of placed) {
    const moved = box.placement;
    // §21.2.2.181 — a label off its slice is tied back to it, where the chart
    // asks for leader lines; one dragged into the pie stands on its slice, and
    // takes none (orderOfCNumFmtElements.xlsx's 67%).
    const off = (box.where === 'out' || moved !== undefined) && across(bx, by, d.cyTop) >= r;
    if (off && chart.dataLabels?.showLeaderLines) {
      const [ex, ey] = d.point(box.index, box.mid, r);
      const tx = Math.min(Math.max(ex, bx - box.width / 2), bx + box.width / 2);
      const ty = Math.min(Math.max(ey, by - box.height / 2), by + box.height / 2);
      if (Math.hypot(tx - ex, ty - ey) > 3) {
        polylines.push({
          points: [
            [ex, ey],
            [tx, ty],
          ],
          strokeHex: LABEL_COLOR,
          widthPt: 0.75,
        });
      }
    }
    // White on the slice where the reader resolved no colour; the chart's own
    // colour off it, on a ring, or for a label the author typed.
    const colorHex =
      box.where === 'out' ||
      moved !== undefined ||
      box.custom ||
      chart.doughnut ||
      chart.text?.dataLabels?.colorHex
        ? face.colorHex
        : 'FFFFFF';
    box.lines.forEach((line, k) => {
      labels.push({
        text: line,
        x: bx,
        y: by - face.sizePt / 3 + ((box.lines.length - 1) / 2 - k) * lineH,
        ...face,
        colorHex,
        align: 'center',
      });
    });
  }

  pushChartTitle(labels, title, wPt, hPt, chart);
  legend.emit(rects, labels);
  return { rects, polylines, wedges, labels, ...(polygons.length > 0 ? { polygons } : {}) };
}

/**
 * The parts of an arc from `from` to `to` (radians, `from` < `to`) that lie
 * on the disc's front half — below its centre in the y-up frame, where sin
 * is negative — each as a [from, to] pair.
 */
function frontArcs(from: number, to: number): Array<readonly [number, number]> {
  const out: Array<readonly [number, number]> = [];
  // The front half is (π, 2π) modulo 2π; walk the arc's turns.
  const TAU = 2 * Math.PI;
  const k0 = Math.floor((from - Math.PI) / TAU) - 1;
  const k1 = Math.ceil((to - Math.PI) / TAU) + 1;
  for (let k = k0; k <= k1; k++) {
    const lo = Math.max(from, Math.PI + k * TAU);
    const hi = Math.min(to, TAU + k * TAU);
    if (hi - lo > 1e-6) out.push([lo, hi]);
  }
  return out;
}

/** Points along an arc from `a0` by its sweep to `a1`, one every 3.75°. */
function arcPoints(
  a0: number,
  a1: number,
  at: (t: number) => readonly [number, number],
): Array<readonly [number, number]> {
  const n = Math.max(2, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 48)));
  const out: Array<readonly [number, number]> = [];
  for (let k = 0; k <= n; k++) out.push(at(a0 + ((a1 - a0) * k) / n));
  return out;
}

/** `hex` with each channel scaled by `f` — a darker shade of it below 1. */
function shadeHex(hex: string, f: number): string {
  const n = parseInt(hex, 16);
  const ch = (shift: number): string =>
    Math.max(0, Math.min(255, Math.round(((n >> shift) & 255) * f)))
      .toString(16)
      .padStart(2, '0');
  return `${ch(16)}${ch(8)}${ch(0)}`.toUpperCase();
}

/** The room kept between a pie and a label set outside it. */
const PIE_LABEL_GAP = 6;

/** The most labels round a pie that are spread apart (see buildPieScene). */
const MOST_SPREAD_LABELS = 64;

/** A slice's label, measured before the pie is sized. */
interface PieLabelBox {
  /** The slice's index. */
  readonly index: number;
  readonly mid: number;
  readonly lines: ReadonlyArray<string>;
  readonly width: number;
  readonly height: number;
  /** Inside the slice — at its middle, its centre or its end — or outside it. */
  readonly where: 'in' | 'ctr' | 'end' | 'out';
  readonly placement?: { readonly x?: number; readonly y?: number; readonly edge?: boolean };
  readonly custom: boolean;
}

/**
 * §21.2.2.48 — each slice's label, measured and placed: inside where the
 * position asks for it, outside for `outEnd`, and for `bestFit`, Excel's
 * default, inside where it fits across the slice and outside where it does
 * not — wrapped at its spaces to under a third of the chart's width, as
 * 45544.xlsx's function names stand beside their slices.
 */
function pieLabelBoxes(
  chart: Chart,
  series: ChartSeries,
  values: ReadonlyArray<number>,
  total: number,
  sweeps: ReadonlyArray<number>,
  starts: ReadonlyArray<number>,
  fullR: number,
  wPt: number,
  measure: MeasureText,
): Array<PieLabelBox> {
  const face = faceOf(chart, 'dataLabels');
  const lineH = face.sizePt * 1.2;
  const percents = wholePercents(values, total);
  const widest = (lines: ReadonlyArray<string>): number =>
    lines.reduce((m, l) => Math.max(m, widthIn(measure, l, face)), 0);
  const out: Array<PieLabelBox> = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i] ?? 0;
    if (v <= 0) continue;
    // A label the author typed wins over the one we would compute — see
    // ChartSeries.pointLabels — and the one we compute says what the chart's
    // switches ask for (sliceLabel), or nothing at all.
    const custom = atPoint(series.pointLabels, i)?.text;
    const text = custom ?? sliceLabel(chart, series, i, v, total, percents[i] ?? 0);
    if (text === undefined || text.length === 0) continue;
    const placement = atPoint(series.pointLabelPlacements, i);
    const position = placement?.position ?? chart.dataLabels?.position ?? 'bestFit';
    const sweep = Math.abs(sweeps[i] ?? 0);
    const mid = (starts[i] ?? 0) - sweep / 2;
    // A separator may break the label into lines (`c:separator` "\n").
    let lines = text.split('\n');
    let width = widest(lines);
    // The room across the slice where its label stands — and across most of
    // the pie for a slice that is more than half of it: orderOfCNumFmtElements
    // .xlsx sets its 79% slice's "110кВ; 26,7млрд.кВтч; 79,2%" on one line.
    const across =
      sweep >= Math.PI ? fullR * 1.6 : 2 * fullR * 0.6 * Math.sin(Math.min(sweep, Math.PI) / 2);
    let fits = width + 4 <= across && lines.length * lineH <= fullR * 0.7;
    // A label the author dragged moved from where it stood unwrapped: the
    // offset is from THAT place (§21.2.2.95), outside its slice where it did
    // not fit — orderOfCNumFmtElements.xlsx's 67% slice's label, dragged up
    // and left from past the slice's end, lands inside it.
    const dragged =
      placement !== undefined &&
      !placement.edge &&
      (placement.x !== undefined || placement.y !== undefined);
    if (!fits && position === 'bestFit' && !dragged) {
      // …wrapped first, as Excel fits a label into its slice before it sets it
      // outside: orderOfCNumFmtElements.xlsx's 67% slice holds its whole
      // "Промышленные потребители; 22,7млрд.кВтч; 67,3%" in four lines.
      const wrapped = lines.flatMap((line) =>
        wrapAtSpaces(line, Math.min(across * 0.9, wPt * 0.3), (t) => widthIn(measure, t, face)),
      );
      const wrappedWidth = widest(wrapped);
      if (wrappedWidth + 4 <= across && wrapped.length * lineH <= fullR * 0.7) {
        lines = wrapped;
        width = wrappedWidth;
        fits = true;
      }
    }
    const where: PieLabelBox['where'] = chart.doughnut
      ? 'in'
      : position === 'outEnd' || (position === 'bestFit' && !fits)
        ? 'out'
        : position === 'ctr' || position === 'inBase'
          ? 'ctr'
          : position === 'inEnd'
            ? 'end'
            : 'in';
    if (where === 'out') {
      const room = Math.max(40, wPt * 0.3);
      lines = lines.flatMap((line) => wrapAtSpaces(line, room, (t) => widthIn(measure, t, face)));
      width = widest(lines);
    }
    out.push({
      index: i,
      mid,
      lines,
      width,
      height: lines.length * lineH,
      where,
      ...(placement && (placement.x !== undefined || placement.y !== undefined)
        ? {
            placement: {
              ...(placement.x !== undefined ? { x: placement.x } : {}),
              ...(placement.y !== undefined ? { y: placement.y } : {}),
              ...(placement.edge ? { edge: true } : {}),
            },
          }
        : {}),
      custom: custom !== undefined,
    });
  }
  return out;
}

/** A line broken at its spaces into lines no wider than `room` where a break allows. */
function wrapAtSpaces(line: string, room: number, width: (t: string) => number): Array<string> {
  const words = line.split(' ');
  const out: Array<string> = [];
  let cur = '';
  for (const word of words) {
    const longer = cur ? `${cur} ${word}` : word;
    if (cur && width(longer) > room) {
      out.push(cur);
      cur = word;
    } else {
      cur = longer;
    }
  }
  out.push(cur);
  return out;
}

function pointColor(series: ChartSeries, idx: number): string | undefined {
  return atPoint(series.pointColors, idx)?.colorHex;
}

/** Each override list's entries by point index, built on first use. */
const overridesByPoint = new WeakMap<
  ReadonlyArray<{ readonly idx: number }>,
  ReadonlyMap<number, { readonly idx: number }>
>();

/**
 * The first of `overrides` for the point at `idx` — looked up, not searched
 * for: a series that overrides many of its points, searched through for every
 * one of them, took the square of their number.
 */
function atPoint<T extends { readonly idx: number }>(
  overrides: ReadonlyArray<T> | undefined,
  idx: number,
): T | undefined {
  if (!overrides) return undefined;
  let byPoint = overridesByPoint.get(overrides);
  if (!byPoint) {
    const index = new Map<number, T>();
    for (const o of overrides) if (!index.has(o.idx)) index.set(o.idx, o);
    overridesByPoint.set(overrides, index);
    byPoint = index;
  }
  return byPoint.get(idx) as T | undefined;
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
  face: TextFace,
  marker: LegendMarker = 'box',
): LegendLayout {
  if (!hasLegend || entries.length === 0) {
    return { rightWidth: 0, bottomHeight: 0, emit: () => {} };
  }
  const sw = face.sizePt; // key height reference
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
  const entryW = (e: LegendEntry): number => keyW + 3 + widthIn(measure, e.name, face) + gap * 2;

  if (pos === 'r' || pos === 'l') {
    const colW = entries.reduce((w, e) => Math.max(w, entryW(e)), 0);
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
            ...face,
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
          ...face,
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
  const n = catCount(chart);
  const flip = (i: number): number => n - 1 - i;
  return {
    ...chart,
    // An empty category list is not a list of empty labels: the builders label
    // an unlabelled axis with the point index — which, run backwards, counts
    // down.
    categories: Array.from({ length: n }, (_, i) =>
      chart.categories.length > 0 ? (chart.categories[flip(i)] ?? '') : String(flip(i) + 1),
    ),
    series: chart.series.map((s) => ({
      ...s,
      values: Array.from({ length: n }, (_, i) => s.values[flip(i)] ?? 0),
      ...(s.pointColors
        ? { pointColors: s.pointColors.map((p) => ({ ...p, idx: flip(p.idx) })) }
        : {}),
      ...(s.pointLabels
        ? { pointLabels: s.pointLabels.map((p) => ({ ...p, idx: flip(p.idx) })) }
        : {}),
      ...(s.pointLabelPlacements
        ? {
            pointLabelPlacements: s.pointLabelPlacements.map((p) => ({ ...p, idx: flip(p.idx) })),
          }
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
    ...(chart.frameLineHex
      ? {
          strokeHex: chart.frameLineHex,
          strokeWidthPt: chart.frameLineWidthPt ?? 0.75,
          ...(chart.frameLineDash ? { strokeDash: chart.frameLineDash } : {}),
        }
      : {}),
  };
  return { ...scene, background: frame };
}
