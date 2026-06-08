// Pure chart geometry: a Chart + box (w×h points) → a ChartScene of rectangles,
// polylines, wedges and labels in a LOCAL y-up frame (origin bottom-left). No
// PDF or font dependency — text widths come through an injected measure fn, so
// this module is unit-testable in isolation. The renderer converts the scene
// to draw commands (rects/polylines/wedges via the vector layer, labels via the
// text pass).

import type { Chart, ChartSeries } from '@/document-model';

export interface ChartRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly fillHex?: string;
  readonly strokeHex?: string;
  readonly strokeWidthPt?: number;
}
export interface ChartPolyline {
  readonly points: ReadonlyArray<readonly [number, number]>;
  readonly strokeHex: string;
  readonly widthPt: number;
}
// A closed, filled polygon (area-chart bands). Drawn before strokes/labels.
export interface ChartPolygon {
  readonly points: ReadonlyArray<readonly [number, number]>;
  readonly fillHex: string;
  readonly strokeHex?: string;
  readonly widthPt?: number;
}
export interface ChartWedge {
  readonly cx: number;
  readonly cy: number;
  readonly r: number;
  readonly startRad: number;
  readonly sweepRad: number;
  readonly fillHex: string;
  readonly strokeHex?: string;
}
export type LabelAlign = 'left' | 'center' | 'right';
export interface ChartLabel {
  readonly text: string;
  readonly x: number; // anchor point; `align` says how text sits relative to it
  readonly y: number; // baseline
  readonly sizePt: number;
  readonly colorHex: string;
  readonly align: LabelAlign;
}
export interface ChartScene {
  readonly rects: ReadonlyArray<ChartRect>;
  readonly polylines: ReadonlyArray<ChartPolyline>;
  readonly wedges: ReadonlyArray<ChartWedge>;
  readonly labels: ReadonlyArray<ChartLabel>;
  readonly polygons?: ReadonlyArray<ChartPolygon>;
}

export type MeasureText = (text: string, sizePt: number) => number;

export const CHART_LABEL_PT = 9;
export const CHART_TITLE_PT = 13;
const AXIS_COLOR = '595959';
const GRID_COLOR = 'D9D9D9';
const LABEL_COLOR = '595959';
const TITLE_COLOR = '404040';

// Office accent cycle for series without an explicit colour.
export const SERIES_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

export const seriesColor = (s: ChartSeries, i: number): string =>
  s.colorHex ?? SERIES_COLORS[i % SERIES_COLORS.length]!;

// ─── value-axis "nice numbers" (Heckbert) ──────────────────────────────────
function niceNum(range: number, round: boolean): number {
  const exp = Math.floor(Math.log10(range));
  const f = range / 10 ** exp;
  let nf: number;
  if (round) nf = f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10;
  else nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * 10 ** exp;
}

export interface Scale {
  readonly min: number;
  readonly max: number;
  readonly step: number;
}

export function niceScale(dataMin: number, dataMax: number, maxTicks = 6): Scale {
  const lo = Math.min(dataMin, dataMax);
  let hi = Math.max(dataMin, dataMax);
  if (lo === hi) {
    hi = lo + 1;
  }
  const range = niceNum(hi - lo, false);
  const step = niceNum(range / Math.max(1, maxTicks - 1), true);
  return { min: Math.floor(lo / step) * step, max: Math.ceil(hi / step) * step, step };
}

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
  // mutable scene chrome the caller appends series geometry to:
  readonly rects: Array<ChartRect>;
  readonly polylines: Array<ChartPolyline>;
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
  const labels: Array<ChartLabel> = [];

  const nCats = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1);
  const allVals = chart.series.flatMap((s) => s.values.slice(0, nCats));
  const [dataMin, dataMax] = opts.dataRange ?? [Math.min(0, ...allVals), Math.max(0, ...allVals)];
  const scale = niceScale(dataMin, dataMax);
  const fmtVal = opts.formatValue ?? ((v: number): string => formatTick(v, scale.step));
  const tickVals = ticks(scale);

  let top = 4;
  if (chart.title) top += CHART_TITLE_PT * 1.6;
  if (chart.valAxisTitle) top += CHART_LABEL_PT * 1.4;
  const legendEntries: Array<LegendEntry> = chart.series
    .map((s, i) => ({ name: s.name ?? '', colorHex: seriesColor(s, i) }))
    .filter((e) => e.name !== '');
  const legend = layoutLegend(
    legendEntries,
    chart.hasLegend,
    chart.legendPos ?? 'b',
    wPt,
    hPt,
    measure,
  );
  const plotRight = wPt - 4 - legend.rightWidth;

  const tickLabelW = Math.max(0, ...tickVals.map((v) => measure(fmtVal(v), CHART_LABEL_PT))) + 4;
  const catLabelH = CHART_LABEL_PT * 1.6;
  const catTitleH = chart.catAxisTitle ? CHART_LABEL_PT * 1.5 : 0;

  const x0 = 4 + tickLabelW;
  const y0 = 4 + legend.bottomHeight + catTitleH + catLabelH;
  const plotW = Math.max(1, plotRight - x0);
  const plotH = Math.max(1, hPt - top - y0);

  const valueOffset = (v: number): number =>
    ((v - scale.min) / (scale.max - scale.min)) * (horizontal ? plotW : plotH);
  const zeroOffset = valueOffset(0);

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
  // Axis titles. The value-axis title is laid out horizontally above the plot
  // (a deliberate simplification — chart text is not rotated); the category-axis
  // title sits centred below the category labels.
  if (chart.valAxisTitle) {
    labels.push({
      text: chart.valAxisTitle,
      x: x0,
      y: hPt - top + 3,
      sizePt: CHART_LABEL_PT,
      colorHex: LABEL_COLOR,
      align: 'left',
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

  for (const v of tickVals) {
    const off = valueOffset(v);
    if (horizontal) {
      const gx = x0 + off;
      polylines.push({
        points: [
          [gx, y0],
          [gx, y0 + plotH],
        ],
        strokeHex: GRID_COLOR,
        widthPt: 0.75,
      });
      labels.push({
        text: fmtVal(v),
        x: gx,
        y: y0 - CHART_LABEL_PT,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'center',
      });
    } else {
      const gy = y0 + off;
      polylines.push({
        points: [
          [x0, gy],
          [x0 + plotW, gy],
        ],
        strokeHex: GRID_COLOR,
        widthPt: 0.75,
      });
      labels.push({
        text: fmtVal(v),
        x: x0 - 3,
        y: gy - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: LABEL_COLOR,
        align: 'right',
      });
    }
  }

  const slot = (horizontal ? plotH : plotW) / nCats;
  for (let c = 0; c < nCats; c++) {
    const cat = chart.categories[c] ?? '';
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

  polylines.push({
    points: [
      [x0, y0],
      [x0, y0 + plotH],
    ],
    strokeHex: AXIS_COLOR,
    widthPt: 1,
  });
  polylines.push({
    points: [
      [x0, y0],
      [x0 + plotW, y0],
    ],
    strokeHex: AXIS_COLOR,
    widthPt: 1,
  });
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
    rects,
    polylines,
    labels,
  };
}

const pctLabel = (v: number): string => `${Math.round(v * 100)}%`;

// A datum's printed value (c:dLbls/showVal): integers as-is, else ≤2 decimals.
const fmtDataLabel = (v: number): string =>
  Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100);

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
  if (g === 'stacked') {
    const t = stackedTotals(chart, nCats);
    return { dataRange: [Math.min(0, t.min), Math.max(0, t.max)] };
  }
  return {};
}

// ─── bar / column chart (clustered, stacked, percentStacked) ────────────────
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

  if (stacked) {
    const groupPad = f.slot * 0.15;
    const barW = f.slot - 2 * groupPad;
    for (let c = 0; c < f.nCats; c++) {
      const along = (horizontal ? f.y0 : f.x0) + c * f.slot + groupPad;
      const denom = percent
        ? chart.series.reduce((a, s) => a + Math.abs(s.values[c] ?? 0), 0) || 1
        : 1;
      let cumPos = 0;
      let cumNeg = 0;
      for (let s = 0; s < chart.series.length; s++) {
        const series = chart.series[s]!;
        const v = (series.values[c] ?? 0) / denom; // fraction when percent, else raw
        const base = v >= 0 ? cumPos : cumNeg;
        const top = base + v;
        const o0 = f.valueOffset(base);
        const o1 = f.valueOffset(top);
        const lo = Math.min(o0, o1);
        const span = Math.abs(o1 - o0);
        const color = pointColor(series, c) ?? seriesColor(series, s);
        if (horizontal) f.rects.push({ x: f.x0 + lo, y: along, w: span, h: barW, fillHex: color });
        else f.rects.push({ x: along, y: f.y0 + lo, w: barW, h: span, fillHex: color });
        if (chart.showValues && span > CHART_LABEL_PT) {
          const raw = series.values[c] ?? 0;
          if (horizontal)
            f.labels.push(
              centeredLabel(fmtDataLabel(raw), f.x0 + lo + span / 2, along + barW * 0.3),
            );
          else
            f.labels.push(
              centeredLabel(fmtDataLabel(raw), along + barW / 2, f.y0 + lo + span / 2 - 3),
            );
        }
        if (v >= 0) cumPos = top;
        else cumNeg = top;
      }
    }
    return { rects: f.rects, polylines: f.polylines, wedges: [], labels: f.labels };
  }

  // clustered: series side by side within each category slot
  const nSer = Math.max(1, chart.series.length);
  const groupPad = f.slot * 0.15;
  const barW = (f.slot - 2 * groupPad) / nSer;
  for (let c = 0; c < f.nCats; c++) {
    const slotStart = (horizontal ? f.y0 : f.x0) + c * f.slot + groupPad;
    for (let s = 0; s < chart.series.length; s++) {
      const series = chart.series[s]!;
      const len = f.valueOffset(series.values[c] ?? 0) - f.zeroOffset; // signed from zero line
      const color = pointColor(series, c) ?? seriesColor(series, s);
      const along = slotStart + s * barW;
      if (horizontal) {
        const bx = f.x0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: bx, y: along, w: Math.abs(len), h: barW * 0.9, fillHex: color });
      } else {
        const by = f.y0 + f.zeroOffset + Math.min(0, len);
        f.rects.push({ x: along, y: by, w: barW * 0.9, h: Math.abs(len), fillHex: color });
      }
      if (chart.showValues) {
        const raw = series.values[c] ?? 0;
        const txt = fmtDataLabel(raw);
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
  return { rects: f.rects, polylines: f.polylines, wedges: [], labels: f.labels };
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
      polygons.push(areaBand(top, base, f, xAt, seriesColor(series, s)));
      for (let c = 0; c < nCats; c++) cum[c] = top[c]!;
    }
  } else {
    // standard: filled to baseline, back-to-front so series 0 stays on top
    const base = new Array<number>(nCats).fill(0);
    for (let s = chart.series.length - 1; s >= 0; s--) {
      const series = chart.series[s]!;
      const top = Array.from({ length: nCats }, (_, c) => series.values[c] ?? 0);
      polygons.push(areaBand(top, base, f, xAt, seriesColor(series, s)));
    }
  }
  return { rects: f.rects, polylines: f.polylines, wedges: [], labels: f.labels, polygons };
}

// ─── scatter chart (numeric X/Y) ────────────────────────────────────────────
export function buildScatterScene(
  chart: Chart,
  wPt: number,
  hPt: number,
  measure: MeasureText,
): ChartScene {
  const rects: Array<ChartRect> = [];
  const polylines: Array<ChartPolyline> = [];
  const labels: Array<ChartLabel> = [];

  const xs: Array<number> = [];
  const ys: Array<number> = [];
  for (const s of chart.series) {
    for (let i = 0; i < s.values.length; i++) {
      xs.push(s.xValues?.[i] ?? i);
      ys.push(s.values[i] ?? 0);
    }
  }
  if (xs.length === 0) return { rects, polylines, wedges: [], labels };

  // Both scatter axes auto-min — neither is forced through 0.
  const xScale = niceScale(Math.min(...xs), Math.max(...xs));
  const yScale = niceScale(Math.min(...ys), Math.max(...ys));
  const xTicks = ticks(xScale);
  const yTicks = ticks(yScale);

  let top = 4;
  if (chart.title) top += CHART_TITLE_PT * 1.6;
  const legendEntries: Array<LegendEntry> = chart.series
    .map((s, i) => ({ name: s.name ?? '', colorHex: seriesColor(s, i) }))
    .filter((e) => e.name !== '');
  const legend = layoutLegend(
    legendEntries,
    chart.hasLegend,
    chart.legendPos ?? 'b',
    wPt,
    hPt,
    measure,
  );
  const tickLabelW =
    Math.max(0, ...yTicks.map((v) => measure(formatTick(v, yScale.step), CHART_LABEL_PT))) + 4;
  const x0 = 4 + tickLabelW;
  const y0 = 4 + legend.bottomHeight + CHART_LABEL_PT * 1.6;
  const plotW = Math.max(1, wPt - 4 - legend.rightWidth - x0);
  const plotH = Math.max(1, hPt - top - y0);
  const xAt = (v: number): number => x0 + ((v - xScale.min) / (xScale.max - xScale.min)) * plotW;
  const yAt = (v: number): number => y0 + ((v - yScale.min) / (yScale.max - yScale.min)) * plotH;

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
  for (const v of yTicks) {
    const gy = yAt(v);
    polylines.push({
      points: [
        [x0, gy],
        [x0 + plotW, gy],
      ],
      strokeHex: GRID_COLOR,
      widthPt: 0.75,
    });
    labels.push({
      text: formatTick(v, yScale.step),
      x: x0 - 3,
      y: gy - CHART_LABEL_PT / 3,
      sizePt: CHART_LABEL_PT,
      colorHex: LABEL_COLOR,
      align: 'right',
    });
  }
  for (const v of xTicks) {
    const gx = xAt(v);
    polylines.push({
      points: [
        [gx, y0],
        [gx, y0 + plotH],
      ],
      strokeHex: GRID_COLOR,
      widthPt: 0.75,
    });
    labels.push({
      text: formatTick(v, xScale.step),
      x: gx,
      y: y0 - CHART_LABEL_PT,
      sizePt: CHART_LABEL_PT,
      colorHex: LABEL_COLOR,
      align: 'center',
    });
  }
  polylines.push({
    points: [
      [x0, y0],
      [x0, y0 + plotH],
    ],
    strokeHex: AXIS_COLOR,
    widthPt: 1,
  });
  polylines.push({
    points: [
      [x0, y0],
      [x0 + plotW, y0],
    ],
    strokeHex: AXIS_COLOR,
    widthPt: 1,
  });

  for (let s = 0; s < chart.series.length; s++) {
    const series = chart.series[s]!;
    const color = seriesColor(series, s);
    for (let i = 0; i < series.values.length; i++) {
      const px = xAt(series.xValues?.[i] ?? i);
      const py = yAt(series.values[i] ?? 0);
      rects.push({ x: px - 2, y: py - 2, w: 4, h: 4, fillHex: color });
    }
  }
  legend.emit(rects, labels);
  return { rects, polylines, wedges: [], labels };
}

// ─── line chart ───────────────────────────────────────────────────────────────
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
  const f = buildFrame(chart, wPt, hPt, measure, false, { dataRange: range });
  for (let s = 0; s < chart.series.length; s++) {
    const series = chart.series[s]!;
    const color = seriesColor(series, s);
    const pts: Array<readonly [number, number]> = [];
    for (let c = 0; c < f.nCats; c++) {
      const x = f.x0 + c * f.slot + f.slot / 2;
      const y = f.y0 + f.valueOffset(series.values[c] ?? 0);
      pts.push([x, y]);
      if (chart.showValues)
        f.labels.push(centeredLabel(fmtDataLabel(series.values[c] ?? 0), x, y + 3));
    }
    if (pts.length >= 2) {
      f.polylines.push({ points: pts, strokeHex: color, widthPt: 1.5 });
    } else if (pts.length === 1) {
      // A single data point: a small marker so it is visible.
      const [px, py] = pts[0]!;
      f.rects.push({ x: px - 1.5, y: py - 1.5, w: 3, h: 3, fillHex: color });
    }
  }
  return { rects: f.rects, polylines: f.polylines, wedges: [], labels: f.labels };
}

// ─── pie chart ──────────────────────────────────────────────────────────────
const sliceColor = (series: ChartSeries, i: number): string =>
  pointColor(series, i) ?? SERIES_COLORS[i % SERIES_COLORS.length]!;

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
  // Pie legend lists categories (each in its slice colour).
  const legendEntries: Array<LegendEntry> = chart.categories.map((c, i) => ({
    name: c,
    colorHex: sliceColor(series, i),
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
      fillHex: sliceColor(series, i),
      strokeHex: 'FFFFFF',
    });
    const mid = ang + sweep / 2;
    const pct = Math.round((v / total) * 100);
    if (pct >= 5) {
      labels.push({
        text: `${pct}%`,
        x: cx + Math.cos(mid) * labelR,
        y: cy + Math.sin(mid) * labelR - CHART_LABEL_PT / 3,
        sizePt: CHART_LABEL_PT,
        colorHex: chart.doughnut ? LABEL_COLOR : 'FFFFFF',
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
}
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
): LegendLayout {
  if (!hasLegend || entries.length === 0) {
    return { rightWidth: 0, bottomHeight: 0, emit: () => {} };
  }
  const sw = CHART_LABEL_PT; // swatch size
  const gap = 4;
  const entryW = (e: LegendEntry): number => sw + 3 + measure(e.name, CHART_LABEL_PT) + gap * 2;

  if (pos === 'r' || pos === 'l') {
    const colW = Math.max(...entries.map(entryW));
    return {
      rightWidth: pos === 'r' ? colW : 0,
      bottomHeight: 0,
      emit: (rects, labels) => {
        const lx = pos === 'r' ? wPt - colW + gap : gap;
        let ly = hPt / 2 + (entries.length * (sw + 4)) / 2 - sw;
        for (const e of entries) {
          rects.push({ x: lx, y: ly, w: sw, h: sw, fillHex: e.colorHex });
          labels.push({
            text: e.name,
            x: lx + sw + 3,
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
        rects.push({ x: lx, y: ly, w: sw, h: sw, fillHex: e.colorHex });
        labels.push({
          text: e.name,
          x: lx + sw + 3,
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

// Dispatch by chart type. Returns null for an unrenderable type (the renderer
// then reserves the box with a light border).
export function buildChartScene(
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
