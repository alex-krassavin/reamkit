import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { Chart } from '@/core/document-model';
import type { ChartScene } from '@/core/drawingml/chart-geometry';
import { FontRegistry } from '@/core/font';
import { pt } from '@/core/ir';
import { layoutStyledDocument } from '@/layout/styled-layout';
import {
  buildAreaScene,
  buildBarScene,
  buildChartScene,
  buildLineScene,
  buildPieScene,
  buildScatterScene,
  formatTick,
  niceScale,
} from '@/core/drawingml/chart-geometry';

// Crude monospace-ish measure for layout tests.
const measure = (t: string, sz: number): number => t.length * sz * 0.5;

const W = 400;
const H = 240;

function inBounds(items: ReadonlyArray<{ x: number; y: number; w?: number; h?: number }>): boolean {
  return items.every(
    (i) => i.x >= -1 && i.y >= -1 && i.x + (i.w ?? 0) <= W + 1 && i.y + (i.h ?? 0) <= H + 1,
  );
}

const barChart = (barDir: 'col' | 'bar'): Chart => ({
  type: 'bar',
  barDir,
  grouping: 'clustered',
  title: 'T',
  categories: ['A', 'B', 'C'],
  hasLegend: true,
  legendPos: 'b',
  series: [
    { name: 'S1', values: [10, 20, 15], colorHex: '4472C4' },
    { name: 'S2', values: [12, 18, 25], colorHex: 'ED7D31' },
  ],
});

describe('niceScale', () => {
  it("scales an axis as Excel does: the ends from the data's spread, the step 1-2-5", () => {
    // Measured against Excel's own PDF of eight column charts (2026-10-01).
    // The near end at zero unless the spread is under a sixth of the far one,
    // 5% of the spread beyond the data, then the smallest 1/2/5 × 10ⁿ step
    // that leaves at most ten intervals.
    expect(niceScale(2336, 3750)).toEqual({ min: 0, max: 4000, step: 500 });
    expect(niceScale(-1, 1.2)).toEqual({ min: -1.5, max: 1.5, step: 0.5 });
    expect(niceScale(0.3, 4.7)).toEqual({ min: 0, max: 5, step: 0.5 });
    expect(niceScale(120, 950)).toEqual({ min: 0, max: 1000, step: 100 });
    expect(niceScale(93, 97)).toEqual({ min: 91, max: 98, step: 1 });
    expect(niceScale(0.012, 0.047)).toEqual({ min: 0, max: 0.05, step: 0.005 });
    // 57362.xlsx's 12-value bar: room above it, the axis labelled to 14.
    expect(niceScale(0, 12)).toEqual({ min: 0, max: 14, step: 2 });
  });

  it('steps more coarsely only where the labels would not fit', () => {
    expect(niceScale(0, 3750, 4)).toEqual({ min: 0, max: 4000, step: 1000 });
  });

  it('handles a flat range', () => {
    const s = niceScale(5, 5);
    expect(s.max).toBeGreaterThan(s.min);
  });
});

describe('formatTick', () => {
  it('keeps integers integral and trims fractional zeros', () => {
    expect(formatTick(100, 20)).toBe('100');
    expect(formatTick(0, 20)).toBe('0');
    expect(formatTick(0.5, 0.5)).toBe('0.5');
    // General shows no trailing zero, whatever the step's decimals.
    expect(formatTick(0.05, 0.005)).toBe('0.05');
    expect(formatTick(0.045, 0.005)).toBe('0.045');
    expect(formatTick(1, 0.5)).toBe('1');
  });
});

describe('font subsetting reaches every string a chart draws', () => {
  it('keeps the glyphs an axis title and a typed data label need', () => {
    // The subset is built from the strings the document draws. Axis titles and
    // author-typed data labels were left out of that walk, so a character
    // appearing ONLY there was dropped from the embedded font and drew blank —
    // with the text layer still claiming it. shape-macro-ext-ref.xlsx printed
    // "Translation X [mm]" as "Translation    mm".
    const registry = FontRegistry.fromBytes({
      regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
    });
    const { title: _title, ...noTitle } = barChart('col');
    const chart: Chart = {
      ...noTitle,
      categories: ['a', 'b', 'c'],
      series: [{ name: 'n', values: [1, 2, 3], pointLabels: [{ idx: 0, text: 'Ж' }] }],
      catAxisTitle: 'Щ',
      valAxisTitle: 'Э',
    };
    const laid = layoutStyledDocument(
      [
        {
          kind: 'chart',
          chart: { chartRelId: 'c1', width: pt(300), height: pt(200), paragraphProperties: {} },
        },
      ],
      {
        registry,
        charts: new Map([['c1', chart]]),
        styles: {
          defaultRunProperties: {},
          defaultParagraphProperties: {},
          styles: new Map(),
        },
      },
    );
    const res = [...laid.fontResources.values()][0]!;
    for (const ch of ['Щ', 'Э', 'Ж']) {
      const gid = res.parsed.glyphForCodepoint(ch.codePointAt(0)!);
      expect(gid).toBeGreaterThan(0);
      expect(res.gids.has(gid)).toBe(true);
    }
  });
});

describe('a crowded category axis', () => {
  it('labels every Nth category so they do not collide', () => {
    // 47813.xlsx plots 1700 points. Drawing every label turned the axis into a
    // solid black bar under the plot; Excel and Calc both thin it.
    const many: Chart = {
      ...barChart('col'),
      categories: Array.from({ length: 400 }, (_, i) => String(i + 1)),
      series: [{ name: 'S', values: Array.from({ length: 400 }, () => 1) }],
    };
    const drawn = buildBarScene(many, W, H, measure).labels.filter((l) =>
      /^\d+$/.test(l.text),
    ).length;
    // Far fewer than 400, and not none.
    expect(drawn).toBeGreaterThan(2);
    expect(drawn).toBeLessThan(120);
    // A short axis still labels every category.
    const few = buildBarScene(barChart('col'), W, H, measure).labels.map((l) => l.text);
    expect(few).toContain('A');
    expect(few).toContain('B');
    expect(few).toContain('C');
  });
});

describe('how many ticks an axis carries', () => {
  // A tall axis takes more of them — a fixed six drew 0/100/200/300 down a plot
  // where both references fit 0/50/…/300 — but not one per 32pt: at that rate a
  // 427pt chart asked for ten, and Excel labels the same 0…5 data 0/1/…/6 where
  // we drew half-steps (chart-texture-bg.pptx).
  const axisLabels = (h: number, values: ReadonlyArray<number>): Array<string> =>
    buildBarScene(
      { ...barChart('col'), series: [{ name: 'S', values }], categories: ['A', 'B', 'C'] },
      W,
      h,
      measure,
    )
      .labels.map((l) => l.text)
      .filter((t) => /^[\d.]+$/.test(t));

  it('gives a tall chart whole steps where the data is small', () => {
    const tall = axisLabels(427, [1.8, 3, 5]);
    expect(tall).toContain('6');
    expect(tall.some((t) => t.includes('.'))).toBe(false); // no 0.5 half-steps
  });

  it('still fills a tall axis when the data is large', () => {
    // The case the height-dependence exists for: 0/50/…/300, not 0/100/200/300.
    expect(axisLabels(400, [120, 250, 300])).toContain('50');
  });

  it('thins out on a short axis', () => {
    expect(axisLabels(90, [120, 250, 300]).length).toBeLessThanOrEqual(6);
  });
});

describe('the chart-space frame (§21.2.2.198)', () => {
  it('draws it first, behind everything else', () => {
    const framed: Chart = { ...barChart('col'), frameFillHex: 'FFFFFF', frameLineHex: 'D9D9D9' };
    const scene = buildChartScene(framed, W, H, measure)!;
    // Its own layer, drawn under the gridlines and the data alike.
    expect(scene.background).toMatchObject({
      x: 0,
      y: 0,
      w: W,
      h: H,
      fillHex: 'FFFFFF',
      strokeHex: 'D9D9D9',
    });
    // A chart without one keeps exactly the rects it had, and no background.
    const plain = buildChartScene(barChart('col'), W, H, measure)!;
    expect(plain.background).toBeUndefined();
    expect(plain.rects).toHaveLength(scene.rects.length);
  });
});

describe('where the category axis crosses (§21.2.2.33, §21.2.2.207)', () => {
  // Values both sides of zero: the category axis lies on the zero line.
  const signed: Chart = {
    ...barChart('col'),
    title: '',
    hasLegend: false,
    series: [{ name: 'S1', values: [-10, 20, 15], colorHex: '4472C4' }],
  };
  const labelY = (scene: ChartScene, text: string): number =>
    scene.labels.find((l) => l.text === text)!.y;
  // The zero line: the top of the bar below it.
  const zeroY = (scene: ChartScene): number =>
    Math.max(...scene.rects.filter((r) => r.fillHex === '4472C4').map((r) => r.y));

  it('labels the categories beside it, on the zero line', () => {
    const scene = buildBarScene(signed, W, H, measure);
    const zero = zeroY(scene);
    expect(labelY(scene, 'A')).toBeLessThan(zero);
    expect(labelY(scene, 'A')).toBeGreaterThan(zero - 2 * 9);
    // …and draws the axis there, not along the plot's foot.
    const flat = scene.polylines.filter(
      (p) => p.points.length === 2 && p.points[0]![1] === p.points[1]![1],
    );
    expect(flat.some((p) => Math.abs(p.points[0]![1] - zero) < 0.01)).toBe(true);
  });

  it('labels them at the low end when the file says so', () => {
    const low = buildBarScene({ ...signed, catTickLabelPos: 'low' }, W, H, measure);
    expect(labelY(low, 'A')).toBeLessThan(zeroY(low) - 2 * 9);
    expect(
      buildBarScene({ ...signed, catTickLabelPos: 'none' }, W, H, measure).labels,
    ).not.toContainEqual(expect.objectContaining({ text: 'A' }));
  });

  it('grows the bars from where it crosses', () => {
    // Crossing at the minimum, the bars all stand on the plot's foot.
    const scene = buildBarScene({ ...signed, catAxisCrosses: 'min' }, W, H, measure);
    const bars = scene.rects.filter((r) => r.fillHex === '4472C4');
    const foot = Math.min(...bars.map((r) => r.y));
    for (const r of bars) expect(r.y).toBeCloseTo(foot, 5);
  });

  it("steps a crowded axis by the labels it draws, a point's index among them", () => {
    // No categories: the axis is labelled 1…716, and the labels must not run
    // into one another.
    const many: Chart = {
      ...signed,
      categories: [],
      series: [{ name: 'S1', values: Array.from({ length: 716 }, (_, i) => Math.sin(i / 50)) }],
    };
    const shown = buildBarScene(many, W, H, measure)
      .labels.filter((l) => /^\d+$/.test(l.text) && l.align === 'center')
      .sort((a, b) => a.x - b.x);
    expect(shown.length).toBeGreaterThan(3);
    for (let i = 1; i < shown.length; i++) {
      const gap = shown[i]!.x - shown[i - 1]!.x;
      expect(gap).toBeGreaterThanOrEqual(
        (measure(shown[i]!.text, 9) + measure(shown[i - 1]!.text, 9)) / 2,
      );
    }
  });
});

describe('a chart title longer than the chart is wide', () => {
  it('wraps at its spaces, each line within four fifths of the width', () => {
    const title = 'Ranking of Washington Counties on Days per Patient (ALOS) in 2015';
    for (const chart of [
      { ...barChart('bar'), title },
      { ...barChart('col'), title },
      { ...barChart('col'), type: 'pie' as const, title },
    ]) {
      const scene = buildChartScene(chart, W, H, measure)!;
      const lines = scene.labels.filter((l) => title.includes(l.text) && l.sizePt > 10);
      expect(lines.length).toBeGreaterThan(1);
      expect(lines.map((l) => l.text).join(' ')).toBe(title);
      for (const l of lines) expect(measure(l.text, l.sizePt)).toBeLessThanOrEqual(W * 0.8);
      // Top line first, each below the one before.
      for (let i = 1; i < lines.length; i++) expect(lines[i]!.y).toBeLessThan(lines[i - 1]!.y);
    }
  });

  it('pushes the plot down by the lines it takes', () => {
    const short = buildBarScene(barChart('col'), W, H, measure);
    const long = buildBarScene(
      { ...barChart('col'), title: 'A title far too long to stand on a single line of this chart' },
      W,
      H,
      measure,
    );
    const top = (s: ChartScene): number =>
      Math.max(...s.rects.filter((r) => r.fillHex === '4472C4').map((r) => r.y + r.h));
    expect(top(long)).toBeLessThan(top(short));
  });
});

describe('a value axis the author fixed (§21.2.2.157)', () => {
  it('draws to the declared max, not to the data', () => {
    // Every value here is 0. Left to the data the axis would run 0…1; the
    // author pinned it at 300 and every reader honours that.
    const flat: Chart = {
      ...barChart('col'),
      series: [{ name: 'S1', values: [0, 0, 0], colorHex: '4472C4' }],
      valAxisMax: 300,
    };
    const ticks = buildBarScene(flat, W, H, measure)
      .labels.map((l) => l.text)
      .filter((t) => /^\d+$/.test(t));
    expect(ticks).toContain('300');
    expect(ticks).toContain('0');

    // Without it, the same chart scales to its (degenerate) data.
    const auto = { ...flat };
    delete (auto as { valAxisMax?: number }).valAxisMax;
    expect(
      buildBarScene(auto, W, H, measure)
        .labels.map((l) => l.text)
        .filter((t) => /^\d+$/.test(t)),
    ).not.toContain('300');
  });
});

describe('buildBarScene', () => {
  it('emits one bar per (series, category) plus legend swatches, in bounds', () => {
    const scene = buildBarScene(barChart('col'), W, H, measure);
    // 2 series × 3 categories = 6 bars + 2 legend swatches.
    expect(scene.rects).toHaveLength(8);
    expect(inBounds(scene.rects)).toBe(true);
    // Bars carry the series fill colours.
    const fills = scene.rects.map((r) => r.fillHex);
    expect(fills).toContain('4472C4');
    expect(fills).toContain('ED7D31');
    // Axis lines + gridlines present.
    expect(scene.polylines.length).toBeGreaterThanOrEqual(2);
    // Title + category + tick + legend labels present.
    expect(scene.labels.some((l) => l.text === 'T')).toBe(true);
    expect(scene.labels.some((l) => l.text === 'A')).toBe(true);
  });

  it('lays out a horizontal bar chart within bounds too', () => {
    const scene = buildBarScene(barChart('bar'), W, H, measure);
    expect(scene.rects).toHaveLength(8);
    expect(inBounds(scene.rects)).toBe(true);
  });

  it("keeps a bar chart's category names and its last value inside the frame", () => {
    // The names stand right-aligned left of the plot, so the left of it is
    // sized by them, not by the value ticks that run along the foot; and the
    // last tick's label, centred under the plot's end, is given room too.
    const named: Chart = {
      ...barChart('bar'),
      categories: ['Grays Harbor', 'Walla Walla', 'Pend Oreille'],
      hasLegend: false,
    };
    const scene = buildBarScene(named, W, H, measure);
    for (const name of named.categories) {
      const label = scene.labels.find((l) => l.text === name)!;
      expect(label.align).toBe('right');
      expect(label.x - measure(name, label.sizePt)).toBeGreaterThanOrEqual(0);
    }
    const ticks = scene.labels.filter((l) => /^\d+$/.test(l.text));
    const last = ticks.reduce((a, b) => (b.x > a.x ? b : a));
    expect(last.x + measure(last.text, last.sizePt) / 2).toBeLessThanOrEqual(W);
  });

  it('runs the categories the other way when the axis says maxMin', () => {
    // §21.2.2.134 — dataValidationTableRange.xlsx ranks 38 counties and writes
    // `maxMin` so the ranking reads top-down; plotted in file order it comes out
    // upside down. A horizontal bar chart's first category is at the BOTTOM by
    // default, so reversing puts it at the top.
    const chart = { ...barChart('bar'), series: [barChart('bar').series[0]!], hasLegend: false };
    // Bars of the one series, top of the plot downwards.
    const barLengths = (c: Chart): Array<number> =>
      (buildChartScene(c, W, H, measure)?.rects ?? [])
        .filter((r) => r.fillHex === '4472C4')
        .sort((a, b) => b.y - a.y)
        .map((r) => Math.round(r.w));
    const labelOrder = (c: Chart): Array<string> =>
      (buildChartScene(c, W, H, measure)?.labels ?? [])
        .filter((l) => ['A', 'B', 'C'].includes(l.text))
        .sort((a, b) => b.y - a.y)
        .map((l) => l.text);

    expect(labelOrder(chart)).toEqual(['C', 'B', 'A']);
    expect(labelOrder({ ...chart, catAxisReversed: true })).toEqual(['A', 'B', 'C']);
    // The bars follow their categories rather than staying put.
    expect(barLengths({ ...chart, catAxisReversed: true })).toEqual(
      [...barLengths(chart)].reverse(),
    );
  });

  it('fills the whole slot when the file asks for no gap', () => {
    // §21.2.2.75 `gapWidth="0"` — both references draw bars that touch.
    const scene = buildBarScene(
      {
        ...barChart('col'),
        series: [barChart('col').series[0]!],
        gapPercent: 0,
        hasLegend: false,
      },
      W,
      H,
      measure,
    );
    const bars = scene.rects.filter((r) => r.fillHex === '4472C4');
    expect(bars).toHaveLength(3);
    const xs = bars.map((b) => b.x).sort((a, b) => a - b);
    // Each bar starts exactly where the one before it ends.
    expect(xs[1]! - xs[0]!).toBeCloseTo(bars[0]!.w, 6);
  });

  it('rescales the value axis to summed totals when stacked', () => {
    // Clustered tops out at the max single value (25). Stacked tops at the max
    // category sum (cat2 = 15+25 = 40).
    const clustered = buildBarScene({ ...barChart('col'), grouping: 'clustered' }, W, H, measure);
    const stacked = buildBarScene({ ...barChart('col'), grouping: 'stacked' }, W, H, measure);
    const topTick = (s: ChartScene): number =>
      Math.max(...s.labels.filter((l) => /^\d+$/u.test(l.text)).map((l) => Number(l.text)));
    // The clustered axis reaches over the tallest BAR and no further; the
    // stacked one has to reach over the tallest SUM.
    expect(topTick(clustered)).toBe(30);
    expect(clustered.labels.some((l) => l.text === '40')).toBe(false);
    expect(topTick(stacked)).toBeGreaterThanOrEqual(40);
    expect(stacked.labels.some((l) => l.text === '40')).toBe(true);
    expect(inBounds(stacked.rects)).toBe(true);
  });

  it('labels the value axis as a percentage when percentStacked', () => {
    const scene = buildBarScene({ ...barChart('col'), grouping: 'percentStacked' }, W, H, measure);
    expect(scene.labels.some((l) => l.text === '100%')).toBe(true);
    expect(inBounds(scene.rects)).toBe(true);
  });
});

const pointsInBounds = (
  polys: ReadonlyArray<{ points: ReadonlyArray<readonly [number, number]> }>,
) => polys.every((p) => p.points.every(([x, y]) => x >= -1 && y >= -1 && x <= W + 1 && y <= H + 1));

describe('chart polish (#57)', () => {
  it('prints datum values as labels when showValues is set', () => {
    // 12 and 25 are data values but not axis ticks (ticks: 0/10/20/30).
    const plain = buildBarScene(barChart('col'), W, H, measure);
    const labelled = buildBarScene({ ...barChart('col'), showValues: true }, W, H, measure);
    expect(plain.labels.some((l) => l.text === '12')).toBe(false);
    expect(labelled.labels.some((l) => l.text === '12')).toBe(true);
    expect(labelled.labels.some((l) => l.text === '25')).toBe(true);
  });

  it('emits axis-title labels', () => {
    const scene = buildBarScene(
      { ...barChart('col'), catAxisTitle: 'Quarter', valAxisTitle: 'Sales' },
      W,
      H,
      measure,
    );
    expect(scene.labels.some((l) => l.text === 'Quarter')).toBe(true);
    expect(scene.labels.some((l) => l.text === 'Sales')).toBe(true);
  });

  it('auto-mins the line value axis away from 0 when data is far from it', () => {
    const line: Chart = {
      type: 'line',
      categories: ['A', 'B', 'C'],
      hasLegend: false,
      series: [{ name: 'S', values: [100, 110, 105], colorHex: '4472C4' }],
    };
    const scene = buildLineScene(line, W, H, measure);
    expect(scene.labels.some((l) => l.text === '100')).toBe(true);
    expect(scene.labels.some((l) => l.text === '0')).toBe(false);
  });
});

describe('buildAreaScene', () => {
  it('emits one filled band per series, in bounds', () => {
    const scene = buildAreaScene({ ...barChart('col'), type: 'area' }, W, H, measure);
    expect(scene.polygons).toHaveLength(2);
    expect(pointsInBounds(scene.polygons!)).toBe(true);
    expect(scene.polygons!.every((p) => p.fillHex.length === 6)).toBe(true);
  });

  it('stacks bands and pins the percent axis at 100%', () => {
    const scene = buildAreaScene(
      { ...barChart('col'), type: 'area', grouping: 'percentStacked' },
      W,
      H,
      measure,
    );
    expect(scene.polygons).toHaveLength(2);
    expect(scene.labels.some((l) => l.text === '100%')).toBe(true);
  });
});

describe('buildScatterScene', () => {
  const scatter: Chart = {
    type: 'scatter',
    categories: [],
    hasLegend: false,
    series: [{ name: 'S', values: [2, 4, 8, 6], xValues: [1, 2, 3, 4], colorHex: '4472C4' }],
  };

  it('plots one marker per (x,y) point, in bounds', () => {
    const scene = buildScatterScene(scatter, W, H, measure);
    expect(scene.rects).toHaveLength(4); // 4 markers, no legend swatches
    expect(inBounds(scene.rects)).toBe(true);
    expect(scene.rects.every((r) => r.fillHex === '4472C4')).toBe(true);
    // Both axes drawn (≥ 2 axis lines + gridlines).
    expect(scene.polylines.length).toBeGreaterThanOrEqual(2);
  });
});

describe('buildLineScene', () => {
  const lineChart: Chart = {
    type: 'line',
    categories: ['A', 'B', 'C'],
    hasLegend: false,
    series: [
      { name: 'S1', values: [1, 2, 3], colorHex: '4472C4' },
      { name: 'S2', values: [3, 1, 2], colorHex: 'ED7D31' },
    ],
  };

  it('emits one polyline per series, each with a point per category', () => {
    const scene = buildLineScene(lineChart, W, H, measure);
    const seriesLines = scene.polylines.filter((p) => p.widthPt === 1.5);
    expect(seriesLines).toHaveLength(2);
    expect(seriesLines[0]!.points).toHaveLength(3);
    expect(seriesLines[0]!.strokeHex).toBe('4472C4');
    expect(seriesLines[1]!.strokeHex).toBe('ED7D31');
    // All points within the box.
    for (const pl of seriesLines) {
      for (const [x, y] of pl.points) {
        expect(x).toBeGreaterThanOrEqual(-1);
        expect(x).toBeLessThanOrEqual(W + 1);
        expect(y).toBeGreaterThanOrEqual(-1);
        expect(y).toBeLessThanOrEqual(H + 1);
      }
    }
  });
});

describe('buildPieScene', () => {
  const pie: Chart = {
    type: 'pie',
    categories: ['A', 'B', 'C', 'D'],
    hasLegend: true,
    legendPos: 'r',
    series: [{ values: [40, 30, 20, 10] }],
  };

  it('emits one wedge per slice, sweeping a full clockwise turn', () => {
    const scene = buildPieScene(pie, W, H, measure);
    expect(scene.wedges).toHaveLength(4);
    const total = scene.wedges.reduce((a, w) => a + w.sweepRad, 0);
    expect(total).toBeCloseTo(-2 * Math.PI, 5); // clockwise
    expect(scene.wedges[0]!.sweepRad).toBeCloseTo(-0.4 * 2 * Math.PI, 5); // 40%
    // Slice colours cycle the accent palette.
    expect(scene.wedges[0]!.fillHex).toBe('4472C4');
    expect(scene.wedges[1]!.fillHex).toBe('ED7D31');
  });

  it('produces no wedges when the total is zero', () => {
    const empty: Chart = { ...pie, series: [{ values: [0, 0] }] };
    expect(buildPieScene(empty, W, H, measure).wedges).toHaveLength(0);
  });

  it('punches a central white hole for a doughnut', () => {
    const plain = buildPieScene(pie, W, H, measure);
    const ring = buildPieScene({ ...pie, doughnut: true }, W, H, measure);
    // The hole is an extra full-circle (|sweep| ≈ 2π) white disc.
    const isHole = (w: { sweepRad: number; fillHex: string }) =>
      Math.abs(w.sweepRad) > 2 * Math.PI - 1e-6 && w.fillHex === 'FFFFFF';
    expect(plain.wedges.some(isHole)).toBe(false);
    expect(ring.wedges.some(isHole)).toBe(true);
    expect(ring.wedges).toHaveLength(5); // 4 slices + hole
  });
});
