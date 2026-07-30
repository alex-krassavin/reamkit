import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { Chart } from '@/core/document-model';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';
import {
  buildAreaScene,
  buildBarScene,
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
  it('produces round ticks covering the data', () => {
    expect(niceScale(0, 95)).toEqual({ min: 0, max: 100, step: 20 });
    expect(niceScale(0, 25)).toEqual({ min: 0, max: 30, step: 10 });
    expect(niceScale(0, 8)).toEqual({ min: 0, max: 8, step: 2 });
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
    const chart: Chart = {
      ...barChart('col'),
      title: undefined,
      categories: ['a', 'b', 'c'],
      series: [{ name: 'n', values: [1, 2, 3], pointLabels: [{ idx: 0, text: 'Ж' }] }],
      catAxisTitle: 'Щ',
      valAxisTitle: 'Э',
    };
    const laid = layoutStyledDocument(
      [
        {
          kind: 'chart',
          chart: { chartRelId: 'c1', width: 300, height: 200, paragraphProperties: {} },
        },
      ],
      { registry, charts: new Map([['c1', chart]]) },
    );
    const res = [...laid.fontResources.values()][0]!;
    for (const ch of ['Щ', 'Э', 'Ж']) {
      const gid = res.parsed.glyphForCodepoint(ch.codePointAt(0)!);
      expect(gid).toBeGreaterThan(0);
      expect(res.gids.has(gid)).toBe(true);
    }
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

  it('rescales the value axis to summed totals when stacked', () => {
    // Clustered tops out at the max single value (25 → tick 30). Stacked tops at
    // the max category sum (cat2 = 15+25 = 40 → tick 40).
    const clustered = buildBarScene({ ...barChart('col'), grouping: 'clustered' }, W, H, measure);
    const stacked = buildBarScene({ ...barChart('col'), grouping: 'stacked' }, W, H, measure);
    expect(clustered.labels.some((l) => l.text === '30')).toBe(true);
    expect(clustered.labels.some((l) => l.text === '40')).toBe(false);
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
