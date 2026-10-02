// §21.2.2.16 `c:bar3DChart` — bars drawn as boxes standing in a box of walls,
// as Excel's own PDF of its probe charts draws them (2026-10-02): four views,
// three depths, stacks, values below zero, deleted axes.

import { describe, expect, it } from 'vitest';

import type { Chart } from '@/core/document-model';
import type { ChartPolygon, ChartScene } from '@/core/drawingml/chart-geometry';
import { buildChartScene } from '@/core/drawingml/chart-geometry';
import { parseChart } from '@/core/drawingml/chart-parser';
import { defaultColorResolver } from '@/core/drawingml/colors';

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const measure = (t: string, sz: number): number => t.length * sz * 0.5;
const enc = new TextEncoder();
const rad = (deg: number): number => (deg * Math.PI) / 180;

const BAR3D = (view: string, group = '', walls = ''): string =>
  `<c:chartSpace ${C_NS}><c:chart>${view}${walls}<c:plotArea><c:bar3DChart>` +
  '<c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>' +
  '<c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt>' +
  '<c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:val>' +
  `</c:ser>${group}<c:axId val="1"/><c:axId val="2"/><c:axId val="0"/></c:bar3DChart>` +
  '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
  '<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>' +
  '</c:plotArea></c:chart></c:chartSpace>';

/** Excel's probe: three columns of one series, seen as `view` says. */
const columns = (
  view: Partial<NonNullable<Chart['bar3D']>> = {},
  extra: Partial<Chart> = {},
): Chart => ({
  type: 'bar',
  barDir: 'col',
  grouping: 'clustered',
  categories: ['A', 'B', 'C'],
  hasLegend: false,
  series: [{ name: 'S', values: [10, 20, 15], colorHex: '156082' }],
  bar3D: { rotX: 15, rotY: 20, depthPercent: 100, gapDepth: 150, ...view },
  ...extra,
});

const scene = (chart: Chart): ChartScene => buildChartScene(chart, 400, 260, measure)!;
const fronts = (s: ChartScene) => s.rects.filter((r) => r.fillHex === '156082');
const shaded = (s: ChartScene, hex: string): Array<ChartPolygon> =>
  (s.polygons ?? []).filter((pg) => pg.fillHex === hex);
const texts = (s: ChartScene): Array<string> => s.labels.map((l) => l.text);

describe('reading a 3-D bar chart (§21.2.2.16)', () => {
  it('reads its view, depth and the room around its bars', () => {
    const chart = parseChart(
      enc.encode(
        BAR3D(
          '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="300"/><c:rAngAx val="1"/></c:view3D>',
          '<c:gapDepth val="80"/>',
        ),
      ),
      defaultColorResolver,
    )!;
    expect(chart.bar3D).toEqual({ rotX: 15, rotY: 20, depthPercent: 300, gapDepth: 80 });
  });

  it("takes the schema's view, depth and room where the part states none", () => {
    // aascu 5864.pptx's view says only that its axes are square.
    const chart = parseChart(
      enc.encode(BAR3D('<c:view3D><c:rAngAx val="1"/></c:view3D>')),
      defaultColorResolver,
    )!;
    expect(chart.bar3D).toEqual({ rotX: 0, rotY: 0, depthPercent: 100, gapDepth: 150 });
    // A flat bar chart has none.
    const flat = BAR3D('')
      .replace(/bar3DChart/g, 'barChart')
      .replace('<c:axId val="0"/>', '');
    expect(parseChart(enc.encode(flat), defaultColorResolver)!.bar3D).toBeUndefined();
  });

  it('reads the floor and walls as the chart fills and rules them', () => {
    // tdf128207.docx clears them all.
    const walls =
      '<c:floor><c:thickness val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:floor>' +
      '<c:backWall><c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr></c:backWall>' +
      '<c:sideWall><c:thickness val="0"/></c:sideWall>';
    const chart = parseChart(
      enc.encode(BAR3D('<c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>', '', walls)),
      defaultColorResolver,
    )!;
    expect(chart.floor).toEqual({ line: { none: true } });
    expect(chart.backWall).toEqual({ fillHex: 'F2F2F2' });
    expect(chart.sideWall).toBeUndefined();
  });
});

describe('drawing a 3-D bar chart as Excel does', () => {
  it('stands each bar a box as deep as it is wide, its top and side shaded', () => {
    // Excel's PDF: depth runs sin rotY across and sin rotX up; the top 0.76 of
    // the bar's colour, the right side 0.635.
    const s = scene(columns());
    expect(fronts(s)).toHaveLength(3);
    const tops = shaded(s, '104963');
    const sides = shaded(s, '0D3D53');
    expect(tops).toHaveLength(3);
    expect(sides).toHaveLength(3);
    const w = fronts(s)[0]!.w;
    const top = tops[0]!.points;
    expect(top[3]![0] - top[0]![0]).toBeCloseTo(w * Math.sin(rad(20)), 6);
    expect(top[3]![1] - top[0]![1]).toBeCloseTo(w * Math.sin(rad(15)), 6);
  });

  it('builds the box two and a half bars deep, the bars standing three quarters of one back', () => {
    // `gapDepth="150"`: the box is the bar's depth and 150% of it; the bar's
    // front stands half that room back from the box's.
    const s = scene(columns());
    const w = fronts(s)[0]!.w;
    const grid = s.gridlines ?? [];
    expect(grid.length).toBeGreaterThan(0);
    for (const g of grid) {
      // Across the side wall, then along the back wall.
      expect(g.points).toHaveLength(3);
      expect(g.points[1]![0] - g.points[0]![0]).toBeCloseTo(2.5 * w * Math.sin(rad(20)), 6);
      expect(g.points[1]![1] - g.points[0]![1]).toBeCloseTo(2.5 * w * Math.sin(rad(15)), 6);
      expect(g.points[2]![1]).toBeCloseTo(g.points[1]![1], 6);
    }
    const floor = (s.walls ?? [])[0]!;
    expect(floor).toMatchObject({ strokeHex: 'D9D9D9' });
    expect(floor.fillHex).toBeUndefined();
    // The bars' feet stand back from the floor's front edge, by 0.75 of a depth.
    expect(fronts(s)[0]!.y - floor.points[0]![1]).toBeCloseTo(0.75 * w * Math.sin(rad(15)), 6);
  });

  it("runs a 3-D chart's value axis to its data and no further", () => {
    // Excel's PDF: 10…20 on 0…20 by 2 in 3-D, where flat columns run to 25.
    const s = scene(columns());
    expect(texts(s)).toContain('20');
    expect(texts(s)).not.toContain('22');
    expect(texts(s)).not.toContain('25');
    const { bar3D, ...flat } = columns();
    expect(bar3D).toBeDefined();
    expect(texts(scene(flat))).toContain('25');
    // …a view of no angle at all is flat, and unpadded still.
    const square = scene(columns({ rotX: 0, rotY: 0 }));
    expect(square.polygons ?? []).toHaveLength(0);
    expect(square.walls ?? []).toHaveLength(0);
    expect(texts(square)).not.toContain('25');
  });

  it('rules the floor only as the chart says', () => {
    const cleared = scene(columns({}, { floor: { line: { none: true } } }));
    expect(cleared.walls ?? []).toHaveLength(0);
    const filled = scene(
      columns({}, { floor: { fillHex: 'EEEEEE' }, backWall: { line: { colorHex: '7F7F7F' } } }),
    );
    expect((filled.walls ?? []).map((w) => [w.fillHex, w.strokeHex])).toEqual([
      [undefined, '7F7F7F'],
      ['EEEEEE', undefined],
    ]);
  });

  it('turned left, labels its values up the back of the box, its bars showing their left', () => {
    // Excel's PDF at rotY 340: the value axis on the back edge, the side wall
    // on the right, each bar's left side at 0.4 of its colour.
    const right = scene(columns());
    const left = scene(columns({ rotY: 340 }));
    const zero = (s: ChartScene) => s.labels.find((l) => l.text === '0')!;
    const floor = (s: ChartScene) => (s.walls ?? [])[0]!.points;
    expect(zero(right).x).toBeCloseTo(floor(right)[0]![0] - 3, 6);
    const backLeft = floor(left)[3]!;
    expect(zero(left).x).toBeCloseTo(backLeft[0] - 3, 6);
    expect(zero(left).y).toBeGreaterThan(zero(right).y);
    expect(shaded(left, '082634')).toHaveLength(3);
    for (const g of left.gridlines ?? []) expect(g.points[2]![0]).toBeGreaterThan(g.points[0]![0]);
  });

  it('seen from below, shows the bars underneath and no axis along its foot', () => {
    // Excel's PDF at rotX −15 draws neither the category axis nor its labels.
    const s = scene(columns({ rotX: -15 }));
    for (const t of ['A', 'B', 'C']) expect(texts(s)).not.toContain(t);
    expect(shaded(s, '082634')).toHaveLength(3);
    expect(texts(scene(columns()))).toContain('A');
  });

  it('lays a stack of bars on its left wall, its gridlines across the bottom', () => {
    const stack: Chart = {
      ...columns(),
      barDir: 'bar',
      grouping: 'stacked',
      series: [
        { name: 'S1', values: [10, 20, 15], colorHex: '156082' },
        { name: 'S2', values: [5, 8, 4], colorHex: 'E97132' },
      ],
    };
    const s = scene(stack);
    const floor = (s.walls ?? [])[0]!.points;
    // The floor stands upright at the value axis's start.
    expect(floor[0]![0]).toBeCloseTo(floor[3]![0], 6);
    for (const g of s.gridlines ?? []) {
      // Up the bottom wall to the back, then up the back wall.
      expect(g.points[1]![1]).toBeGreaterThan(g.points[0]![1]);
      expect(g.points[2]![0]).toBeCloseTo(g.points[1]![0], 6);
    }
    // A stack is one bar: 0.4 of its slot at a gap of 150.
    const bars = s.rects.filter((r) => r.fillHex === '156082');
    const pitch = Math.abs(bars[1]!.y - bars[0]!.y);
    expect(bars[0]!.h / pitch).toBeCloseTo(0.4, 6);
  });

  it('draws the boxes back to front', () => {
    // Turned right, the right-hand bar is nearer: its faces go on last.
    const s = scene(columns());
    const tops = shaded(s, '104963').map((pg) => pg.points[0]![0]);
    expect(tops).toEqual([...tops].sort((a, b) => a - b));
    const left = shaded(scene(columns({ rotY: 340 })), '082634').map((pg) => pg.points[0]![0]);
    expect(left).toEqual([...left].sort((a, b) => b - a));
  });

  it("labels a column over its top's far edge and a bar past its end", () => {
    // Excel's PDF centres a column's value on the far edge of its top, over
    // its front, and starts a bar's halfway along its side, level with its front.
    const s = scene(columns({}, { showValues: true }));
    const bar = fronts(s).find(
      (r) => Math.abs(r.h - Math.max(...fronts(s).map((f) => f.h))) < 1e-9,
    )!;
    const twenty = s.labels.find((l) => l.text === '20' && l.x > bar.x && l.x < bar.x + bar.w)!;
    const rise = bar.w * Math.sin(rad(15));
    expect(twenty.y + twenty.sizePt * 0.35).toBeCloseTo(bar.y + bar.h + rise, 6);
    const lying = scene(columns({}, { showValues: true, barDir: 'bar' }));
    const long = lying.rects
      .filter((r) => r.fillHex === '156082')
      .reduce((a, b) => (b.w > a.w ? b : a));
    const label = lying.labels.find((l) => l.text === '20' && l.align === 'left')!;
    expect(label.align).toBe('left');
    expect(label.x).toBeCloseTo(long.x + long.w + 3 + (long.h * Math.sin(rad(20))) / 2, 6);
    expect(label.y + label.sizePt * 0.35).toBeCloseTo(long.y + long.h / 2, 6);
  });
});
