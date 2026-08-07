import { describe, expect, it } from 'vitest';

import type { CustomGeometry } from '@/core/document-model';
import type { PathSegment } from '@/pdf/vector-graphics';
import { customPaths, presetPaths } from '@/core/drawingml/preset-geometry';

const W = 100;
const H = 60;

const KNOWN = [
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'rtTriangle',
  'diamond',
  'parallelogram',
  'trapezoid',
  'pentagon',
  'hexagon',
  'line',
  'straightConnector1',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'star4',
  'star5',
  'star6',
  'star7',
  'star8',
  'star10',
  'star12',
  'star16',
  'star24',
  'star32',
  'round1Rect',
  'round2SameRect',
  'round2DiagRect',
  'snip1Rect',
  'snip2SameRect',
  'snip2DiagRect',
  'snipRoundRect',
  // The batch the pptx corpus asked for, most-used first.
  'chevron',
  'homePlate',
  'pie',
  'blockArc',
  'corner',
  'leftRightArrow',
  'notchedRightArrow',
  'stripedRightArrow',
  'bentUpArrow',
  'gear6',
  'gear9',
  'actionButtonBlank',
  'flowChartMerge',
  'flowChartExtract',
  'flowChartDelay',
  'flowChartMagneticDisk',
  'flowChartMagneticTape',
  'flowChartPunchedTape',
  'circularArrow',
  'leftCircularArrow',
  'leftBracket',
  'rightBracket',
  'leftBrace',
  'rightBrace',
];

function coords(segs: ReadonlyArray<PathSegment>): Array<number> {
  const out: Array<number> = [];
  for (const s of segs) {
    if (s.op === 'move' || s.op === 'line') out.push(s.x, s.y);
    else if (s.op === 'cubic') out.push(s.x1, s.y1, s.x2, s.y2, s.x, s.y);
  }
  return out;
}

describe('presetPaths', () => {
  it('produces a finite, in-bounds path for every supported preset', () => {
    for (const name of KNOWN) {
      const paths = presetPaths(name, W, H, new Map());
      expect(paths, name).not.toBeNull();
      expect(paths!.length, name).toBeGreaterThan(0);
      const segs = paths!.flatMap((p) => p.segments);
      expect(segs.length, name).toBeGreaterThan(0);
      for (const v of coords(segs)) {
        expect(Number.isFinite(v), name).toBe(true);
        expect(v, name).toBeGreaterThanOrEqual(-0.001);
        expect(v, name).toBeLessThanOrEqual(Math.max(W, H) + 0.001);
      }
    }
  });

  it("runs a callout's tail OUTSIDE its box, which is what a callout is", () => {
    // The default tail points below the rectangle; only the body is in bounds.
    const paths = presetPaths('wedgeRoundRectCallout', W, H, new Map());
    expect(paths).toHaveLength(2);
    const body = coords(paths![0]!.segments);
    for (const v of body) expect(v).toBeGreaterThanOrEqual(-0.001);
    const tail = coords(paths![1]!.segments);
    expect(tail.some((v) => v < 0)).toBe(true);
    for (const v of tail) expect(Number.isFinite(v)).toBe(true);
  });

  it('points a chevron and notches its back, where a home plate has none', () => {
    // Both are the same pentagon; the chevron carries one point more, the
    // notch that lets a row of them interlock.
    const chevron = presetPaths('chevron', W, H, new Map())![0]!.segments;
    const plate = presetPaths('homePlate', W, H, new Map())![0]!.segments;
    expect(chevron.length).toBe(plate.length + 1);
    // The tip of each sits mid-height at the right edge.
    expect(coords(plate)).toContain(W);
    expect(coords(chevron)).toContain(H / 2);
  });

  it('cuts a pie through its centre and a block arc through its inner edge', () => {
    const pie = presetPaths('pie', W, H, new Map())![0]!.segments;
    // …ending at the centre before it closes.
    const last = pie[pie.length - 2];
    expect(last?.op === 'line' ? [last.x, last.y] : []).toEqual([W / 2, H / 2]);
    const arc = presetPaths('blockArc', W, H, new Map())![0]!.segments;
    expect(arc.filter((s) => s.op === 'cubic').length).toBeGreaterThan(
      pie.filter((s) => s.op === 'cubic').length,
    );
  });

  it('gives a gear one tooth per name, four corners each', () => {
    for (const [preset, teeth] of [
      ['gear6', 6],
      ['gear9', 9],
    ] as const) {
      const segs = presetPaths(preset, W, H, new Map())![0]!.segments;
      // move + 4 points per tooth (the first is the move) + close.
      expect(segs.length, preset).toBe(teeth * 4 + 1);
    }
  });

  it('leaves a brace and a bracket open, so only the stroke shows', () => {
    for (const preset of ['leftBrace', 'rightBrace', 'leftBracket', 'rightBracket']) {
      const segs = presetPaths(preset, W, H, new Map())![0]!.segments;
      expect(
        segs.some((s) => s.op === 'close'),
        preset,
      ).toBe(false);
    }
  });

  it('returns null for an unknown preset (caller falls back to rect)', () => {
    expect(presetPaths('cloudCallout', W, H, new Map())).toBeNull();
    expect(presetPaths('cube', W, H, new Map())).toBeNull();
  });

  it('a star alternates outer and inner vertices and spans the box', () => {
    const segs = presetPaths('star5', W, H, new Map())![0]!.segments;
    // 5 outer + 5 inner points, then close.
    expect(segs.map((s) => s.op)).toEqual(['move', ...Array<'line'>(9).fill('line'), 'close']);
    const xs = coords(segs).filter((_, i) => i % 2 === 0);
    const ys = coords(segs).filter((_, i) => i % 2 === 1);
    expect(Math.min(...xs)).toBeCloseTo(0, 2);
    expect(Math.max(...xs)).toBeCloseTo(W, 2);
    expect(Math.min(...ys)).toBeCloseTo(0, 2);
    expect(Math.max(...ys)).toBeCloseTo(H, 2);
    // The first vertex is the topmost point, dead centre.
    const first = segs[0]!;
    if (first.op !== 'move') throw new Error('unreachable');
    expect(first.x).toBeCloseTo(W / 2, 6);
    expect(first.y).toBeCloseTo(H, 6);
  });

  it('a snipped corner cuts the corner off and leaves the rest square', () => {
    // adj1 (default 16667) cuts the TOP pair; adj2 defaults to 0, so the bottom
    // corners stay square. The cut is that fraction of the SHORTER side (60).
    const segs = presetPaths('snip2SameRect', W, H, new Map())![0]!.segments;
    expect(segs.map((s) => s.op)).toEqual([
      'move',
      'line',
      'line',
      'line',
      'line',
      'line',
      'close',
    ]);
    const want = [0, 50, 10, 60, 90, 60, 100, 50, 100, 0, 0, 0];
    coords(segs).forEach((v, i) => expect(v).toBeCloseTo(want[i]!, 3));
  });

  it('a rounded corner arcs where a snipped one cuts', () => {
    const snip = presetPaths('snip1Rect', W, H, new Map())![0]!.segments;
    const round = presetPaths('round1Rect', W, H, new Map())![0]!.segments;
    expect(snip.map((s) => s.op)).toEqual(['move', 'line', 'line', 'line', 'line', 'close']);
    expect(round.map((s) => s.op)).toEqual(['move', 'line', 'cubic', 'line', 'line', 'close']);
  });

  it('triangle is three line segments closed', () => {
    const paths = presetPaths('triangle', W, H, new Map())!;
    expect(paths[0]!.segments.map((s) => s.op)).toEqual(['move', 'line', 'line', 'close']);
  });

  it('triangle apex honours the adj guide', () => {
    const segs = presetPaths('triangle', W, H, new Map([['adj', 75000]]))![0]!.segments;
    const apex = segs[2]!;
    if (apex.op !== 'line') throw new Error('unreachable');
    expect(apex.x).toBeCloseTo(75, 6); // 0.75 * 100
    expect(apex.y).toBeCloseTo(60, 6);
  });

  it('line/connector is an open two-point path (no close)', () => {
    const segs = presetPaths('line', W, H, new Map())![0]!.segments;
    expect(segs.map((s) => s.op)).toEqual(['move', 'line']);
  });

  it('block arrows are 7-point polygons', () => {
    for (const dir of ['rightArrow', 'leftArrow', 'upArrow', 'downArrow']) {
      const segs = presetPaths(dir, W, H, new Map())![0]!.segments;
      // 1 move + 6 line + close = 8
      expect(segs.length, dir).toBe(8);
      expect(segs[segs.length - 1]!.op, dir).toBe('close');
    }
  });
});

describe('customPaths (custGeom)', () => {
  it('scales path-space and flips to y-up', () => {
    // Triangle in a 100×100 path box → rendered into a 200×100 shape.
    const geom: CustomGeometry = {
      pathWidth: 100,
      pathHeight: 100,
      commands: [
        { cmd: 'move', x: 0, y: 0 },
        { cmd: 'line', x: 100, y: 0 },
        { cmd: 'line', x: 50, y: 100 },
        { cmd: 'close' },
      ],
    };
    const segs = customPaths(geom, 200, 100)[0]!.segments;
    expect(segs).toEqual([
      { op: 'move', x: 0, y: 100 }, // (0,0) top-left → y-up top-left
      { op: 'line', x: 200, y: 100 },
      { op: 'line', x: 100, y: 0 }, // (50,100) bottom-centre → y-up bottom
      { op: 'close' },
    ]);
  });

  it('elevates a quadratic to a single cubic', () => {
    const geom: CustomGeometry = {
      pathWidth: 100,
      pathHeight: 100,
      commands: [
        { cmd: 'move', x: 0, y: 0 },
        { cmd: 'quad', x1: 0, y1: 100, x: 100, y: 100 },
      ],
    };
    const segs = customPaths(geom, 100, 100)[0]!.segments;
    const cubic = segs.find((s) => s.op === 'cubic');
    expect(cubic).toBeDefined();
    if (cubic?.op !== 'cubic') throw new Error('unreachable');
    expect(cubic.x).toBeCloseTo(100, 6); // endpoint (100,100) → (100,0)
    expect(cubic.y).toBeCloseTo(0, 6);
    expect(segs.filter((s) => s.op === 'cubic')).toHaveLength(1);
  });

  it('decomposes an arcTo into cubics ending at the swept point', () => {
    // start (100,50); centre (50,50); 90° clockwise sweep → end (50,100).
    const geom: CustomGeometry = {
      pathWidth: 100,
      pathHeight: 100,
      commands: [
        { cmd: 'move', x: 100, y: 50 },
        { cmd: 'arc', wR: 50, hR: 50, stAng: 0, swAng: 5400000 },
      ],
    };
    const segs = customPaths(geom, 100, 100)[0]!.segments;
    const cubics = segs.filter((s) => s.op === 'cubic');
    expect(cubics.length).toBeGreaterThanOrEqual(1);
    const last = cubics[cubics.length - 1]!;
    expect(last.x).toBeCloseTo(50, 4); // (50,100) path-space → y-up (50,0)
    expect(last.y).toBeCloseTo(0, 4);
  });
});

// §20.1.10.55 — the gallery shapes a deck reaches for that used to degrade to
// the bounding rectangle. tdf114848.pptx draws one of each.
describe('the gallery presets that were rectangles', () => {
  const NEW = [
    'bevel',
    'halfFrame',
    'heptagon',
    'octagon',
    'decagon',
    'dodecagon',
    'donut',
    'frame',
    'diagStripe',
    'teardrop',
    'can',
    'plaque',
  ];

  it('knows every one of them', () => {
    for (const p of NEW) expect(presetPaths(p, W, H, new Map()), p).not.toBeNull();
  });

  it('gives a polygon its side count, and stands an even one on a flat edge', () => {
    const corners = (p: string): number =>
      presetPaths(p, W, H, new Map())![0]!.segments.filter((s) => s.op === 'line').length + 1;
    expect(corners('heptagon')).toBe(7);
    expect(corners('decagon')).toBe(10);
    expect(corners('dodecagon')).toBe(12);
    // An even polygon is turned half a step so it rests on an edge: no vertex
    // sits at the very top, where an odd one's does.
    const top = (p: string): number =>
      Math.max(...presetPaths(p, W, H, new Map())![0]!.segments.map((s) => ('y' in s ? s.y : 0)));
    expect(top('heptagon')).toBeCloseTo(H, 4);
    expect(top('decagon')).toBeLessThan(H);
  });

  it('cuts a hole in the shapes that have one', () => {
    for (const p of ['donut', 'frame']) {
      const path = presetPaths(p, W, H, new Map())![0]!;
      // Two subpaths, and the inner one only counts as a hole under even-odd.
      expect(path.fillRule, p).toBe('evenodd');
      expect(path.segments.filter((s) => s.op === 'move').length, p).toBe(2);
    }
  });

  it("bites a plaque's corners inwards where a round rectangle cuts them off", () => {
    const at = (p: string): ReadonlyArray<{ x: number; y: number }> =>
      presetPaths(p, W, H, new Map())![0]!.segments.flatMap((s) =>
        'x' in s ? [{ x: s.x, y: s.y }] : [],
      );
    // Both leave the corner alone at the same distance; the plaque's curve then
    // runs INTO the shape, so its points stay nearer the middle than the round
    // rectangle's, which bulge back out to the edge.
    const near = (pts: ReadonlyArray<{ x: number; y: number }>): number =>
      Math.min(...pts.filter((q) => q.y < H / 2).map((q) => q.x + q.y));
    expect(near(at('plaque'))).toBeGreaterThan(near(at('roundRect')));
  });

  it('mitres a half frame where its two bars meet', () => {
    // Two bars — the top-left half of a frame — joined at the box's own
    // diagonal, so a wide box mitres at a shallower angle than a tall one.
    const pts = presetPaths('halfFrame', W, H, new Map())![0]!.segments.flatMap((s) =>
      'x' in s ? [{ x: s.x, y: s.y }] : [],
    );
    expect(pts).toHaveLength(6);
    // The outer corner is the box's own, and the inner one is set in by the
    // thickness of both bars.
    expect(pts.some((q) => q.x === 0 && q.y === 0)).toBe(true);
    expect(pts.some((q) => q.x === W && q.y === H)).toBe(true);
    const inner = Math.min(...pts.filter((q) => q.x > 0).map((q) => q.x));
    expect(inner).toBeCloseTo(Math.min(W, H) / 3, 3);
  });

  it("builds a bevel's four faces round its inner rectangle", () => {
    // PowerPoint shades each face to make the block read as raised; a shape
    // here takes one fill, so what carries is the structure.
    const paths = presetPaths('bevel', W, H, new Map())!;
    expect(paths).toHaveLength(5);
    const t = 0.125 * Math.min(W, H);
    const last = paths[4]!.segments.flatMap((s) => ('x' in s ? [{ x: s.x, y: s.y }] : []));
    // The inner rectangle is inset by the chamfer on every side.
    expect(Math.min(...last.map((q) => q.x))).toBeCloseTo(t, 3);
    expect(Math.max(...last.map((q) => q.x))).toBeCloseTo(W - t, 3);
    expect(Math.max(...last.map((q) => q.y))).toBeCloseTo(H - t, 3);
  });

  it('fills the lid of a cylinder instead of punching it out', () => {
    // Both paths wind the same way round: wound the other way a nonzero fill
    // cancels the two and the top comes out hollow — which is how the
    // flowchart disk had been drawn all along.
    for (const preset of ['can', 'flowChartMagneticDisk']) {
      const [b1, l1] = presetPaths(preset, W, H, new Map())!;
      expect(b1, preset).toBeDefined();
      expect(l1, preset).toBeDefined();
    }
    const [body, lid] = presetPaths('can', W, H, new Map())!;
    const turn = (p: { segments: ReadonlyArray<PathSegment> }): number => {
      const pts = p.segments.flatMap((s) => ('x' in s ? [[s.x, s.y] as const] : []));
      let a = 0;
      for (let i = 0; i < pts.length; i++) {
        const [x1, y1] = pts[i]!;
        const [x2, y2] = pts[(i + 1) % pts.length]!;
        a += x1 * y2 - x2 * y1;
      }
      return Math.sign(a);
    };
    expect(turn(body!)).toBe(turn(lid!));
  });

  it("bulges a cylinder's base downwards, not up", () => {
    const [body] = presetPaths('flowChartMagneticDisk', W, H, new Map())!;
    const ys = body!.segments.flatMap((s) => ('y' in s ? [s.y] : []));
    // The base reaches the very bottom of the box; drawn the other way round it
    // arched up into the body instead.
    expect(Math.min(...ys)).toBeCloseTo(0, 3);
  });
});
