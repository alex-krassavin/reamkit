import { describe, expect, it } from 'vitest';

import type { CustomGeometry } from '@/document-model';
import type { PathSegment } from '@/pdf/vector-graphics';
import { customPaths, presetPaths } from '@/pdf/preset-geometry';

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

  it('returns null for an unknown preset (caller falls back to rect)', () => {
    expect(presetPaths('cloudCallout', W, H, new Map())).toBeNull();
    expect(presetPaths('star5', W, H, new Map())).toBeNull();
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
