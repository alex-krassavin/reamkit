import { describe, expect, it } from 'vitest';

import type { PathSegment } from '@/pdf/vector-graphics';
import { arcPoint, arcToBeziers, ellipseSegments, roundRectSegments } from '@/pdf/arc-to-bezier';

const cubics = (segs: ReadonlyArray<PathSegment>) => segs.filter((s) => s.op === 'cubic');

function bounds(segs: ReadonlyArray<PathSegment>) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const see = (x: number, y: number) => {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  };
  for (const s of segs) {
    if (s.op === 'move' || s.op === 'line') see(s.x, s.y);
    else if (s.op === 'cubic') {
      see(s.x1, s.y1);
      see(s.x2, s.y2);
      see(s.x, s.y);
    }
  }
  return { minX, minY, maxX, maxY };
}

describe('arcToBeziers', () => {
  it('approximates a 90° unit arc with one cubic and the kappa control points', () => {
    const segs = arcToBeziers(0, 0, 1, 1, 0, Math.PI / 2);
    expect(segs).toHaveLength(1);
    const c = segs[0]!;
    if (c.op !== 'cubic') throw new Error('expected cubic');
    // start (1,0) → end (0,1), handles at κ ≈ 0.5523
    expect(c.x1).toBeCloseTo(1, 6);
    expect(c.y1).toBeCloseTo(0.55228, 4);
    expect(c.x2).toBeCloseTo(0.55228, 4);
    expect(c.y2).toBeCloseTo(1, 6);
    expect(c.x).toBeCloseTo(0, 6);
    expect(c.y).toBeCloseTo(1, 6);
  });

  it('splits a full turn into four ≤90° cubics', () => {
    expect(arcToBeziers(0, 0, 1, 1, 0, 2 * Math.PI)).toHaveLength(4);
  });

  it('returns nothing for a degenerate (zero-radius / zero-sweep) arc', () => {
    expect(arcToBeziers(0, 0, 0, 0, 0, Math.PI / 2)).toHaveLength(0);
    expect(arcToBeziers(0, 0, 1, 1, 0, 0)).toHaveLength(0);
  });

  it('arcPoint lands on the ellipse', () => {
    const [x, y] = arcPoint(2, 1, 2, 1, Math.PI / 2);
    expect(x).toBeCloseTo(2, 6);
    expect(y).toBeCloseTo(2, 6);
  });
});

describe('ellipseSegments', () => {
  it('is a closed path of move + 4 cubics, starting at the rightmost point', () => {
    const segs = ellipseSegments(4, 2);
    expect(segs[0]).toEqual({ op: 'move', x: 4, y: 1 });
    expect(cubics(segs)).toHaveLength(4);
    expect(segs[segs.length - 1]).toEqual({ op: 'close' });
    const b = bounds(segs);
    expect(b.minX).toBeCloseTo(0, 5);
    expect(b.maxX).toBeCloseTo(4, 5);
    expect(b.minY).toBeCloseTo(0, 5);
    expect(b.maxY).toBeCloseTo(2, 5);
  });
});

describe('roundRectSegments', () => {
  it('rounds the corners and stays within the box', () => {
    const segs = roundRectSegments(100, 50, 10);
    expect(segs[0]).toEqual({ op: 'move', x: 10, y: 0 });
    expect(cubics(segs).length).toBe(4); // one cubic per corner
    expect(segs[segs.length - 1]).toEqual({ op: 'close' });
    const b = bounds(segs);
    expect(b.minX).toBeGreaterThanOrEqual(-1e-6);
    expect(b.maxX).toBeLessThanOrEqual(100 + 1e-6);
    expect(b.minY).toBeGreaterThanOrEqual(-1e-6);
    expect(b.maxY).toBeLessThanOrEqual(50 + 1e-6);
  });

  it('clamps the radius to half the shorter side', () => {
    // r=40 on a 50-tall box clamps to 25 → straight bottom edge from x=25.
    const segs = roundRectSegments(100, 50, 40);
    expect(segs[0]).toEqual({ op: 'move', x: 25, y: 0 });
  });

  it('degenerates to a plain rectangle when r=0', () => {
    const segs = roundRectSegments(100, 50, 0);
    expect(cubics(segs)).toHaveLength(0);
    expect(segs).toEqual([
      { op: 'move', x: 0, y: 0 },
      { op: 'line', x: 100, y: 0 },
      { op: 'line', x: 100, y: 50 },
      { op: 'line', x: 0, y: 50 },
      { op: 'close' },
    ]);
  });
});
