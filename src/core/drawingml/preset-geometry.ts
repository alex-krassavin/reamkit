// ECMA-376 §20.1.10.55 — preset shape geometries → vector paths.
//
// Produces paths in a LOCAL y-up frame (origin bottom-left) sized w×h points,
// matching the vector-graphics emit layer. The y-down→y-up flip therefore
// lives here (and in custGeom), never at emit time. Adjust values are the raw
// a:gd numbers (Word's preset guides are mostly thousandths-of-a-percent).

import type { CustomGeometry } from '@/core/document-model';
import type { PathSegment, VectorPath } from '@/core/vector';
import { arcToBeziers, ellipseSegments, roundRectSegments } from '@/core/arc-to-bezier';
import { PathBuilder } from '@/core/vector';

/**
 * The four closed segments of the `w`×`h` bounding rectangle, in the local y-up
 * frame (origin bottom-left). The fallback geometry for unknown presets.
 */
export function rectSegments(w: number, h: number): ReadonlyArray<PathSegment> {
  return new PathBuilder().moveTo(0, 0).lineTo(w, 0).lineTo(w, h).lineTo(0, h).close().build()
    .segments;
}

/** The `w`×`h` bounding rectangle as a {@link VectorPath}. */
export function rectPath(w: number, h: number): VectorPath {
  return { segments: rectSegments(w, h) };
}

// Closed polygon from a point list (y-up, points in path order).
function polygon(points: ReadonlyArray<readonly [number, number]>): VectorPath {
  const b = new PathBuilder();
  points.forEach(([x, y], i) => (i === 0 ? b.moveTo(x, y) : b.lineTo(x, y)));
  return b.close().build();
}

const clamp = (v: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, v));

// Adjust guide as a fraction (raw a:gd val ÷ 100000), with a default.
const frac = (adjust: ReadonlyMap<string, number>, key: string, def: number): number =>
  (adjust.get(key) ?? def) / 100000;

/**
 * Map a preset shape name (§20.1.10.55) to its vector path(s) in the local y-up
 * frame, sized `w`×`h` points and shaped by the `adjust` guides. Returns `null`
 * for an unknown preset so the caller falls back to the bounding rectangle —
 * graceful degradation that keeps even unimplemented presets visible (with the
 * right fill/line) and never throws.
 *
 * @param preset The preset geometry name (e.g. `'roundRect'`, `'rightArrow'`).
 * @param w      Box width in points.
 * @param h      Box height in points.
 * @param adjust Raw `a:gd` adjust guides by name (thousandths of a percent).
 * @returns The path(s), or `null` for an unrecognised preset.
 */
export function presetPaths(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
): Array<VectorPath> | null {
  switch (preset) {
    case 'rect':
      return [rectPath(w, h)];
    case 'roundRect': {
      // adj is a fraction of the shorter side (default 16667 ⇒ ~1/6).
      return [{ segments: roundRectSegments(w, h, frac(adjust, 'adj', 16667) * Math.min(w, h)) }];
    }
    case 'ellipse':
      return [{ segments: ellipseSegments(w, h) }];
    case 'triangle': {
      // adj = horizontal apex position (default centred).
      const apexX = clamp(frac(adjust, 'adj', 50000), 0, 1) * w;
      return [
        polygon([
          [0, 0],
          [w, 0],
          [apexX, h],
        ]),
      ];
    }
    case 'rtTriangle':
      return [
        polygon([
          [0, 0],
          [w, 0],
          [0, h],
        ]),
      ];
    case 'diamond':
      return [
        polygon([
          [w / 2, 0],
          [w, h / 2],
          [w / 2, h],
          [0, h / 2],
        ]),
      ];
    case 'parallelogram': {
      const s = clamp(frac(adjust, 'adj', 25000) * w, 0, w);
      return [
        polygon([
          [0, 0],
          [w - s, 0],
          [w, h],
          [s, h],
        ]),
      ];
    }
    case 'trapezoid': {
      const t = clamp(frac(adjust, 'adj', 25000) * w, 0, w / 2);
      return [
        polygon([
          [0, 0],
          [w, 0],
          [w - t, h],
          [t, h],
        ]),
      ];
    }
    case 'pentagon':
      return [regularPolygon(w, h, 5)];
    case 'hexagon': {
      const inset = clamp(frac(adjust, 'adj', 25000) * w, 0, w / 2);
      return [
        polygon([
          [inset, 0],
          [w - inset, 0],
          [w, h / 2],
          [w - inset, h],
          [inset, h],
          [0, h / 2],
        ]),
      ];
    }
    case 'line':
    case 'straightConnector1':
      // Box diagonal (top-left → bottom-right in y-down ⇒ (0,h)→(w,0) here).
      // Open path: a connector is stroked, never filled.
      return [new PathBuilder().moveTo(0, h).lineTo(w, 0).build()];
    case 'rightArrow':
      return [blockArrow(w, h, adjust, 'right')];
    case 'leftArrow':
      return [blockArrow(w, h, adjust, 'left')];
    case 'upArrow':
      return [blockArrow(w, h, adjust, 'up')];
    case 'downArrow':
      return [blockArrow(w, h, adjust, 'down')];
    default:
      return null;
  }
}

/**
 * Convert a custom geometry (§20.1.9) to a vector path. Path-space coordinates
 * (y-down, origin top-left) are scaled to the shape box and flipped to the local
 * y-up frame; quadratics are elevated to cubics; `arcTo` is decomposed via the
 * shared arc helper. Angles are 1/60000° clockwise in y-down, which becomes
 * negative (CCW) once y is flipped.
 *
 * @param geom The parsed custom geometry (path size + draw commands).
 * @param wPt  Box width in points.
 * @param hPt  Box height in points.
 * @returns A single-element array holding the built path.
 */
export function customPaths(geom: CustomGeometry, wPt: number, hPt: number): Array<VectorPath> {
  const sx = geom.pathWidth > 0 ? wPt / geom.pathWidth : 1;
  const sy = geom.pathHeight > 0 ? hPt / geom.pathHeight : 1;
  const tx = (x: number): number => x * sx;
  const ty = (y: number): number => hPt - y * sy;

  const b = new PathBuilder();
  // Current point in path-space (needed to derive arcTo's ellipse centre).
  let curX = 0;
  let curY = 0;
  for (const cmd of geom.commands) {
    switch (cmd.cmd) {
      case 'move':
        b.moveTo(tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'line':
        b.lineTo(tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'cubic':
        b.cubicTo(tx(cmd.x1), ty(cmd.y1), tx(cmd.x2), ty(cmd.y2), tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'quad': {
        // Elevate (P0=cur, ctrl, P2) to a cubic in path-space, then transform.
        const c1x = curX + (2 / 3) * (cmd.x1 - curX);
        const c1y = curY + (2 / 3) * (cmd.y1 - curY);
        const c2x = cmd.x + (2 / 3) * (cmd.x1 - cmd.x);
        const c2y = cmd.y + (2 / 3) * (cmd.y1 - cmd.y);
        b.cubicTo(tx(c1x), ty(c1y), tx(c2x), ty(c2y), tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      }
      case 'arc': {
        const st = (cmd.stAng / 60000) * (Math.PI / 180);
        const sw = (cmd.swAng / 60000) * (Math.PI / 180);
        // The current point is the arc start; derive the ellipse centre in
        // path-space, then map to local (negating angles for the y-flip).
        const cxP = curX - cmd.wR * Math.cos(st);
        const cyP = curY - cmd.hR * Math.sin(st);
        b.append(arcToBeziers(tx(cxP), ty(cyP), cmd.wR * sx, cmd.hR * sy, -st, -sw));
        curX = cxP + cmd.wR * Math.cos(st + sw);
        curY = cyP + cmd.hR * Math.sin(st + sw);
        break;
      }
      case 'close':
        b.close();
        break;
    }
  }
  return [b.build()];
}

// Regular n-gon inscribed in the box, first vertex at the top (pointing up).
function regularPolygon(w: number, h: number, n: number): VectorPath {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const pts: Array<readonly [number, number]> = [];
  for (let i = 0; i < n; i++) {
    const a = Math.PI / 2 + (i * 2 * Math.PI) / n;
    pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)]);
  }
  return polygon(pts);
}

// Single-headed block arrow. adj1 = body thickness fraction (default 0.5),
// adj2 = head length fraction (default 0.5).
function blockArrow(
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
  dir: 'right' | 'left' | 'up' | 'down',
): VectorPath {
  const thick = clamp(frac(adjust, 'adj1', 50000), 0, 1);
  const head = clamp(frac(adjust, 'adj2', 50000), 0, 1);
  if (dir === 'right' || dir === 'left') {
    const sb = (h * (1 - thick)) / 2;
    const st = (h * (1 + thick)) / 2;
    const hl = head * w;
    if (dir === 'right') {
      const bx = w - hl;
      return polygon([
        [0, sb],
        [bx, sb],
        [bx, 0],
        [w, h / 2],
        [bx, h],
        [bx, st],
        [0, st],
      ]);
    }
    const bx = hl;
    return polygon([
      [w, sb],
      [bx, sb],
      [bx, 0],
      [0, h / 2],
      [bx, h],
      [bx, st],
      [w, st],
    ]);
  }
  const sl = (w * (1 - thick)) / 2;
  const sr = (w * (1 + thick)) / 2;
  const hh = head * h;
  if (dir === 'up') {
    const by = h - hh;
    return polygon([
      [sl, 0],
      [sr, 0],
      [sr, by],
      [w, by],
      [w / 2, h],
      [0, by],
      [sl, by],
    ]);
  }
  const ty = hh;
  return polygon([
    [sl, h],
    [sr, h],
    [sr, ty],
    [w, ty],
    [w / 2, 0],
    [0, ty],
    [sl, ty],
  ]);
}
