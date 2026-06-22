// Elliptical-arc → cubic-Bézier decomposition (the standard "kappa" method).
//
// Works in a clean math convention: a LOCAL y-up frame, angles measured in
// radians counter-clockwise from the +x axis. A point on the ellipse at angle
// θ is (cx + rx·cosθ, cy + ry·sinθ). Callers that come from DrawingML space
// (y-down, clockwise angles) convert to this convention before calling in —
// the y-up winding lives here and nowhere else.
//
// Shared by ellipse / roundRect presets and by custGeom's <a:arcTo>.

import type { PathSegment } from '@/core/vector';
import { PathBuilder } from '@/core/vector';

/**
 * Point on an axis-aligned ellipse at angle θ (radians, CCW from +x): returns
 * `[cx + rx·cosθ, cy + ry·sinθ]`.
 */
export function arcPoint(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  angleRad: number,
): readonly [number, number] {
  return [cx + rx * Math.cos(angleRad), cy + ry * Math.sin(angleRad)];
}

/**
 * Decompose an arc into cubic segments, ≤90° each. Returns ONLY `cubic`
 * segments — the pen is assumed to already sit at the arc's start point
 * (`arcPoint(..., startAngleRad)`); the caller emits a leading move/line.
 * Degenerate input (non-positive radius or zero sweep) yields no segments,
 * so no NaN can reach the serializer.
 *
 * @param startAngleRad Start angle in radians (CCW from +x).
 * @param sweepAngleRad Signed angular extent in radians (sign = winding direction).
 */
export function arcToBeziers(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  startAngleRad: number,
  sweepAngleRad: number,
): Array<PathSegment> {
  const segments: Array<PathSegment> = [];
  if (!(rx > 0) || !(ry > 0) || sweepAngleRad === 0 || !Number.isFinite(sweepAngleRad)) {
    return segments;
  }
  // Split into equal pieces no wider than 90°. The +1e-9 avoids an extra
  // segment from floating error when the sweep is an exact multiple of π/2.
  const nSegs = Math.max(1, Math.ceil(Math.abs(sweepAngleRad) / (Math.PI / 2 + 1e-9)));
  const delta = sweepAngleRad / nSegs;
  // Control-point handle length for a delta-wide arc.
  const alpha = (4 / 3) * Math.tan(delta / 4);

  let theta = startAngleRad;
  for (let i = 0; i < nSegs; i++) {
    const t2 = theta + delta;
    const cos1 = Math.cos(theta);
    const sin1 = Math.sin(theta);
    const cos2 = Math.cos(t2);
    const sin2 = Math.sin(t2);
    // Tangent at θ is (-rx·sinθ, ry·cosθ); P1 = P0 + α·T(θ1), P2 = P3 − α·T(θ2).
    segments.push({
      op: 'cubic',
      x1: cx + rx * cos1 - alpha * rx * sin1,
      y1: cy + ry * sin1 + alpha * ry * cos1,
      x2: cx + rx * cos2 + alpha * rx * sin2,
      y2: cy + ry * sin2 - alpha * ry * cos2,
      x: cx + rx * cos2,
      y: cy + ry * sin2,
    });
    theta = t2;
  }
  return segments;
}

/**
 * Full ellipse inscribed in the `(0,0)`–`(w,h)` box, as a closed path. Starts at
 * the rightmost point and sweeps a full turn CCW.
 */
export function ellipseSegments(w: number, h: number): ReadonlyArray<PathSegment> {
  const rx = w / 2;
  const ry = h / 2;
  const cx = rx;
  const cy = ry;
  const start = arcPoint(cx, cy, rx, ry, 0);
  return new PathBuilder()
    .moveTo(start[0], start[1])
    .append(arcToBeziers(cx, cy, rx, ry, 0, 2 * Math.PI))
    .close()
    .build().segments;
}

/**
 * Rounded rectangle in the `(0,0)`–`(w,h)` box (y-up) with uniform corner radius
 * `r` (clamped to half the shorter side). Traversed CCW: bottom edge → BR corner
 * → right edge → TR corner → top edge → TL corner → left edge → BL corner.
 */
export function roundRectSegments(w: number, h: number, r: number): ReadonlyArray<PathSegment> {
  const rr = Math.max(0, Math.min(r, Math.min(w, h) / 2));
  const b = new PathBuilder();
  if (rr === 0) {
    return b.moveTo(0, 0).lineTo(w, 0).lineTo(w, h).lineTo(0, h).close().build().segments;
  }
  const Q = Math.PI / 2;
  b.moveTo(rr, 0);
  b.lineTo(w - rr, 0);
  b.append(arcToBeziers(w - rr, rr, rr, rr, -Q, Q)); // bottom-right
  b.lineTo(w, h - rr);
  b.append(arcToBeziers(w - rr, h - rr, rr, rr, 0, Q)); // top-right
  b.lineTo(rr, h);
  b.append(arcToBeziers(rr, h - rr, rr, rr, Q, Q)); // top-left
  b.lineTo(0, rr);
  b.append(arcToBeziers(rr, rr, rr, rr, Math.PI, Q)); // bottom-left
  b.close();
  return b.build().segments;
}
