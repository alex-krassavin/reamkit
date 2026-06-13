// E-PDF EP16b — emit a DrawingML gradient fill as a PDF shading pattern
// (ISO 32000-1 §8.7.4.5 PatternType 2 + §8.7.4.5.x axial/radial shadings). The
// pattern's /Matrix is the shape's CTM and the shading /Coords are in the
// shape's local space, so the gradient maps through the same transform as the
// path — handled like an ordinary non-stroking colour (`/Pattern cs /Pn scn`).
//
// The colour ramp is an exact stitching function (type 3 of type-2 exponential
// segments), so the stops round-trip without sampling. DeviceRGB is used, so the
// caller must keep the solid-colour fallback for PDF/A (where a device colour
// space needs the OutputIntent).

import type { ShapeGradient, VectorShape } from '@/core/vector';
import type { PdfDict, PdfRef, PdfValue } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';

import { dict, name, ref } from '@/pdf/objects';

interface Bbox {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
}

// The local-space bounding box of a shape's paths (y-up). Returns undefined for
// an empty/degenerate shape (no gradient can be placed).
export function shapeBbox(shape: VectorShape): Bbox | undefined {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const add = (x: number, y: number): void => {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  };
  for (const path of shape.paths) {
    for (const seg of path.segments) {
      if (seg.op === 'move' || seg.op === 'line') add(seg.x, seg.y);
      else if (seg.op === 'cubic') {
        add(seg.x1, seg.y1);
        add(seg.x2, seg.y2);
        add(seg.x, seg.y);
      }
    }
  }
  return Number.isFinite(minX) && maxX > minX && maxY > minY
    ? { minX, minY, maxX, maxY }
    : undefined;
}

// Build the PatternType-2 shading pattern object and return its ref.
export function buildGradientPattern(
  doc: PdfDocument,
  gradient: ShapeGradient,
  bbox: Bbox,
  ctm: readonly [number, number, number, number, number, number],
): PdfRef {
  const fnRef = doc.add(buildRamp(gradient.stops));
  const shading =
    gradient.kind === 'radial'
      ? radialShading(bbox, fnRef)
      : axialShading(gradient.angle ?? 0, bbox, fnRef);
  return doc.add(
    dict({ Type: name('Pattern'), PatternType: 2, Shading: shading, Matrix: [...ctm] }),
  );
}

// §8.7.4.5.3 axial shading. The local frame is y-up, so a DrawingML angle θ
// (clockwise from +x in a y-down frame) points along (cos θ, −sin θ). The axis
// spans the bbox: project its corners onto the direction for the extent.
function axialShading(angleDeg: number, b: Bbox, fnRef: PdfRef): PdfDict {
  const rad = (angleDeg * Math.PI) / 180;
  const dx = Math.cos(rad);
  const dy = -Math.sin(rad);
  const cx = (b.minX + b.maxX) / 2;
  const cy = (b.minY + b.maxY) / 2;
  const cproj = cx * dx + cy * dy;
  let tmin = Infinity;
  let tmax = -Infinity;
  for (const [x, y] of [
    [b.minX, b.minY],
    [b.maxX, b.minY],
    [b.minX, b.maxY],
    [b.maxX, b.maxY],
  ] as const) {
    const t = x * dx + y * dy;
    tmin = Math.min(tmin, t);
    tmax = Math.max(tmax, t);
  }
  return dict({
    ShadingType: 2,
    ColorSpace: name('DeviceRGB'),
    Coords: [
      cx + (tmin - cproj) * dx,
      cy + (tmin - cproj) * dy,
      cx + (tmax - cproj) * dx,
      cy + (tmax - cproj) * dy,
    ],
    Function: ref(fnRef.id),
    Extend: [true, true],
  });
}

// §8.7.4.5.4 radial shading: a point at the centre (r=0) growing to a circle at
// half the bbox diagonal. Param 0 (first stop) is the centre, matching a:path.
function radialShading(b: Bbox, fnRef: PdfRef): PdfDict {
  const cx = (b.minX + b.maxX) / 2;
  const cy = (b.minY + b.maxY) / 2;
  const r = Math.hypot(b.maxX - b.minX, b.maxY - b.minY) / 2;
  return dict({
    ShadingType: 3,
    ColorSpace: name('DeviceRGB'),
    Coords: [cx, cy, 0, cx, cy, r],
    Function: ref(fnRef.id),
    Extend: [true, true],
  });
}

// §7.10.4 — the colour ramp. One stop → a constant type-2 function; otherwise a
// type-3 stitching function over type-2 (linear) segments between the stops.
function buildRamp(stopsIn: ShapeGradient['stops']): PdfDict {
  const stops = normalizeStops(stopsIn);
  if (stops.length <= 2) {
    // One stop → a constant colour; two stops → a single linear segment.
    return dict({
      FunctionType: 2,
      Domain: [0, 1],
      C0: rgb01(stops[0]!.colorHex),
      C1: rgb01(stops[stops.length - 1]!.colorHex),
      N: 1,
    });
  }
  const functions: Array<PdfValue> = [];
  const bounds: Array<number> = [];
  const encode: Array<number> = [];
  for (let i = 0; i < stops.length - 1; i++) {
    functions.push(
      dict({
        FunctionType: 2,
        Domain: [0, 1],
        C0: rgb01(stops[i]!.colorHex),
        C1: rgb01(stops[i + 1]!.colorHex),
        N: 1,
      }),
    );
    if (i > 0) bounds.push(stops[i]!.offset);
    encode.push(0, 1);
  }
  return dict({
    FunctionType: 3,
    Domain: [0, 1],
    Functions: functions,
    Bounds: bounds,
    Encode: encode,
  });
}

// Sort by offset, drop NaNs, clamp into [0,1], and pin the endpoints to 0 and 1
// so the function domain is fully covered.
function normalizeStops(
  stops: ShapeGradient['stops'],
): Array<{ offset: number; colorHex: string }> {
  const out = stops
    .map((s) => ({ offset: Math.max(0, Math.min(1, s.offset)), colorHex: s.colorHex }))
    .sort((a, b) => a.offset - b.offset);
  if (out.length === 0) return [{ offset: 0, colorHex: '000000' }];
  out[0] = { offset: 0, colorHex: out[0]!.colorHex };
  out[out.length - 1] = { offset: 1, colorHex: out[out.length - 1]!.colorHex };
  return out;
}

function rgb01(hex: string): Array<number> {
  const n = parseInt(hex, 16);
  if (Number.isNaN(n)) return [0, 0, 0];
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
}
