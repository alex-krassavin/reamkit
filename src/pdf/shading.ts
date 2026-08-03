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

import { dict, name, ref, stream } from '@/pdf/objects';

interface Bbox {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
}

/**
 * The local-space bounding box of a shape's paths (y-up).
 *
 * @returns The bounding box, or `undefined` for an empty/degenerate shape (where
 *   no gradient can be placed).
 */
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

/**
 * Build the PatternType-2 shading pattern object (axial or radial) for a
 * DrawingML gradient (EP16b) and add it to `doc`. The pattern's `/Matrix` is the
 * shape's CTM and the shading `/Coords` are in the shape's local space.
 *
 * @param gradient The gradient fill (kind, angle, colour stops).
 * @param bbox     The shape's local-space bounding box, from {@link shapeBbox}.
 * @param ctm      The shape's current transformation matrix, used as `/Matrix`.
 * @returns A reference to the `/Pattern` object, for a `/Pattern cs /Pn scn` fill.
 */
export function buildGradientPattern(
  doc: PdfDocument,
  gradient: ShapeGradient,
  bbox: Bbox,
  ctm: readonly [number, number, number, number, number, number],
): PdfRef {
  const shading =
    gradient.kind === 'radial' && gradient.sweep === 'rect'
      ? rectShading(doc, bbox, gradient.stops, gradient.center)
      : gradient.kind === 'radial'
        ? radialShading(bbox, doc.add(buildRamp(gradient.stops)), gradient.center)
        : axialShading(gradient.angle ?? 0, bbox, doc.add(buildRamp(gradient.stops)));
  return doc.add(
    dict({ Type: name('Pattern'), PatternType: 2, Shading: shading, Matrix: [...ctm] }),
  );
}

// §8.7.4.5.3 axial shading. The local frame is y-up, so a DrawingML angle θ
// (clockwise from +x in a y-down frame) points along (cos θ, −sin θ). The axis
// spans the bbox: project its corners onto the direction for the extent.
function axialShading(angleDeg: number, b: Bbox, fnRef: PdfRef, cs = 'DeviceRGB'): PdfDict {
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
    ColorSpace: name(cs),
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
function radialShading(
  b: Bbox,
  fnRef: PdfRef,
  center?: { readonly x: number; readonly y: number },
  cs = 'DeviceRGB',
): PdfDict {
  // The centre is given in the box's own fractions, y DOWN (a `0,0` centre is
  // the top-left corner); this frame is y-up.
  const cx = center ? b.minX + center.x * (b.maxX - b.minX) : (b.minX + b.maxX) / 2;
  const cy = center ? b.maxY - center.y * (b.maxY - b.minY) : (b.minY + b.maxY) / 2;
  // Half the diagonal, wherever the centre sits: the sweep out of a corner
  // reaches the middle of the box and the last stop holds the rest, which is
  // what both references draw for fill.docx.
  const r = Math.hypot(b.maxX - b.minX, b.maxY - b.minY) / 2;
  return dict({
    ShadingType: 3,
    ColorSpace: name(cs),
    Coords: [cx, cy, 0, cx, cy, r],
    Function: ref(fnRef.id),
    Extend: [true, true],
  });
}

// §8.7.4.5.2 function-based shading — the sweep VML asks for, whose contours
// are RECTANGLES growing out of the focus until they cover the box. PDF has no
// shading type for it (2 is a line, 3 is a circle), so the field is written out
// as a calculator function: the parameter at a point is how far it stands from
// the focus along whichever axis is further, and the ramp is evaluated on that.
function rectShading(
  doc: PdfDocument,
  b: Bbox,
  stopsIn: ShapeGradient['stops'],
  center?: { readonly x: number; readonly y: number },
): PdfDict {
  // The centre is given in the box's own fractions, y DOWN; this frame is y-up.
  const cx = center ? b.minX + center.x * (b.maxX - b.minX) : (b.minX + b.maxX) / 2;
  const cy = center ? b.maxY - center.y * (b.maxY - b.minY) : (b.minY + b.maxY) / 2;
  // The rectangle reaches its last stop half a box from the focus, wherever
  // that focus sits — the same reach the centred circle has always had, and
  // the extent both references draw for fill.docx's corner sweep.
  const halfW = Math.max((b.maxX - b.minX) / 2, 1e-6);
  const halfH = Math.max((b.maxY - b.minY) / 2, 1e-6);
  const domain = [b.minX, b.maxX, b.minY, b.maxY];
  const fn = doc.add(
    stream(
      { FunctionType: 4, Domain: domain, Range: [0, 1, 0, 1, 0, 1] },
      new TextEncoder().encode(rectSweepProgram(cx, cy, 1 / halfW, 1 / halfH, stopsIn)),
    ),
  );
  return dict({
    ShadingType: 1,
    ColorSpace: name('DeviceRGB'),
    Domain: domain,
    Function: ref(fn.id),
  });
}

// The calculator program (§7.10.5): x y in, r g b out. Written by hand rather
// than sampled, so the stops stay exact.
function rectSweepProgram(
  cx: number,
  cy: number,
  sx: number,
  sy: number,
  stopsIn: ShapeGradient['stops'],
): string {
  const stops = normalizeStops(stopsIn);
  const body = [
    '{',
    // x y → the distance from the focus along each axis, in units of the far
    // side, then the larger of the two (that is the rectangle the point is on).
    `${ps(cy)} sub abs ${ps(sy)} mul`,
    `exch ${ps(cx)} sub abs ${ps(sx)} mul`,
    '2 copy lt { exch } if pop',
    'dup 1 gt { pop 1 } if',
    rampProgram(stops, 0),
    '}',
  ];
  return body.join('\n');
}

// The ramp as nested conditionals over the stop offsets: `t` on the stack in,
// three colour components out.
function rampProgram(
  stops: ReadonlyArray<{ offset: number; colorHex: string }>,
  i: number,
): string {
  const last = i >= stops.length - 2;
  const a = stops[i]!;
  const b = stops[i + 1]!;
  const seg = segmentProgram(a, b);
  if (last) return seg;
  return `dup ${ps(b.offset)} le { ${seg} } { ${rampProgram(stops, i + 1)} } ifelse`;
}

// One linear segment: `t` in, the colour between the two stops out.
function segmentProgram(
  a: { offset: number; colorHex: string },
  b: { offset: number; colorHex: string },
): string {
  const c0 = rgb01(a.colorHex);
  const c1 = rgb01(b.colorHex);
  const span = b.offset - a.offset;
  if (span <= 1e-9) return `pop ${c1.map(ps).join(' ')}`;
  return [
    `${ps(a.offset)} sub ${ps(1 / span)} mul`,
    'dup 0 lt { pop 0 } if dup 1 gt { pop 1 } if',
    'dup dup',
    `${ps(c1[0]! - c0[0]!)} mul ${ps(c0[0]!)} add`,
    '3 1 roll',
    `${ps(c1[1]! - c0[1]!)} mul ${ps(c0[1]!)} add`,
    'exch',
    `${ps(c1[2]! - c0[2]!)} mul ${ps(c0[2]!)} add`,
  ].join(' ');
}

// A number the calculator will read: fixed notation, never an exponent.
function ps(n: number): string {
  if (!Number.isFinite(n)) return '0';
  const s = n.toFixed(6).replace(/0+$/u, '').replace(/\.$/u, '');
  return s === '' || s === '-0' ? '0' : s;
}

/**
 * §11.6.5.2 — the luminosity soft mask a gradient with per-stop TRANSPARENCY is
 * painted through.
 *
 * A PDF shading has one colour per point and no alpha, so a gradient that fades
 * out is drawn as the colour shading masked by a second shading of the same
 * geometry whose GREY is the transparency: white where the stop is opaque,
 * black where it is clear. 45541_Footer's master lays a white band down the
 * left of every slide that starts half transparent, and painted flat it covered
 * the artwork behind it.
 *
 * @param doc      The document to add the objects to.
 * @param gradient The gradient fill.
 * @param bbox     The shape's local-space bounding box.
 * @param ctm      The shape's CTM, as the form's `/Matrix`.
 * @returns The mask form XObject, or `undefined` when every stop is opaque or
 *          the sweep has no shading type that can carry a mask.
 */
export function buildGradientAlphaMask(
  doc: PdfDocument,
  gradient: ShapeGradient,
  bbox: Bbox,
  ctm: readonly [number, number, number, number, number, number],
): PdfRef | undefined {
  if (!gradient.stops.some((s) => (s.alpha ?? 1) < 1)) return undefined;
  // A `rect` sweep is a calculator function over RGB; masking it would mean a
  // second one over grey, which no corpus file has asked for.
  if (gradient.kind === 'radial' && gradient.sweep === 'rect') return undefined;
  const fn = doc.add(buildAlphaRamp(gradient.stops));
  const shading =
    gradient.kind === 'radial'
      ? radialShading(bbox, fn, gradient.center, 'DeviceGray')
      : axialShading(gradient.angle ?? 0, bbox, fn, 'DeviceGray');
  return doc.add(
    stream(
      {
        Type: name('XObject'),
        Subtype: name('Form'),
        BBox: [bbox.minX, bbox.minY, bbox.maxX, bbox.maxY],
        Matrix: [...ctm],
        Group: dict({ S: name('Transparency'), CS: name('DeviceGray') }),
        Resources: dict({ Shading: dict({ Sm: doc.add(shading) }) }),
      },
      new TextEncoder().encode('q /Sm sh Q'),
    ),
  );
}

// The same stitching as the colour ramp, over one grey component: the stop's
// own transparency.
function buildAlphaRamp(stopsIn: ShapeGradient['stops']): PdfDict {
  const stops = normalizeStops(stopsIn);
  if (stops.length <= 2) {
    return dict({
      FunctionType: 2,
      Domain: [0, 1],
      C0: [stops[0]!.alpha],
      C1: [stops[stops.length - 1]!.alpha],
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
        C0: [stops[i]!.alpha],
        C1: [stops[i + 1]!.alpha],
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
): Array<{ offset: number; colorHex: string; alpha: number }> {
  const out = stops
    .map((s) => ({
      offset: Math.max(0, Math.min(1, s.offset)),
      colorHex: s.colorHex,
      alpha: Math.max(0, Math.min(1, s.alpha ?? 1)),
    }))
    .sort((a, b) => a.offset - b.offset);
  if (out.length === 0) return [{ offset: 0, colorHex: '000000', alpha: 1 }];
  out[0] = { ...out[0]!, offset: 0 };
  out[out.length - 1] = { ...out[out.length - 1]!, offset: 1 };
  return out;
}

function rgb01(hex: string): Array<number> {
  const n = parseInt(hex, 16);
  if (Number.isNaN(n)) return [0, 0, 0];
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
}
