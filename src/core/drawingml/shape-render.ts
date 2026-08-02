// Shape model → vector primitives (paths, stroke, placement matrix) — the
// format-agnostic half of shape rendering, shared by the layout engine (PDF /
// SVG via PageDoc) and the HTML writer (inline <svg>). Pure math: no fonts,
// no pagination.

import type { ShapeDash, ShapeGeometry, ShapeLine } from '@/core/document-model';
import type { PathSegment, ShapeGradient, StrokeStyle, VectorPath } from '@/core/vector';

import { customPaths, presetPaths, rectPath } from '@/core/drawingml/preset-geometry';

/** EMU per point: 1 inch = 914400 EMU = 72 pt, so 1 pt = 12700 EMU. */
export const EMU_PER_PT = 12700;

// a:ln default width when @w is absent (9525 EMU = 0.75pt).
const DEFAULT_LINE_WIDTH_EMU = 9525;

/** Word's default left/right text-box inset (§20.1.2.1) — 0.1", in points. */
export const DEFAULT_INSET_LR_PT = 91440 / EMU_PER_PT;
/** Word's default top/bottom text-box inset (§20.1.2.1) — 0.05", in points. */
export const DEFAULT_INSET_TB_PT = 45720 / EMU_PER_PT;

/**
 * Build the vector path(s) for a shape's geometry, sized `widthPt`×`heightPt`.
 * Dispatches preset geometries to {@link presetPaths} (falling back to the
 * bounding rectangle for unknown presets) and custom geometries to
 * {@link customPaths}.
 *
 * @param geometry The shape geometry (preset or custom).
 * @param widthPt  Box width in points.
 * @param heightPt Box height in points.
 * @returns The path(s) in the local y-up frame.
 */
export function buildShapePaths(
  geometry: ShapeGeometry,
  widthPt: number,
  heightPt: number,
): Array<VectorPath> {
  if (geometry.kind === 'preset') {
    const paths = presetPaths(
      geometry.preset ?? 'rect',
      widthPt,
      heightPt,
      geometry.adjust ?? new Map(),
    );
    return paths ?? [rectPath(widthPt, heightPt)];
  }
  if (geometry.custom) return customPaths(geometry.custom, widthPt, heightPt);
  return [rectPath(widthPt, heightPt)];
}

/**
 * A gradient's solid approximation: the per-channel average of its stop colours
 * (EP16). Writers without gradient support (the plain PDF emitter) paint this.
 *
 * @param gradient The gradient fill.
 * @returns The averaged colour as uppercase RRGGBB (`'000000'` if no valid stops).
 */
export function gradientToSolid(gradient: ShapeGradient): string {
  let r = 0;
  let g = 0;
  let b = 0;
  let n = 0;
  for (const stop of gradient.stops) {
    const num = parseInt(stop.colorHex, 16);
    if (Number.isNaN(num)) continue;
    r += (num >> 16) & 255;
    g += (num >> 8) & 255;
    b += num & 255;
    n++;
  }
  if (n === 0) return '000000';
  const hx = (x: number): string =>
    Math.round(x / n)
      .toString(16)
      .padStart(2, '0');
  return (hx(r) + hx(g) + hx(b)).toUpperCase();
}

/**
 * An SVG `<linearGradient>` / `<radialGradient>` definition for a gradient fill
 * (EP16), shared by the SVG and HTML writers. The linear vector is expressed in
 * `objectBoundingBox` space; the angle is negated because the shape's own path
 * transform flips y (local y-up → page y-down).
 *
 * @param id The gradient element id (referenced by `fill="url(#id)"`).
 * @param g  The gradient fill.
 * @returns The `<linearGradient>` or `<radialGradient>` markup.
 */
export function gradientSvgDef(id: string, g: ShapeGradient): string {
  const n = (x: number): string => String(Math.round(x * 1e4) / 1e4);
  const stops = g.stops
    .map((s) => `<stop offset="${n(s.offset)}" stop-color="#${s.colorHex}"/>`)
    .join('');
  if (g.kind === 'radial') {
    // The centre is in the box's own fractions, which is exactly what SVG's
    // objectBoundingBox units are; the sweep reaches half the diagonal, as the
    // centred case does.
    const c = g.center;
    if (!c) return `<radialGradient id="${id}">${stops}</radialGradient>`;
    const r = Math.SQRT2 / 2;
    return (
      `<radialGradient id="${id}" cx="${n(c.x)}" cy="${n(c.y)}" r="${n(r)}">` +
      `${stops}</radialGradient>`
    );
  }
  const rad = (-(g.angle ?? 0) * Math.PI) / 180;
  const dx = Math.cos(rad) / 2;
  const dy = Math.sin(rad) / 2;
  return (
    `<linearGradient id="${id}" x1="${n(0.5 - dx)}" y1="${n(0.5 - dy)}" ` +
    `x2="${n(0.5 + dx)}" y2="${n(0.5 + dy)}">${stops}</linearGradient>`
  );
}

// §20.1.10.34/§20.1.10.35 — the three size steps an end decoration comes in,
// as multiples of the pen. A hairline arrow drawn at 3× a 0.75pt pen is a
// speck, so the pen is taken as at least a point: the head then lands near the
// 4-5pt Word draws, well short of LibreOffice's 9.
const LINE_END_STEPS: Record<'sm' | 'med' | 'lg', number> = { sm: 3, med: 4.5, lg: 6 };
const MIN_LINE_END_PEN_PT = 1;

/**
 * §20.1.8.24 / §20.1.8.42 — the arrowheads a line asks for at its ends, as
 * filled paths in the same local frame as the shape's own geometry. The head
 * rides the first point of the first subpath and the tail the last point of
 * the last one, each turned along the segment that reaches it.
 *
 * @param paths        The shape's geometry paths.
 * @param line         The parsed line (its `headEnd`/`tailEnd` are read).
 * @param strokeWidthPt The pen width the decorations are sized against.
 * @returns The decoration paths (empty when the line asks for none).
 */
export function lineEndPaths(
  paths: ReadonlyArray<VectorPath>,
  line: ShapeLine | undefined,
  strokeWidthPt: number,
): Array<VectorPath> {
  if (!line || (!line.headEnd && !line.tailEnd) || line.fill === 'none') return [];
  const pts = pathPoints(paths);
  if (pts.length < 2) return [];
  const out: Array<VectorPath> = [];
  const pen = Math.max(strokeWidthPt, MIN_LINE_END_PEN_PT);
  if (line.headEnd) {
    const p = pts[0]!;
    const q = pts[1]!;
    const path = endPath(line.headEnd, p, q, pen);
    if (path) out.push(path);
  }
  if (line.tailEnd) {
    const p = pts[pts.length - 1]!;
    const q = pts[pts.length - 2]!;
    const path = endPath(line.tailEnd, p, q, pen);
    if (path) out.push(path);
  }
  return out;
}

// The path's points in order, ignoring which operator produced them: enough to
// know where a line starts and ends and which way it is going there.
function pathPoints(paths: ReadonlyArray<VectorPath>): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  for (const path of paths) {
    for (const seg of path.segments) {
      if (seg.op === 'close') continue;
      out.push({ x: seg.x, y: seg.y });
    }
  }
  return out;
}

// One decoration at `tip`, pointing away from `from`.
function endPath(
  end: NonNullable<ShapeLine['headEnd']>,
  tip: { x: number; y: number },
  from: { x: number; y: number },
  pen: number,
): VectorPath | undefined {
  const dx = tip.x - from.x;
  const dy = tip.y - from.y;
  const len = Math.hypot(dx, dy);
  if (len < 1e-6) return undefined;
  // Unit vector along the line (ux, uy) and its normal (nx, ny).
  const ux = dx / len;
  const uy = dy / len;
  const nx = -uy;
  const ny = ux;
  const L = LINE_END_STEPS[end.length ?? 'med'] * pen;
  const W = LINE_END_STEPS[end.width ?? 'med'] * pen;
  const at = (along: number, across: number): { x: number; y: number } => ({
    x: tip.x - ux * along + nx * across,
    y: tip.y - uy * along + ny * across,
  });
  const segments: Array<PathSegment> = [];
  const move = (p: { x: number; y: number }): void => {
    segments.push({ op: 'move', x: p.x, y: p.y });
  };
  const line2 = (p: { x: number; y: number }): void => {
    segments.push({ op: 'line', x: p.x, y: p.y });
  };
  if (end.type === 'diamond') {
    move(at(0, 0));
    line2(at(L / 2, W / 2));
    line2(at(L, 0));
    line2(at(L / 2, -W / 2));
  } else if (end.type === 'oval') {
    // A circle centred half a length back, drawn as four arcs' worth of
    // Béziers would be exact; the octagon below is within a pen width of it.
    const r = W / 2;
    const c = at(L / 2, 0);
    for (let i = 0; i < 8; i++) {
      const a = (i / 8) * 2 * Math.PI;
      const p = { x: c.x + r * Math.cos(a), y: c.y + r * Math.sin(a) };
      if (i === 0) move(p);
      else line2(p);
    }
  } else if (end.type === 'stealth') {
    // A triangle with its back notched forward — the concave arrowhead.
    move(at(0, 0));
    line2(at(L, W / 2));
    line2(at(L * 0.6, 0));
    line2(at(L, -W / 2));
  } else {
    // 'triangle' and 'arrow' both draw as a filled triangle; Word's open
    // `arrow` is two strokes, which at these sizes reads the same.
    move(at(0, 0));
    line2(at(L, W / 2));
    line2(at(L, -W / 2));
  }
  segments.push({ op: 'close' });
  return { segments };
}

/**
 * Build a {@link StrokeStyle} from a shape's `a:ln` line, resolving the default
 * width (0.75pt), dash pattern and line cap. Returns `undefined` for no line or
 * an explicit no-fill stroke.
 *
 * @param line The parsed line, or `undefined`.
 * @returns The stroke style, or `undefined` when nothing should be stroked.
 */
export function buildStroke(line: ShapeLine | undefined): StrokeStyle | undefined {
  if (!line || line.fill === 'none') return undefined;
  const widthPt = line.width ?? DEFAULT_LINE_WIDTH_EMU / EMU_PER_PT;
  // §20.1.8.21 — the author's own pattern states its lengths as multiples of
  // the line's width, and it wins over any preset beside it.
  const custom = line.customDash?.map((n) => Math.max(0.01, n * widthPt));
  const dash =
    custom ?? (line.dash && line.dash !== 'solid' ? dashPattern(line.dash, widthPt) : undefined);
  // DrawingML 'flat' cap is PDF butt; round/square map straight through.
  const cap: StrokeStyle['cap'] | undefined = line.cap === 'flat' ? 'butt' : line.cap;
  return {
    colorHex: line.colorHex ?? '000000',
    widthPt,
    ...(dash ? { dash } : {}),
    ...(cap ? { cap } : {}),
  };
}

// Dash patterns expressed in multiples of the line width (a common rendering
// convention), in points. 'solid' has no pattern.
function dashPattern(dash: ShapeDash, w: number): Array<number> | undefined {
  const u = Math.max(w, 0.1);
  switch (dash) {
    case 'solid':
      return undefined;
    case 'dot':
      return [u, 2 * u];
    case 'dash':
      return [4 * u, 3 * u];
    case 'dashDot':
      return [4 * u, 3 * u, u, 3 * u];
    case 'lgDash':
      return [8 * u, 3 * u];
    case 'lgDashDot':
      return [8 * u, 3 * u, u, 3 * u];
    case 'sysDash':
      return [3 * u, u];
    case 'sysDot':
      return [u, u];
  }
}

/**
 * Build the 2×3 affine placement matrix `[a, b, c, d, e, f]` that positions a
 * shape's local y-up box at `(pageX, pageY)`, applying rotation and h/v flips
 * about the box centre. DrawingML `rot` is clockwise in y-down space ⇒ a
 * negative angle in PDF y-up.
 *
 * @param pageX       Box left, in page points.
 * @param pageY       Box bottom, in page points.
 * @param widthPt     Box width in points.
 * @param heightPt    Box height in points.
 * @param rotation60k Rotation in 1/60000°, clockwise.
 * @param flipH       Mirror horizontally.
 * @param flipV       Mirror vertically.
 * @returns The affine matrix as `[a, b, c, d, e, f]`.
 */
export function buildShapeTransform(
  pageX: number,
  pageY: number,
  widthPt: number,
  heightPt: number,
  rotation60k: number,
  flipH: boolean,
  flipV: boolean,
): [number, number, number, number, number, number] {
  const theta = (-rotation60k / 60000) * (Math.PI / 180);
  const sx = flipH ? -1 : 1;
  const sy = flipV ? -1 : 1;
  const cos = Math.cos(theta);
  const sin = Math.sin(theta);
  const a = sx * cos;
  const b = sx * sin;
  const c = -sy * sin;
  const d = sy * cos;
  const cxL = widthPt / 2;
  const cyL = heightPt / 2;
  const centerX = pageX + cxL;
  const centerY = pageY + cyL;
  const e = centerX - (a * cxL + c * cyL);
  const f = centerY - (b * cxL + d * cyL);
  return [a, b, c, d, e, f];
}
