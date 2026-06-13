// Shape model → vector primitives (paths, stroke, placement matrix) — the
// format-agnostic half of shape rendering, shared by the layout engine (PDF /
// SVG via PageDoc) and the HTML writer (inline <svg>). Pure math: no fonts,
// no pagination.

import type { ShapeDash, ShapeGeometry, ShapeLine } from '@/core/document-model';
import type { ShapeGradient, StrokeStyle, VectorPath } from '@/core/vector';

import { customPaths, presetPaths, rectPath } from '@/core/drawingml/preset-geometry';

// 1 inch = 914400 EMU = 72 pt, so 1 pt = 12700 EMU.
export const EMU_PER_PT = 12700;

// a:ln default width when @w is absent (9525 EMU = 0.75pt).
const DEFAULT_LINE_WIDTH_EMU = 9525;

// Word's default text-box insets (§20.1.2.1) — 0.1" L/R, 0.05" T/B in EMU.
export const DEFAULT_INSET_LR_PT = 91440 / EMU_PER_PT;
export const DEFAULT_INSET_TB_PT = 45720 / EMU_PER_PT;

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

// A gradient's solid approximation: the per-channel average of its stop colours
// (§EP16). Writers without gradient support (the plain PDF emitter) paint this.
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

// An SVG `<linearGradient>` / `<radialGradient>` definition for a gradient fill
// (EP16), shared by the SVG and HTML writers. The linear vector is expressed in
// objectBoundingBox space; the angle is negated because the shape's own path
// transform flips y (local y-up → page y-down).
export function gradientSvgDef(id: string, g: ShapeGradient): string {
  const n = (x: number): string => String(Math.round(x * 1e4) / 1e4);
  const stops = g.stops
    .map((s) => `<stop offset="${n(s.offset)}" stop-color="#${s.colorHex}"/>`)
    .join('');
  if (g.kind === 'radial') return `<radialGradient id="${id}">${stops}</radialGradient>`;
  const rad = (-(g.angle ?? 0) * Math.PI) / 180;
  const dx = Math.cos(rad) / 2;
  const dy = Math.sin(rad) / 2;
  return (
    `<linearGradient id="${id}" x1="${n(0.5 - dx)}" y1="${n(0.5 - dy)}" ` +
    `x2="${n(0.5 + dx)}" y2="${n(0.5 + dy)}">${stops}</linearGradient>`
  );
}

export function buildStroke(line: ShapeLine | undefined): StrokeStyle | undefined {
  if (!line || line.fill === 'none') return undefined;
  const widthPt = line.width ?? DEFAULT_LINE_WIDTH_EMU / EMU_PER_PT;
  const dash = line.dash && line.dash !== 'solid' ? dashPattern(line.dash, widthPt) : undefined;
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

// DrawingML rot is clockwise in y-down space ⇒ a negative angle in PDF y-up.
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
