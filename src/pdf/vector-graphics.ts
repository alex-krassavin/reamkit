// PDF vector emission (ISO 32000 §8.5): VectorShape → content-stream
// operators. The model itself is format-agnostic and lives in core/vector.

import type { PathSegment, StrokeStyle, VectorPath, VectorShape } from '@/core/vector';

// Re-export the model so existing '@/pdf/vector-graphics' imports keep
// working; new code should import the model from '@/core/vector'.
export type { PathSegment, StrokeStyle, VectorPath, VectorShape } from '@/core/vector';
export { PathBuilder } from '@/core/vector';

/**
 * Emit a {@link VectorShape} as PDF content-stream operators (ISO 32000 §8.5),
 * wrapped in a `q`/`Q` save-restore with the shape's transform as the CTM. The
 * subpaths are painted with a single operator, so the winding rule is decided
 * once (even-odd if any subpath asks).
 *
 * @param shape       The shape to draw.
 * @param patternName When set, the fill is this named shading pattern from the
 *   page's `/Pattern` resources (a gradient, EP16b) rather than a solid colour;
 *   only used if the shape actually carries a `fillGradient`.
 * @param alphaStateName When set, the shape's shadow is drawn first, under this
 *   named `/ExtGState` (its constant alpha) from the page's resources.
 * @returns The content-stream operator lines.
 */
/**
 * §20.1.8.40 `blurRad` — how a blurred shadow is drawn: as `count` copies of
 * the silhouette, each at `alpha` transparency, growing through the blur's
 * width. Shared with the emitter that registers the /ExtGState for it, so the
 * alpha it registers is the one the layers are drawn at.
 *
 * @param shadow The shadow (its `blurPt` and `alpha` are read).
 * @returns The number of layers and the transparency of each.
 */
export function shadowBlurLayers(shadow: { readonly blurPt: number; readonly alpha: number }): {
  count: number;
  alpha: number;
} {
  // A hard shadow is one copy at its own alpha. Past that, one layer per two
  // points of blur — enough for the eye at any size a document asks for, and
  // few enough that a page of shadowed shapes stays small.
  if (!(shadow.blurPt > 0)) return { count: 1, alpha: shadow.alpha };
  const count = Math.min(8, Math.max(2, Math.round(shadow.blurPt / 2)));
  // Layers COMPOSITE, they do not add: `n` copies at `a` leave `1 − (1 − a)^n`
  // of the shadow, so the share each one carries is the root of what is left,
  // not the quotient. Divided instead, the middle of a shadow never reached the
  // transparency it asked for — a 50% shadow under tdf128596's tile settled at
  // 44% and the shape came out a step lighter than either reference.
  const alpha = 1 - Math.pow(1 - shadow.alpha, 1 / count);
  return { count, alpha };
}

// The local-space bounds of the shape's paths — the frame the blur grows in.
function shadowBbox(
  shape: VectorShape,
): { minX: number; minY: number; maxX: number; maxY: number } | undefined {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const path of shape.paths) {
    for (const seg of path.segments) {
      if (seg.op === 'close') continue;
      minX = Math.min(minX, seg.x);
      minY = Math.min(minY, seg.y);
      maxX = Math.max(maxX, seg.x);
      maxY = Math.max(maxY, seg.y);
    }
  }
  return Number.isFinite(minX) && maxX > minX && maxY > minY
    ? { minX, minY, maxX, maxY }
    : undefined;
}

export function emitVectorShape(
  shape: VectorShape,
  patternName?: string,
  alphaStateName?: string,
  fillAlphaStateName?: string,
): Array<string> {
  const out: Array<string> = [];
  const [a, b, c, d, e, f] = shape.transform;
  // §20.1.8.40 — the shadow is the same geometry, offset, drawn first so the
  // shape lands on top of it. PDF has no blur short of a soft-masked image, so
  // the edge is hard where the source asked for `blurRad`; the displacement,
  // the colour and the transparency are the source's own.
  const shadow = shape.shadow;
  if (shadow) {
    const [sr, sg, sb] = hexToRgb01(shadow.colorHex);
    // §20.1.8.40 `blurRad` — the soft edge, built as a stack of copies of the
    // same silhouette growing through the blur's width, each drawn at a
    // fraction of the shadow's transparency: where they overlap the colour
    // builds back up to what the shadow asked for, and the rim fades out over
    // the radius. PDF has no blur operator short of a soft-masked image.
    const bbox = shadowBbox(shape);
    const scale = Math.sqrt(Math.abs(a * d - b * c)) || 1;
    const { count } = shadowBlurLayers(shadow);
    for (let i = 0; i < count; i++) {
      out.push('q');
      if (alphaStateName !== undefined) out.push(`/${alphaStateName} gs`);
      // The stored CTM maps the shape's local frame onto a y-UP page, so a
      // shadow that falls DOWN the page moves in -y.
      out.push(
        `${num(a)} ${num(b)} ${num(c)} ${num(d)} ` +
          `${num(e + shadow.dxPt)} ${num(f - shadow.dyPt)} cm`,
      );
      if (count > 1 && bbox) {
        // Grow the silhouette about its own centre, from half the blur inside
        // to half outside, in the shape's local units.
        const local = shadow.blurPt / scale;
        const grow = -local / 2 + ((i + 0.5) * local) / count;
        const halfW = Math.max((bbox.maxX - bbox.minX) / 2, 1e-6);
        const halfH = Math.max((bbox.maxY - bbox.minY) / 2, 1e-6);
        const cx = (bbox.minX + bbox.maxX) / 2;
        const cy = (bbox.minY + bbox.maxY) / 2;
        const sx = Math.max((halfW + grow) / halfW, 0);
        const sy = Math.max((halfH + grow) / halfH, 0);
        out.push(`${num(sx)} 0 0 ${num(sy)} ${num(cx * (1 - sx))} ${num(cy * (1 - sy))} cm`);
      }
      out.push(`${num(sr)} ${num(sg)} ${num(sb)} rg`);
      let shadowEvenodd = false;
      for (const path of shape.paths) {
        if (path.fillRule === 'evenodd') shadowEvenodd = true;
        for (const seg of path.segments) out.push(emitSegment(seg));
      }
      out.push(shadowEvenodd ? 'f*' : 'f');
      out.push('Q');
    }
  }
  out.push('q');
  // §20.1.2.3.1 — a fill the document made TRANSPARENT is drawn transparent,
  // not composited over the paper: what stands behind the shape shows through
  // it, which is the whole point of the fill saying so.
  if (fillAlphaStateName !== undefined) out.push(`/${fillAlphaStateName} gs`);
  out.push(`${num(a)} ${num(b)} ${num(c)} ${num(d)} ${num(e)} ${num(f)} cm`);

  const stroke = shape.stroke;
  if (stroke) {
    out.push(`${num(stroke.widthPt)} w`);
    if (stroke.cap !== undefined) out.push(`${capCode(stroke.cap)} J`);
    if (stroke.join !== undefined) out.push(`${joinCode(stroke.join)} j`);
    if (stroke.dash && stroke.dash.length > 0) {
      out.push(`[${stroke.dash.map(num).join(' ')}] 0 d`);
    }
    const [r, g, bl] = hexToRgb01(stroke.colorHex);
    out.push(`${num(r)} ${num(g)} ${num(bl)} RG`);
  }
  // A gradient (EP16b) is a shading pattern set as the non-stroking colour;
  // otherwise a solid colour. Either way the path is painted with the same fill
  // operator below.
  const usePattern = patternName !== undefined && shape.fillGradient !== undefined;
  if (usePattern) {
    out.push(`/Pattern cs /${patternName} scn`);
  } else if (shape.fillColorHex) {
    const [r, g, bl] = hexToRgb01(shape.fillColorHex);
    out.push(`${num(r)} ${num(g)} ${num(bl)} rg`);
  }

  // A single painting operator covers every subpath constructed since the last
  // paint, so the winding rule is decided once: evenodd if any subpath asks.
  let evenodd = false;
  for (const path of shape.paths) {
    if (path.fillRule === 'evenodd') evenodd = true;
    for (const seg of path.segments) out.push(emitSegment(seg));
  }
  out.push(paintOp(usePattern || shape.fillColorHex !== undefined, stroke !== undefined, evenodd));
  out.push('Q');
  return out;
}

/**
 * The operators that make `paths` the current CLIP, under `transform` — the
 * same local→y-up-page matrix {@link emitVectorShape} takes. Emitted inside the
 * caller's own `q`…`Q`, and followed by the inverse matrix so what comes after
 * is back in page coordinates.
 *
 * @param paths     The clip outline, in the shape's local frame.
 * @param transform The local→page matrix.
 * @returns The operators: `cm`, the path, `W n`, and the inverse `cm`.
 */
export function emitClipPath(
  paths: ReadonlyArray<VectorPath>,
  transform: readonly [number, number, number, number, number, number],
): Array<string> {
  const [a, b, c, d, e, f] = transform;
  const det = a * d - b * c;
  if (det === 0) return [];
  const out: Array<string> = [`${num(a)} ${num(b)} ${num(c)} ${num(d)} ${num(e)} ${num(f)} cm`];
  let evenodd = false;
  for (const path of paths) {
    if (path.fillRule === 'evenodd') evenodd = true;
    for (const seg of path.segments) out.push(emitSegment(seg));
  }
  out.push(evenodd ? 'W* n' : 'W n');
  // …and back to page coordinates for whatever the caller draws next.
  const ia = d / det;
  const ib = -b / det;
  const ic = -c / det;
  const id = a / det;
  out.push(
    `${num(ia)} ${num(ib)} ${num(ic)} ${num(id)} ` +
      `${num(-(e * ia + f * ic))} ${num(-(e * ib + f * id))} cm`,
  );
  return out;
}

function emitSegment(seg: PathSegment): string {
  switch (seg.op) {
    case 'move':
      return `${num(seg.x)} ${num(seg.y)} m`;
    case 'line':
      return `${num(seg.x)} ${num(seg.y)} l`;
    case 'cubic':
      return (
        `${num(seg.x1)} ${num(seg.y1)} ${num(seg.x2)} ${num(seg.y2)} ` +
        `${num(seg.x)} ${num(seg.y)} c`
      );
    case 'close':
      return 'h';
  }
}

// ISO 32000-1 §8.5.3 — pick the path-painting operator.
//   fill+stroke → B/B*   fill → f/f*   stroke → S   neither → n (no-op)
function paintOp(fill: boolean, stroke: boolean, evenodd: boolean): string {
  if (fill && stroke) return evenodd ? 'B*' : 'B';
  if (fill) return evenodd ? 'f*' : 'f';
  if (stroke) return 'S';
  return 'n';
}

const capCode = (c: 'butt' | 'round' | 'square'): number =>
  c === 'round' ? 1 : c === 'square' ? 2 : 0;

const joinCode = (j: 'miter' | 'round' | 'bevel'): number =>
  j === 'round' ? 1 : j === 'bevel' ? 2 : 0;

// Content-stream number formatter: fixed-precision, trailing zeros trimmed.
// Six decimals keep rotation-matrix coefficients (cos/sin) accurate to well
// under a point over a full page. Guards non-finite values, which would make
// the serializer throw and produce an invalid stream.
function num(n: number): string {
  if (!Number.isFinite(n)) return '0';
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
}

function hexToRgb01(hex: string): readonly [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  return [r, g, b];
}
