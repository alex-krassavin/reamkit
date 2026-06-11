// PDF vector emission (ISO 32000 §8.5): VectorShape → content-stream
// operators. The model itself is format-agnostic and lives in core/vector.

import type { PathSegment, StrokeStyle, VectorShape } from '@/core/vector';

// Re-export the model so existing '@/pdf/vector-graphics' imports keep
// working; new code should import the model from '@/core/vector'.
export type { PathSegment, StrokeStyle, VectorPath, VectorShape } from '@/core/vector';
export { PathBuilder } from '@/core/vector';

export function emitVectorShape(shape: VectorShape): Array<string> {
  const out: Array<string> = [];
  out.push('q');
  const [a, b, c, d, e, f] = shape.transform;
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
  if (shape.fillColorHex) {
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
  out.push(paintOp(shape.fillColorHex !== undefined, stroke !== undefined, evenodd));
  out.push('Q');
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
