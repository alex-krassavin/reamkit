// Vector-graphics emit layer (ISO 32000-1 §8.5 — Path Construction & Painting).
//
// This is the reusable foundation under DrawingML shapes (and, later, charts).
// It owns nothing about OOXML: callers hand it geometry already reduced to
// straight + cubic-Bézier segments in a LOCAL frame (bottom-left origin,
// y-up — the same orientation as PDF user space), plus a paint description and
// an affine `transform` that places that frame on the page.
//
// The y-down (DrawingML) → y-up (PDF) flip happens ONCE, upstream, in the
// geometry layer (preset-geometry / custGeom). Nothing here flips again, so a
// segment's coordinates map directly to `m`/`l`/`c` operands.

// ISO 32000-1 §8.5.2 — path construction operators.
export type PathSegment =
  | { readonly op: 'move'; readonly x: number; readonly y: number } // m
  | { readonly op: 'line'; readonly x: number; readonly y: number } // l
  | {
      readonly op: 'cubic'; // c
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: 'close' }; // h

export interface VectorPath {
  readonly segments: ReadonlyArray<PathSegment>;
  // Winding rule for filling (§8.5.3.3). Defaults to nonzero. evenodd selects
  // the `*` painting variants (f*/B*).
  readonly fillRule?: 'nonzero' | 'evenodd';
}

export interface StrokeStyle {
  readonly colorHex: string; // 6-hex, no leading '#'
  readonly widthPt: number;
  readonly cap?: 'butt' | 'round' | 'square'; // §8.4.3.3 → 0|1|2 J
  readonly join?: 'miter' | 'round' | 'bevel'; // §8.4.3.4 → 0|1|2 j
  // Dash pattern in points (§8.4.3.6). Empty/omitted = solid line.
  readonly dash?: ReadonlyArray<number>;
}

export interface VectorShape {
  readonly paths: ReadonlyArray<VectorPath>;
  // Non-stroking fill colour (6-hex). Omitted = no fill.
  readonly fillColorHex?: string;
  // Stroke description. Omitted = no stroke.
  readonly stroke?: StrokeStyle;
  // CTM applied via `cm` (§8.3.4): maps the local frame onto the page. The
  // 6-tuple is [a b c d e f] of the matrix [[a b 0][c d 0][e f 1]].
  readonly transform: readonly [number, number, number, number, number, number];
}

// Fluent builder so geometry modules read like the shape they describe.
export class PathBuilder {
  private readonly segments: Array<PathSegment> = [];

  moveTo(x: number, y: number): this {
    this.segments.push({ op: 'move', x, y });
    return this;
  }

  lineTo(x: number, y: number): this {
    this.segments.push({ op: 'line', x, y });
    return this;
  }

  cubicTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): this {
    this.segments.push({ op: 'cubic', x1, y1, x2, y2, x, y });
    return this;
  }

  // Append already-built segments (e.g. an arc decomposed into cubics).
  append(segments: ReadonlyArray<PathSegment>): this {
    for (const s of segments) this.segments.push(s);
    return this;
  }

  close(): this {
    this.segments.push({ op: 'close' });
    return this;
  }

  build(fillRule?: 'nonzero' | 'evenodd'): VectorPath {
    return { segments: this.segments.slice(), ...(fillRule ? { fillRule } : {}) };
  }
}

// Emit the content-stream operators for one shape as a list of lines. Fully
// self-contained inside its own q…Q so stroke/dash/colour state never leaks
// into a neighbouring pass.
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
