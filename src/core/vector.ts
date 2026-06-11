// Format-agnostic vector model (paths, strokes, shapes) — the shared
// vocabulary of every writer, plus the path-data/flip helpers every emitter
// of that vocabulary needs. The PDF operator emission lives in
// pdf/vector-graphics (oop-design §8, C9 — this move cut the svg→pdf model
// dependency).

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

// Compose the page flip with a local→page CTM built in the y-up frame, so the
// stored transform targets the top-left frame the PageDoc schema froze on.
// The flip is an involution: the PDF emitter applies the same operation to
// recover the y-up matrix. Negating the linear part only flips sign bits
// (exact in IEEE 754); the y-translation is the one component that re-rounds.
export function flipTransform(
  m: readonly [number, number, number, number, number, number],
  pageHeight: number,
): [number, number, number, number, number, number] {
  return [m[0], -m[1], m[2], -m[3], m[4], pageHeight - m[5]];
}

// SVG path data from segments, raw coordinates (the caller's transform maps
// the local frame into the target one).
export function svgPathData(
  segments: ReadonlyArray<PathSegment>,
  fmt: (n: number) => string,
): string {
  const parts: Array<string> = [];
  for (const s of segments) {
    if (s.op === 'move') parts.push(`M ${fmt(s.x)} ${fmt(s.y)}`);
    else if (s.op === 'line') parts.push(`L ${fmt(s.x)} ${fmt(s.y)}`);
    else if (s.op === 'cubic')
      parts.push(`C ${fmt(s.x1)} ${fmt(s.y1)} ${fmt(s.x2)} ${fmt(s.y2)} ${fmt(s.x)} ${fmt(s.y)}`);
    else parts.push('Z');
  }
  return parts.join(' ');
}
