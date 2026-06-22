// E-PDF EP16c — read a PDF axial/radial shading back into a gradient fill. A
// shape filled with a shading pattern (`/Pattern cs /Pn scn` … fill) is the
// path-bounded case the interpreter captures; this resolves the page's
// /Pattern resources into ShapeGradients keyed by resource name, evaluating the
// shading's /Function (type 2 exponential, type 3 stitching, type 0 sampled) for
// the colour stops. The bare `sh` operator (clip-bounded) is not captured.

import type { GradientStop, ShapeGradient } from '@/core/vector';
import type { PdfDict, PdfValue } from '@/pdf/objects';

import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfStream } from '@/pdf/objects';

/**
 * Resolve a page's `/Pattern` resources into gradient fills (E-PDF EP16c, ISO
 * 32000-1 §8.7.4.5). Every `PatternType` 2 (shading) pattern is evaluated — its
 * `/Shading` type 2 (axial) or 3 (radial) plus the `/Function` colour stops —
 * and keyed by resource name; the interpreter looks the name up when a shape is
 * filled with `/Pattern cs /Pn scn`. The bare `sh` operator (clip-bounded) is
 * not captured.
 *
 * @returns A map from pattern resource name to its {@link ShapeGradient}.
 */
export function buildShadingMap(file: PdfFile, page: PdfPage): Map<string, ShapeGradient> {
  const out = new Map<string, ShapeGradient>();
  if (!page.resources) return out;
  const patterns = file.get(page.resources, 'Pattern');
  if (!(patterns instanceof Map)) return out;
  for (const [nm, value] of patterns) {
    const pat = file.resolve(value);
    if (!(pat instanceof Map)) continue;
    const shading = dictOf(file.resolve(pat.get('Shading') ?? PDF_NULL));
    if (!shading) continue;
    const gradient = parseShading(file, shading);
    if (gradient) out.set(nm, gradient);
  }
  return out;
}

function parseShading(file: PdfFile, sh: PdfDict): ShapeGradient | undefined {
  const type = numOf(file.get(sh, 'ShadingType'));
  if (type !== 2 && type !== 3) return undefined; // only axial (2) / radial (3)
  const stops = parseFunction(file, sh.get('Function'));
  if (!stops || stops.length === 0) return undefined;
  if (type === 3) return { kind: 'radial', stops };
  // Axial: the angle is the Coords direction, with y negated (PDF y-up → the
  // DrawingML y-down convention the model stores).
  const c = numArray(file, sh.get('Coords'));
  const angle =
    c && c.length >= 4
      ? ((((Math.atan2(-(c[3]! - c[1]!), c[2]! - c[0]!) * 180) / Math.PI) % 360) + 360) % 360
      : 0;
  return { kind: 'linear', angle, stops };
}

// A 1-in / n-out function → colour stops. Recurses for the type-3 stitching case.
function parseFunction(
  file: PdfFile,
  value: PdfValue | undefined,
): Array<GradientStop> | undefined {
  const resolved = value !== undefined ? file.resolve(value) : undefined;
  const dict = dictOf(resolved);
  if (!dict) return undefined;
  const type = numOf(file.get(dict, 'FunctionType'));

  if (type === 2) {
    const c0 = colorOf(numArray(file, dict.get('C0')) ?? [0]);
    const c1 = colorOf(numArray(file, dict.get('C1')) ?? [1]);
    return [
      { offset: 0, colorHex: c0 },
      { offset: 1, colorHex: c1 },
    ];
  }

  if (type === 3) {
    const fns = dict.get('Functions');
    const subs = Array.isArray(fns) ? fns : undefined;
    if (!subs) return undefined;
    const bounds = numArray(file, dict.get('Bounds')) ?? [];
    const domain = numArray(file, dict.get('Domain')) ?? [0, 1];
    const d0 = domain[0] ?? 0;
    const d1 = domain[domain.length - 1] ?? 1;
    const edges = [d0, ...bounds, d1];
    const stops: Array<GradientStop> = [];
    const span = d1 - d0 || 1;
    for (let i = 0; i < subs.length; i++) {
      const sub = parseFunction(file, subs[i]);
      if (!sub || sub.length < 2) continue;
      pushStop(stops, ((edges[i] ?? d0) - d0) / span, sub[0]!.colorHex);
      pushStop(stops, ((edges[i + 1] ?? d1) - d0) / span, sub[sub.length - 1]!.colorHex);
    }
    return stops.length > 0 ? stops : undefined;
  }

  if (type === 0 && resolved instanceof PdfStream) {
    return sampleFunction(file, resolved);
  }
  return undefined; // type 4 (PostScript calculator) — not evaluated
}

// §7.10.2 sampled function — read the table and sample it at a few offsets.
function sampleFunction(file: PdfFile, stream: PdfStream): Array<GradientStop> | undefined {
  const d = stream.dict;
  const size = numArray(file, d.get('Size'));
  const range = numArray(file, d.get('Range'));
  const bps = numOf(file.get(d, 'BitsPerSample'));
  if (!size || !range || size.length < 1 || range.length < 2 || bps < 1) return undefined;
  const n = size[0]!;
  const comps = range.length / 2;
  const data = file.streamData(stream);
  const maxv = 2 ** bps - 1;
  const bitAt = (sampleIdx: number, comp: number): number => {
    let bit = (sampleIdx * comps + comp) * bps;
    let v = 0;
    for (let k = 0; k < bps; k++) {
      const byte = data[bit >> 3] ?? 0;
      v = (v << 1) | ((byte >> (7 - (bit & 7))) & 1);
      bit++;
    }
    return v;
  };
  const count = Math.min(Math.max(n, 2), 16);
  const stops: Array<GradientStop> = [];
  for (let s = 0; s < count; s++) {
    const off = s / (count - 1);
    const j = Math.round(off * (n - 1));
    const c: Array<number> = [];
    for (let comp = 0; comp < comps; comp++) {
      const lo = range[comp * 2]!;
      const hi = range[comp * 2 + 1]!;
      c.push(lo + (bitAt(j, comp) / maxv) * (hi - lo));
    }
    pushStop(stops, off, colorOf(c));
  }
  return stops.length > 0 ? stops : undefined;
}

// --- helpers ----------------------------------------------------------------

function dictOf(v: PdfValue | undefined): PdfDict | undefined {
  if (v instanceof PdfStream) return v.dict;
  if (v instanceof Map) return v;
  return undefined;
}

function numOf(v: PdfValue | undefined): number {
  return typeof v === 'number' ? v : 0;
}

function numArray(file: PdfFile, v: PdfValue | undefined): Array<number> | undefined {
  const r = v !== undefined ? file.resolve(v) : undefined;
  if (!Array.isArray(r)) return undefined;
  return r.map((x) => (typeof x === 'number' ? x : 0));
}

// Colour components (0..1, in the shading colour space) → a 6-hex sRGB string.
function colorOf(c: ReadonlyArray<number>): string {
  if (c.length >= 4) {
    const k = c[3]!;
    return hex255(
      255 * (1 - c[0]!) * (1 - k),
      255 * (1 - c[1]!) * (1 - k),
      255 * (1 - c[2]!) * (1 - k),
    );
  }
  if (c.length === 3) return hex255(c[0]! * 255, c[1]! * 255, c[2]! * 255);
  const g = (c[0] ?? 0) * 255;
  return hex255(g, g, g);
}

function hex255(r: number, g: number, b: number): string {
  const h = (x: number): string =>
    Math.max(0, Math.min(255, Math.round(x)))
      .toString(16)
      .padStart(2, '0');
  return (h(r) + h(g) + h(b)).toUpperCase();
}

// Append a stop, coalescing one at the same offset (keeps the gradient monotone).
function pushStop(stops: Array<GradientStop>, offset: number, colorHex: string): void {
  const o = Math.max(0, Math.min(1, offset));
  const last = stops[stops.length - 1];
  if (last && Math.abs(last.offset - o) < 1e-6) {
    if (last.colorHex === colorHex) return;
  }
  stops.push({ offset: o, colorHex });
}
