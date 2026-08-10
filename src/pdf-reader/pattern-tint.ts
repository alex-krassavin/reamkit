// §8.7.3 — what a tiling pattern comes to when it is used as a TINT.
//
// A run carries one colour, not a content stream, so type filled with a tiling
// pattern cannot be filled with the pattern here. Its COLOUR can be told (see
// `./text`), and its density can be measured: a pattern that covers a third of
// its cell reads as a third-strength tint, and painting it at full strength is
// as wrong in the other direction as painting it black was.
//
// The density cannot be had from the marks' own areas. A tile is clipped to its
// `/BBox` and repeats on `/XStep` × `/YStep`, and the two need not agree with
// either the marks or each other: ContentStreamCycleType3insideType3.pdf tiles
// on 55 × 32 and paints two 300 × 300 squares, one of them starting 400 units
// out. Summed, the ink is a hundred times the cell and every pattern would read
// as solid. So the cell is SAMPLED instead — a grid of points, each asked
// whether any mark covers it — which clips and overlaps for free.

import { interpretContent } from './content';
import { buildFonts } from './text';
import type { PathSeg, VectorPlacement } from './content';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { PdfFile } from './document';
import { PDF_NULL, PdfStream } from '@/pdf/objects';

/** How many points across the cell are asked. 48² is 2304 samples per mark. */
const GRID = 48;

/** How deep the walk follows a tile that draws its marks through Type 3 glyphs. */
const MAX_DEPTH = 4;

/** A tiling pattern read as a tint: the colour it paints, and how much of the cell. */
export interface PatternTint {
  readonly colorHex: string;
  /** 0–1 — the share of the cell the pattern's marks cover. */
  readonly coverage: number;
}

/**
 * Read a named tiling pattern as a tint.
 *
 * @param file      The owning file.
 * @param resources The resources the pattern is named in.
 * @param name      The pattern's resource name.
 * @returns Its colour and coverage, or `undefined` when neither can be told.
 */
export function patternTint(
  file: PdfFile,
  resources: PdfDict | undefined,
  name: string,
): PatternTint | undefined {
  if (!resources) return undefined;
  const patterns = file.get(resources, 'Pattern');
  if (!(patterns instanceof Map)) return undefined;
  const stream = file.resolve(patterns.get(name) ?? PDF_NULL);
  if (!(stream instanceof PdfStream)) return undefined;
  const cell = cellOf(file, stream.dict);
  if (!cell) return undefined;

  const marks: Array<VectorPlacement> = [];
  const colours: Array<string> = [];
  collect(file, stream, resources, marks, colours, 0, new Set());
  const colorHex = colours[0];
  if (colorHex === undefined) return undefined;
  return { colorHex, coverage: sample(marks, cell) };
}

/** The cell a tile repeats on: `/XStep` × `/YStep` from the `/BBox`'s corner. */
function cellOf(
  file: PdfFile,
  dict: PdfDict,
): { x: number; y: number; w: number; h: number } | undefined {
  const bbox = numbers(file, dict.get('BBox'));
  if (!bbox || bbox.length < 4) return undefined;
  const x = Math.min(bbox[0]!, bbox[2]!);
  const y = Math.min(bbox[1]!, bbox[3]!);
  const bw = Math.abs(bbox[2]! - bbox[0]!);
  const bh = Math.abs(bbox[3]! - bbox[1]!);
  // §8.7.3.1 — a step wider than the box leaves gaps, a narrower one overlaps;
  // either way the CELL is what repeats, and what is outside it is another
  // tile's business.
  const w = Math.abs(asNumber(file.resolve(dict.get('XStep') ?? PDF_NULL))) || bw;
  const h = Math.abs(asNumber(file.resolve(dict.get('YStep') ?? PDF_NULL))) || bh;
  return w > 0 && h > 0 ? { x, y, w, h } : undefined;
}

/** Every mark the tile paints, in pattern space, following its Type 3 glyphs. */
function collect(
  file: PdfFile,
  stream: PdfStream,
  fallback: PdfDict | undefined,
  marks: Array<VectorPlacement>,
  colours: Array<string>,
  depth: number,
  visiting: Set<PdfStream>,
): void {
  if (depth > MAX_DEPTH || visiting.has(stream)) return;
  visiting.add(stream);
  const own = file.get(stream.dict, 'Resources');
  const resources = own instanceof Map ? own : fallback;
  const result = interpretContent(file.streamData(stream), buildFonts(file, resources));
  for (const v of result.vectors) {
    marks.push(v);
    const hex = v.fillHex ?? v.strokeHex;
    if (hex !== undefined && hex !== 'FFFFFF') colours.push(hex);
  }
  // A tile may draw its marks as type, and that type may be a Type 3 face whose
  // glyphs are paths again — which is how this file paints its tile.
  for (const run of result.texts) if (run.colorHex !== 'FFFFFF') colours.push(run.colorHex);
  for (const glyph of result.glyphs) {
    const inner = interpretContent(file.streamData(glyph.stream), new Map(), glyph.ctm);
    for (const v of inner.vectors) {
      marks.push(v);
      const hex = v.fillHex ?? v.strokeHex;
      if (hex !== undefined && hex !== 'FFFFFF') colours.push(hex);
    }
  }
  visiting.delete(stream);
}

/** The share of the cell any mark covers, asked at a grid of points. */
function sample(
  marks: ReadonlyArray<VectorPlacement>,
  cell: { x: number; y: number; w: number; h: number },
): number {
  if (marks.length === 0) return 0;
  const flat = marks.map((m) => ({
    lines: flatten(m.segs),
    filled: m.fillHex !== undefined || m.gradient !== undefined,
    half: m.strokeHex !== undefined ? Math.max(m.lineWidth ?? 1, 0.1) / 2 : 0,
  }));
  let hit = 0;
  for (let iy = 0; iy < GRID; iy++) {
    const py = cell.y + ((iy + 0.5) / GRID) * cell.h;
    for (let ix = 0; ix < GRID; ix++) {
      const px = cell.x + ((ix + 0.5) / GRID) * cell.w;
      if (flat.some((m) => covers(m, px, py))) hit++;
    }
  }
  return hit / (GRID * GRID);
}

interface FlatMark {
  readonly lines: ReadonlyArray<readonly [number, number, number, number]>;
  readonly filled: boolean;
  readonly half: number;
}

/** Whether one mark paints the point — inside its outline, or under its pen. */
function covers(mark: FlatMark, px: number, py: number): boolean {
  if (mark.half > 0) {
    const limit = mark.half * mark.half;
    for (const [x0, y0, x1, y1] of mark.lines) {
      if (distanceSquared(px, py, x0, y0, x1, y1) <= limit) return true;
    }
  }
  if (!mark.filled) return false;
  // Crossing parity: the even-odd rule, which agrees with the non-zero one for
  // every outline that does not cross itself, and most do not.
  let inside = false;
  for (const [x0, y0, x1, y1] of mark.lines) {
    if (y0 > py !== y1 > py && px < ((x1 - x0) * (py - y0)) / (y1 - y0) + x0) inside = !inside;
  }
  return inside;
}

/** The squared distance from a point to a segment. */
function distanceSquared(
  px: number,
  py: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): number {
  const dx = x1 - x0;
  const dy = y1 - y0;
  const len = dx * dx + dy * dy;
  const t = len > 0 ? Math.max(0, Math.min(1, ((px - x0) * dx + (py - y0) * dy) / len)) : 0;
  const qx = x0 + t * dx - px;
  const qy = y0 + t * dy - py;
  return qx * qx + qy * qy;
}

/** A path as straight segments; a curve becomes the chords of its control net. */
function flatten(segs: ReadonlyArray<PathSeg>): Array<readonly [number, number, number, number]> {
  const out: Array<readonly [number, number, number, number]> = [];
  let sx = 0;
  let sy = 0;
  let cx = 0;
  let cy = 0;
  for (const seg of segs) {
    switch (seg.op) {
      case 'move':
        sx = cx = seg.x;
        sy = cy = seg.y;
        break;
      case 'line':
        out.push([cx, cy, seg.x, seg.y]);
        cx = seg.x;
        cy = seg.y;
        break;
      case 'cubic': {
        const steps = 8;
        let lx = cx;
        let ly = cy;
        for (let i = 1; i <= steps; i++) {
          const t = i / steps;
          const u = 1 - t;
          const x =
            u * u * u * cx + 3 * u * u * t * seg.x1 + 3 * u * t * t * seg.x2 + t * t * t * seg.x;
          const y =
            u * u * u * cy + 3 * u * u * t * seg.y1 + 3 * u * t * t * seg.y2 + t * t * t * seg.y;
          out.push([lx, ly, x, y]);
          lx = x;
          ly = y;
        }
        cx = seg.x;
        cy = seg.y;
        break;
      }
      case 'close':
        out.push([cx, cy, sx, sy]);
        cx = sx;
        cy = sy;
        break;
    }
  }
  return out;
}

function numbers(file: PdfFile, value: PdfValue | undefined): Array<number> | undefined {
  const v = file.resolve(value ?? PDF_NULL);
  if (!Array.isArray(v)) return undefined;
  const out = v.map((n) => asNumber(file.resolve(n)));
  return out.some((n) => !Number.isFinite(n)) ? undefined : out;
}

function asNumber(v: PdfValue): number {
  return typeof v === 'number' ? v : NaN;
}
