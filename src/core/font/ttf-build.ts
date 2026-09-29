// ISO/IEC 14496-22 — a TrueType font BUILT from outlines, for a document that
// embeds the faces it was drawn in.
//
// A PDF carries its faces as programs in three formats (ISO 32000 §9.9):
// TrueType, CFF and Type 1, the last two drawn in cubic curves. A word
// processor embeds TrueType (ECMA-376 §17.15.1.42 `w:embedTrueTypeFonts`), and
// a PDF's program is a SUBSET whose codes are the PDF's own in any case — no
// `cmap` a word processor could reach it through, glyphs named `g24`. So the
// face is rebuilt: each character the document shows, mapped to the outline the
// page drew for it and set with the advance the page set it with.
//
// The em is a thousand units, which is the unit a PDF states its widths in, so
// an advance goes through unrounded wherever the file wrote a whole number.

import { directoryGeometry, paddedChecksum } from '@/core/font/ttf-subset';

/** One piece of a glyph's outline, in a one-unit em with y up. */
export type GlyphSeg =
  | { readonly op: 'move'; readonly x: number; readonly y: number }
  | { readonly op: 'line'; readonly x: number; readonly y: number }
  | {
      readonly op: 'cubic';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: 'close' };

/** One glyph of a {@link BuiltFace}. */
export interface BuiltGlyph {
  /** The characters that select it, as code points. */
  readonly codePoints: ReadonlyArray<number>;
  /** Its contours, filled by the nonzero rule; empty for a blank glyph. */
  readonly outline: ReadonlyArray<GlyphSeg>;
  /** How far the pen moves after it, in thousandths of an em. */
  readonly advance: number;
}

/** Everything {@link buildTrueType} needs to know about the face it builds. */
export interface BuiltFace {
  /** The family the font is installed under (`name` ID 1). */
  readonly family: string;
  /** The style within the family: its bold and italic slot (`name` ID 2). */
  readonly bold: boolean;
  readonly italic: boolean;
  /** The PostScript name (`name` ID 6); only its printable ASCII is kept. */
  readonly postScriptName: string;
  readonly glyphs: ReadonlyArray<BuiltGlyph>;
  /** Above and below the baseline, in thousandths of an em — `descent` negative. */
  readonly ascent: number;
  readonly descent: number;
  readonly capHeight?: number;
  readonly xHeight?: number;
  /** Degrees counterclockwise from the vertical; a face slanted right is negative. */
  readonly italicAngle?: number;
  readonly fixedPitch?: boolean;
  /**
   * OS/2 `fsType` — the embedding the face's licence allows, carried over from
   * the program the outlines came from.
   */
  readonly fsType: number;
}

/** OS/2 `fsType` bits (ISO/IEC 14496-22, OS/2 table). */
const USAGE_BITS = 0x000e;
const EDITABLE_EMBEDDING = 0x0008;
const NO_SUBSETTING = 0x0100;
const BITMAP_ONLY = 0x0200;

/**
 * Whether a face's licence lets it be embedded, as outlines and subset, in a
 * document that stays EDITABLE: OS/2 `fsType` states installable embedding (no
 * usage bit) or grants editable embedding, and neither forbids subsetting nor
 * allows bitmaps only. Restricted-licence and preview-and-print faces are
 * refused — the second may travel only in a document opened read-only.
 *
 * Several usage bits at once is an old font's way of stating permissions, and
 * the least restrictive of them applies (OS/2 versions 0–2).
 *
 * @param fsType The field, or `undefined` where the program states none —
 *               which restricts nothing.
 * @returns Whether an editable document may carry the face.
 */
export function editableEmbedding(fsType: number | undefined): boolean {
  if (fsType === undefined) return true;
  if ((fsType & (NO_SUBSETTING | BITMAP_ONLY)) !== 0) return false;
  const usage = fsType & USAGE_BITS;
  return usage === 0 || (usage & EDITABLE_EMBEDDING) !== 0;
}

/** Units per em. */
const UPEM = 1000;

/**
 * How far, in font units, a quadratic may stray from the cubic it replaces.
 * Half a unit is the rounding every coordinate takes anyway.
 */
const CURVE_TOLERANCE = 0.5;

/** No cubic is cut into more quadratics than this. */
const MAX_CURVE_PIECES = 16;

/**
 * The most glyphs a built face holds. A `cmap` format 4 subtable is limited to
 * 64 KB, and past this many characters it may not fit; a face that large is a
 * whole CJK font, which is no subset.
 */
export const MAX_BUILT_GLYPHS = 8000;

/**
 * Build a TrueType font from outlines (ISO/IEC 14496-22): `.notdef` first, then
 * one glyph per {@link BuiltGlyph} in the order given, reachable through a
 * Unicode `cmap`.
 *
 * Cubic curves become quadratic ones. A cubic that is a raised quadratic — what
 * a TrueType source's outline arrives as — comes back as exactly that
 * quadratic; any other is cut into as many quadratics as keep within half a
 * unit of it. Each glyph's contours are turned to run clockwise round their
 * ink, as TrueType's are.
 *
 * @param face The face: its names, metrics, licence and glyphs.
 * @returns The font program's bytes.
 * @throws When the face holds more than {@link MAX_BUILT_GLYPHS} glyphs.
 */
export function buildTrueType(face: BuiltFace): Uint8Array {
  if (face.glyphs.length > MAX_BUILT_GLYPHS) {
    throw new Error(`A built face holds at most ${MAX_BUILT_GLYPHS} glyphs`);
  }
  const encoded = [emptyGlyph(), ...face.glyphs.map((g) => encodeGlyph(contoursOf(g.outline)))];
  const advances = [0, ...face.glyphs.map((g) => clamp(Math.round(g.advance), 0, 0xffff))];
  const inked = encoded.filter((g) => g.contours > 0);
  const bounds = {
    xMin: Math.min(0, ...inked.map((g) => g.xMin)),
    yMin: Math.min(0, ...inked.map((g) => g.yMin)),
    xMax: Math.max(0, ...inked.map((g) => g.xMax)),
    yMax: Math.max(0, ...inked.map((g) => g.yMax)),
  };
  const ascent = clamp(Math.round(face.ascent), 1, 0x7fff);
  const descent = clamp(Math.round(face.descent), -0x7fff, 0);
  const metrics: FaceMetrics = {
    ascent,
    descent,
    // Windows clips what stands outside these, so they reach the ink.
    winAscent: Math.max(ascent, bounds.yMax),
    winDescent: Math.max(-descent, -bounds.yMin),
    bounds,
  };
  const cmap = cmapTable(face.glyphs);
  const tables: Array<[string, Uint8Array]> = [
    ['OS/2', os2Table(face, encoded, advances, metrics)],
    ['cmap', cmap],
    ['glyf', concat(encoded.map((g) => g.bytes))],
    ['head', headTable(face, bounds)],
    ['hhea', hheaTable(face, encoded, advances, metrics)],
    ['hmtx', hmtxTable(encoded, advances)],
    ['loca', locaTable(encoded)],
    ['maxp', maxpTable(encoded)],
    ['name', nameTable(face)],
    ['post', postTable(face)],
  ];
  return assembleSfnt(tables);
}

/** The vertical metrics every table that states them agrees on. */
interface FaceMetrics {
  readonly ascent: number;
  readonly descent: number;
  readonly winAscent: number;
  readonly winDescent: number;
  readonly bounds: { xMin: number; yMin: number; xMax: number; yMax: number };
}

/** A point of a TrueType contour: on the curve, or its quadratic control. */
interface Point {
  x: number;
  y: number;
  on: boolean;
}

/**
 * The outline as TrueType contours, in font units and rounded: each closed,
 * quadratic, and running clockwise round its ink.
 */
function contoursOf(outline: ReadonlyArray<GlyphSeg>): Array<Array<Point>> {
  const out: Array<Array<Point>> = [];
  let current: Array<Point> = [];
  let x = 0;
  let y = 0;
  // Where the subpath began, which is where `close` leaves the pen.
  let startX = 0;
  let startY = 0;
  const finish = (): void => {
    const contour = tidy(current);
    if (contour.length >= 3) out.push(contour);
    current = [];
  };
  for (const seg of outline) {
    if (seg.op === 'move') {
      finish();
      x = seg.x * UPEM;
      y = seg.y * UPEM;
      startX = x;
      startY = y;
      current.push({ x, y, on: true });
    } else if (seg.op === 'line') {
      if (current.length === 0) current.push({ x, y, on: true });
      x = seg.x * UPEM;
      y = seg.y * UPEM;
      current.push({ x, y, on: true });
    } else if (seg.op === 'cubic') {
      if (current.length === 0) current.push({ x, y, on: true });
      const to = { x: seg.x * UPEM, y: seg.y * UPEM };
      quadraticsOf(
        { x, y },
        { x: seg.x1 * UPEM, y: seg.y1 * UPEM },
        { x: seg.x2 * UPEM, y: seg.y2 * UPEM },
        to,
        current,
      );
      x = to.x;
      y = to.y;
    } else {
      finish();
      x = startX;
      y = startY;
    }
  }
  finish();
  // TrueType runs its outer contours clockwise, PostScript counter-clockwise;
  // the fill is nonzero either way, but a rasteriser's dropout control reads
  // the direction. The biggest contour is an outer one, and says which way
  // this glyph runs.
  let biggest = 0;
  for (const contour of out) {
    const a = area(contour);
    if (Math.abs(a) > Math.abs(biggest)) biggest = a;
  }
  if (biggest > 0) for (const contour of out) contour.reverse();
  return out;
}

/**
 * One cubic as quadratics, appended to `out` (its start is already there).
 *
 * A cubic raised from a quadratic has both its controls two thirds of the way
 * to the same point, and that point is the quadratic's control. Any other
 * cubic is cut evenly into pieces, each replaced by the quadratic through its
 * ends whose control is the average of the two its controls imply — which
 * strays from the piece by at most √3/36 of |P3 − 3C2 + 3C1 − P0|, a length
 * that shrinks with the cube of the piece.
 */
function quadraticsOf(
  p0: { x: number; y: number },
  c1: { x: number; y: number },
  c2: { x: number; y: number },
  p3: { x: number; y: number },
  out: Array<Point>,
): void {
  const q1 = { x: (3 * c1.x - p0.x) / 2, y: (3 * c1.y - p0.y) / 2 };
  const q2 = { x: (3 * c2.x - p3.x) / 2, y: (3 * c2.y - p3.y) / 2 };
  if (Math.hypot(q1.x - q2.x, q1.y - q2.y) <= RAISED_QUADRATIC) {
    out.push({ x: (q1.x + q2.x) / 2, y: (q1.y + q2.y) / 2, on: false });
    out.push({ x: p3.x, y: p3.y, on: true });
    return;
  }
  const dx = p3.x - 3 * c2.x + 3 * c1.x - p0.x;
  const dy = p3.y - 3 * c2.y + 3 * c1.y - p0.y;
  const stray = (Math.hypot(dx, dy) * Math.sqrt(3)) / 36;
  const pieces = clamp(Math.ceil(Math.cbrt(stray / CURVE_TOLERANCE)), 1, MAX_CURVE_PIECES);
  const at = (t: number): { x: number; y: number } => {
    const u = 1 - t;
    return {
      x: u * u * u * p0.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t * t * t * p3.x,
      y: u * u * u * p0.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t * t * t * p3.y,
    };
  };
  // The derivative at t, scaled by a third of the piece: the step from a
  // piece's end to the control beside it.
  const lean = (t: number): { x: number; y: number } => {
    const u = 1 - t;
    const k = 1 / pieces;
    return {
      x: k * (u * u * (c1.x - p0.x) + 2 * u * t * (c2.x - c1.x) + t * t * (p3.x - c2.x)),
      y: k * (u * u * (c1.y - p0.y) + 2 * u * t * (c2.y - c1.y) + t * t * (p3.y - c2.y)),
    };
  };
  for (let i = 0; i < pieces; i++) {
    const t0 = i / pieces;
    const t1 = (i + 1) / pieces;
    const a = at(t0);
    const b = i + 1 === pieces ? p3 : at(t1);
    const la = lean(t0);
    const lb = lean(t1);
    const s1 = { x: a.x + la.x, y: a.y + la.y };
    const s2 = { x: b.x - lb.x, y: b.y - lb.y };
    out.push({
      x: (3 * (s1.x + s2.x) - (a.x + b.x)) / 4,
      y: (3 * (s1.y + s2.y) - (a.y + b.y)) / 4,
      on: false,
    });
    out.push({ x: b.x, y: b.y, on: true });
  }
}

/** How near, in font units, a cubic's two implied controls must be to be one. */
const RAISED_QUADRATIC = 0.01;

/**
 * A contour rounded to whole units, with the points rounding merges into one
 * and the closing point that repeats the first taken out.
 */
function tidy(points: ReadonlyArray<Point>): Array<Point> {
  const out: Array<Point> = [];
  for (const p of points) {
    const q = {
      x: clamp(Math.round(p.x), -0x8000, 0x7fff),
      y: clamp(Math.round(p.y), -0x8000, 0x7fff),
      on: p.on,
    };
    const last = out[out.length - 1];
    // A control that sits on its end point bends nothing: the two are a
    // corner, which is a point on the curve.
    if (last && last.x === q.x && last.y === q.y) last.on = last.on || q.on;
    else out.push(q);
  }
  while (out.length > 1) {
    const first = out[0]!;
    const last = out[out.length - 1]!;
    if (first.x !== last.x || first.y !== last.y) break;
    first.on = first.on || last.on;
    out.pop();
  }
  return out;
}

/** Twice the signed area a contour's points enclose — positive counter-clockwise. */
function area(points: ReadonlyArray<Point>): number {
  let sum = 0;
  for (let i = 0; i < points.length; i++) {
    const a = points[i]!;
    const b = points[(i + 1) % points.length]!;
    sum += a.x * b.y - b.x * a.y;
  }
  return sum;
}

/** One glyph's `glyf` entry and what the other tables need to know of it. */
interface EncodedGlyph {
  readonly bytes: Uint8Array;
  readonly contours: number;
  readonly points: number;
  readonly xMin: number;
  readonly yMin: number;
  readonly xMax: number;
  readonly yMax: number;
}

function emptyGlyph(): EncodedGlyph {
  return { bytes: new Uint8Array(0), contours: 0, points: 0, xMin: 0, yMin: 0, xMax: 0, yMax: 0 };
}

// `glyf` flag bits.
const ON_CURVE = 0x01;
const X_SHORT = 0x02;
const Y_SHORT = 0x04;
const REPEAT = 0x08;
const X_SAME_OR_POSITIVE = 0x10;
const Y_SAME_OR_POSITIVE = 0x20;

/**
 * A simple glyph as `glyf` stores it: the contours' last points, no
 * instructions, then flags and coordinate deltas packed as small as they go.
 */
function encodeGlyph(contours: ReadonlyArray<ReadonlyArray<Point>>): EncodedGlyph {
  const points = contours.flat();
  if (points.length === 0) return emptyGlyph();
  const xs = points.map((p) => p.x);
  const ys = points.map((p) => p.y);
  const box = {
    xMin: Math.min(...xs),
    yMin: Math.min(...ys),
    xMax: Math.max(...xs),
    yMax: Math.max(...ys),
  };
  const flags: Array<number> = [];
  const xBytes: Array<number> = [];
  const yBytes: Array<number> = [];
  let px = 0;
  let py = 0;
  const delta = (d: number, short: number, same: number, bytes: Array<number>): number => {
    if (d === 0) return same;
    if (Math.abs(d) <= 0xff) {
      bytes.push(Math.abs(d));
      return short | (d > 0 ? same : 0);
    }
    bytes.push((d >> 8) & 0xff, d & 0xff);
    return 0;
  };
  for (const p of points) {
    let flag = p.on ? ON_CURVE : 0;
    flag |= delta(p.x - px, X_SHORT, X_SAME_OR_POSITIVE, xBytes);
    flag |= delta(p.y - py, Y_SHORT, Y_SAME_OR_POSITIVE, yBytes);
    flags.push(flag);
    px = p.x;
    py = p.y;
  }
  const packed: Array<number> = [];
  for (let i = 0; i < flags.length; ) {
    let run = 1;
    while (i + run < flags.length && flags[i + run] === flags[i] && run < 256) run++;
    if (run > 1) {
      packed.push(flags[i]! | REPEAT, run - 1);
    } else {
      packed.push(flags[i]!);
    }
    i += run;
  }
  const size = 10 + contours.length * 2 + 2 + packed.length + xBytes.length + yBytes.length;
  const out = new Writer(size + ((4 - (size % 4)) % 4));
  out.i16(contours.length);
  out.i16(box.xMin);
  out.i16(box.yMin);
  out.i16(box.xMax);
  out.i16(box.yMax);
  let end = -1;
  for (const contour of contours) {
    end += contour.length;
    out.u16(end);
  }
  out.u16(0); // no instructions
  out.bytes(packed);
  out.bytes(xBytes);
  out.bytes(yBytes);
  return {
    bytes: out.data,
    contours: contours.length,
    points: points.length,
    ...box,
  };
}

/**
 * `cmap`: the BMP through format 4, under both Unicode encodings a reader looks
 * for, and — where the face reaches past it — every character through format 12.
 */
function cmapTable(glyphs: ReadonlyArray<BuiltGlyph>): Uint8Array {
  const pairs: Array<[number, number]> = [];
  glyphs.forEach((g, i) => {
    for (const cp of g.codePoints) pairs.push([cp, i + 1]);
  });
  pairs.sort((a, b) => a[0] - b[0]);
  const unique = pairs.filter((p, i) => i === 0 || p[0] !== pairs[i - 1]![0]);
  const bmp = unique.filter(([cp]) => cp <= 0xffff && (cp < 0xd800 || cp > 0xdfff));
  const format4 = cmapFormat4(bmp);
  const wide = unique.some(([cp]) => cp > 0xffff) ? cmapFormat12(unique) : undefined;
  // Encoding records, sorted by platform and then encoding.
  type EncodingRecord = readonly [number, number, 'bmp' | 'wide'];
  const records: ReadonlyArray<EncodingRecord> = [
    [0, 3, 'bmp'],
    ...(wide ? [[0, 4, 'wide'] as const] : []),
    [3, 1, 'bmp'],
    ...(wide ? [[3, 10, 'wide'] as const] : []),
  ];
  const headerSize = 4 + records.length * 8;
  const out = new Writer(headerSize + format4.length + (wide?.length ?? 0));
  out.u16(0);
  out.u16(records.length);
  for (const [platform, encoding, which] of records) {
    out.u16(platform);
    out.u16(encoding);
    out.u32(which === 'bmp' ? headerSize : headerSize + format4.length);
  }
  out.bytes(format4);
  if (wide) out.bytes(wide);
  return out.data;
}

/** Format 4: segments of consecutive characters on consecutive glyphs. */
function cmapFormat4(pairs: ReadonlyArray<[number, number]>): Uint8Array {
  const segments: Array<{ start: number; end: number; gid: number }> = [];
  for (const [cp, gid] of pairs) {
    const last = segments[segments.length - 1];
    if (last && cp === last.end + 1 && gid === last.gid + (cp - last.start)) last.end = cp;
    else segments.push({ start: cp, end: cp, gid });
  }
  // The table ends on a segment for U+FFFF that maps it to nothing.
  if (segments[segments.length - 1]?.end !== 0xffff) {
    segments.push({ start: 0xffff, end: 0xffff, gid: 0 });
  }
  const count = segments.length;
  const entrySelector = Math.floor(Math.log2(count));
  const searchRange = 2 * 2 ** entrySelector;
  const length = 16 + count * 8;
  const out = new Writer(length);
  out.u16(4);
  out.u16(length);
  out.u16(0); // language
  out.u16(count * 2);
  out.u16(searchRange);
  out.u16(entrySelector);
  out.u16(count * 2 - searchRange);
  for (const s of segments) out.u16(s.end);
  out.u16(0); // reservedPad
  for (const s of segments) out.u16(s.start);
  for (const s of segments) {
    // U+FFFF's own segment maps to glyph 0: 0xFFFF + 1 wraps to it.
    out.u16(s.gid === 0 ? 1 : (s.gid - s.start) & 0xffff);
  }
  for (let i = 0; i < count; i++) out.u16(0); // idRangeOffset
  return out.data;
}

/** Format 12: groups of consecutive characters on consecutive glyphs. */
function cmapFormat12(pairs: ReadonlyArray<[number, number]>): Uint8Array {
  const groups: Array<{ start: number; end: number; gid: number }> = [];
  for (const [cp, gid] of pairs) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const last = groups[groups.length - 1];
    if (last && cp === last.end + 1 && gid === last.gid + (cp - last.start)) last.end = cp;
    else groups.push({ start: cp, end: cp, gid });
  }
  const length = 16 + groups.length * 12;
  const out = new Writer(length);
  out.u16(12);
  out.u16(0);
  out.u32(length);
  out.u32(0); // language
  out.u32(groups.length);
  for (const g of groups) {
    out.u32(g.start);
    out.u32(g.end);
    out.u32(g.gid);
  }
  return out.data;
}

/** `head` — version 1.0, a thousand units to the em, long `loca` offsets. */
function headTable(face: BuiltFace, bounds: FaceMetrics['bounds']): Uint8Array {
  const out = new Writer(54);
  out.u32(0x00010000); // version
  out.u32(0x00010000); // fontRevision
  out.u32(0); // checkSumAdjustment, set once the whole font is assembled
  out.u32(0x5f0f3cf5); // magicNumber
  // Baseline at y = 0, left side bearing point at x = 0, integer scaling.
  out.u16(0x000b);
  out.u16(UPEM);
  // created and modified: none stated, so the same face builds to the same bytes.
  out.u32(0);
  out.u32(0);
  out.u32(0);
  out.u32(0);
  out.i16(bounds.xMin);
  out.i16(bounds.yMin);
  out.i16(bounds.xMax);
  out.i16(bounds.yMax);
  out.u16((face.bold ? 1 : 0) | (face.italic ? 2 : 0)); // macStyle
  out.u16(8); // lowestRecPPEM
  out.i16(2); // fontDirectionHint
  out.i16(1); // indexToLocFormat: long
  out.i16(0); // glyphDataFormat
  return out.data;
}

/** `hhea` — the horizontal header, one metric per glyph. */
function hheaTable(
  face: BuiltFace,
  glyphs: ReadonlyArray<EncodedGlyph>,
  advances: ReadonlyArray<number>,
  metrics: FaceMetrics,
): Uint8Array {
  const inked = glyphs
    .map((g, i) => ({ g, advance: advances[i]! }))
    .filter((e) => e.g.contours > 0);
  const slope = slopeOf(face.italicAngle ?? 0);
  const out = new Writer(36);
  out.u32(0x00010000);
  out.i16(metrics.ascent);
  out.i16(metrics.descent);
  out.i16(0); // lineGap
  out.u16(Math.max(0, ...advances));
  out.i16(Math.min(0, ...inked.map((e) => e.g.xMin)));
  out.i16(Math.min(0, ...inked.map((e) => e.advance - e.g.xMax)));
  out.i16(Math.max(0, ...inked.map((e) => e.g.xMax)));
  out.i16(slope.rise);
  out.i16(slope.run);
  out.i16(0); // caretOffset
  for (let i = 0; i < 4; i++) out.i16(0);
  out.i16(0); // metricDataFormat
  out.u16(glyphs.length);
  return out.data;
}

/** The caret's slope for a face slanted by `angle` degrees. */
function slopeOf(angle: number): { rise: number; run: number } {
  if (!Number.isFinite(angle) || angle === 0) return { rise: 1, run: 0 };
  return { rise: UPEM, run: Math.round(-Math.tan((angle * Math.PI) / 180) * UPEM) };
}

function hmtxTable(
  glyphs: ReadonlyArray<EncodedGlyph>,
  advances: ReadonlyArray<number>,
): Uint8Array {
  const out = new Writer(glyphs.length * 4);
  glyphs.forEach((g, i) => {
    out.u16(advances[i]!);
    out.i16(g.contours > 0 ? g.xMin : 0);
  });
  return out.data;
}

function locaTable(glyphs: ReadonlyArray<EncodedGlyph>): Uint8Array {
  const out = new Writer((glyphs.length + 1) * 4);
  let at = 0;
  for (const g of glyphs) {
    out.u32(at);
    at += g.bytes.length;
  }
  out.u32(at);
  return out.data;
}

/** `maxp` version 1.0: the most points and contours any glyph holds, and no hinting. */
function maxpTable(glyphs: ReadonlyArray<EncodedGlyph>): Uint8Array {
  const out = new Writer(32);
  out.u32(0x00010000);
  out.u16(glyphs.length);
  out.u16(Math.max(0, ...glyphs.map((g) => g.points)));
  out.u16(Math.max(0, ...glyphs.map((g) => g.contours)));
  out.u16(0); // maxCompositePoints
  out.u16(0); // maxCompositeContours
  out.u16(2); // maxZones
  for (let i = 0; i < 8; i++) out.u16(0);
  return out.data;
}

/** OS/2 version 4 — the metrics and the licence Windows reads. */
function os2Table(
  face: BuiltFace,
  glyphs: ReadonlyArray<EncodedGlyph>,
  advances: ReadonlyArray<number>,
  metrics: FaceMetrics,
): Uint8Array {
  const codePoints = face.glyphs.flatMap((g) => g.codePoints);
  const inked = advances.filter((a, i) => a > 0 && (glyphs[i]?.contours ?? 0) > 0);
  const average = inked.length > 0 ? inked.reduce((s, a) => s + a, 0) / inked.length : UPEM / 2;
  const bmp = codePoints.filter((cp) => cp <= 0xffff);
  const out = new Writer(96);
  out.u16(4); // version
  out.i16(Math.round(average));
  out.u16(face.bold ? 700 : 400);
  out.u16(5); // usWidthClass: medium
  out.u16(face.fsType & 0xffff);
  // Subscript and superscript sizes and offsets.
  for (const v of [650, 600, 0, 75, 650, 600, 0, 350]) out.i16(v);
  out.i16(50); // yStrikeoutSize
  out.i16(Math.round((face.xHeight ?? 500) / 2)); // yStrikeoutPosition
  out.i16(0); // sFamilyClass
  out.bytes([0, 0, 0, 0, 0, 0, 0, 0, 0, 0]); // panose: any
  const ranges = unicodeRanges(codePoints);
  for (const r of ranges) out.u32(r);
  out.bytes([0x20, 0x20, 0x20, 0x20]); // achVendID
  out.u16(
    (face.italic ? 0x01 : 0) | (face.bold ? 0x20 : 0) | (!face.bold && !face.italic ? 0x40 : 0),
  );
  out.u16(bmp.length > 0 ? Math.min(...bmp) : 0x20);
  out.u16(bmp.length > 0 ? Math.max(...bmp) : 0x20);
  out.i16(metrics.ascent);
  out.i16(metrics.descent);
  out.i16(0); // sTypoLineGap
  out.u16(clamp(metrics.winAscent, 0, 0xffff));
  out.u16(clamp(metrics.winDescent, 0, 0xffff));
  const pages = codePages(codePoints);
  out.u32(pages[0]);
  out.u32(pages[1]);
  out.i16(Math.round(face.xHeight ?? 500));
  out.i16(Math.round(face.capHeight ?? 700));
  out.u16(0); // usDefaultChar
  out.u16(0x20); // usBreakChar
  out.u16(0); // usMaxContext
  return out.data;
}

/** OS/2 `ulUnicodeRange` bits for the blocks a face's characters fall in. */
const UNICODE_RANGES: ReadonlyArray<readonly [number, number, number]> = [
  [0, 0x0000, 0x007f],
  [1, 0x0080, 0x00ff],
  [2, 0x0100, 0x017f],
  [3, 0x0180, 0x024f],
  [4, 0x0250, 0x02af],
  [5, 0x02b0, 0x02ff],
  [6, 0x0300, 0x036f],
  [7, 0x0370, 0x03ff],
  [9, 0x0400, 0x052f],
  [10, 0x0530, 0x058f],
  [11, 0x0590, 0x05ff],
  [13, 0x0600, 0x06ff],
  [24, 0x0e00, 0x0e7f],
  [29, 0x1e00, 0x1eff],
  [30, 0x1f00, 0x1fff],
  [31, 0x2000, 0x206f],
  [32, 0x2070, 0x209f],
  [33, 0x20a0, 0x20cf],
  [35, 0x2100, 0x214f],
  [36, 0x2150, 0x218f],
  [37, 0x2190, 0x21ff],
  [38, 0x2200, 0x22ff],
  [39, 0x2300, 0x23ff],
  [43, 0x2500, 0x257f],
  [44, 0x2580, 0x259f],
  [45, 0x25a0, 0x25ff],
  [46, 0x2600, 0x26ff],
  [47, 0x2700, 0x27bf],
  [48, 0x3000, 0x303f],
  [49, 0x3040, 0x309f],
  [50, 0x30a0, 0x30ff],
  [56, 0xac00, 0xd7af],
  [59, 0x4e00, 0x9fff],
  [60, 0xe000, 0xf8ff],
  [62, 0xfb00, 0xfb4f],
  [68, 0xff00, 0xffef],
];

function unicodeRanges(codePoints: ReadonlyArray<number>): [number, number, number, number] {
  const words: [number, number, number, number] = [0, 0, 0, 0];
  const set = (bit: number): void => {
    words[bit >> 5] = (words[bit >> 5]! | (1 << (bit & 31))) >>> 0;
  };
  for (const cp of codePoints) {
    if (cp > 0xffff) set(57);
    for (const [bit, from, to] of UNICODE_RANGES) if (cp >= from && cp <= to) set(bit);
  }
  return words;
}

/**
 * OS/2 `ulCodePageRange` — the Windows code pages a face's characters serve.
 * Latin 1 stands where nothing else does: a face that claims no code page at
 * all is one Windows may decline to use.
 */
function codePages(codePoints: ReadonlyArray<number>): [number, number] {
  let low = 0;
  const has = (from: number, to: number): boolean =>
    codePoints.some((cp) => cp >= from && cp <= to);
  if (has(0x20, 0xff)) low |= 1 << 0; // 1252 Latin 1
  if (has(0x100, 0x17f)) low |= (1 << 1) | (1 << 4) | (1 << 7); // 1250, 1254, 1257
  if (has(0x400, 0x4ff)) low |= 1 << 2; // 1251 Cyrillic
  if (has(0x370, 0x3ff)) low |= 1 << 3; // 1253 Greek
  if (has(0x590, 0x5ff)) low |= 1 << 5; // 1255 Hebrew
  if (has(0x600, 0x6ff)) low |= 1 << 6; // 1256 Arabic
  if (has(0x1ea0, 0x1ef9)) low |= 1 << 8; // 1258 Vietnamese
  if (has(0xe00, 0xe7f)) low |= 1 << 16; // 874 Thai
  if (has(0x3040, 0x30ff)) low |= 1 << 17; // 932 Japanese
  if (has(0x4e00, 0x9fff)) low |= (1 << 18) | (1 << 20); // 936, 950 Chinese
  if (has(0xac00, 0xd7af)) low |= 1 << 19; // 949 Korean
  if (low === 0) low = 1;
  return [low >>> 0, 0];
}

/** `name` — the family, style, full and PostScript names, in Windows' encoding. */
function nameTable(face: BuiltFace): Uint8Array {
  const style =
    face.bold && face.italic
      ? 'Bold Italic'
      : face.bold
        ? 'Bold'
        : face.italic
          ? 'Italic'
          : 'Regular';
  const family = face.family.trim() || 'Embedded';
  const postScript =
    face.postScriptName.replace(/[^!-~]|[[\](){}<>/%]/gu, '').slice(0, 63) ||
    family.replace(/[^A-Za-z0-9-]/gu, '');
  const full = style === 'Regular' ? family : `${family} ${style}`;
  const records: Array<[number, string]> = [
    [1, family],
    [2, style],
    [3, `${postScript};${style}`],
    [4, full],
    [5, 'Version 1.000'],
    [6, postScript],
  ];
  const strings = records.map(([, text]) => utf16be(text));
  const storage = 6 + records.length * 12;
  const out = new Writer(storage + strings.reduce((s, b) => s + b.length, 0));
  out.u16(0); // format
  out.u16(records.length);
  out.u16(storage);
  let offset = 0;
  records.forEach(([id], i) => {
    out.u16(3); // Windows
    out.u16(1); // Unicode BMP
    out.u16(0x0409); // English (United States)
    out.u16(id);
    out.u16(strings[i]!.length);
    out.u16(offset);
    offset += strings[i]!.length;
  });
  for (const s of strings) out.bytes(s);
  return out.data;
}

/** `post` version 3.0: no glyph names, only the slant and the pitch. */
function postTable(face: BuiltFace): Uint8Array {
  const out = new Writer(32);
  out.u32(0x00030000);
  out.u32(Math.round((face.italicAngle ?? 0) * 65536) >>> 0);
  out.i16(-100); // underlinePosition
  out.i16(50); // underlineThickness
  out.u32(face.fixedPitch === true ? 1 : 0);
  for (let i = 0; i < 4; i++) out.u32(0);
  return out.data;
}

/**
 * An sfnt from its tables: the directory sorted by tag, each table on a
 * four-byte boundary, and `head.checkSumAdjustment` set so the whole file sums
 * to the magic number.
 */
function assembleSfnt(tables: ReadonlyArray<[string, Uint8Array]>): Uint8Array {
  const sorted = [...tables].sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  const directory = 12 + sorted.length * 16;
  const placed: Array<{ tag: string; data: Uint8Array; offset: number }> = [];
  let cursor = directory;
  for (const [tag, data] of sorted) {
    placed.push({ tag, data, offset: cursor });
    cursor += data.length + ((4 - (data.length % 4)) % 4);
  }
  const out = new Uint8Array(cursor);
  const view = new DataView(out.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, sorted.length);
  const geometry = directoryGeometry(sorted.length);
  view.setUint16(6, geometry.searchRange);
  view.setUint16(8, geometry.entrySelector);
  view.setUint16(10, geometry.rangeShift);
  placed.forEach((t, i) => {
    const at = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[at + k] = t.tag.charCodeAt(k);
    view.setUint32(at + 4, paddedChecksum(t.data));
    view.setUint32(at + 8, t.offset);
    view.setUint32(at + 12, t.data.length);
    out.set(t.data, t.offset);
  });
  const head = placed.find((t) => t.tag === 'head');
  if (head) view.setUint32(head.offset + 8, (0xb1b0afba - paddedChecksum(out)) >>> 0);
  return out;
}

/** Big-endian bytes, written in order into a buffer of known size. */
class Writer {
  readonly data: Uint8Array;
  private readonly view: DataView;
  private at = 0;

  constructor(size: number) {
    this.data = new Uint8Array(size);
    this.view = new DataView(this.data.buffer);
  }

  /** Write an unsigned 16-bit value. */
  u16(v: number): void {
    this.view.setUint16(this.at, v & 0xffff);
    this.at += 2;
  }

  /** Write a signed 16-bit value. */
  i16(v: number): void {
    this.view.setInt16(this.at, clamp(Math.round(v), -0x8000, 0x7fff));
    this.at += 2;
  }

  /** Write an unsigned 32-bit value. */
  u32(v: number): void {
    this.view.setUint32(this.at, v >>> 0);
    this.at += 4;
  }

  /** Write raw bytes. */
  bytes(values: ArrayLike<number>): void {
    for (let i = 0; i < values.length; i++) this.data[this.at + i] = values[i]!;
    this.at += values.length;
  }
}

function utf16be(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    out[i * 2] = text.charCodeAt(i) >> 8;
    out[i * 2 + 1] = text.charCodeAt(i) & 0xff;
  }
  return out;
}

function concat(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v));
}
