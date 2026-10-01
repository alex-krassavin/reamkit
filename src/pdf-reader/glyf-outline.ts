// §9.6.6 — the OUTLINE of one glyph in an embedded TrueType program.
//
// A composite font may be read by glyph index and say nothing about what its
// characters are: `/Encoding /Identity-H` makes every code a CID, a subset
// program has had its `cmap` stripped, and `/ToUnicode` is either absent or
// covers other codes entirely. Eleven files of the pdf.js corpus are exactly
// this, and every one of them reconstructed to a blank sheet —
// complex_ttf_font.pdf is eight lines of Arabic.
//
// There is no character to recover, and inventing one is worse than nothing: a
// private-use codepoint would carry straight into Markdown and HTML as mojibake.
// What CAN be recovered is the shape, which is what the page actually showed.
// So the glyph is drawn — its contours lifted out of `glyf` and painted as a
// path, the way a Type 3 glyph's procedure already is.
//
// The outline comes back in a ONE-UNIT em (each coordinate divided by
// `unitsPerEm`), so the caller places it with the same matrix it would use for
// a Type 3 font whose `/FontMatrix` is `[1/upem 0 0 1/upem 0 0]`.

import { macGlyphName } from './encodings';
import type { PathSeg } from './content';

/** A reader for one font program's outlines, with its tables located once. */
export interface OutlineSource {
  /** The glyph's contours in a one-unit em, or `undefined` for an empty one. */
  readonly path: (gid: number) => Array<PathSeg> | undefined;
  /**
   * How many glyphs the program holds. A glyph INSIDE that count which draws
   * nothing is a blank one — a space — and not a glyph the program lacks.
   */
  readonly count: number;
  /**
   * Whether an sfnt program carries a `cmap` at all. Without one there is no
   * way to reach a glyph by character (§9.6.6.4), which says what the file
   * means by its codes: they are glyph INDICES. Absent for a bare CFF, which
   * has no such table and is addressed by NAME through its charset.
   */
  readonly cmap?: boolean;
}

/**
 * Prepare an embedded TrueType program for outline reading (§9.6.6).
 *
 * @param program The raw `/FontFile2` bytes.
 * @returns A reader, or `undefined` where the program carries no `glyf`/`loca`
 *          pair — an OpenType/CFF program under `/FontFile2` is one such, and
 *          its outlines are charstrings this does not read.
 */
export function outlineSource(program: Uint8Array): OutlineSource | undefined {
  let tables: Map<string, { offset: number; length: number }>;
  try {
    tables = sfntTables(program);
  } catch {
    return undefined;
  }
  const head = tables.get('head');
  const loca = tables.get('loca');
  const glyf = tables.get('glyf');
  if (!head || !loca || !glyf) return undefined;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  const unitsPerEm = safeUint16(view, head.offset + 18) || 1000;
  // `head.indexToLocFormat`: 0 is short (halved offsets), 1 is long.
  const longLoca = safeInt16(view, head.offset + 50) === 1;
  const cache = new Map<number, Array<PathSeg> | undefined>();

  const glyphRange = (gid: number): { start: number; end: number } | undefined => {
    const stride = longLoca ? 4 : 2;
    const at = loca.offset + gid * stride;
    if (at + stride * 2 > loca.offset + loca.length) return undefined;
    const read = (o: number): number => (longLoca ? safeUint32(view, o) : safeUint16(view, o) * 2);
    const start = read(at);
    let end = read(at + stride);
    if (end === start) return undefined; // an empty glyph — a space, say
    if (end < start) {
      // A subsetter that drops glyphs may write their `loca` entries as zero,
      // and the glyph before a dropped one then seems to end before it
      // begins. TCPDF does: bug1650302_reduced.pdf's dotless i and caron both
      // came back empty, and with them the í's stem and the whole of its ř.
      // The glyph runs to the next entry that lies past its start — or to the
      // end of the table, its own counts saying where it stops.
      end = glyf.length;
      for (let o = at + stride * 2; o + stride <= loca.offset + loca.length; o += stride) {
        const next = read(o);
        if (next > start) {
          end = Math.min(next, glyf.length);
          break;
        }
      }
      if (end <= start) return undefined;
    }
    if (glyf.offset + end > program.length) return undefined;
    return { start: glyf.offset + start, end: glyf.offset + end };
  };

  const read = (gid: number, depth: number): Array<PathSeg> | undefined => {
    const range = glyphRange(gid);
    if (!range) return undefined;
    const contours = safeInt16(view, range.start);
    if (contours >= 0) return simpleGlyph(view, range.start, contours, range.end);
    if (depth >= MAX_COMPONENT_DEPTH) return undefined;
    return compositeGlyph(view, range.start + 10, range.end, (sub) => read(sub, depth + 1));
  };

  return {
    count: Math.max(0, Math.floor(loca.length / (longLoca ? 4 : 2)) - 1),
    cmap: (tables.get('cmap')?.length ?? 0) > 0,
    path: (gid: number): Array<PathSeg> | undefined => {
      const had = cache.get(gid);
      if (had !== undefined || cache.has(gid)) return had;
      let out: Array<PathSeg> | undefined;
      try {
        const raw = read(gid, 0);
        out = raw ? raw.map((seg) => scaleSeg(seg, 1 / unitsPerEm)) : undefined;
      } catch {
        out = undefined;
      }
      cache.set(gid, out);
      return out;
    },
  };
}

/** A composite glyph may name another composite; this is where that stops. */
const MAX_COMPONENT_DEPTH = 5;

/**
 * §post format 2.0 — the NAME each glyph of a TrueType program goes by.
 *
 * A legacy eight-bit font is reached through this table and nothing else. Its
 * program carries no `cmap`, so a code cannot be turned into a character and
 * then into a glyph; what it can be turned into is a glyph NAME, through the
 * encoding the font dictionary states (Annex D.2) — and the name is looked up
 * here. TrueType_without_cmap.pdf is an Armenian face, Masis, whose `i` draws
 * ի: read as text its line came back "'>in", and drawn by name it is the line
 * the page shows.
 *
 * @param program The raw `/FontFile2` bytes.
 * @returns Glyph name → index, or `undefined` where the table is missing or
 *          states no names of its own (formats 1.0 and 3.0).
 */
export function postGlyphNames(program: Uint8Array): Map<string, number> | undefined {
  let tables: Map<string, { offset: number; length: number }>;
  try {
    tables = sfntTables(program);
  } catch {
    return undefined;
  }
  const post = tables.get('post');
  if (!post || post.length < 34) return undefined;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  if (safeUint32(view, post.offset) !== POST_NAMED) return undefined;
  const count = safeUint16(view, post.offset + 32);
  const end = post.offset + post.length;
  if (post.offset + 34 + count * 2 > end) return undefined;
  // The Pascal strings after the index array, in the order they are written:
  // index 258 is the first of them, 259 the second, and so on.
  const own: Array<string> = [];
  for (let at = post.offset + 34 + count * 2; at < end; ) {
    const length = program[at] ?? 0;
    if (at + 1 + length > end) break;
    own.push(String.fromCharCode(...program.subarray(at + 1, at + 1 + length)));
    at += 1 + length;
  }
  const out = new Map<string, number>();
  for (let gid = 0; gid < count; gid++) {
    const index = safeUint16(view, post.offset + 34 + gid * 2);
    const name = index < MAC_ORDER_SIZE ? macGlyphName(index) : own[index - MAC_ORDER_SIZE];
    // The first glyph to claim a name keeps it: a subset that leaves several
    // glyphs `.notdef` would otherwise hand the name to the last of them.
    if (name !== undefined && name.length > 0 && !out.has(name)) out.set(name, gid);
  }
  return out.size > 0 ? out : undefined;
}

/**
 * OS/2 `fsType` — the embedding a TrueType or OpenType program's licence allows
 * (ISO/IEC 14496-22, OS/2 table).
 *
 * @param program The raw sfnt bytes.
 * @returns The field, or `undefined` where the program carries no OS/2 table —
 *          a subsetter may drop it, and then the program states nothing.
 */
export function sfntFsType(program: Uint8Array): number | undefined {
  let tables: Map<string, { offset: number; length: number }>;
  try {
    tables = sfntTables(program);
  } catch {
    return undefined;
  }
  const os2 = tables.get('OS/2');
  if (!os2 || os2.length < 10 || os2.offset + 10 > program.length) return undefined;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  return view.getUint16(os2.offset + 8);
}

/**
 * How far the pen moves after an sfnt program's space, in thousandths of an em:
 * the glyph its `cmap` (or its `post` names) gives U+0020, measured in `hmtx`.
 * A subset keeps the metrics of glyphs it dropped the outlines of, and the
 * width a page's words were set apart by is the one a writer should give the
 * space the page never showed.
 *
 * @param program The raw sfnt bytes.
 * @returns The advance, or `undefined` where the program says nothing of a space.
 */
export function sfntSpaceAdvance(program: Uint8Array): number | undefined {
  let tables: Map<string, { offset: number; length: number }>;
  try {
    tables = sfntTables(program);
  } catch {
    return undefined;
  }
  const head = tables.get('head');
  const hhea = tables.get('hhea');
  const hmtx = tables.get('hmtx');
  if (!head || !hhea || !hmtx) return undefined;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  const cmaps = cmapSubtables(program);
  const viaCmap = [
    cmaps.get('3,1')?.(0x20),
    cmaps.get('0,3')?.(0x20),
    cmaps.get('3,0')?.(0xf020),
    cmaps.get('3,0')?.(0x20),
    cmaps.get('1,0')?.(0x20),
  ].find((gid) => gid !== undefined && gid > 0);
  const gid = viaCmap ?? postGlyphNames(program)?.get('space');
  if (gid === undefined || gid <= 0) return undefined;
  const unitsPerEm = safeUint16(view, head.offset + 18) || 1000;
  const metrics = safeUint16(view, hhea.offset + 34);
  if (metrics === 0) return undefined;
  const at = hmtx.offset + Math.min(gid, metrics - 1) * 4;
  if (at + 2 > hmtx.offset + hmtx.length) return undefined;
  const advance = safeUint16(view, at);
  return advance > 0 ? (advance / unitsPerEm) * 1000 : undefined;
}

/**
 * Every `cmap` subtable an sfnt program carries, keyed `platform,encoding` —
 * `3,1` for Windows Unicode, `3,0` for Windows symbol, `1,0` for Mac Roman.
 *
 * A simple TrueType font reaches its glyphs through whichever of these it has,
 * and which one decides what a code is looked up AS (§9.6.6.4): a character, a
 * symbol code, a Mac Roman byte.
 *
 * @param program The raw sfnt bytes.
 * @returns Subtable → code-to-glyph lookup (0 where it maps nothing). Formats
 *          0, 4, 6 and 12 are read; any other is left out.
 */
export function cmapSubtables(program: Uint8Array): Map<string, (code: number) => number> {
  const out = new Map<string, (code: number) => number>();
  let tables: Map<string, { offset: number; length: number }>;
  try {
    tables = sfntTables(program);
  } catch {
    return out;
  }
  const cmap = tables.get('cmap');
  if (!cmap) return out;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  const count = safeUint16(view, cmap.offset + 2);
  for (let i = 0; i < count; i++) {
    const at = cmap.offset + 4 + i * 8;
    const key = `${safeUint16(view, at)},${safeUint16(view, at + 2)}`;
    if (out.has(key)) continue;
    const lookup = cmapLookup(view, cmap.offset + safeUint32(view, at + 4));
    if (lookup) out.set(key, lookup);
  }
  return out;
}

/** One `cmap` subtable as a lookup, for the formats a simple font's program uses. */
function cmapLookup(view: DataView, at: number): ((code: number) => number) | undefined {
  const format = safeUint16(view, at);
  if (format === 0) {
    return (code) => (code >= 0 && code < 256 ? safeUint8(view, at + 6 + code) : 0);
  }
  if (format === 6) {
    const first = safeUint16(view, at + 6);
    const entries = safeUint16(view, at + 8);
    return (code) =>
      code >= first && code < first + entries ? safeUint16(view, at + 10 + (code - first) * 2) : 0;
  }
  if (format === 4) {
    const segments = safeUint16(view, at + 6) >> 1;
    const ends = at + 14;
    const starts = ends + segments * 2 + 2;
    const deltas = starts + segments * 2;
    const ranges = deltas + segments * 2;
    return (code) => {
      if (code < 0 || code > 0xffff) return 0;
      for (let s = 0; s < segments; s++) {
        if (safeUint16(view, ends + s * 2) < code) continue;
        const start = safeUint16(view, starts + s * 2);
        if (start > code) return 0;
        const delta = safeUint16(view, deltas + s * 2);
        const range = safeUint16(view, ranges + s * 2);
        if (range === 0) return (code + delta) & 0xffff;
        const gid = safeUint16(view, ranges + s * 2 + range + (code - start) * 2);
        return gid === 0 ? 0 : (gid + delta) & 0xffff;
      }
      return 0;
    };
  }
  if (format === 12) {
    const groups = safeUint32(view, at + 12);
    return (code) => {
      for (let g = 0; g < groups; g++) {
        const entry = at + 16 + g * 12;
        const start = safeUint32(view, entry);
        const end = safeUint32(view, entry + 4);
        if (code >= start && code <= end) return safeUint32(view, entry + 8) + (code - start);
      }
      return 0;
    };
  }
  return undefined;
}

/** `post` version 2.0 — the only one that states names of its own. */
const POST_NAMED = 0x0002_0000;

/** How many names the Macintosh standard order holds (§post). */
const MAC_ORDER_SIZE = 258;

/** The sfnt table directory: tag → where that table is. */
function sfntTables(program: Uint8Array): Map<string, { offset: number; length: number }> {
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  const out = new Map<string, { offset: number; length: number }>();
  const count = safeUint16(view, 4);
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    if (at + 16 > program.length) break;
    const tag = String.fromCharCode(
      program[at]!,
      program[at + 1]!,
      program[at + 2]!,
      program[at + 3]!,
    );
    out.set(tag, { offset: safeUint32(view, at + 8), length: safeUint32(view, at + 12) });
  }
  return out;
}

// A simple glyph: end points per contour, then the flags and the deltas, which
// are packed — a flag says whether each coordinate is one byte or two and
// whether it repeats the last.
function simpleGlyph(
  view: DataView,
  start: number,
  contours: number,
  end: number,
): Array<PathSeg> | undefined {
  let at = start + 10;
  const ends: Array<number> = [];
  for (let i = 0; i < contours; i++, at += 2) ends.push(safeUint16(view, at));
  const points = contours > 0 ? (ends[contours - 1] ?? -1) + 1 : 0;
  if (points <= 0 || points > MAX_POINTS) return undefined;
  at += 2 + safeUint16(view, at); // skip the hinting instructions
  const flags = new Uint8Array(points);
  for (let i = 0; i < points && at < end; ) {
    const flag = safeUint8(view, at++);
    flags[i++] = flag;
    if ((flag & 0x08) !== 0) {
      // REPEAT: the next byte says how many more points share this flag.
      let repeats = safeUint8(view, at++);
      while (repeats-- > 0 && i < points) flags[i++] = flag;
    }
  }
  const xs = new Int16Array(points);
  let x = 0;
  for (let i = 0; i < points; i++) {
    const flag = flags[i]!;
    if ((flag & 0x02) !== 0) {
      const d = safeUint8(view, at++);
      x += (flag & 0x10) !== 0 ? d : -d;
    } else if ((flag & 0x10) === 0) {
      x += safeInt16(view, at);
      at += 2;
    }
    xs[i] = x;
  }
  const ys = new Int16Array(points);
  let y = 0;
  for (let i = 0; i < points; i++) {
    const flag = flags[i]!;
    if ((flag & 0x04) !== 0) {
      const d = safeUint8(view, at++);
      y += (flag & 0x20) !== 0 ? d : -d;
    } else if ((flag & 0x20) === 0) {
      y += safeInt16(view, at);
      at += 2;
    }
    ys[i] = y;
  }
  const out: Array<PathSeg> = [];
  let from = 0;
  for (const last of ends) {
    if (last >= points) break;
    emitContour(out, flags, xs, ys, from, last);
    from = last + 1;
  }
  return out.length > 0 ? out : undefined;
}

/** How many points one glyph may hold — a guard, not a format limit. */
const MAX_POINTS = 10000;

// One closed contour. TrueType curves are QUADRATIC and a contour may hold two
// off-curve points in a row, which implies an on-curve point halfway between
// them; the quadratics become cubics because that is what a path carries.
function emitContour(
  out: Array<PathSeg>,
  flags: Uint8Array,
  xs: Int16Array,
  ys: Int16Array,
  from: number,
  last: number,
): void {
  const n = last - from + 1;
  if (n <= 0) return;
  const on = (i: number): boolean => (flags[from + (((i % n) + n) % n)]! & 0x01) !== 0;
  const px = (i: number): number => xs[from + (((i % n) + n) % n)]!;
  const py = (i: number): number => ys[from + (((i % n) + n) % n)]!;
  // The contour starts at its first on-curve point; where every point is off
  // the curve, it starts at the midpoint between the last and the first.
  let startIdx = 0;
  while (startIdx < n && !on(startIdx)) startIdx++;
  let cx: number;
  let cy: number;
  if (startIdx === n) {
    startIdx = 0;
    cx = (px(0) + px(n - 1)) / 2;
    cy = (py(0) + py(n - 1)) / 2;
  } else {
    cx = px(startIdx);
    cy = py(startIdx);
  }
  out.push({ op: 'move', x: cx, y: cy });
  let i = startIdx + 1;
  const stop = startIdx + n;
  while (i <= stop) {
    if (on(i)) {
      out.push({ op: 'line', x: px(i), y: py(i) });
      cx = px(i);
      cy = py(i);
      i++;
      continue;
    }
    const qx = px(i);
    const qy = py(i);
    // The point after a control point ends the curve — or, where it is another
    // control point, the implied midpoint between the two does.
    const nextOn = on(i + 1);
    const ex = nextOn ? px(i + 1) : (qx + px(i + 1)) / 2;
    const ey = nextOn ? py(i + 1) : (qy + py(i + 1)) / 2;
    out.push({
      op: 'cubic',
      x1: cx + (2 / 3) * (qx - cx),
      y1: cy + (2 / 3) * (qy - cy),
      x2: ex + (2 / 3) * (qx - ex),
      y2: ey + (2 / 3) * (qy - ey),
      x: ex,
      y: ey,
    });
    cx = ex;
    cy = ey;
    i += nextOn ? 2 : 1;
  }
  out.push({ op: 'close' });
}

// A composite glyph places other glyphs, each with its own offset and optional
// 2×2 transform. The offsets may be point INDICES rather than coordinates,
// which this does not follow — that form is vanishingly rare and placing the
// component at the origin is closer than dropping the glyph.
function compositeGlyph(
  view: DataView,
  start: number,
  end: number,
  component: (gid: number) => Array<PathSeg> | undefined,
): Array<PathSeg> | undefined {
  const out: Array<PathSeg> = [];
  let at = start;
  for (let guard = 0; guard < MAX_COMPONENTS && at + 4 <= end; guard++) {
    const flags = safeUint16(view, at);
    const gid = safeUint16(view, at + 2);
    at += 4;
    let dx = 0;
    let dy = 0;
    if ((flags & 0x0001) !== 0) {
      // ARG_1_AND_2_ARE_WORDS
      if ((flags & 0x0002) !== 0) {
        dx = safeInt16(view, at);
        dy = safeInt16(view, at + 2);
      }
      at += 4;
    } else {
      if ((flags & 0x0002) !== 0) {
        dx = safeInt8(view, at);
        dy = safeInt8(view, at + 1);
      }
      at += 2;
    }
    let a = 1;
    let b = 0;
    let c = 0;
    let d = 1;
    if ((flags & 0x0008) !== 0) {
      a = d = f2dot14(view, at);
      at += 2;
    } else if ((flags & 0x0040) !== 0) {
      a = f2dot14(view, at);
      d = f2dot14(view, at + 2);
      at += 4;
    } else if ((flags & 0x0080) !== 0) {
      a = f2dot14(view, at);
      b = f2dot14(view, at + 2);
      c = f2dot14(view, at + 4);
      d = f2dot14(view, at + 6);
      at += 8;
    }
    const sub = component(gid);
    if (sub) for (const seg of sub) out.push(transformSeg(seg, a, b, c, d, dx, dy));
    if ((flags & 0x0020) === 0) break; // MORE_COMPONENTS
  }
  return out.length > 0 ? out : undefined;
}

/** How many components one composite glyph may place. */
const MAX_COMPONENTS = 32;

function transformSeg(
  seg: PathSeg,
  a: number,
  b: number,
  c: number,
  d: number,
  dx: number,
  dy: number,
): PathSeg {
  const at = (x: number, y: number): [number, number] => [a * x + c * y + dx, b * x + d * y + dy];
  if (seg.op === 'close') return seg;
  if (seg.op === 'cubic') {
    const [x1, y1] = at(seg.x1, seg.y1);
    const [x2, y2] = at(seg.x2, seg.y2);
    const [x, y] = at(seg.x, seg.y);
    return { op: 'cubic', x1, y1, x2, y2, x, y };
  }
  const [x, y] = at(seg.x, seg.y);
  return { op: seg.op, x, y };
}

function scaleSeg(seg: PathSeg, k: number): PathSeg {
  if (seg.op === 'close') return seg;
  if (seg.op === 'cubic') {
    return {
      op: 'cubic',
      x1: seg.x1 * k,
      y1: seg.y1 * k,
      x2: seg.x2 * k,
      y2: seg.y2 * k,
      x: seg.x * k,
      y: seg.y * k,
    };
  }
  return { op: seg.op, x: seg.x * k, y: seg.y * k };
}

/** The 2.14 fixed-point a component transform is stated in. */
function f2dot14(view: DataView, at: number): number {
  return safeInt16(view, at) / 16384;
}

// A truncated font is a fact of life in this corpus; reading past the end
// returns zero rather than throwing, and a glyph built from zeros is empty.
function safeUint8(view: DataView, at: number): number {
  return at >= 0 && at < view.byteLength ? view.getUint8(at) : 0;
}
function safeInt8(view: DataView, at: number): number {
  return at >= 0 && at < view.byteLength ? view.getInt8(at) : 0;
}
function safeUint16(view: DataView, at: number): number {
  return at >= 0 && at + 2 <= view.byteLength ? view.getUint16(at) : 0;
}
function safeInt16(view: DataView, at: number): number {
  return at >= 0 && at + 2 <= view.byteLength ? view.getInt16(at) : 0;
}
function safeUint32(view: DataView, at: number): number {
  return at >= 0 && at + 4 <= view.byteLength ? view.getUint32(at) : 0;
}
