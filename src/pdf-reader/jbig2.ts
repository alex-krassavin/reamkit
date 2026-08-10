// E-PDF — JBIG2 (ISO/IEC 14492 / ITU-T T.88), the embedded profile PDF's
// `/JBIG2Decode` filter carries (ISO 32000-1 §7.4.7).
//
// JBIG2 is how a scanner stores a page of text: a bilevel image coded either as
// one generic region of arithmetic-coded pixels, or — the reason the format
// exists — as a DICTIONARY of the glyph shapes on the page plus a list of where
// each one is placed, so the letter "e" is stored once for the whole document.
//
// Unread, such a page is not degraded but ABSENT: the filter passed through and
// the image came back empty, so a quarter of the pdf.js corpus (96 files) scored
// a flat 1.000 against its own rendering.
//
// This decodes:
//   §6.2  generic region — arithmetic templates 0-3 with their AT pixels, the
//         typical-prediction skip (TPGDON), and MMR (which is T.6, so the CCITT
//         decoder does it)
//   §6.3  refinement region — templates 0-1, TPGRON
//   §7.4  page composition — OR / AND / XOR / XNOR / REPLACE, striped pages
//
// The symbol dictionary, text region and halftone region are a later stage.

import { decodeCcitt } from './ccitt';

/** A bilevel raster, one byte per pixel, 1 = black — JBIG2's own sense. */
export interface Jbig2Bitmap {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

function makeBitmap(width: number, height: number, fill = 0): Jbig2Bitmap {
  const data = new Uint8Array(width * height);
  if (fill) data.fill(1);
  return { width, height, data };
}

// ---------------------------------------------------------------- MQ decoder

/**
 * §E.1 Table E.1 — the MQ-coder's probability estimation state machine, as
 * `[Qe, NMPS, NLPS, SWITCH]` per state. The same table drives JPEG 2000.
 */
const QE: ReadonlyArray<readonly [number, number, number, number]> = [
  [0x5601, 1, 1, 1],
  [0x3401, 2, 6, 0],
  [0x1801, 3, 9, 0],
  [0x0ac1, 4, 12, 0],
  [0x0521, 5, 29, 0],
  [0x0221, 38, 33, 0],
  [0x5601, 7, 6, 1],
  [0x5401, 8, 14, 0],
  [0x4801, 9, 14, 0],
  [0x3801, 10, 14, 0],
  [0x3001, 11, 17, 0],
  [0x2401, 12, 18, 0],
  [0x1c01, 13, 20, 0],
  [0x1601, 29, 21, 0],
  [0x5601, 15, 14, 1],
  [0x5401, 16, 14, 0],
  [0x5101, 17, 15, 0],
  [0x4801, 18, 16, 0],
  [0x3801, 19, 17, 0],
  [0x3401, 20, 18, 0],
  [0x3001, 21, 19, 0],
  [0x2801, 22, 19, 0],
  [0x2401, 23, 20, 0],
  [0x2201, 24, 21, 0],
  [0x1c01, 25, 22, 0],
  [0x1801, 26, 23, 0],
  [0x1601, 27, 24, 0],
  [0x1401, 28, 25, 0],
  [0x1201, 29, 26, 0],
  [0x1101, 30, 27, 0],
  [0x0ac1, 31, 28, 0],
  [0x09c1, 32, 29, 0],
  [0x08a1, 33, 30, 0],
  [0x0521, 34, 31, 0],
  [0x0441, 35, 32, 0],
  [0x02a1, 36, 33, 0],
  [0x0221, 37, 34, 0],
  [0x0141, 38, 35, 0],
  [0x0111, 39, 36, 0],
  [0x0085, 40, 37, 0],
  [0x0049, 41, 38, 0],
  [0x0025, 42, 39, 0],
  [0x0015, 43, 40, 0],
  [0x0009, 44, 41, 0],
  [0x0005, 45, 42, 0],
  [0x0001, 45, 43, 0],
  [0x5601, 46, 46, 0],
];

/** A context's state: the Table E.1 index and which symbol is currently more probable. */
interface Cx {
  i: Uint8Array;
  mps: Uint8Array;
}

/** §E.3.5 — a fresh set of `size` contexts, all at state 0 with MPS 0. */
export function newContexts(size: number): Cx {
  return { i: new Uint8Array(size), mps: new Uint8Array(size) };
}

/**
 * §E.3 — the MQ arithmetic decoder.
 *
 * Written as a class with the spec's own register names (`c`, `a`, `ct`)
 * because the procedures below are transcribed from its flow charts, and a
 * renaming makes them impossible to check against it.
 */
export class MQDecoder {
  private readonly data: Uint8Array;
  private bp: number;
  private readonly end: number;
  private c = 0;
  private a = 0;
  private ct = 0;

  constructor(data: Uint8Array, start = 0, end = data.length) {
    this.data = data;
    this.bp = start;
    this.end = end;
    // INITDEC (§E.3.5)
    this.c = (this.byteAt(this.bp) << 16) >>> 0;
    this.byteIn();
    this.c = (this.c << 7) >>> 0;
    this.ct -= 7;
    this.a = 0x8000;
  }

  private byteAt(i: number): number {
    return i < this.end ? (this.data[i] ?? 0xff) : 0xff;
  }

  /** BYTEIN (§E.3.4) — 0xFF is the marker escape, so it feeds a shorter byte. */
  private byteIn(): void {
    if (this.byteAt(this.bp) === 0xff) {
      if (this.byteAt(this.bp + 1) > 0x8f) {
        this.c = (this.c + 0xff00) >>> 0;
        this.ct = 8;
      } else {
        this.bp++;
        this.c = (this.c + (this.byteAt(this.bp) << 9)) >>> 0;
        this.ct = 7;
      }
    } else {
      this.bp++;
      this.c = (this.c + (this.byteAt(this.bp) << 8)) >>> 0;
      this.ct = 8;
    }
  }

  /**
   * DECODE (§E.3.2) — one bit in context `k`.
   *
   * @param cx The context set the caller keeps for this decoding procedure.
   * @param k  Which context within it.
   */
  decode(cx: Cx, k: number): number {
    let i = cx.i[k] ?? 0;
    let mps = cx.mps[k] ?? 0;
    const entry = QE[i] ?? QE[0]!;
    const [qe, nmps, nlps, sw] = entry;
    let d: number;
    this.a = (this.a - qe) & 0xffff;
    if (((this.c >>> 16) & 0xffff) < qe) {
      // LPS exchange / MPS exchange (§E.3.2), then RENORMD.
      if (this.a < qe) {
        this.a = qe;
        d = mps;
        i = nmps;
      } else {
        this.a = qe;
        d = 1 - mps;
        if (sw === 1) mps = 1 - mps;
        i = nlps;
      }
    } else {
      this.c = (this.c - (qe << 16)) >>> 0;
      if ((this.a & 0x8000) !== 0) return mps;
      if (this.a < qe) {
        d = 1 - mps;
        if (sw === 1) mps = 1 - mps;
        i = nlps;
      } else {
        d = mps;
        i = nmps;
      }
    }
    // RENORMD (§E.3.3)
    do {
      if (this.ct === 0) this.byteIn();
      this.a = (this.a << 1) & 0xffff;
      this.c = (this.c << 1) >>> 0;
      this.ct--;
    } while ((this.a & 0x8000) === 0);
    cx.i[k] = i;
    cx.mps[k] = mps;
    return d;
  }
}

// ------------------------------------------------------------ generic region

/** An adaptive template pixel — a coordinate the coder may move (§6.2.5.3). */
interface At {
  readonly x: number;
  readonly y: number;
}

/** The nominal AT positions per template (§6.2.5.3 figures 8-11). */
const NOMINAL_AT: ReadonlyArray<ReadonlyArray<At>> = [
  [
    { x: 3, y: -1 },
    { x: -3, y: -1 },
    { x: 2, y: -2 },
    { x: -2, y: -2 },
  ],
  [{ x: 3, y: -1 }],
  [{ x: 2, y: -1 }],
  [{ x: 2, y: -1 }],
];

/**
 * §6.2.5.7 — decode a generic region with the arithmetic coder.
 *
 * The context for a pixel is the pattern of already-decoded neighbours the
 * template names, read MSB first in the order the figures give. `TPGDON` adds a
 * per-row bit that says "this row is the same as the one above", which is what
 * makes a page of white space nearly free.
 *
 * @param mq       The decoder, positioned at the region's data.
 * @param cx       The 2^16 generic contexts (shared across a segment's regions).
 * @param width    Region width in pixels.
 * @param height   Region height in pixels.
 * @param template Which of the four templates (0-3).
 * @param at       The adaptive pixels, one for templates 1-3 and four for 0.
 * @param tpgdon   Whether typical prediction is on.
 * @param skip     A bitmap of pixels to leave white without decoding (halftone).
 */
export function decodeGenericRegion(
  mq: MQDecoder,
  cx: Cx,
  width: number,
  height: number,
  template: number,
  at: ReadonlyArray<At>,
  tpgdon: boolean,
  skip?: Jbig2Bitmap,
): Jbig2Bitmap {
  const bmp = makeBitmap(width, height);
  const px = (x: number, y: number): number =>
    x < 0 || x >= width || y < 0 ? 0 : (bmp.data[y * width + x] ?? 0);
  // §6.2.5.7 — the context the TPGDON bit is decoded in, per template.
  const TPGD_CX = [0x9b25, 0x0795, 0x00e5, 0x0195];
  let ltp = 0;
  for (let y = 0; y < height; y++) {
    if (tpgdon) {
      ltp ^= mq.decode(cx, TPGD_CX[template] ?? 0);
      if (ltp === 1) {
        // The row repeats the one above it — copy and decode nothing.
        if (y > 0) bmp.data.copyWithin(y * width, (y - 1) * width, y * width);
        continue;
      }
    }
    for (let x = 0; x < width; x++) {
      if (skip && skip.data[y * width + x] === 1) continue;
      let ctx: number;
      const a = at;
      switch (template) {
        case 0:
          ctx =
            (px(x - 1, y) << 0) |
            (px(x - 2, y) << 1) |
            (px(x - 3, y) << 2) |
            (px(x - 4, y) << 3) |
            (px(x + (a[0]?.x ?? 3), y + (a[0]?.y ?? -1)) << 4) |
            (px(x + 2, y - 1) << 5) |
            (px(x + 1, y - 1) << 6) |
            (px(x, y - 1) << 7) |
            (px(x - 1, y - 1) << 8) |
            (px(x - 2, y - 1) << 9) |
            (px(x + (a[1]?.x ?? -3), y + (a[1]?.y ?? -1)) << 10) |
            (px(x + (a[2]?.x ?? 2), y + (a[2]?.y ?? -2)) << 11) |
            (px(x + 1, y - 2) << 12) |
            (px(x, y - 2) << 13) |
            (px(x - 1, y - 2) << 14) |
            (px(x + (a[3]?.x ?? -2), y + (a[3]?.y ?? -2)) << 15);
          break;
        case 1:
          ctx =
            (px(x - 1, y) << 0) |
            (px(x - 2, y) << 1) |
            (px(x - 3, y) << 2) |
            (px(x + (a[0]?.x ?? 3), y + (a[0]?.y ?? -1)) << 3) |
            (px(x + 2, y - 1) << 4) |
            (px(x + 1, y - 1) << 5) |
            (px(x, y - 1) << 6) |
            (px(x - 1, y - 1) << 7) |
            (px(x - 2, y - 1) << 8) |
            (px(x + 2, y - 2) << 9) |
            (px(x + 1, y - 2) << 10) |
            (px(x, y - 2) << 11) |
            (px(x - 1, y - 2) << 12);
          break;
        case 2:
          ctx =
            (px(x - 1, y) << 0) |
            (px(x - 2, y) << 1) |
            (px(x + (a[0]?.x ?? 2), y + (a[0]?.y ?? -1)) << 2) |
            (px(x + 1, y - 1) << 3) |
            (px(x, y - 1) << 4) |
            (px(x - 1, y - 1) << 5) |
            (px(x - 2, y - 1) << 6) |
            (px(x + 1, y - 2) << 7) |
            (px(x, y - 2) << 8) |
            (px(x - 1, y - 2) << 9);
          break;
        default:
          ctx =
            (px(x - 1, y) << 0) |
            (px(x - 2, y) << 1) |
            (px(x - 3, y) << 2) |
            (px(x - 4, y) << 3) |
            (px(x + (a[0]?.x ?? 2), y + (a[0]?.y ?? -1)) << 4) |
            (px(x + 1, y - 1) << 5) |
            (px(x, y - 1) << 6) |
            (px(x - 1, y - 1) << 7) |
            (px(x - 2, y - 1) << 8) |
            (px(x - 3, y - 1) << 9);
          break;
      }
      bmp.data[y * width + x] = mq.decode(cx, ctx);
    }
  }
  return bmp;
}

// --------------------------------------------------------- refinement region

/** The nominal refinement AT positions (§6.3.5.3). */
const NOMINAL_AT_REFINE: ReadonlyArray<At> = [
  { x: -1, y: -1 },
  { x: -1, y: -1 },
];

/**
 * §6.3.5.6 — refine `reference` into a bitmap of `width` × `height`.
 *
 * Refinement codes a bitmap against one already decoded — the same glyph at a
 * lower resolution, or the same region in a previous pass — so only where they
 * differ costs anything.
 *
 * @param dx Where the reference sits relative to the region (GRREFERENCEDX).
 * @param dy Likewise, vertically.
 */
/**
 * §6.3.5.3 figures 12-14 — the refinement templates, as the pixels they read
 * and the ORDER they read them in: every pixel of the bitmap being coded, then
 * every pixel of the reference, most significant bit first. Template 0 takes an
 * adaptive pixel in each list; template 1 is fixed.
 */
const REFINE_TEMPLATE: ReadonlyArray<{
  readonly coding: ReadonlyArray<At>;
  readonly reference: ReadonlyArray<At>;
}> = [
  {
    coding: [
      { x: 0, y: -1 },
      { x: 1, y: -1 },
      { x: -1, y: 0 },
    ],
    reference: [
      { x: 0, y: -1 },
      { x: 1, y: -1 },
      { x: -1, y: 0 },
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: -1, y: 1 },
      { x: 0, y: 1 },
      { x: 1, y: 1 },
    ],
  },
  {
    coding: [
      { x: -1, y: -1 },
      { x: 0, y: -1 },
      { x: 1, y: -1 },
      { x: -1, y: 0 },
    ],
    reference: [
      { x: 0, y: -1 },
      { x: -1, y: 0 },
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: 0, y: 1 },
      { x: 1, y: 1 },
    ],
  },
];

/** §6.3.5.6 — the context the TPGRON bit is read in, per template. */
const REFINE_TPGR_CX = [0x0020, 0x0008];

/**
 * §6.3.5.6 — refine `reference` into a bitmap of `width` × `height`.
 *
 * Refinement codes a bitmap against one already decoded — the same glyph at a
 * lower resolution, or the same region in an earlier pass — so only where the
 * two differ costs anything.
 *
 * @param dx Where the reference sits relative to the region (GRREFERENCEDX).
 * @param dy Likewise, vertically.
 */
export function decodeRefinement(
  mq: MQDecoder,
  cx: Cx,
  width: number,
  height: number,
  template: number,
  reference: Jbig2Bitmap,
  dx: number,
  dy: number,
  at: ReadonlyArray<At>,
  tpgron: boolean,
): Jbig2Bitmap {
  const bmp = makeBitmap(width, height);
  const t = REFINE_TEMPLATE[template === 0 ? 0 : 1]!;
  // Template 0 takes one adaptive pixel into each list, at its end.
  const coding = template === 0 ? [...t.coding, at[0] ?? NOMINAL_AT_REFINE[0]!] : t.coding;
  const ref = template === 0 ? [...t.reference, at[1] ?? NOMINAL_AT_REFINE[1]!] : t.reference;

  const px = (x: number, y: number): number =>
    x < 0 || x >= width || y < 0 || y >= height ? 0 : (bmp.data[y * width + x] ?? 0);
  const rf = (x: number, y: number): number => {
    const rx = x - dx;
    const ry = y - dy;
    return rx < 0 || rx >= reference.width || ry < 0 || ry >= reference.height
      ? 0
      : (reference.data[ry * reference.width + rx] ?? 0);
  };

  let ltp = 0;
  for (let y = 0; y < height; y++) {
    if (tpgron) ltp ^= mq.decode(cx, REFINE_TPGR_CX[template === 0 ? 0 : 1]!);
    for (let x = 0; x < width; x++) {
      if (ltp === 1) {
        // §6.3.5.6 — where the reference's 3×3 neighbourhood is all one value,
        // the refined pixel takes it and is not coded at all.
        const s =
          rf(x - 1, y - 1) +
          rf(x, y - 1) +
          rf(x + 1, y - 1) +
          rf(x - 1, y) +
          rf(x, y) +
          rf(x + 1, y) +
          rf(x - 1, y + 1) +
          rf(x, y + 1) +
          rf(x + 1, y + 1);
        if (s === 0 || s === 9) {
          bmp.data[y * width + x] = s === 9 ? 1 : 0;
          continue;
        }
      }
      let ctx = 0;
      for (const p of coding) ctx = (ctx << 1) | px(x + p.x, y + p.y);
      for (const p of ref) ctx = (ctx << 1) | rf(x + p.x, y + p.y);
      bmp.data[y * width + x] = mq.decode(cx, ctx);
    }
  }
  return bmp;
}

// ------------------------------------------------- arithmetic integer decoding

/** §A.2 returns this instead of a number when the coder signals out-of-band. */
const OOB = Symbol('OOB');

/**
 * §A.2 — the integer decoding procedure.
 *
 * Every count, coordinate and delta in a symbol dictionary or text region comes
 * through here: a sign bit, then a prefix that says how many magnitude bits
 * follow and what to add to them. The context is a walk down a binary tree of
 * the bits read so far, which is why each kind of integer (IADH, IADW, IADT …)
 * needs a context set of its own.
 */
class IntDecoder {
  private readonly cx = newContexts(512);
  constructor(private readonly mq: MQDecoder) {}

  decode(): number | typeof OOB {
    let prev = 1;
    const bit = (): number => {
      const b = this.mq.decode(this.cx, prev);
      prev = prev < 256 ? (prev << 1) | b : ((((prev << 1) | b) & 511) | 256) >>> 0;
      return b;
    };
    const sign = bit();
    // §A.2 — the prefix is a run of up to five 1-bits, and each one read means
    // a longer magnitude field with a bigger base added to it.
    const FIELDS: ReadonlyArray<readonly [number, number]> = [
      [2, 0],
      [4, 4],
      [6, 20],
      [8, 84],
      [12, 340],
      [32, 4436],
    ];
    let level = 0;
    while (level < 5 && bit() === 1) level++;
    const [n, offset] = FIELDS[level]!;
    let v = 0;
    for (let i = 0; i < n; i++) v = (v << 1) | bit();
    v += offset;
    if (sign === 0) return v;
    // §A.2 — a negative zero is the out-of-band signal that ends a run.
    return v === 0 ? OOB : -v;
  }
}

/** §A.3 — the symbol ID decoder, a fixed-width code down its own context tree. */
class IdDecoder {
  private readonly cx: Cx;
  constructor(
    private readonly mq: MQDecoder,
    private readonly codeLen: number,
  ) {
    this.cx = newContexts(1 << (codeLen + 1));
  }

  decode(): number {
    let prev = 1;
    for (let i = 0; i < this.codeLen; i++) prev = (prev << 1) | this.mq.decode(this.cx, prev);
    return prev - (1 << this.codeLen);
  }
}

// --------------------------------------------------- symbol dictionary + text

/** How a text region places each instance against the point it names (§6.4.5). */
const enum Corner {
  BottomLeft = 0,
  TopLeft = 1,
  BottomRight = 2,
  TopRight = 3,
}

/** The parameters a text region needs, however they were coded. */
interface TextRegionParams {
  readonly width: number;
  readonly height: number;
  readonly instances: number;
  readonly stripT: number;
  readonly refCorner: Corner;
  readonly transposed: boolean;
  readonly combOp: number;
  readonly defPixel: number;
  readonly dsOffset: number;
  readonly refine: boolean;
  readonly rTemplate: number;
  readonly rAt: ReadonlyArray<At>;
}

/**
 * §6.4.5 — draw a text region: a run of strips, each holding instances of the
 * symbols in `symbols`, placed by running coordinates rather than absolute ones.
 *
 * This is what JBIG2 is for. A scanned page is not stored as pixels but as "the
 * shape called 37, here; the shape called 12, four pixels on" — so the letter
 * "e" costs its bitmap once and a few bits per occurrence after that.
 */
function decodeTextRegion(
  mq: MQDecoder,
  symbols: ReadonlyArray<Jbig2Bitmap>,
  p: TextRegionParams,
  stripsLog: number,
): Jbig2Bitmap {
  const bmp = makeBitmap(p.width, p.height, p.defPixel);
  const strips = 1 << stripsLog;
  const codeLen = Math.max(1, Math.ceil(Math.log2(Math.max(symbols.length, 1))));
  const iadt = new IntDecoder(mq);
  const iafs = new IntDecoder(mq);
  const iads = new IntDecoder(mq);
  const iait = new IntDecoder(mq);
  const iari = new IntDecoder(mq);
  const iardw = new IntDecoder(mq);
  const iardh = new IntDecoder(mq);
  const iardx = new IntDecoder(mq);
  const iardy = new IntDecoder(mq);
  const iaid = new IdDecoder(mq, codeLen);
  const refineCx = newContexts(1 << 13);

  const num = (v: number | typeof OOB): number => (v === OOB ? 0 : v);
  let stripT = -num(iadt.decode()) * strips;
  let firstS = 0;
  let placed = 0;

  while (placed < p.instances) {
    stripT += num(iadt.decode()) * strips;
    firstS += num(iafs.decode());
    let curS = firstS;
    for (;;) {
      const t = strips === 1 ? 0 : num(iait.decode());
      const tt = stripT + t;
      const id = iaid.decode();
      let sym = symbols[id] ?? makeBitmap(1, 1);
      if (p.refine && num(iari.decode()) !== 0) {
        // §6.4.11 — this instance is not the dictionary's shape but a
        // refinement of it, which is how a scan keeps the one "e" that smudged.
        const rdw = num(iardw.decode());
        const rdh = num(iardh.decode());
        const rdx = num(iardx.decode());
        const rdy = num(iardy.decode());
        const w = sym.width + rdw;
        const h = sym.height + rdh;
        if (w > 0 && h > 0 && w < 1 << 14 && h < 1 << 14) {
          sym = decodeRefinement(
            mq,
            refineCx,
            w,
            h,
            p.rTemplate,
            sym,
            (rdw >> 1) + rdx,
            (rdh >> 1) + rdy,
            p.rAt,
            false,
          );
        }
      }
      const sw = sym.width;
      const sh = sym.height;
      // §6.4.5 step 3(c)(x) — where the instance's own box goes, given which
      // of its corners the running coordinates name.
      let x: number;
      let y: number;
      if (p.transposed) {
        x = tt;
        y = curS;
        if (p.refCorner === Corner.BottomLeft || p.refCorner === Corner.BottomRight) {
          // nothing: in transposed placement S runs down and T across
        }
        if (p.refCorner === Corner.TopRight || p.refCorner === Corner.BottomRight) x = tt - sw + 1;
      } else {
        x = curS;
        y = tt;
        if (p.refCorner === Corner.BottomLeft || p.refCorner === Corner.BottomRight) {
          y = tt - sh + 1;
        }
      }
      compose(bmp, sym, x, y, p.combOp);
      curS += (p.transposed ? sh : sw) - 1;
      placed++;
      if (placed >= p.instances) break;
      const ids = iads.decode();
      if (ids === OOB) break; // the strip ends
      curS += ids + p.dsOffset;
    }
  }
  return bmp;
}

/**
 * §6.5 — decode a symbol dictionary: the shapes a text region will place.
 *
 * Shapes come in HEIGHT CLASSES, tallest dimension first: one delta says how
 * much taller this class is than the last, then a run of widths within it, each
 * ending when the width decoder signals out-of-band.
 *
 * @param input The symbols this dictionary inherits from the ones it refers to.
 */
function decodeSymbolDictionary(
  mq: MQDecoder,
  input: ReadonlyArray<Jbig2Bitmap>,
  newCount: number,
  exportCount: number,
  template: number,
  at: ReadonlyArray<At>,
  refAgg: boolean,
  rTemplate: number,
  rAt: ReadonlyArray<At>,
): Array<Jbig2Bitmap> {
  const iadh = new IntDecoder(mq);
  const iadw = new IntDecoder(mq);
  const iaex = new IntDecoder(mq);
  const iaai = new IntDecoder(mq);
  const iardx = new IntDecoder(mq);
  const iardy = new IntDecoder(mq);
  const genericCx = newContexts(1 << 16);
  const refineCx = newContexts(1 << 13);
  const newSymbols: Array<Jbig2Bitmap> = [];
  const num = (v: number | typeof OOB): number => (v === OOB ? 0 : v);

  let height = 0;
  let guard = 0;
  while (newSymbols.length < newCount && guard++ < 10000) {
    height += num(iadh.decode());
    let width = 0;
    for (;;) {
      const dw = iadw.decode();
      if (dw === OOB) break;
      width += dw;
      if (newSymbols.length >= newCount) break;
      if (width <= 0 || height <= 0 || width > 1 << 14 || height > 1 << 14) {
        newSymbols.push(makeBitmap(1, 1));
        continue;
      }
      if (!refAgg) {
        newSymbols.push(decodeGenericRegion(mq, genericCx, width, height, template, at, false));
      } else {
        const instances = num(iaai.decode());
        if (instances === 1) {
          // §6.5.8.2.2 — one instance: the new shape is a refinement of an
          // existing one, which is how a dictionary stores a family of glyphs.
          const all = [...input, ...newSymbols];
          const codeLen = Math.max(1, Math.ceil(Math.log2(Math.max(all.length + newCount, 1))));
          const id = new IdDecoder(mq, codeLen).decode();
          const rdx = num(iardx.decode());
          const rdy = num(iardy.decode());
          newSymbols.push(
            decodeRefinement(
              mq,
              refineCx,
              width,
              height,
              rTemplate,
              all[id] ?? makeBitmap(1, 1),
              rdx,
              rdy,
              rAt,
              false,
            ),
          );
        } else {
          // §6.5.8.2 — several instances: the shape is a little text region
          // drawn from the symbols known so far.
          const all = [...input, ...newSymbols];
          newSymbols.push(
            decodeTextRegion(
              mq,
              all,
              {
                width,
                height,
                instances,
                stripT: 0,
                refCorner: Corner.TopLeft,
                transposed: false,
                combOp: 0,
                defPixel: 0,
                dsOffset: 0,
                refine: true,
                rTemplate,
                rAt,
              },
              0,
            ),
          );
        }
      }
    }
  }

  // §6.5.10 — the export flags: alternating runs of "not exported" and
  // "exported" over the input symbols followed by the new ones.
  const all = [...input, ...newSymbols];
  const exported: Array<Jbig2Bitmap> = [];
  let i = 0;
  let exporting = false;
  let steps = 0;
  while (i < all.length && exported.length < exportCount && steps++ < 10000) {
    const run = num(iaex.decode());
    if (exporting) for (let k = 0; k < run && i + k < all.length; k++) exported.push(all[i + k]!);
    i += run;
    exporting = !exporting;
    if (run === 0 && steps > all.length * 2 + 4) break;
  }
  return exported.length > 0 ? exported : newSymbols;
}

// ------------------------------------------------------ halftone + patterns

/**
 * §6.7 — a pattern dictionary: the cells a halftone region tiles a picture out
 * of, from all-white through all-black.
 *
 * They are coded as ONE wide generic region — every pattern side by side — and
 * cut apart afterwards, which is why the first adaptive pixel sits a whole
 * pattern-width back: it makes each cell predict from the one before it.
 */
function decodePatternDictionary(
  data: Uint8Array,
  mmr: boolean,
  template: number,
  pw: number,
  ph: number,
  grayMax: number,
): Array<Jbig2Bitmap> {
  const count = grayMax + 1;
  const width = count * pw;
  let collective: Jbig2Bitmap;
  if (mmr) {
    const packed = decodeCcitt(data, { k: -1, columns: width, rows: ph, byteAlign: false });
    if (!packed) return [];
    collective = makeBitmap(width, ph);
    const rowBytes = (width + 7) >> 3;
    for (let y = 0; y < ph; y++) {
      for (let x = 0; x < width; x++) {
        collective.data[y * width + x] = (packed[y * rowBytes + (x >> 3)]! >> (7 - (x & 7))) & 1;
      }
    }
  } else {
    collective = decodeGenericRegion(
      new MQDecoder(data),
      newContexts(1 << 16),
      width,
      ph,
      template,
      [
        { x: -pw, y: 0 },
        { x: -3, y: -1 },
        { x: 2, y: -2 },
        { x: -2, y: -2 },
      ],
      false,
    );
  }
  const out: Array<Jbig2Bitmap> = [];
  for (let i = 0; i < count; i++) {
    const cell = makeBitmap(pw, ph);
    for (let y = 0; y < ph; y++) {
      for (let x = 0; x < pw; x++) {
        cell.data[y * pw + x] = collective.data[y * width + i * pw + x] ?? 0;
      }
    }
    out.push(cell);
  }
  return out;
}

/**
 * §C.5 — the grey-scale image a halftone region is built from: one bitmap per
 * BIT of the value, most significant first, Gray-coded so neighbouring levels
 * differ in one plane and the planes stay smooth enough to code well.
 */
function decodeGrayScale(
  mq: MQDecoder | undefined,
  data: Uint8Array,
  mmr: boolean,
  template: number,
  at: ReadonlyArray<At>,
  width: number,
  height: number,
  bits: number,
  skip: Jbig2Bitmap | undefined,
): Array<number> {
  const cx = newContexts(1 << 16);
  const planes: Array<Jbig2Bitmap> = new Array<Jbig2Bitmap>(bits);
  // MMR codes every plane into ONE stream, one after another; the arithmetic
  // form shares a decoder and its contexts the same way.
  let mmrOffset = 0;
  const decoder = mq ?? new MQDecoder(data);
  for (let j = bits - 1; j >= 0; j--) {
    if (mmr) {
      const packed = decodeCcitt(data.subarray(mmrOffset), {
        k: -1,
        columns: width,
        rows: height,
        byteAlign: false,
      });
      const plane = makeBitmap(width, height);
      if (packed) {
        const rowBytes = (width + 7) >> 3;
        for (let y = 0; y < height; y++) {
          for (let x = 0; x < width; x++) {
            plane.data[y * width + x] = (packed[y * rowBytes + (x >> 3)]! >> (7 - (x & 7))) & 1;
          }
        }
        mmrOffset += packed.length;
      }
      planes[j] = plane;
    } else {
      planes[j] = decodeGenericRegion(decoder, cx, width, height, template, at, false, skip);
    }
  }
  // Undo the Gray code from the top plane down, accumulating the value.
  const values = new Array<number>(width * height).fill(0);
  for (let i = 0; i < width * height; i++) {
    let bit = planes[bits - 1]?.data[i] ?? 0;
    let v = bit;
    for (let j = bits - 2; j >= 0; j--) {
      bit ^= planes[j]?.data[i] ?? 0;
      v = v * 2 + bit;
    }
    values[i] = v;
  }
  return values;
}

// ------------------------------------------------------------------ segments

/** §7.2 — one segment's header, and where its data lies in the stream. */
interface Segment {
  readonly number: number;
  readonly type: number;
  readonly referred: ReadonlyArray<number>;
  readonly page: number;
  readonly start: number;
  readonly end: number;
}

class Reader {
  pos = 0;
  constructor(readonly d: Uint8Array) {}
  u8(): number {
    return this.d[this.pos++] ?? 0;
  }
  u16(): number {
    return (this.u8() << 8) | this.u8();
  }
  u32(): number {
    return ((this.u8() << 24) | (this.u16() << 8) | this.u8()) >>> 0;
  }
  i8(): number {
    const v = this.u8();
    return v > 127 ? v - 256 : v;
  }
  get done(): boolean {
    return this.pos >= this.d.length;
  }
}

/** §7.2 — walk the segment headers of an embedded (PDF) JBIG2 stream. */
function parseSegments(data: Uint8Array): Array<Segment> {
  const r = new Reader(data);
  const out: Array<Segment> = [];
  while (r.pos + 11 <= data.length) {
    const number = r.u32();
    const flags = r.u8();
    const type = flags & 0x3f;
    const pageSize = (flags & 0x40) !== 0 ? 4 : 1;
    // §7.2.4 — a count of 7 in the top three bits escapes to a long form.
    const rt = r.u8();
    let count = rt >> 5;
    if (count === 7) {
      r.pos--;
      count = r.u32() & 0x1fffffff;
      r.pos += Math.ceil((count + 1) / 8); // retain-bits
    }
    // §7.2.5 — a referred-to number is as wide as it needs to be for THIS one.
    const refSize = number <= 256 ? 1 : number <= 65536 ? 2 : 4;
    const referred: Array<number> = [];
    for (let i = 0; i < count; i++) {
      referred.push(refSize === 1 ? r.u8() : refSize === 2 ? r.u16() : r.u32());
    }
    const page = pageSize === 1 ? r.u8() : r.u32();
    const length = r.u32();
    const start = r.pos;
    // §7.2.7 — 0xFFFFFFFF is "unknown", legal only for an immediate generic
    // region, whose data then runs to the end of the stream.
    const end = length === 0xffffffff ? data.length : Math.min(data.length, start + length);
    out.push({ number, type, referred, page, start, end });
    if (end <= start && length !== 0) break;
    r.pos = end;
    if (length === 0xffffffff) break;
  }
  return out;
}

/** §7.4.1 — the rectangle a region segment paints into, and how it combines. */
interface RegionInfo {
  readonly width: number;
  readonly height: number;
  readonly x: number;
  readonly y: number;
  readonly combOp: number;
}

function readRegionInfo(r: Reader): RegionInfo {
  const width = r.u32();
  const height = r.u32();
  const x = r.u32();
  const y = r.u32();
  const combOp = r.u8() & 7;
  return { width, height, x, y, combOp };
}

/** §6.2.5.1 / §8.2 — paint `src` onto `dst` at (x, y) under a composition op. */
function compose(dst: Jbig2Bitmap, src: Jbig2Bitmap, x0: number, y0: number, op: number): void {
  for (let y = 0; y < src.height; y++) {
    const dy = y0 + y;
    if (dy < 0 || dy >= dst.height) continue;
    for (let x = 0; x < src.width; x++) {
      const dx = x0 + x;
      if (dx < 0 || dx >= dst.width) continue;
      const s = src.data[y * src.width + x] ?? 0;
      const i = dy * dst.width + dx;
      const d = dst.data[i] ?? 0;
      dst.data[i] =
        op === 0 ? d | s : op === 1 ? d & s : op === 2 ? d ^ s : op === 3 ? d ^ s ^ 1 : s;
    }
  }
}

/** §7.4.8 — the page's own size, default pixel value and default combination op. */
interface PageInfo {
  width: number;
  height: number;
  defPixel: number;
  defCombOp: number;
  striped: boolean;
}

/**
 * Decode an embedded JBIG2 image.
 *
 * @param data    The `/JBIG2Decode` stream's own segments.
 * @param globals The `/JBIG2Globals` stream's segments, when the image names one.
 * @param width   The image's `/Width`, which the page info may not state.
 * @param height  Its `/Height`.
 * @returns A packed 1-bit-per-pixel bitmap (`rowBytes × height`, bit 1 = black,
 *   MSB first) — the same shape {@link decodeCcitt} returns — or `undefined`
 *   when nothing in the stream could be decoded.
 */
export function decodeJbig2(
  data: Uint8Array,
  globals: Uint8Array | undefined,
  width: number,
  height: number,
): Uint8Array | undefined {
  if (width <= 0 || height <= 0 || width > 1 << 16 || height > 1 << 16) return undefined;
  const page: PageInfo = {
    width,
    height,
    defPixel: 0,
    defCombOp: 0,
    striped: false,
  };
  let canvas = makeBitmap(width, height);
  // A counter rather than a flag: the assignment happens inside the closure
  // below, where narrowing cannot see it.
  const painted = { regions: 0 };
  // §7.4.1 — an INTERMEDIATE region (types 36, 40) is not painted: it is kept
  // for whichever later segment refers to it. bitmap-refine.pdf is exactly
  // that shape — an intermediate generic region, then a refinement of it — and
  // painting the intermediate onto the page put the unrefined bitmap there and
  // then refined the page against itself.
  const buffers = new Map<number, Jbig2Bitmap>();
  // §7.4.3 — what each symbol dictionary segment exported, for the text
  // regions that name it.
  const symbolsBySegment = new Map<number, ReadonlyArray<Jbig2Bitmap>>();
  // §7.4.4 — what each pattern dictionary segment holds, for the halftone
  // regions that name it.
  const patternsBySegment = new Map<number, ReadonlyArray<Jbig2Bitmap>>();

  const run = (bytes: Uint8Array): void => {
    for (const seg of parseSegments(bytes)) {
      const r = new Reader(bytes.subarray(seg.start, seg.end));
      try {
        if (seg.type === 48) {
          // §7.4.8 page information
          const pw = r.u32();
          const ph = r.u32();
          r.u32(); // x resolution
          r.u32(); // y resolution
          const flags = r.u8();
          page.defPixel = (flags >> 2) & 1;
          page.defCombOp = (flags >> 3) & 3;
          const striping = r.u16();
          page.striped = (striping & 0x8000) !== 0;
          // A page states its own size unless it is striped, where the height
          // is 0xFFFFFFFF until an end-of-stripe says otherwise. The image's
          // own /Width and /Height are the authority either way.
          if (pw > 0 && pw !== 0xffffffff && ph > 0 && ph !== 0xffffffff) {
            page.width = Math.min(pw, width);
            page.height = Math.min(ph, height);
          }
          canvas = makeBitmap(width, height, page.defPixel);
          continue;
        }
        // §7.4.4 pattern dictionary — the cells a halftone tiles out of.
        if (seg.type === 16) {
          const flags = r.u8();
          const mmr = (flags & 1) !== 0;
          const template = (flags >> 1) & 3;
          const pw = r.u8();
          const ph = r.u8();
          const grayMax = r.u32();
          if (pw <= 0 || ph <= 0 || grayMax > 10000) continue;
          patternsBySegment.set(
            seg.number,
            decodePatternDictionary(
              bytes.subarray(seg.start + r.pos, seg.end),
              mmr,
              template,
              pw,
              ph,
              grayMax,
            ),
          );
          continue;
        }
        // §7.4.5 halftone region — a picture rebuilt out of those cells, one
        // per grid position, chosen by a grey level.
        if (seg.type === 20 || seg.type === 22 || seg.type === 23) {
          const info = readRegionInfo(r);
          const flags = r.u8();
          const mmr = (flags & 1) !== 0;
          const template = (flags >> 1) & 3;
          const enableSkip = (flags & 8) !== 0;
          const combOp = (flags >> 4) & 7;
          const defPixel = (flags >> 7) & 1;
          const gw = r.u32();
          const gh = r.u32();
          const gx = r.u32() | 0;
          const gy = r.u32() | 0;
          const rx = r.u16();
          const ry = r.u16();
          const patterns = seg.referred.flatMap((n) => patternsBySegment.get(n) ?? []);
          const w = Math.min(info.width, 1 << 16);
          const h = Math.min(info.height, 1 << 16);
          if (patterns.length === 0 || w <= 0 || h <= 0) continue;
          if (gw <= 0 || gh <= 0 || gw * gh > 1 << 24) continue;
          const bmp = makeBitmap(w, h, defPixel);
          const pw = patterns[0]!.width;
          const ph = patterns[0]!.height;
          // §6.6.5.1 — a cell whose whole footprint falls outside the region is
          // not coded at all, which is what the skip bitmap says.
          let skip: Jbig2Bitmap | undefined;
          if (enableSkip) {
            skip = makeBitmap(gw, gh);
            for (let m = 0; m < gh; m++) {
              for (let n = 0; n < gw; n++) {
                const x = (gx + m * ry + n * rx) >> 8;
                const y = (gy + m * rx - n * ry) >> 8;
                if (x + pw <= 0 || x >= w || y + ph <= 0 || y >= h) skip.data[m * gw + n] = 1;
              }
            }
          }
          const bits = Math.max(1, Math.ceil(Math.log2(patterns.length)));
          const gray = decodeGrayScale(
            undefined,
            bytes.subarray(seg.start + r.pos, seg.end),
            mmr,
            template,
            [
              { x: template <= 1 ? 3 : 2, y: -1 },
              { x: -3, y: -1 },
              { x: 2, y: -2 },
              { x: -2, y: -2 },
            ],
            gw,
            gh,
            bits,
            skip,
          );
          for (let m = 0; m < gh; m++) {
            for (let n = 0; n < gw; n++) {
              if (skip?.data[m * gw + n] === 1) continue;
              const level = Math.min(gray[m * gw + n] ?? 0, patterns.length - 1);
              const x = (gx + m * ry + n * rx) >> 8;
              const y = (gy + m * rx - n * ry) >> 8;
              compose(bmp, patterns[level]!, x, y, combOp);
            }
          }
          if (seg.type === 20) buffers.set(seg.number, bmp);
          else {
            compose(canvas, bmp, info.x, info.y, info.combOp);
            painted.regions++;
          }
          continue;
        }
        // §7.4.3 symbol dictionary — the shapes, kept for the text regions
        // that refer to this segment.
        if (seg.type === 0) {
          const flags = r.u16();
          const huff = (flags & 1) !== 0;
          const refAgg = (flags & 2) !== 0;
          const template = (flags >> 10) & 3;
          const rTemplate = (flags >> 12) & 1;
          const ctxUsed = (flags & 0x0100) !== 0;
          const at: Array<At> = [];
          if (!huff) {
            const n = template === 0 ? 4 : 1;
            for (let i = 0; i < n; i++) at.push({ x: r.i8(), y: r.i8() });
          }
          const rAt: Array<At> = [];
          if (refAgg && rTemplate === 0) {
            for (let i = 0; i < 2; i++) rAt.push({ x: r.i8(), y: r.i8() });
          }
          const exportCount = r.u32();
          const newCount = r.u32();
          // Huffman-coded dictionaries are a later stage; without their tables
          // nothing here can be read, and guessing would draw rubbish.
          if (huff || newCount > 10000) continue;
          const inherited = seg.referred.flatMap((n) => symbolsBySegment.get(n) ?? []);
          void ctxUsed;
          symbolsBySegment.set(
            seg.number,
            decodeSymbolDictionary(
              new MQDecoder(bytes.subarray(seg.start + r.pos, seg.end)),
              inherited,
              newCount,
              exportCount,
              template,
              at.length > 0 ? at : (NOMINAL_AT[template] ?? []),
              refAgg,
              rTemplate,
              rAt.length > 0 ? rAt : NOMINAL_AT_REFINE,
            ),
          );
          continue;
        }
        // §7.4.4 text region — where each of those shapes goes.
        if (seg.type === 4 || seg.type === 6 || seg.type === 7) {
          const info = readRegionInfo(r);
          const flags = r.u16();
          const huff = (flags & 1) !== 0;
          const refine = (flags & 2) !== 0;
          const stripsLog = (flags >> 2) & 3;
          const refCorner = (flags >> 4) & 3;
          const transposed = (flags & 0x40) !== 0;
          const combOp = (flags >> 7) & 3;
          const defPixel = (flags >> 9) & 1;
          let dsOffset = (flags >> 10) & 0x1f;
          if (dsOffset > 15) dsOffset -= 32;
          const rTemplate = (flags >> 15) & 1;
          if (huff) continue; // Huffman-coded text regions are a later stage.
          const rAt: Array<At> = [];
          if (refine && rTemplate === 0) {
            for (let i = 0; i < 2; i++) rAt.push({ x: r.i8(), y: r.i8() });
          }
          const instances = r.u32();
          const symbols = seg.referred.flatMap((n) => symbolsBySegment.get(n) ?? []);
          if (symbols.length === 0 || instances > 100000) continue;
          const w = Math.min(info.width, 1 << 16);
          const h = Math.min(info.height, 1 << 16);
          if (w <= 0 || h <= 0) continue;
          const bmp = decodeTextRegion(
            new MQDecoder(bytes.subarray(seg.start + r.pos, seg.end)),
            symbols,
            {
              width: w,
              height: h,
              instances,
              stripT: 0,
              refCorner,
              transposed,
              combOp,
              defPixel,
              dsOffset,
              refine,
              rTemplate,
              rAt: rAt.length > 0 ? rAt : NOMINAL_AT_REFINE,
            },
            stripsLog,
          );
          if (seg.type === 4) buffers.set(seg.number, bmp);
          else {
            compose(canvas, bmp, info.x, info.y, info.combOp);
            painted.regions++;
          }
          continue;
        }
        // Immediate generic region (38, 39) and its intermediate form (36).
        if (seg.type === 36 || seg.type === 38 || seg.type === 39) {
          const info = readRegionInfo(r);
          const flags = r.u8();
          const mmr = (flags & 1) !== 0;
          const template = (flags >> 1) & 3;
          const tpgdon = (flags & 8) !== 0;
          const at: Array<At> = [];
          if (!mmr) {
            const n = template === 0 ? 4 : 1;
            for (let i = 0; i < n; i++) at.push({ x: r.i8(), y: r.i8() });
          }
          const body = bytes.subarray(seg.start + r.pos, seg.end);
          const w = Math.min(info.width, 1 << 16);
          const h = Math.min(info.height === 0xffffffff ? height : info.height, 1 << 16);
          if (w <= 0 || h <= 0) continue;
          let bmp: Jbig2Bitmap;
          if (mmr) {
            // §6.2.6 — MMR is T.6, which the fax decoder already speaks.
            const packed = decodeCcitt(body, { k: -1, columns: w, rows: h, byteAlign: false });
            if (!packed) continue;
            bmp = makeBitmap(w, h);
            const rowBytes = (w + 7) >> 3;
            for (let y = 0; y < h; y++) {
              for (let x = 0; x < w; x++) {
                bmp.data[y * w + x] = (packed[y * rowBytes + (x >> 3)]! >> (7 - (x & 7))) & 1;
              }
            }
          } else {
            bmp = decodeGenericRegion(
              new MQDecoder(body),
              newContexts(1 << 16),
              w,
              h,
              template,
              at.length > 0 ? at : (NOMINAL_AT[template] ?? []),
              tpgdon,
            );
          }
          if (seg.type === 36) buffers.set(seg.number, bmp);
          else {
            compose(canvas, bmp, info.x, info.y, info.combOp);
            painted.regions++;
          }
          continue;
        }
        // Immediate refinement region (42, 43) and its intermediate form (40).
        if (seg.type === 40 || seg.type === 42 || seg.type === 43) {
          const info = readRegionInfo(r);
          const flags = r.u8();
          const template = flags & 1;
          const tpgron = (flags & 2) !== 0;
          const at: Array<At> = [];
          if (template === 0) for (let i = 0; i < 2; i++) at.push({ x: r.i8(), y: r.i8() });
          const w = Math.min(info.width, 1 << 16);
          const h = Math.min(info.height, 1 << 16);
          if (w <= 0 || h <= 0) continue;
          // §6.3.2 — the reference is the buffer of the intermediate region
          // this segment refers to. With none named it is the page area the
          // refinement stands on, which is how a page is refined in place:
          // bitmap-refine.pdf takes the first path and bitmap-refine-page.pdf
          // the second, and they draw the same picture.
          let ref = seg.referred
            .map((n) => buffers.get(n))
            .find((b): b is Jbig2Bitmap => b !== undefined);
          if (!ref) {
            const crop = makeBitmap(w, h);
            for (let y = 0; y < h; y++) {
              for (let x = 0; x < w; x++) {
                const sx = info.x + x;
                const sy = info.y + y;
                crop.data[y * w + x] =
                  sx < canvas.width && sy < canvas.height
                    ? (canvas.data[sy * canvas.width + sx] ?? 0)
                    : 0;
              }
            }
            ref = crop;
          }
          const bmp = decodeRefinement(
            new MQDecoder(bytes.subarray(seg.start + r.pos, seg.end)),
            newContexts(1 << 13),
            w,
            h,
            template,
            ref,
            0,
            0,
            at.length > 0 ? at : NOMINAL_AT_REFINE,
            tpgron,
          );
          if (seg.type === 40) buffers.set(seg.number, bmp);
          else {
            compose(canvas, bmp, info.x, info.y, info.combOp);
            painted.regions++;
          }
          continue;
        }
      } catch {
        // A malformed segment stops that segment, not the page: JBIG2 streams
        // in the wild are truncated often, and what decoded before the fault is
        // still the picture.
        continue;
      }
    }
  };

  if (globals) run(globals);
  run(data);
  if (painted.regions === 0) return undefined;

  const rowBytes = (width + 7) >> 3;
  const packed = new Uint8Array(rowBytes * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (canvas.data[y * width + x] === 1) {
        packed[y * rowBytes + (x >> 3)]! |= 0x80 >> (x & 7);
      }
    }
  }
  return packed;
}
