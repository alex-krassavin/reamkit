// Baseline JPEG decoder (ITU-T T.81) — enough to turn a DCTDecode stream back
// into pixels.
//
// Ream embeds a JPEG in its output verbatim, which is why nothing here was
// needed until a picture had to be CHANGED rather than carried: an `/SMask`
// (ISO 32000-1 §8.9.5.4) is a second image supplying the alpha, and folding it
// in means having the samples of both. 22060_A1_01_Plans.pdf draws its floor
// plans that way — the JPEG is the wash and the mask is the line work — and
// unmasked it renders as a dark rectangle.
//
// Scope is the sequential baseline (SOF0/SOF1), 8-bit, one to three components,
// which is what a PDF producer writes. Progressive (SOF2), arithmetic coding
// and the lossless modes decline, and the caller carries the JPEG through
// untouched exactly as before.

/** A decoded JPEG: 8-bit interleaved samples, one or three components. */
export interface DecodedJpeg {
  readonly width: number;
  readonly height: number;
  /** 1 = grayscale, 3 = RGB (already converted from YCbCr where it applies). */
  readonly components: 1 | 3;
  /** `width * height * components` bytes, row-major. */
  readonly samples: Uint8Array;
}

/** §A.3.6 — the zig-zag order the coefficients of a block are coded in. */
const ZIGZAG = new Uint8Array([
  0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5, 12, 19, 26, 33, 40, 48, 41, 34, 27, 20,
  13, 6, 7, 14, 21, 28, 35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51, 58, 59, 52,
  45, 38, 31, 39, 46, 53, 60, 61, 54, 47, 55, 62, 63,
]);

/** A Huffman table as the flat code→value walk §F.2.2.3 describes. */
interface HuffTable {
  /** `maxcode[l]` is the largest code of length `l`, or −1 when there is none. */
  readonly maxcode: Int32Array;
  readonly valptr: Int32Array;
  readonly mincode: Int32Array;
  readonly values: Uint8Array;
}

interface Component {
  readonly id: number;
  /** §A.1.1 horizontal and vertical sampling factors. */
  readonly h: number;
  readonly v: number;
  readonly tq: number;
  dc: number; // DC table selector, set by SOS
  ac: number; // AC table selector, set by SOS
  pred: number; // §F.1.2.1 the DC predictor, reset at every restart interval
  blocksPerLine: number;
  blocksPerCol: number;
  plane: Uint8ClampedArray;
  stride: number;
}

const MAX_PIXELS = 50_000_000; // guard: a decoded plane is one byte per sample

/**
 * Decode a baseline JPEG to interleaved 8-bit samples.
 *
 * @param bytes The JPEG stream, starting at its SOI marker.
 * @returns The decoded image, or `undefined` when the stream is not a baseline
 *          JPEG this decoder handles (progressive, arithmetic, 12-bit, CMYK) or
 *          is malformed — the caller then carries the original bytes through.
 */
export function decodeJpeg(bytes: Uint8Array): DecodedJpeg | undefined {
  try {
    return decode(bytes);
  } catch {
    // A truncated or malformed stream is not a crash: the picture simply
    // travels as it arrived.
    return undefined;
  }
}

function decode(bytes: Uint8Array): DecodedJpeg | undefined {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;

  const quant: Array<Int32Array | undefined> = [];
  const dcTables: Array<HuffTable | undefined> = [];
  const acTables: Array<HuffTable | undefined> = [];
  let frame:
    | { width: number; height: number; comps: Array<Component>; maxH: number; maxV: number }
    | undefined;
  let restartInterval = 0;
  // §A.11 Adobe APP14: transform 0 means the three components are RGB already.
  let adobeTransform: number | undefined;

  let p = 2;
  while (p < bytes.length) {
    if (bytes[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = bytes[p + 1]!;
    p += 2;
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (marker === 0xd9) break; // EOI
    const length = be16(bytes, p);
    const segEnd = p + length;

    switch (marker) {
      case 0xdb: // DQT §B.2.4.1
        readQuantTables(bytes, p + 2, segEnd, quant);
        break;
      case 0xc4: // DHT §B.2.4.2
        readHuffTables(bytes, p + 2, segEnd, dcTables, acTables);
        break;
      case 0xdd: // DRI §B.2.4.4
        restartInterval = be16(bytes, p + 2);
        break;
      case 0xee: // APP14
        if (length >= 14 && str(bytes, p + 2, 5) === 'Adobe') adobeTransform = bytes[p + 13];
        break;
      case 0xc0: // SOF0 baseline
      case 0xc1: // SOF1 extended sequential, Huffman — same bitstream
        frame = readFrame(bytes, p + 2);
        if (!frame) return undefined;
        break;
      case 0xc2: // SOF2 progressive
      case 0xc3:
      case 0xc5:
      case 0xc6:
      case 0xc7:
      case 0xc9:
      case 0xca:
      case 0xcb:
      case 0xcd:
      case 0xce:
      case 0xcf:
        return undefined; // progressive / lossless / arithmetic — not this decoder
      case 0xda: {
        // SOS §B.2.3 — the scan header, then the entropy-coded data.
        if (!frame) return undefined;
        const scanComps = readScanHeader(bytes, p + 2, frame.comps);
        if (!scanComps) return undefined;
        p = decodeScan(bytes, segEnd, frame, scanComps, dcTables, acTables, quant, restartInterval);
        continue;
      }
      default:
        break;
    }
    p = segEnd;
  }

  if (!frame) return undefined;
  return assemble(frame, adobeTransform);
}

// --- headers ----------------------------------------------------------------

function readFrame(
  bytes: Uint8Array,
  p: number,
):
  | { width: number; height: number; comps: Array<Component>; maxH: number; maxV: number }
  | undefined {
  if (bytes[p] !== 8) return undefined; // sample precision; only 8-bit here
  const height = be16(bytes, p + 1);
  const width = be16(bytes, p + 3);
  const count = bytes[p + 5]!;
  if (width <= 0 || height <= 0 || width * height > MAX_PIXELS) return undefined;
  if (count !== 1 && count !== 3) return undefined; // grayscale or YCbCr/RGB only

  const comps: Array<Component> = [];
  let maxH = 1;
  let maxV = 1;
  for (let i = 0; i < count; i++) {
    const at = p + 6 + i * 3;
    const h = bytes[at + 1]! >> 4;
    const v = bytes[at + 1]! & 15;
    if (h < 1 || h > 4 || v < 1 || v > 4) return undefined;
    maxH = Math.max(maxH, h);
    maxV = Math.max(maxV, v);
    comps.push({
      id: bytes[at]!,
      h,
      v,
      tq: bytes[at + 2]!,
      dc: 0,
      ac: 0,
      pred: 0,
      blocksPerLine: 0,
      blocksPerCol: 0,
      plane: new Uint8ClampedArray(0),
      stride: 0,
    });
  }

  // §A.2.3 — each component is coded in whole MCUs, so its plane is rounded up
  // to the MCU grid and cropped only at the end.
  const mcusPerLine = Math.ceil(width / (8 * maxH));
  const mcusPerCol = Math.ceil(height / (8 * maxV));
  for (const c of comps) {
    c.blocksPerLine = mcusPerLine * c.h;
    c.blocksPerCol = mcusPerCol * c.v;
    c.stride = c.blocksPerLine * 8;
    c.plane = new Uint8ClampedArray(c.stride * c.blocksPerCol * 8);
  }
  return { width, height, comps, maxH, maxV };
}

function readScanHeader(
  bytes: Uint8Array,
  p: number,
  comps: ReadonlyArray<Component>,
): Array<Component> | undefined {
  const count = bytes[p]!;
  const out: Array<Component> = [];
  for (let i = 0; i < count; i++) {
    const id = bytes[p + 1 + i * 2]!;
    const tables = bytes[p + 2 + i * 2]!;
    const comp = comps.find((c) => c.id === id);
    if (!comp) return undefined;
    comp.dc = tables >> 4;
    comp.ac = tables & 15;
    out.push(comp);
  }
  return out;
}

function readQuantTables(
  bytes: Uint8Array,
  p: number,
  end: number,
  into: Array<Int32Array | undefined>,
): void {
  while (p < end) {
    const pq = bytes[p]! >> 4;
    const tq = bytes[p]! & 15;
    p++;
    const table = new Int32Array(64);
    for (let i = 0; i < 64; i++) {
      table[ZIGZAG[i]!] = pq === 0 ? bytes[p + i]! : be16(bytes, p + i * 2);
    }
    p += pq === 0 ? 64 : 128;
    into[tq] = table;
  }
}

function readHuffTables(
  bytes: Uint8Array,
  p: number,
  end: number,
  dc: Array<HuffTable | undefined>,
  ac: Array<HuffTable | undefined>,
): void {
  while (p < end) {
    const tc = bytes[p]! >> 4;
    const th = bytes[p]! & 15;
    p++;
    const counts = bytes.subarray(p, p + 16);
    p += 16;
    let total = 0;
    for (const n of counts) total += n;
    const values = bytes.slice(p, p + total);
    p += total;
    const table = buildHuffTable(counts, values);
    if (tc === 0) dc[th] = table;
    else ac[th] = table;
  }
}

/** §F.2.2.3 — the decoding tables, built from BITS and HUFFVAL. */
function buildHuffTable(counts: Uint8Array, values: Uint8Array): HuffTable {
  const maxcode = new Int32Array(18).fill(-1);
  const mincode = new Int32Array(18);
  const valptr = new Int32Array(18);
  let code = 0;
  let k = 0;
  for (let l = 1; l <= 16; l++) {
    const n = counts[l - 1]!;
    if (n === 0) {
      maxcode[l] = -1;
    } else {
      valptr[l] = k;
      mincode[l] = code;
      code += n;
      k += n;
      maxcode[l] = code - 1;
    }
    code <<= 1;
  }
  return { maxcode, mincode, valptr, values };
}

// --- entropy-coded data ------------------------------------------------------

/**
 * Decode one scan's entropy-coded segment into the components' planes.
 *
 * @returns The offset of the marker that ends the scan.
 */
function decodeScan(
  bytes: Uint8Array,
  start: number,
  frame: { width: number; height: number; comps: Array<Component>; maxH: number; maxV: number },
  scan: ReadonlyArray<Component>,
  dcTables: ReadonlyArray<HuffTable | undefined>,
  acTables: ReadonlyArray<HuffTable | undefined>,
  quant: ReadonlyArray<Int32Array | undefined>,
  restartInterval: number,
): number {
  let at = start;
  let bitBuf = 0;
  let bitCount = 0;

  const nextBit = (): number => {
    if (bitCount === 0) {
      let b = bytes[at++] ?? 0;
      if (b === 0xff) {
        const next = bytes[at] ?? 0;
        // §B.1.1.5 — a 0xFF in the data is stuffed with a following 0x00.
        if (next === 0x00) at++;
        else {
          // A marker: the scan is over, feed zeros to finish the block.
          at--;
          b = 0;
        }
      }
      bitBuf = b;
      bitCount = 8;
    }
    bitCount--;
    return (bitBuf >> bitCount) & 1;
  };

  const decodeHuff = (table: HuffTable | undefined): number => {
    if (!table) return 0;
    let code = nextBit();
    let l = 1;
    while (l <= 16 && (table.maxcode[l] === -1 || code > table.maxcode[l]!)) {
      code = (code << 1) | nextBit();
      l++;
    }
    if (l > 16) return 0;
    return table.values[table.valptr[l]! + code - table.mincode[l]!] ?? 0;
  };

  // §F.2.2.1 — a magnitude of `s` bits, sign-extended.
  const receiveExtend = (s: number): number => {
    if (s === 0) return 0;
    let v = 0;
    for (let i = 0; i < s; i++) v = (v << 1) | nextBit();
    return v < 1 << (s - 1) ? v - (1 << s) + 1 : v;
  };

  const block = new Int32Array(64);
  const decodeBlock = (comp: Component, row: number, col: number): void => {
    block.fill(0);
    const q = quant[comp.tq];
    const t = decodeHuff(dcTables[comp.dc]);
    comp.pred += receiveExtend(t);
    block[0] = comp.pred * (q?.[0] ?? 1);
    let i = 1;
    while (i < 64) {
      const rs = decodeHuff(acTables[comp.ac]);
      const run = rs >> 4;
      const size = rs & 15;
      if (size === 0) {
        if (run !== 15) break; // EOB
        i += 16;
        continue;
      }
      i += run;
      if (i > 63) break;
      const z = ZIGZAG[i]!;
      block[z] = receiveExtend(size) * (q?.[z] ?? 1);
      i++;
    }
    idctBlock(block, comp.plane, row * 8 * comp.stride + col * 8, comp.stride);
  };

  const mcusPerLine = Math.ceil(frame.width / (8 * frame.maxH));
  const mcusPerCol = Math.ceil(frame.height / (8 * frame.maxV));
  const single = scan.length === 1;
  const total = single
    ? Math.ceil(scan[0]!.blocksPerLine * scan[0]!.blocksPerCol)
    : mcusPerLine * mcusPerCol;
  const perLine = single ? scan[0]!.blocksPerLine : mcusPerLine;

  let untilRestart = restartInterval || total;
  for (let n = 0; n < total; n++) {
    if (single) {
      const c = scan[0]!;
      decodeBlock(c, Math.floor(n / perLine), n % perLine);
    } else {
      const mcuRow = Math.floor(n / mcusPerLine);
      const mcuCol = n % mcusPerLine;
      for (const c of scan) {
        for (let v = 0; v < c.v; v++) {
          for (let h = 0; h < c.h; h++) {
            decodeBlock(c, mcuRow * c.v + v, mcuCol * c.h + h);
          }
        }
      }
    }
    if (--untilRestart === 0 && n + 1 < total) {
      // §B.2.1 — RSTn: the bit reader realigns and every predictor resets.
      bitCount = 0;
      while (
        at + 1 < bytes.length &&
        !(bytes[at] === 0xff && bytes[at + 1]! >= 0xd0 && bytes[at + 1]! <= 0xd7)
      ) {
        at++;
      }
      at += 2;
      for (const c of scan) c.pred = 0;
      untilRestart = restartInterval;
    }
  }

  // Walk to the marker that ends the scan.
  while (at + 1 < bytes.length && !(bytes[at] === 0xff && bytes[at + 1] !== 0x00)) at++;
  return at;
}

// --- IDCT --------------------------------------------------------------------

/** §A.3.3 — cos((2x+1)uπ/16) · C(u)/2, precomputed. */
const COS = (() => {
  const t = new Float32Array(64);
  for (let u = 0; u < 8; u++) {
    const c = u === 0 ? Math.SQRT1_2 : 1;
    for (let x = 0; x < 8; x++) t[u * 8 + x] = (c * Math.cos(((2 * x + 1) * u * Math.PI) / 16)) / 2;
  }
  return t;
})();

/** The separable 2-D inverse DCT of one block, levelshifted into `out`. */
function idctBlock(
  block: Int32Array,
  out: Uint8ClampedArray,
  offset: number,
  stride: number,
): void {
  const rows = new Float32Array(64);
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      let sum = 0;
      for (let u = 0; u < 8; u++) sum += COS[u * 8 + x]! * block[y * 8 + u]!;
      rows[y * 8 + x] = sum;
    }
  }
  for (let x = 0; x < 8; x++) {
    for (let y = 0; y < 8; y++) {
      let sum = 0;
      for (let v = 0; v < 8; v++) sum += COS[v * 8 + y]! * rows[v * 8 + x]!;
      // §A.3.1 — the level shift back to unsigned.
      out[offset + y * stride + x] = sum + 128;
    }
  }
}

// --- assembly ----------------------------------------------------------------

/** Upsample every component to full size and convert to gray or RGB. */
function assemble(
  frame: { width: number; height: number; comps: Array<Component>; maxH: number; maxV: number },
  adobeTransform: number | undefined,
): DecodedJpeg {
  const { width, height, comps, maxH, maxV } = frame;
  const n = comps.length;
  const samples = new Uint8Array(width * height * n);
  for (let ci = 0; ci < n; ci++) {
    const c = comps[ci]!;
    const sx = c.h / maxH;
    const sy = c.v / maxV;
    for (let y = 0; y < height; y++) {
      const srcRow = Math.min(c.blocksPerCol * 8 - 1, (y * sy) | 0) * c.stride;
      for (let x = 0; x < width; x++) {
        samples[(y * width + x) * n + ci] = c.plane[srcRow + Math.min(c.stride - 1, (x * sx) | 0)]!;
      }
    }
  }
  if (n === 1) return { width, height, components: 1, samples };
  // §A.11 — three components are YCbCr unless Adobe says transform 0 (RGB).
  if (adobeTransform !== 0) ycbcrToRgb(samples);
  return { width, height, components: 3, samples };
}

/** ITU-T T.871 §7 — the JFIF YCbCr → RGB conversion, in place. */
function ycbcrToRgb(s: Uint8Array): void {
  for (let i = 0; i < s.length; i += 3) {
    const y = s[i]!;
    const cb = s[i + 1]! - 128;
    const cr = s[i + 2]! - 128;
    s[i] = clamp(y + 1.402 * cr);
    s[i + 1] = clamp(y - 0.344136 * cb - 0.714136 * cr);
    s[i + 2] = clamp(y + 1.772 * cb);
  }
}

const clamp = (v: number): number => (v < 0 ? 0 : v > 255 ? 255 : v | 0);

const be16 = (b: Uint8Array, p: number): number => ((b[p] ?? 0) << 8) | (b[p + 1] ?? 0);

const str = (b: Uint8Array, p: number, n: number): string =>
  String.fromCharCode(...b.subarray(p, p + n));
