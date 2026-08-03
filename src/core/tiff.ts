// Baseline TIFF 6.0 — decode a stripped or tiled raster into 8-bit samples.
//
// A .docx may embed a TIFF as plainly as a PNG (fdo77476.docx, tdf115094.docx),
// and no PDF reader decodes one: /TIFFDecode does not exist, so the picture has
// to arrive as samples. This is the baseline reader that produces them.
//
// What it reads: both byte orders; PhotometricInterpretation WhiteIsZero,
// BlackIsZero, RGB, palette and CMYK (§8); BitsPerSample 1/2/4/8/16 (16
// truncated to its high byte, as the PNG path does); SamplesPerPixel 1/3/4 with
// an ExtraSamples alpha channel split off as a soft mask; Compression none,
// LZW (§13), PackBits (§9) and Deflate (Adobe TIFF Technical Note 2);
// Predictor 1 and 2 (§14 horizontal differencing); strips and tiles (§15);
// PlanarConfiguration chunky. Multi-page files give their FIRST image, the page
// a document that embeds one means.
//
// Everything else — planar separation, JPEG-in-TIFF, fax compressions, floating
// point samples — throws, so the caller can record the loss rather than draw a
// wrong picture.

import { unzlibSync } from 'fflate';

import { lzwDecodeMsb } from '@/core/lzw';

/** One decoded TIFF image: 8-bit chunky samples plus an optional alpha plane. */
export interface DecodedTiff {
  readonly width: number;
  readonly height: number;
  /** Samples per pixel in {@link data} — 1 (grey), 3 (RGB) or 4 (CMYK). */
  readonly channels: 1 | 3 | 4;
  /** Row-major 8-bit samples, `width * height * channels` long. */
  readonly data: Uint8Array;
  /** One byte per pixel when the file carries an alpha channel. */
  readonly alpha?: Uint8Array;
}

const MAX_PIXELS = 40_000_000; // DoS guard (~40 MP), as the PDF image path uses

// §2 — the tag numbers this reader acts on.
const TAG = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
  fillOrder: 266,
  stripOffsets: 273,
  samplesPerPixel: 277,
  rowsPerStrip: 278,
  stripByteCounts: 279,
  planarConfig: 284,
  predictor: 317,
  colorMap: 320,
  tileWidth: 322,
  tileLength: 323,
  tileOffsets: 324,
  tileByteCounts: 325,
  extraSamples: 338,
  sampleFormat: 339,
} as const;

/** Whether the bytes open with a TIFF header (§2: `II*\0` or `MM\0*`). */
export function isTiff(bytes: Uint8Array): boolean {
  if (bytes.length < 8) return false;
  const le = bytes[0] === 0x49 && bytes[1] === 0x49;
  const be = bytes[0] === 0x4d && bytes[1] === 0x4d;
  if (!le && !be) return false;
  const magic = le ? bytes[2]! | (bytes[3]! << 8) : (bytes[2]! << 8) | bytes[3]!;
  return magic === 42;
}

/**
 * Decode the first image of a baseline TIFF into 8-bit chunky samples.
 *
 * @throws Error when the file is malformed or uses a feature outside baseline
 * TIFF (planar separation, JPEG/fax compression, floating-point samples).
 */
export function decodeTiff(bytes: Uint8Array): DecodedTiff {
  if (!isTiff(bytes)) throw new Error('TIFF: not a TIFF file');
  const le = bytes[0] === 0x49;
  const r = new Reader(bytes, le);
  const ifd = r.readIfd(r.u32(4));

  const width = num(ifd, TAG.width, 0);
  const height = num(ifd, TAG.height, 0);
  if (width <= 0 || height <= 0) throw new Error('TIFF: no image dimensions');
  if (width * height > MAX_PIXELS) throw new Error('TIFF: image too large to decode');

  const samples = num(ifd, TAG.samplesPerPixel, 1);
  const bitsList = list(ifd, TAG.bitsPerSample, [1]);
  const bits = bitsList[0] ?? 1;
  if (bitsList.some((b) => b !== bits)) throw new Error('TIFF: mixed sample depths');
  if (![1, 2, 4, 8, 16].includes(bits)) throw new Error(`TIFF: ${String(bits)}-bit samples`);
  if (num(ifd, TAG.planarConfig, 1) !== 1) throw new Error('TIFF: planar separation');
  const format = list(ifd, TAG.sampleFormat, [1]);
  if (format.some((f) => f !== 1 && f !== 2)) throw new Error('TIFF: non-integer samples');

  const photometric = num(ifd, TAG.photometric, samples >= 3 ? 2 : 1);
  const compression = num(ifd, TAG.compression, 1);
  const predictor = num(ifd, TAG.predictor, 1);
  if (predictor !== 1 && predictor !== 2) throw new Error('TIFF: floating-point predictor');

  // §15 — the picture arrives in strips (rows) or tiles (rectangles). Both are
  // independently compressed blocks laid over the image in reading order, so
  // one loop covers them once each block knows its own box.
  const blocks = blockLayout(ifd, width, height);
  const bytesPerRow = Math.ceil((blocks.blockWidth * samples * bits) / 8);
  const raw = new Uint8Array(width * height * samples * (bits === 16 ? 2 : 1));
  const rawRowBytes = width * samples * (bits === 16 ? 2 : 1);

  const blockBytes = bytesPerRow * blocks.blockHeight;
  for (const b of blocks.blocks) {
    // A file may leave the byte counts out when nothing compresses them (§15
    // makes them required, and producers still forget): a whole block is then
    // exactly what a block holds.
    const count = b.byteCount > 0 ? b.byteCount : blockBytes;
    const block = decodeBlock(bytes.subarray(b.offset, b.offset + count), compression, blockBytes);
    const rows = unpredict(block, predictor, bytesPerRow, samples, bits, le);
    // Place the block's rows into the full image, cropping the padding a block
    // on the right or bottom edge carries (§15: blocks are whole even when the
    // image is not).
    for (let y = 0; y < blocks.blockHeight; y++) {
      const destY = b.y + y;
      if (destY >= height) break;
      const src = y * bytesPerRow;
      if (src >= rows.length) break;
      const row = expandRow(
        rows.subarray(src, src + bytesPerRow),
        bits,
        blocks.blockWidth * samples,
      );
      const keep = Math.min(row.length, (width - b.x) * samples * (bits === 16 ? 2 : 1));
      raw.set(row.subarray(0, keep), destY * rawRowBytes + b.x * samples * (bits === 16 ? 2 : 1));
    }
  }

  const eight = bits === 16 ? highBytes(raw, le) : raw;
  return toChannels(eight, {
    width,
    height,
    samples,
    bits,
    photometric,
    palette: list(ifd, TAG.colorMap, []),
    extra: list(ifd, TAG.extraSamples, []),
  });
}

// ——— the IFD ———————————————————————————————————————————————————————————

interface Entry {
  readonly type: number;
  readonly count: number;
  /** Every value read as a number: TIFF's integer types are all this reader needs. */
  readonly values: ReadonlyArray<number>;
}

type Ifd = ReadonlyMap<number, Entry>;

const TYPE_SIZE: ReadonlyMap<number, number> = new Map([
  [1, 1], // BYTE
  [2, 1], // ASCII
  [3, 2], // SHORT
  [4, 4], // LONG
  [5, 8], // RATIONAL
  [6, 1], // SBYTE
  [7, 1], // UNDEFINED
  [8, 2], // SSHORT
  [9, 4], // SLONG
  [10, 8], // SRATIONAL
  [11, 4], // FLOAT
  [12, 8], // DOUBLE
]);

class Reader {
  constructor(
    private readonly b: Uint8Array,
    private readonly le: boolean,
  ) {}

  u16(o: number): number {
    if (o + 2 > this.b.length) throw new Error('TIFF: truncated');
    return this.le ? this.b[o]! | (this.b[o + 1]! << 8) : (this.b[o]! << 8) | this.b[o + 1]!;
  }

  u32(o: number): number {
    if (o + 4 > this.b.length) throw new Error('TIFF: truncated');
    return this.le
      ? (this.b[o]! | (this.b[o + 1]! << 8) | (this.b[o + 2]! << 16)) + this.b[o + 3]! * 0x1000000
      : (this.b[o + 3]! | (this.b[o + 2]! << 8) | (this.b[o + 1]! << 16)) + this.b[o]! * 0x1000000;
  }

  /** §2 — the directory at `offset`: a count, then 12-byte entries. */
  readIfd(offset: number): Ifd {
    const count = this.u16(offset);
    const out = new Map<number, Entry>();
    for (let i = 0; i < count; i++) {
      const e = offset + 2 + i * 12;
      const tag = this.u16(e);
      const type = this.u16(e + 2);
      const n = this.u32(e + 4);
      const size = TYPE_SIZE.get(type) ?? 0;
      if (size === 0 || n > 1_000_000) continue; // a tag this reader cannot use
      // Values up to four bytes live in the entry itself; longer ones at an offset.
      const at = size * n <= 4 ? e + 8 : this.u32(e + 8);
      out.set(tag, { type, count: n, values: this.readValues(type, n, at) });
    }
    return out;
  }

  private readValues(type: number, count: number, at: number): Array<number> {
    const size = TYPE_SIZE.get(type) ?? 1;
    const out: Array<number> = [];
    for (let i = 0; i < count; i++) {
      const o = at + i * size;
      if (o + size > this.b.length) break;
      switch (type) {
        case 1:
        case 2:
        case 6:
        case 7:
          out.push(this.b[o]!);
          break;
        case 3:
          out.push(this.u16(o));
          break;
        case 8: {
          const v = this.u16(o);
          out.push(v > 0x7fff ? v - 0x10000 : v);
          break;
        }
        case 4:
          out.push(this.u32(o));
          break;
        case 9: {
          const v = this.u32(o);
          out.push(v > 0x7fffffff ? v - 0x100000000 : v);
          break;
        }
        case 5:
        case 10: {
          const d = this.u32(o + 4);
          out.push(d === 0 ? 0 : this.u32(o) / d);
          break;
        }
        default:
          out.push(0);
      }
    }
    return out;
  }
}

function num(ifd: Ifd, tag: number, fallback: number): number {
  return ifd.get(tag)?.values[0] ?? fallback;
}

function list(ifd: Ifd, tag: number, fallback: ReadonlyArray<number>): ReadonlyArray<number> {
  const e = ifd.get(tag);
  return e && e.values.length > 0 ? e.values : fallback;
}

// ——— strips and tiles ——————————————————————————————————————————————————

interface Block {
  readonly offset: number;
  readonly byteCount: number;
  /** Top-left of the block within the image, in pixels. */
  readonly x: number;
  readonly y: number;
}

function blockLayout(
  ifd: Ifd,
  width: number,
  height: number,
): { blocks: Array<Block>; blockWidth: number; blockHeight: number } {
  const tileW = num(ifd, TAG.tileWidth, 0);
  const tileH = num(ifd, TAG.tileLength, 0);
  if (tileW > 0 && tileH > 0) {
    const offsets = list(ifd, TAG.tileOffsets, []);
    const counts = list(ifd, TAG.tileByteCounts, []);
    const across = Math.ceil(width / tileW);
    const blocks = offsets.map((offset, i) => ({
      offset,
      byteCount: counts[i] ?? 0,
      x: (i % across) * tileW,
      y: Math.floor(i / across) * tileH,
    }));
    return { blocks, blockWidth: tileW, blockHeight: tileH };
  }
  const rows = num(ifd, TAG.rowsPerStrip, height) || height;
  const offsets = list(ifd, TAG.stripOffsets, []);
  const counts = list(ifd, TAG.stripByteCounts, []);
  if (offsets.length === 0) throw new Error('TIFF: no image data');
  const blocks = offsets.map((offset, i) => ({
    offset,
    byteCount: counts[i] ?? 0,
    x: 0,
    y: i * rows,
  }));
  return { blocks, blockWidth: width, blockHeight: rows };
}

// §7/§9/§13 — one block's bytes, decompressed. `expected` is what a whole block
// holds, which bounds LZW and sizes the PackBits output.
function decodeBlock(data: Uint8Array, compression: number, expected: number): Uint8Array {
  switch (compression) {
    case 1:
      return data;
    case 5:
      return lzwDecodeMsb(data, { limit: expected });
    case 8:
    case 32946:
      try {
        return unzlibSync(data);
      } catch {
        throw new Error('TIFF: corrupt Deflate strip');
      }
    case 32773:
      return packBits(data, expected);
    default:
      throw new Error(`TIFF: compression ${String(compression)}`);
  }
}

// §9 PackBits: a signed count byte — n ≥ 0 copies n+1 literal bytes, n < 0
// repeats the next byte 1−n times, −128 is a no-op.
function packBits(data: Uint8Array, expected: number): Uint8Array {
  const out = new Uint8Array(expected);
  let o = 0;
  for (let i = 0; i < data.length && o < expected; ) {
    const n = data[i++]!;
    if (n === 0x80) continue;
    if (n < 0x80) {
      const run = Math.min(n + 1, data.length - i, expected - o);
      out.set(data.subarray(i, i + run), o);
      i += n + 1;
      o += run;
    } else {
      const byte = data[i++] ?? 0;
      const run = Math.min(257 - n, expected - o);
      out.fill(byte, o, o + run);
      o += run;
    }
  }
  return out.subarray(0, o);
}

// §14 Predictor 2 — horizontal differencing, undone per row.
function unpredict(
  data: Uint8Array,
  predictor: number,
  rowBytes: number,
  samples: number,
  bits: number,
  le: boolean,
): Uint8Array {
  if (predictor !== 2 || rowBytes <= 0) return data;
  if (bits !== 8 && bits !== 16) return data; // §14 defines it for whole bytes
  const out = data.slice();
  const rows = Math.floor(out.length / rowBytes);
  if (bits === 8) {
    for (let r = 0; r < rows; r++) {
      const off = r * rowBytes;
      for (let i = samples; i < rowBytes; i++)
        out[off + i] = (out[off + i]! + out[off + i - samples]!) & 0xff;
    }
    return out;
  }
  const stride = samples * 2;
  const hi = le ? 1 : 0;
  const lo = le ? 0 : 1;
  for (let r = 0; r < rows; r++) {
    const off = r * rowBytes;
    for (let i = stride; i + 1 < rowBytes; i += 2) {
      const prev = (out[off + i - stride + hi]! << 8) | out[off + i - stride + lo]!;
      const v = (((out[off + i + hi]! << 8) | out[off + i + lo]!) + prev) & 0xffff;
      out[off + i + hi] = v >> 8;
      out[off + i + lo] = v & 0xff;
    }
  }
  return out;
}

// A packed row of sub-byte samples, stretched so every sample owns a byte and
// spans the full range (a 1-bit sample becomes 0 or 255). 8- and 16-bit rows
// are already byte-aligned and pass through.
function expandRow(row: Uint8Array, bits: number, samplesInRow: number): Uint8Array {
  if (bits >= 8) return row;
  const out = new Uint8Array(samplesInRow);
  const max = (1 << bits) - 1;
  for (let i = 0; i < samplesInRow; i++) {
    const bit = i * bits;
    const byte = row[bit >> 3] ?? 0;
    const shift = 8 - bits - (bit & 7);
    out[i] = Math.round((((byte >> shift) & max) * 255) / max);
  }
  return out;
}

// 16-bit samples, truncated to their high byte — the PNG path's choice, and the
// one that costs a picture nothing a printed page can show.
function highBytes(raw: Uint8Array, le: boolean): Uint8Array {
  const out = new Uint8Array(raw.length >> 1);
  const hi = le ? 1 : 0;
  for (let i = 0; i < out.length; i++) out[i] = raw[i * 2 + hi]!;
  return out;
}

// ——— colour ————————————————————————————————————————————————————————————

interface ColorInfo {
  readonly width: number;
  readonly height: number;
  readonly samples: number;
  readonly bits: number;
  readonly photometric: number;
  /** §8 ColorMap: all reds, then greens, then blues, each 0…65535. */
  readonly palette: ReadonlyArray<number>;
  readonly extra: ReadonlyArray<number>;
}

// §8 PhotometricInterpretation — the samples as the colour they mean, plus the
// alpha channel ExtraSamples marks (association 1 associated / 2 unassociated).
function toChannels(eight: Uint8Array, info: ColorInfo): DecodedTiff {
  const pixels = info.width * info.height;
  const { photometric, samples } = info;

  if (photometric === 3) {
    const rgb = new Uint8Array(pixels * 3);
    const entries = info.palette.length / 3;
    for (let i = 0; i < pixels; i++) {
      const idx = eight[i * samples] ?? 0;
      // A palette read from a sub-byte sample arrives stretched to 0…255, so
      // the index is scaled back to the entry it names.
      const e = info.bits < 8 ? Math.round((idx * ((1 << info.bits) - 1)) / 255) : idx;
      rgb[i * 3] = (info.palette[e] ?? 0) >> 8;
      rgb[i * 3 + 1] = (info.palette[entries + e] ?? 0) >> 8;
      rgb[i * 3 + 2] = (info.palette[entries * 2 + e] ?? 0) >> 8;
    }
    return { width: info.width, height: info.height, channels: 3, data: rgb };
  }

  const color = photometric === 5 ? 4 : photometric === 2 ? 3 : 1;
  if (samples < color) throw new Error('TIFF: too few samples for its colour');
  const alphaAt = info.extra.length > 0 && samples > color ? color : -1;
  const data =
    samples === color ? eight.subarray(0, pixels * color) : new Uint8Array(pixels * color);
  const alpha = alphaAt >= 0 ? new Uint8Array(pixels) : undefined;
  if (samples !== color) {
    for (let i = 0; i < pixels; i++) {
      for (let c = 0; c < color; c++) data[i * color + c] = eight[i * samples + c] ?? 0;
      if (alpha) alpha[i] = eight[i * samples + alphaAt] ?? 255;
    }
  }
  // §8 WhiteIsZero — the greys run the other way, so invert them into the
  // DeviceGray a PDF paints.
  if (photometric === 0) for (let i = 0; i < data.length; i++) data[i] = 255 - data[i]!;
  return {
    width: info.width,
    height: info.height,
    channels: color === 4 ? 4 : color === 3 ? 3 : 1,
    data,
    ...(alpha ? { alpha } : {}),
  };
}
