// Raster-image experts shared by every writer (split out of the PDF image
// XObject module at stage 6.4 so the layout engine and the PageDoc schema do
// not depend on pdf/).
//
// JPEG passes through as-is (readers decode DCT themselves); JPEG 2000
// likewise (wavelet codestream). PNG is decoded with zlib (RFC 1950) and
// per-scanline filters reversed, then re-compressed; RGBA / Gray+Alpha PNGs
// split their alpha channel into a separate soft-mask plane.
//
// Every PNG colour type is read: greyscale, RGB, palette (expanded to RGB, its
// `tRNS` becoming the soft mask) and both alpha forms, at bit depths 1/2/4/8
// (stretched to 8) and 16 (truncated to its high byte), interlaced or not.
// Throws on a malformed input so the caller can decide to fall back or surface
// the error.

import { unzlibSync, zlibSync } from 'fflate';

/** The raster formats this module recognizes and can prepare for embedding. */
export type ImageFormat = 'jpeg' | 'png' | 'jpeg2000';

/**
 * Sniff the raster format from a file's leading magic bytes (JPEG SOI, the PNG
 * signature, or a JP2 box / raw JPEG 2000 codestream). Returns the
 * {@link ImageFormat}, or `null` when none matches.
 */
export function detectImageFormat(bytes: Uint8Array): ImageFormat | null {
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'jpeg';
  }
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  ) {
    return 'png';
  }
  // JPEG 2000: either the JP2 box signature (00 00 00 0C 6A 50 20 20 0D 0A 87 0A)
  // or a raw codestream starting SOC+SIZ (FF 4F FF 51).
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x00 &&
    bytes[1] === 0x00 &&
    bytes[2] === 0x00 &&
    bytes[3] === 0x0c &&
    bytes[4] === 0x6a &&
    bytes[5] === 0x50 &&
    bytes[6] === 0x20 &&
    bytes[7] === 0x20 &&
    bytes[8] === 0x0d &&
    bytes[9] === 0x0a &&
    bytes[10] === 0x87 &&
    bytes[11] === 0x0a
  ) {
    return 'jpeg2000';
  }
  if (
    bytes.length >= 4 &&
    bytes[0] === 0xff &&
    bytes[1] === 0x4f &&
    bytes[2] === 0xff &&
    bytes[3] === 0x51
  ) {
    return 'jpeg2000';
  }
  return null;
}

/** Options controlling how {@link prepareImage} emits an image. */
export interface EmbedImageOptions {
  /**
   * PDF/A-1 forbids transparency (soft masks). When true, PNG alpha is
   * composited onto an opaque white background and no `/SMask` is emitted.
   */
  readonly flattenAlpha?: boolean;
}

/**
 * The ready-to-emit result of decoding and validating one image, from the
 * prepare/add split (oop-design §3.1): {@link prepareImage} is the pure expert —
 * decode, validate (throws on unsupported/corrupt input) and produce the
 * stream bytes; `addImage` only creates the PDF objects. Layout probes with
 * `prepareImage` (no throwaway document), the emit phase replays the prepared
 * result, and other writers (SVG) reuse the mime/dimensions.
 */
export interface PreparedImage {
  readonly format: ImageFormat;
  readonly mimeType: 'image/jp2' | 'image/jpeg' | 'image/png';
  readonly widthPx: number;
  readonly heightPx: number;
  /**
   * ColorSpace/BitsPerComponent are absent for JPEG 2000 (carried inside the
   * JPX codestream).
   */
  readonly colorSpace?: 'DeviceGray' | 'DeviceRGB';
  readonly bitsPerComponent?: number;
  /** The PDF stream filter the `data` bytes are encoded with. */
  readonly filter: 'DCTDecode' | 'FlateDecode' | 'JPXDecode';
  readonly data: Uint8Array;
  /** PNG alpha channel, already FlateDecode-compressed (DeviceGray, 8 bpc). */
  readonly smaskData?: Uint8Array;
}

/**
 * Decode and validate one image into a {@link PreparedImage} ready to embed.
 * JPEG and JPEG 2000 pass through verbatim (readers decode them); PNG is
 * inflated, de-filtered and re-compressed, splitting any alpha into a soft mask.
 *
 * @throws Error when the format is unrecognized, unsupported, or corrupt.
 */
export function prepareImage(bytes: Uint8Array, options: EmbedImageOptions = {}): PreparedImage {
  const format = detectImageFormat(bytes);
  if (format === 'jpeg') return prepareJpeg(bytes);
  if (format === 'png') return preparePng(bytes, options);
  if (format === 'jpeg2000') return prepareJpeg2000(bytes);
  throw new Error('Unsupported image format');
}

// JPEG 2000 (ISO/IEC 15444) goes in verbatim via /JPXDecode — like JPEG via
// /DCTDecode, PDF readers decode the wavelet codestream themselves, so we only
// read the dimensions. NB /JPXDecode is permitted in PDF/A-2/3 but NOT PDF/A-1.
function prepareJpeg2000(bytes: Uint8Array): PreparedImage {
  const { width, height } = readJpeg2000Info(bytes);
  return {
    format: 'jpeg2000',
    mimeType: 'image/jp2',
    widthPx: width,
    heightPx: height,
    filter: 'JPXDecode',
    data: bytes,
  };
}

function readU32(b: Uint8Array, o: number): number {
  return b[o]! * 0x1000000 + (b[o + 1]! << 16) + (b[o + 2]! << 8) + b[o + 3]!;
}

const boxType = (b: Uint8Array, o: number): string =>
  String.fromCharCode(b[o]!, b[o + 1]!, b[o + 2]!, b[o + 3]!);

// Dimensions from the JP2 'jp2h'→'ihdr' box (HEIGHT, WIDTH as u32), or — for a
// raw codestream / when no ihdr — from the SIZ marker (Xsiz/Ysiz − offsets).
function readJpeg2000Info(bytes: Uint8Array): { width: number; height: number } {
  if (bytes[0] === 0xff && bytes[1] === 0x4f) return readSiz(bytes, 2);

  let p = 0;
  while (p + 8 <= bytes.length) {
    let len = readU32(bytes, p);
    const type = boxType(bytes, p + 4);
    let contentStart = p + 8;
    if (len === 1) {
      // 64-bit extended length (low 32 bits suffice for any real image box).
      len = readU32(bytes, p + 12);
      contentStart = p + 16;
    }
    const end = len === 0 ? bytes.length : p + len;
    if (type === 'jp2h') {
      let q = contentStart;
      while (q + 8 <= end) {
        const clen = readU32(bytes, q);
        if (boxType(bytes, q + 4) === 'ihdr') {
          return { height: readU32(bytes, q + 8), width: readU32(bytes, q + 12) };
        }
        if (clen === 0) break;
        q += clen;
      }
    }
    if (len === 0) break;
    p = end;
  }

  // Fallback: scan for the SIZ marker (FF 51) inside the codestream box.
  for (let i = 0; i + 22 < bytes.length; i++) {
    if (bytes[i] === 0xff && bytes[i + 1] === 0x51) return readSiz(bytes, i);
  }
  throw new Error('JPEG 2000: no ihdr/SIZ found');
}

// SIZ marker (T.800 §A.5.1): FF51, Lsiz, Rsiz, then Xsiz/Ysiz/XOsiz/YOsiz.
function readSiz(bytes: Uint8Array, sizOffset: number): { width: number; height: number } {
  const xsiz = readU32(bytes, sizOffset + 6);
  const ysiz = readU32(bytes, sizOffset + 10);
  const xo = readU32(bytes, sizOffset + 14);
  const yo = readU32(bytes, sizOffset + 18);
  const width = xsiz - xo;
  const height = ysiz - yo;
  if (width <= 0 || height <= 0) throw new Error('JPEG 2000: invalid SIZ dimensions');
  return { width, height };
}

function prepareJpeg(bytes: Uint8Array): PreparedImage {
  const info = readJpegInfo(bytes);
  return {
    format: 'jpeg',
    mimeType: 'image/jpeg',
    widthPx: info.width,
    heightPx: info.height,
    colorSpace: info.numComponents === 1 ? 'DeviceGray' : 'DeviceRGB',
    bitsPerComponent: info.precision,
    filter: 'DCTDecode',
    data: bytes,
  };
}

interface JpegInfo {
  readonly width: number;
  readonly height: number;
  readonly precision: number;
  readonly numComponents: number;
}

// ISO/IEC 10918 — JPEG. Walk the marker stream until we find a Start-Of-Frame
// marker (SOFn, except SOF4 = DHT and SOF8 = JPG which are not frames).
function readJpegInfo(bytes: Uint8Array): JpegInfo {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    throw new Error('Not a JPEG (missing SOI)');
  }
  let i = 2;
  while (i + 9 < bytes.length) {
    if (bytes[i] !== 0xff) {
      i++;
      continue;
    }
    let marker = bytes[i + 1]!;
    while (marker === 0xff) {
      i++;
      marker = bytes[i + 1]!;
    }
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
      i += 2;
      continue;
    }
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return {
        precision: bytes[i + 4]!,
        height: (bytes[i + 5]! << 8) | bytes[i + 6]!,
        width: (bytes[i + 7]! << 8) | bytes[i + 8]!,
        numComponents: bytes[i + 9]!,
      };
    }
    const length = (bytes[i + 2]! << 8) | bytes[i + 3]!;
    i += 2 + length;
  }
  throw new Error('JPEG SOFn marker not found');
}

interface DecodedPng {
  readonly width: number;
  readonly height: number;
  readonly raw: Uint8Array;
  readonly colorSpace: 'DeviceRGB' | 'DeviceGray';
  readonly bitsPerComponent: number;
  readonly smaskRaw?: Uint8Array;
}

function preparePng(bytes: Uint8Array, options: EmbedImageOptions = {}): PreparedImage {
  let decoded = decodePng(bytes);
  if (options.flattenAlpha && decoded.smaskRaw) {
    decoded = flattenAlphaOnWhite(decoded);
  }
  return {
    format: 'png',
    mimeType: 'image/png',
    widthPx: decoded.width,
    heightPx: decoded.height,
    colorSpace: decoded.colorSpace,
    bitsPerComponent: decoded.bitsPerComponent,
    filter: 'FlateDecode',
    data: zlibSync(decoded.raw),
    ...(decoded.smaskRaw ? { smaskData: zlibSync(decoded.smaskRaw) } : {}),
  };
}

function decodePng(bytes: Uint8Array): DecodedPng {
  // Chunks: 4-byte length, 4-byte type, length-bytes data, 4-byte CRC.
  // IHDR (header), IDAT (data), IEND (end), and many ancillary chunks.
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlaceMethod = 0;
  let palette: Uint8Array | undefined;
  let paletteAlpha: Uint8Array | undefined;
  const idatChunks: Array<Uint8Array> = [];
  let pos = 8;
  while (pos + 12 <= bytes.length) {
    const len = readU32BE(bytes, pos);
    const type = chunkType(bytes, pos + 4);
    const dataOff = pos + 8;
    if (type === 'IHDR') {
      width = readU32BE(bytes, dataOff);
      height = readU32BE(bytes, dataOff + 4);
      bitDepth = bytes[dataOff + 8]!;
      colorType = bytes[dataOff + 9]!;
      interlaceMethod = bytes[dataOff + 12]!;
    } else if (type === 'PLTE') {
      // §11.2.3 — the palette, three bytes per entry.
      palette = bytes.subarray(dataOff, dataOff + len);
    } else if (type === 'tRNS') {
      // §11.3.2.1 — for a palette image, one alpha byte per entry (entries past
      // the end are opaque).
      paletteAlpha = bytes.subarray(dataOff, dataOff + len);
    } else if (type === 'IDAT') {
      idatChunks.push(bytes.subarray(dataOff, dataOff + len));
    } else if (type === 'IEND') {
      break;
    }
    pos = dataOff + len + 4;
  }
  if (width === 0 || height === 0) throw new Error('PNG: missing IHDR');
  if (idatChunks.length === 0) throw new Error('PNG: no IDAT');
  if (![1, 2, 4, 8, 16].includes(bitDepth)) {
    throw new Error(`PNG: bit depth ${bitDepth} not supported`);
  }

  const channels = pngChannels(colorType);
  if (channels === 0) throw new Error(`PNG color type ${colorType} not supported`);

  const inflated = unzlibSync(concatBytes(idatChunks));
  const raw = decodeSamples(
    inflated,
    width,
    height,
    channels,
    bitDepth,
    interlaceMethod === 1,
    // A palette sample is an INDEX; scaling it to the byte range would look up
    // the wrong colour. Every other sample is an intensity, and a 1/2/4-bit one
    // has to be stretched to the 8 bits the PDF image carries.
    colorType !== 3,
  );

  return splitChannels(width, height, colorType, raw, palette, paletteAlpha);
}

/**
 * §7.2 / §9 — the image's samples, one byte each, in `width × height ×
 * channels` order: the IDAT stream inflated, each scanline unfiltered against
 * the one before it, and (when the image is interlaced) each of Adam7's seven
 * passes scattered back into the full-size grid.
 *
 * @param inflated    The inflated IDAT stream.
 * @param width       Image width in pixels.
 * @param height      Image height in pixels.
 * @param channels    Samples per pixel for the colour type.
 * @param bitDepth    Bits per sample (1, 2, 4, 8 or 16).
 * @param interlaced  True for interlace method 1 (Adam7).
 * @param scaleToByte Stretch a sub-byte sample to 0…255 — right for an
 *                    intensity, wrong for a palette index.
 * @returns One byte per sample.
 */
function decodeSamples(
  inflated: Uint8Array,
  width: number,
  height: number,
  channels: number,
  bitDepth: number,
  interlaced: boolean,
  scaleToByte: boolean,
): Uint8Array {
  const out = new Uint8Array(width * height * channels);
  const passes = interlaced ? ADAM7_PASSES : [{ xStart: 0, yStart: 0, xStep: 1, yStep: 1 }];
  // A sub-byte sample counts 0…2ⁿ−1 and the PDF image carries bytes, so the
  // value is stretched over the full range (a 1-bit 1 is white, not 1/255).
  const maxSample = (1 << Math.min(bitDepth, 8)) - 1;
  const bpp = Math.max(1, Math.ceil((channels * bitDepth) / 8));
  let off = 0;
  for (const p of passes) {
    const pw = Math.ceil((width - p.xStart) / p.xStep);
    const ph = Math.ceil((height - p.yStart) / p.yStep);
    if (pw <= 0 || ph <= 0) continue;
    const lineBytes = Math.ceil((pw * channels * bitDepth) / 8);
    let prev: Uint8Array | null = null;
    for (let y = 0; y < ph; y++) {
      if (off + 1 + lineBytes > inflated.length) throw new Error('PNG: truncated image data');
      const filterType = inflated[off]!;
      const line = unfilterScanline(inflated, off + 1, lineBytes, bpp, filterType, prev);
      off += 1 + lineBytes;
      const destY = p.yStart + y * p.yStep;
      for (let px = 0; px < pw; px++) {
        const destX = p.xStart + px * p.xStep;
        const dest = (destY * width + destX) * channels;
        for (let ch = 0; ch < channels; ch++) {
          const v = sampleAt(line, px * channels + ch, bitDepth);
          out[dest + ch] = scaleToByte && bitDepth < 8 ? Math.round((v * 255) / maxSample) : v;
        }
      }
      prev = line;
    }
  }
  return out;
}

/**
 * §9.2 — Adam7's seven passes, each a sub-grid of the image: where the pass
 * starts and how far apart its pixels sit.
 */
const ADAM7_PASSES: ReadonlyArray<{
  xStart: number;
  yStart: number;
  xStep: number;
  yStep: number;
}> = [
  { xStart: 0, yStart: 0, xStep: 8, yStep: 8 },
  { xStart: 4, yStart: 0, xStep: 8, yStep: 8 },
  { xStart: 0, yStart: 4, xStep: 4, yStep: 8 },
  { xStart: 2, yStart: 0, xStep: 4, yStep: 4 },
  { xStart: 0, yStart: 2, xStep: 2, yStep: 4 },
  { xStart: 1, yStart: 0, xStep: 2, yStep: 2 },
  { xStart: 0, yStart: 1, xStep: 1, yStep: 2 },
];

// §9.2 — one scanline, its filter reversed against the reconstructed line above.
function unfilterScanline(
  src: Uint8Array,
  from: number,
  lineBytes: number,
  bpp: number,
  filterType: number,
  prev: Uint8Array | null,
): Uint8Array {
  const decoded = new Uint8Array(lineBytes);
  for (let x = 0; x < lineBytes; x++) {
    const filt = src[from + x]!;
    const a = x < bpp ? 0 : decoded[x - bpp]!;
    const b = prev ? prev[x]! : 0;
    const c = !prev || x < bpp ? 0 : prev[x - bpp]!;
    let v: number;
    switch (filterType) {
      case 0:
        v = filt;
        break;
      case 1:
        v = filt + a;
        break;
      case 2:
        v = filt + b;
        break;
      case 3:
        v = filt + Math.floor((a + b) / 2);
        break;
      case 4:
        v = filt + paethPredictor(a, b, c);
        break;
      default:
        throw new Error(`PNG: unknown filter type ${filterType}`);
    }
    decoded[x] = v & 0xff;
  }
  return decoded;
}

// Sample `i` of an unfiltered scanline. Sub-byte samples pack big-endian-most
// significant first; a 16-bit one keeps its high byte (the PDF image is 8-bit).
function sampleAt(line: Uint8Array, i: number, bitDepth: number): number {
  if (bitDepth === 8) return line[i]!;
  if (bitDepth === 16) return line[i * 2]!;
  const per = 8 / bitDepth;
  const byte = line[Math.floor(i / per)] ?? 0;
  const shift = 8 - bitDepth * ((i % per) + 1);
  return (byte >> shift) & ((1 << bitDepth) - 1);
}

// Composite a decoded image with an alpha channel onto an opaque white
// background: out = src·α + 255·(1−α). Produces an alpha-free image suitable
// for PDF/A-1 (which bans soft masks). Operates per colour channel.
function flattenAlphaOnWhite(decoded: DecodedPng): DecodedPng {
  const alpha = decoded.smaskRaw;
  if (!alpha) return decoded;
  const channels = decoded.colorSpace === 'DeviceRGB' ? 3 : 1;
  const pixelCount = decoded.width * decoded.height;
  const out = new Uint8Array(pixelCount * channels);
  for (let i = 0; i < pixelCount; i++) {
    const a = alpha[i]! / 255;
    for (let c = 0; c < channels; c++) {
      const src = decoded.raw[i * channels + c]!;
      out[i * channels + c] = Math.round(src * a + 255 * (1 - a));
    }
  }
  return {
    width: decoded.width,
    height: decoded.height,
    raw: out,
    colorSpace: decoded.colorSpace,
    bitsPerComponent: decoded.bitsPerComponent,
    // smaskRaw intentionally dropped.
  };
}

function splitChannels(
  width: number,
  height: number,
  colorType: number,
  raw: Uint8Array,
  palette?: Uint8Array,
  paletteAlpha?: Uint8Array,
): DecodedPng {
  const pixelCount = width * height;
  if (colorType === 0) {
    return { width, height, raw, colorSpace: 'DeviceGray', bitsPerComponent: 8 };
  }
  // §11.2.3 — an indexed image: each sample is a palette entry. Expanded to RGB
  // here rather than carried as a PDF /Indexed space, so one code path serves
  // every writer; `tRNS` gives the entries their alpha and becomes the soft
  // mask. Left unread, 45 of the corpus's images drew as nothing at all.
  if (colorType === 3) {
    if (!palette || palette.length < 3) throw new Error('PNG: palette image without PLTE');
    const entries = Math.floor(palette.length / 3);
    const rgb = new Uint8Array(pixelCount * 3);
    const alpha = paletteAlpha && paletteAlpha.length > 0 ? new Uint8Array(pixelCount) : undefined;
    for (let i = 0; i < pixelCount; i++) {
      const idx = raw[i]!;
      const src = (idx < entries ? idx : 0) * 3;
      rgb[i * 3] = palette[src]!;
      rgb[i * 3 + 1] = palette[src + 1]!;
      rgb[i * 3 + 2] = palette[src + 2]!;
      if (alpha) alpha[i] = paletteAlpha![idx] ?? 255;
    }
    return {
      width,
      height,
      raw: rgb,
      colorSpace: 'DeviceRGB',
      bitsPerComponent: 8,
      ...(alpha ? { smaskRaw: alpha } : {}),
    };
  }
  if (colorType === 2) {
    return { width, height, raw, colorSpace: 'DeviceRGB', bitsPerComponent: 8 };
  }
  if (colorType === 4) {
    const gray = new Uint8Array(pixelCount);
    const alpha = new Uint8Array(pixelCount);
    for (let i = 0; i < pixelCount; i++) {
      gray[i] = raw[i * 2]!;
      alpha[i] = raw[i * 2 + 1]!;
    }
    return {
      width,
      height,
      raw: gray,
      colorSpace: 'DeviceGray',
      bitsPerComponent: 8,
      smaskRaw: alpha,
    };
  }
  if (colorType === 6) {
    const rgb = new Uint8Array(pixelCount * 3);
    const alpha = new Uint8Array(pixelCount);
    for (let i = 0; i < pixelCount; i++) {
      rgb[i * 3] = raw[i * 4]!;
      rgb[i * 3 + 1] = raw[i * 4 + 1]!;
      rgb[i * 3 + 2] = raw[i * 4 + 2]!;
      alpha[i] = raw[i * 4 + 3]!;
    }
    return {
      width,
      height,
      raw: rgb,
      colorSpace: 'DeviceRGB',
      bitsPerComponent: 8,
      smaskRaw: alpha,
    };
  }
  throw new Error(`PNG color type ${colorType} not supported`);
}

function pngChannels(colorType: number): number {
  if (colorType === 0) return 1;
  if (colorType === 2) return 3;
  if (colorType === 3) return 1;
  if (colorType === 4) return 2;
  if (colorType === 6) return 4;
  return 0;
}

function paethPredictor(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}

function readU32BE(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]! * 0x1000000 +
    ((bytes[offset + 1]! << 16) | (bytes[offset + 2]! << 8) | bytes[offset + 3]!)
  );
}

function chunkType(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(
    bytes[offset]!,
    bytes[offset + 1]!,
    bytes[offset + 2]!,
    bytes[offset + 3]!,
  );
}

function concatBytes(parts: ReadonlyArray<Uint8Array>): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
