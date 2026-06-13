// E-PDF EP6 — decode a PDF image XObject (ISO 32000-1 §8.9) into a standalone
// raster file the FlowDoc resource store can hold. JPEG (/DCTDecode) and JPEG
// 2000 (/JPXDecode) pass through verbatim — readers decode those wavelet/DCT
// codestreams themselves, so we only strip any filters layered before them.
// Everything else is decoded to raw samples and re-wrapped as PNG (png-encode).
//
// Colour spaces: DeviceGray/RGB/CMYK, CalGray/CalRGB, ICCBased (by /N component
// count), Indexed (expanded against its palette). Filters: Flate, LZW (EP12),
// RunLength, ASCII85, ASCIIHex, plus PNG/TIFF predictors on Flate and LZW. Bit
// depths 1/2/4/8/16. An /SMask becomes the PNG alpha channel. Unsupported inputs
// — stencil /ImageMask, Separation/DeviceN/Lab, CCITT/JBIG2 fax — return a typed
// reason so the caller records a loss instead of emitting a broken image.

import { unzlibSync } from 'fflate';

import { encodePng } from './png-encode';
import { reversePredictor } from './predictor';
import type { PngColor } from './png-encode';
import type { PdfDict, PdfValue } from '@/pdf/objects';

import type { PdfFile } from './document';
import { PDF_NULL, PdfHexString, PdfName, PdfStream } from '@/pdf/objects';

export type DecodedImage =
  | {
      readonly ok: true;
      readonly bytes: Uint8Array;
      readonly format: 'png' | 'jpeg' | 'jpeg2000';
      readonly widthPx: number;
      readonly heightPx: number;
      readonly degraded?: string; // a partial loss (e.g. a JPEG's alpha dropped)
    }
  | { readonly ok: false; readonly severity: 'dropped' | 'degraded'; readonly detail: string };

const MAX_PIXELS = 40_000_000; // DoS guard (~40 MP)

export function decodePdfImage(file: PdfFile, stream: PdfStream): DecodedImage {
  const d = stream.dict;
  const width = intOf(file.get(d, 'Width')) || intOf(file.get(d, 'W'));
  const height = intOf(file.get(d, 'Height')) || intOf(file.get(d, 'H'));
  if (width <= 0 || height <= 0) return fail('dropped', 'image with no dimensions');
  if (width * height > MAX_PIXELS) return fail('dropped', 'image too large to decode');

  if (boolOf(d.get('ImageMask')) || boolOf(d.get('IM'))) {
    return fail('dropped', 'stencil image mask not reconstructed');
  }

  const filters = filterNames(file, d);
  const last = filters[filters.length - 1];

  // JPEG / JPEG 2000 ride through verbatim — decode only the filters that wrap them.
  if (last === 'DCTDecode' || last === 'DCT') {
    const degraded = hasSMask(file, d)
      ? 'image transparency dropped (JPEG carries no alpha)'
      : undefined;
    return {
      ok: true,
      bytes: applyChainExceptLast(filters, stream.data),
      format: 'jpeg',
      widthPx: width,
      heightPx: height,
      ...(degraded ? { degraded } : {}),
    };
  }
  if (last === 'JPXDecode') {
    return {
      ok: true,
      bytes: applyChainExceptLast(filters, stream.data),
      format: 'jpeg2000',
      widthPx: width,
      heightPx: height,
      degraded: 'JPEG 2000 image — limited viewer support',
    };
  }
  if (last === 'CCITTFaxDecode' || last === 'CCF' || last === 'JBIG2Decode') {
    return fail('dropped', 'fax-encoded (CCITT/JBIG2) image not decoded');
  }

  const decoded = decodeToSamples(file, stream, filters, width, height);
  if (typeof decoded === 'string') return fail('dropped', decoded);

  // Fold an /SMask in as the alpha channel (PNG path only).
  const alpha = decodeSMask(file, d, width, height);
  const { color, samples } = alpha ? combineAlpha(decoded, alpha) : decoded;
  return {
    ok: true,
    bytes: encodePng(width, height, color, samples),
    format: 'png',
    widthPx: width,
    heightPx: height,
  };
}

// --- the sample (PNG) path --------------------------------------------------

interface RawColor {
  readonly color: 'gray' | 'rgb';
  readonly samples: Uint8Array; // 8-bit interleaved, width*height*(1|3)
}

// Returns the decoded colour image, or a human reason string on failure.
function decodeToSamples(
  file: PdfFile,
  stream: PdfStream,
  filters: ReadonlyArray<string>,
  width: number,
  height: number,
): RawColor | string {
  const d = stream.dict;
  const raw = decodeChain(file, stream, filters);
  if (!raw) return 'undecodable image stream';
  const cs = resolveColorSpace(file, d.get('ColorSpace') ?? d.get('CS'));
  if (!cs) return 'unsupported image colour space';
  const bpc = intOf(file.get(d, 'BitsPerComponent')) || intOf(file.get(d, 'BPC')) || 8;
  const decodeArr = decodeArrayOf(file, d);
  const integers = unpackSamples(raw, width, height, cs.components, bpc);
  return toColor(cs, integers, width * height, bpc, decodeArr);
}

interface ColorSpace {
  readonly kind: 'gray' | 'rgb' | 'cmyk' | 'indexed';
  readonly components: number;
  readonly base?: ColorSpace; // indexed
  readonly hival?: number;
  readonly lookup?: Uint8Array;
}

function resolveColorSpace(file: PdfFile, csVal: PdfValue | undefined): ColorSpace | undefined {
  const cs = file.resolve(csVal ?? PDF_NULL);
  if (cs instanceof PdfName) return namedColorSpace(cs.value);
  if (!Array.isArray(cs) || cs.length === 0) return undefined;
  const head = file.resolve(cs[0]!);
  const tag = head instanceof PdfName ? head.value : '';
  if (tag === 'ICCBased') {
    const profile = file.resolve(cs[1]!);
    if (profile instanceof PdfStream) {
      const n = intOf(file.get(profile.dict, 'N'));
      if (n === 1) return { kind: 'gray', components: 1 };
      if (n === 3) return { kind: 'rgb', components: 3 };
      if (n === 4) return { kind: 'cmyk', components: 4 };
      // No /N: fall back to the alternate space.
      const alt = resolveColorSpace(file, profile.dict.get('Alternate'));
      if (alt) return alt;
    }
    return undefined;
  }
  if (tag === 'CalGray' || tag === 'G') return { kind: 'gray', components: 1 };
  if (tag === 'CalRGB' || tag === 'RGB') return { kind: 'rgb', components: 3 };
  if (tag === 'Indexed' || tag === 'I') {
    const base = resolveColorSpace(file, cs[1]);
    const hival = intOf(file.resolve(cs[2] ?? PDF_NULL));
    const lookup = valueToBytes(file, cs[3]);
    if (!base || base.kind === 'indexed' || !lookup) return undefined;
    return { kind: 'indexed', components: 1, base, hival, lookup };
  }
  // Separation / DeviceN / Lab / Pattern — tint transforms not reconstructed.
  return undefined;
}

function namedColorSpace(name: string): ColorSpace | undefined {
  if (name === 'DeviceGray' || name === 'G' || name === 'CalGray')
    return { kind: 'gray', components: 1 };
  if (name === 'DeviceRGB' || name === 'RGB' || name === 'CalRGB')
    return { kind: 'rgb', components: 3 };
  if (name === 'DeviceCMYK' || name === 'CMYK') return { kind: 'cmyk', components: 4 };
  return undefined;
}

// Map raw integer samples (0..2^bpc−1) through the colour space to 8-bit gray/RGB.
function toColor(
  cs: ColorSpace,
  s: Uint16Array,
  px: number,
  bpc: number,
  decode: ReadonlyArray<number> | undefined,
): RawColor {
  const maxv = (1 << Math.min(bpc, 15)) * (bpc >= 16 ? 2 : 1) - 1; // 2^bpc − 1
  const c01 = (v: number, comp: number): number => {
    const t = maxv > 0 ? v / maxv : 0;
    if (decode && decode.length >= (comp + 1) * 2) {
      const dmin = decode[comp * 2]!;
      const dmax = decode[comp * 2 + 1]!;
      return clamp01(dmin + t * (dmax - dmin));
    }
    return t;
  };
  const to8 = (v: number, comp: number): number => Math.round(c01(v, comp) * 255);

  if (cs.kind === 'gray') {
    const out = new Uint8Array(px);
    for (let i = 0; i < px; i++) out[i] = to8(s[i]!, 0);
    return { color: 'gray', samples: out };
  }
  if (cs.kind === 'rgb') {
    const out = new Uint8Array(px * 3);
    for (let i = 0; i < px; i++) {
      out[i * 3] = to8(s[i * 3]!, 0);
      out[i * 3 + 1] = to8(s[i * 3 + 1]!, 1);
      out[i * 3 + 2] = to8(s[i * 3 + 2]!, 2);
    }
    return { color: 'rgb', samples: out };
  }
  if (cs.kind === 'cmyk') {
    const out = new Uint8Array(px * 3);
    for (let i = 0; i < px; i++) {
      const c = c01(s[i * 4]!, 0);
      const m = c01(s[i * 4 + 1]!, 1);
      const y = c01(s[i * 4 + 2]!, 2);
      const k = c01(s[i * 4 + 3]!, 3);
      out[i * 3] = Math.round(255 * (1 - c) * (1 - k));
      out[i * 3 + 1] = Math.round(255 * (1 - m) * (1 - k));
      out[i * 3 + 2] = Math.round(255 * (1 - y) * (1 - k));
    }
    return { color: 'rgb', samples: out };
  }
  // indexed
  const base = cs.base!;
  const lookup = cs.lookup!;
  const hival = cs.hival ?? 255;
  const bn = base.components;
  if (base.kind === 'gray') {
    const out = new Uint8Array(px);
    for (let i = 0; i < px; i++) out[i] = lookup[Math.min(s[i]!, hival) * bn] ?? 0;
    return { color: 'gray', samples: out };
  }
  const out = new Uint8Array(px * 3);
  for (let i = 0; i < px; i++) {
    const off = Math.min(s[i]!, hival) * bn;
    const [r, g, b] = paletteRgb(base, lookup, off);
    out[i * 3] = r;
    out[i * 3 + 1] = g;
    out[i * 3 + 2] = b;
  }
  return { color: 'rgb', samples: out };
}

// One palette entry (8-bit per base component) → RGB.
function paletteRgb(base: ColorSpace, lookup: Uint8Array, off: number): [number, number, number] {
  if (base.kind === 'rgb') return [lookup[off] ?? 0, lookup[off + 1] ?? 0, lookup[off + 2] ?? 0];
  if (base.kind === 'cmyk') {
    const c = (lookup[off] ?? 0) / 255;
    const m = (lookup[off + 1] ?? 0) / 255;
    const y = (lookup[off + 2] ?? 0) / 255;
    const k = (lookup[off + 3] ?? 0) / 255;
    return [
      Math.round(255 * (1 - c) * (1 - k)),
      Math.round(255 * (1 - m) * (1 - k)),
      Math.round(255 * (1 - y) * (1 - k)),
    ];
  }
  const g = lookup[off] ?? 0;
  return [g, g, g];
}

// Unpack row-byte-aligned samples (§7.4.4 image data is padded per scanline).
function unpackSamples(
  raw: Uint8Array,
  width: number,
  height: number,
  ncomp: number,
  bpc: number,
): Uint16Array {
  const perRow = width * ncomp;
  const out = new Uint16Array(perRow * height);
  const rowBytes = Math.ceil((perRow * bpc) / 8);
  for (let y = 0; y < height; y++) {
    const rowOff = y * rowBytes;
    const base = y * perRow;
    if (bpc === 8) {
      for (let i = 0; i < perRow; i++) out[base + i] = raw[rowOff + i] ?? 0;
    } else if (bpc === 16) {
      for (let i = 0; i < perRow; i++) {
        out[base + i] = ((raw[rowOff + 2 * i] ?? 0) << 8) | (raw[rowOff + 2 * i + 1] ?? 0);
      }
    } else {
      const mask = (1 << bpc) - 1;
      let bit = 0;
      for (let i = 0; i < perRow; i++) {
        const bytePos = rowOff + (bit >> 3);
        const shift = 8 - bpc - (bit & 7);
        out[base + i] = (((raw[bytePos] ?? 0) >> shift) & mask) >>> 0;
        bit += bpc;
      }
    }
  }
  return out;
}

// --- /SMask alpha -----------------------------------------------------------

interface Alpha {
  readonly data: Uint8Array; // 8-bit, width*height
}

function decodeSMask(file: PdfFile, d: PdfDict, width: number, height: number): Alpha | undefined {
  const sm = file.resolve(d.get('SMask') ?? PDF_NULL);
  if (!(sm instanceof PdfStream)) return undefined;
  const sw = intOf(file.get(sm.dict, 'Width'));
  const sh = intOf(file.get(sm.dict, 'Height'));
  if (sw <= 0 || sh <= 0) return undefined;
  const decoded = decodeToSamples(file, sm, filterNames(file, sm.dict), sw, sh);
  if (typeof decoded === 'string') return undefined;
  const ch = decoded.color === 'rgb' ? 3 : 1;
  // The mask is grayscale; take its first channel and resample to the image grid.
  const gray = new Uint8Array(sw * sh);
  for (let i = 0; i < sw * sh; i++) gray[i] = decoded.samples[i * ch]!;
  if (sw === width && sh === height) return { data: gray };
  const out = new Uint8Array(width * height);
  for (let y = 0; y < height; y++) {
    const sy = Math.min(sh - 1, Math.floor((y * sh) / height));
    for (let x = 0; x < width; x++) {
      const sx = Math.min(sw - 1, Math.floor((x * sw) / width));
      out[y * width + x] = gray[sy * sw + sx]!;
    }
  }
  return { data: out };
}

function combineAlpha(color: RawColor, alpha: Alpha): { color: PngColor; samples: Uint8Array } {
  const px = alpha.data.length;
  if (color.color === 'gray') {
    const out = new Uint8Array(px * 2);
    for (let i = 0; i < px; i++) {
      out[i * 2] = color.samples[i]!;
      out[i * 2 + 1] = alpha.data[i]!;
    }
    return { color: 'gray-alpha', samples: out };
  }
  const out = new Uint8Array(px * 4);
  for (let i = 0; i < px; i++) {
    out[i * 4] = color.samples[i * 3]!;
    out[i * 4 + 1] = color.samples[i * 3 + 1]!;
    out[i * 4 + 2] = color.samples[i * 3 + 2]!;
    out[i * 4 + 3] = alpha.data[i]!;
  }
  return { color: 'rgba', samples: out };
}

// --- filters ----------------------------------------------------------------

function filterNames(file: PdfFile, d: PdfDict): Array<string> {
  const f = file.resolve(d.get('Filter') ?? PDF_NULL);
  const arr: ReadonlyArray<PdfValue> = Array.isArray(f) ? f : [f];
  const out: Array<string> = [];
  for (const x of arr) {
    const r = file.resolve(x);
    if (r instanceof PdfName) out.push(r.value);
  }
  return out;
}

// Decode every filter (the sample path never sees DCT/JPX/CCITT — those are
// handled before this is called), then reverse any /Predictor. Flate and LZW
// (EP12) both carry the predictor in /DecodeParms.
function decodeChain(
  file: PdfFile,
  stream: PdfStream,
  filters: ReadonlyArray<string>,
): Uint8Array | undefined {
  let data = stream.data;
  let mayPredict = false;
  for (const f of filters) {
    if (f === 'FlateDecode' || f === 'Fl') {
      try {
        data = unzlibSync(data);
        mayPredict = true;
      } catch {
        return undefined;
      }
    } else if (f === 'LZWDecode' || f === 'LZW') {
      const dec = lzwDecode(data, lzwEarlyChange(file, stream.dict));
      if (!dec) return undefined;
      data = dec;
      mayPredict = true;
    } else if (f === 'RunLengthDecode' || f === 'RL') {
      data = runLengthDecode(data);
    } else if (f === 'ASCII85Decode' || f === 'A85') {
      data = ascii85Decode(data);
    } else if (f === 'ASCIIHexDecode' || f === 'AHx') {
      data = asciiHexDecode(data);
    } else {
      return undefined;
    }
  }
  return mayPredict ? applyPredictor(file, stream.dict, data) : data;
}

function applyChainExceptLast(filters: ReadonlyArray<string>, raw: Uint8Array): Uint8Array {
  let data = raw;
  for (let i = 0; i < filters.length - 1; i++) {
    const f = filters[i]!;
    if (f === 'FlateDecode' || f === 'Fl') {
      try {
        data = unzlibSync(data);
      } catch {
        /* leave undecoded */
      }
    } else if (f === 'LZWDecode' || f === 'LZW') data = lzwDecode(data, 1) ?? data;
    else if (f === 'RunLengthDecode' || f === 'RL') data = runLengthDecode(data);
    else if (f === 'ASCII85Decode' || f === 'A85') data = ascii85Decode(data);
    else if (f === 'ASCIIHexDecode' || f === 'AHx') data = asciiHexDecode(data);
  }
  return data;
}

// §7.4.4.4 — reverse a /Predictor from the stream's /DecodeParms (shared math
// in predictor.ts).
function applyPredictor(file: PdfFile, d: PdfDict, data: Uint8Array): Uint8Array {
  const parms = decodeParmsOf(file, d);
  if (!parms) return data;
  const predictor = intOf(file.get(parms, 'Predictor'));
  if (predictor < 2) return data;
  return reversePredictor(data, {
    predictor,
    colors: intOf(file.get(parms, 'Colors')) || 1,
    bitsPerComponent: intOf(file.get(parms, 'BitsPerComponent')) || 8,
    columns: intOf(file.get(parms, 'Columns')) || 1,
  });
}

// §7.4.4.2 — the PDF/TIFF variant of LZW. Variable-width codes 9→12 bits, a
// clear-table code (256) and an end-of-data code (257); /EarlyChange (default 1)
// widens the code one step before the table fills. The KwKwK case (a code not
// yet in the table) repeats the previous string plus its own first byte. Output
// is bounded by the per-image pixel guard to cap a decompression bomb.
const MAX_LZW_OUT = MAX_PIXELS * 4;

function lzwDecode(data: Uint8Array, earlyChange: number): Uint8Array | undefined {
  let out = new Uint8Array(1 << 12); // grows by doubling
  let outLen = 0;
  const emit = (e: Uint8Array): boolean => {
    if (outLen + e.length > MAX_LZW_OUT) return false;
    if (outLen + e.length > out.length) {
      let cap = out.length * 2;
      while (cap < outLen + e.length) cap *= 2;
      const grown = new Uint8Array(cap);
      grown.set(out.subarray(0, outLen));
      out = grown;
    }
    out.set(e, outLen);
    outLen += e.length;
    return true;
  };

  let bitBuffer = 0;
  let bitCount = 0;
  let pos = 0;
  let codeLength = 9;
  const readCode = (): number => {
    while (bitCount < codeLength) {
      if (pos >= data.length) return -1;
      bitBuffer = ((bitBuffer << 8) | data[pos++]!) >>> 0;
      bitCount += 8;
    }
    bitCount -= codeLength;
    return (bitBuffer >>> bitCount) & ((1 << codeLength) - 1);
  };

  const dict = new Array<Uint8Array>(4096);
  let nextCode = 258;
  let prev = -1;
  const reset = (): void => {
    for (let i = 0; i < 256; i++) dict[i] = Uint8Array.of(i);
    nextCode = 258;
    codeLength = 9;
    prev = -1;
  };
  reset();

  for (;;) {
    const code = readCode();
    if (code < 0 || code === 257) break; // out of data / end-of-data
    if (code === 256) {
      reset();
      continue;
    }
    if (prev < 0) {
      const first = dict[code];
      if (!first || !emit(first)) break;
      prev = code;
      continue;
    }
    const prevEntry = dict[prev]!;
    // A known code uses its entry; an as-yet-unassigned code is KwKwK.
    let entry = code < nextCode ? dict[code] : undefined;
    if (!entry) {
      entry = new Uint8Array(prevEntry.length + 1);
      entry.set(prevEntry);
      entry[prevEntry.length] = prevEntry[0]!;
    }
    if (!emit(entry)) break;
    if (nextCode < 4096) {
      const added = new Uint8Array(prevEntry.length + 1);
      added.set(prevEntry);
      added[prevEntry.length] = entry[0]!;
      dict[nextCode++] = added;
      if (nextCode + earlyChange === 512) codeLength = 10;
      else if (nextCode + earlyChange === 1024) codeLength = 11;
      else if (nextCode + earlyChange === 2048) codeLength = 12;
    }
    prev = code;
  }
  return out.subarray(0, outLen);
}

// /DecodeParms /EarlyChange — 1 (the default) or 0; absent means 1.
function lzwEarlyChange(file: PdfFile, d: PdfDict): number {
  const parms = decodeParmsOf(file, d);
  const ec = parms?.get('EarlyChange');
  return ec !== undefined && file.resolve(ec) === 0 ? 0 : 1;
}

function runLengthDecode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let i = 0;
  while (i < data.length) {
    const len = data[i++]!;
    if (len === 128) break; // EOD
    if (len < 128) {
      for (let j = 0; j <= len && i < data.length; j++) out.push(data[i++]!);
    } else {
      const b = data[i++] ?? 0;
      for (let j = 0; j < 257 - len; j++) out.push(b);
    }
  }
  return Uint8Array.from(out);
}

function ascii85Decode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let tuple = 0;
  let count = 0;
  for (const c of data) {
    if (c === 0x7e) break; // ~> terminator
    if (c <= 0x20) continue; // whitespace
    if (c === 0x7a && count === 0) {
      out.push(0, 0, 0, 0);
      continue;
    }
    if (c < 0x21 || c > 0x75) continue;
    tuple = tuple * 85 + (c - 0x21);
    if (++count === 5) {
      out.push((tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff);
      tuple = 0;
      count = 0;
    }
  }
  if (count > 0) {
    for (let i = count; i < 5; i++) tuple = tuple * 85 + 84;
    for (let i = 0; i < count - 1; i++) out.push((tuple >>> (24 - i * 8)) & 0xff);
  }
  return Uint8Array.from(out);
}

function asciiHexDecode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let hi = -1;
  for (const c of data) {
    if (c === 0x3e) break; // '>'
    const v = hexVal(c);
    if (v < 0) continue;
    if (hi < 0) hi = v;
    else {
      out.push((hi << 4) | v);
      hi = -1;
    }
  }
  if (hi >= 0) out.push(hi << 4);
  return Uint8Array.from(out);
}

// --- small helpers ----------------------------------------------------------

function decodeParmsOf(file: PdfFile, d: PdfDict): PdfDict | undefined {
  const p = file.resolve(d.get('DecodeParms') ?? d.get('DP') ?? PDF_NULL);
  if (p instanceof Map) return p;
  if (Array.isArray(p)) {
    for (const e of p) {
      const r = file.resolve(e);
      if (r instanceof Map) return r; // the predictor parms (Flate / LZW carry them)
    }
  }
  return undefined;
}

function decodeArrayOf(file: PdfFile, d: PdfDict): Array<number> | undefined {
  const v = file.resolve(d.get('Decode') ?? d.get('D') ?? PDF_NULL);
  if (!Array.isArray(v)) return undefined;
  const nums = v.map((x) => (typeof x === 'number' ? x : NaN));
  return nums.some((n) => !Number.isFinite(n)) ? undefined : nums;
}

function hasSMask(file: PdfFile, d: PdfDict): boolean {
  return file.resolve(d.get('SMask') ?? PDF_NULL) instanceof PdfStream;
}

function valueToBytes(file: PdfFile, v: PdfValue | undefined): Uint8Array | undefined {
  const r = file.resolve(v ?? PDF_NULL);
  if (r instanceof PdfHexString) return r.bytes;
  if (typeof r === 'string') {
    const out = new Uint8Array(r.length);
    for (let i = 0; i < r.length; i++) out[i] = r.charCodeAt(i) & 0xff;
    return out;
  }
  if (r instanceof PdfStream) return file.streamData(r);
  return undefined;
}

function intOf(v: PdfValue | undefined): number {
  return typeof v === 'number' ? Math.round(v) : 0;
}

function boolOf(v: PdfValue | undefined): boolean {
  return v === true;
}

function clamp01(x: number): number {
  return x < 0 ? 0 : x > 1 ? 1 : x;
}

function hexVal(c: number): number {
  if (c >= 0x30 && c <= 0x39) return c - 0x30;
  if (c >= 0x41 && c <= 0x46) return c - 0x41 + 10;
  if (c >= 0x61 && c <= 0x66) return c - 0x61 + 10;
  return -1;
}

function fail(severity: 'dropped' | 'degraded', detail: string): DecodedImage {
  return { ok: false, severity, detail };
}
