// MS-WMF §2.2.2.9 / MS-EMF §2.2.2 — the device-independent bitmap a blit
// record carries, decoded to straight RGBA.
//
// A DIB is a header, a colour table and the pixels, and the pixels are stored
// BOTTOM row first with every row padded out to a four-byte boundary. Both
// metafile dialects embed them the same way, so both readers decode through
// here; what comes out is fed to `encodePng` and travels on as an ordinary
// picture resource, which is how the rest of the pipeline already carries a
// raster.
//
// What a metafile actually contains (measured over the corpus's 35 files with
// a blit): 1/4/8/16/24/32-bit BI_RGB and nothing else. BI_BITFIELDS costs three
// lines so it is here; the run-length forms are not, and a caller that meets
// one is told so rather than shown a wrong picture.

import { encodePng } from '@/core/png-encode';

/** A decoded bitmap: straight (un-premultiplied) RGBA, TOP row first. */
export interface DibImage {
  readonly width: number;
  readonly height: number;
  readonly rgba: Uint8Array;
  /**
   * Whether the source stored its rows bottom-first — which is the usual way,
   * and which a caller needs to know to read a source rectangle: a blit states
   * one in the BITMAP's own coordinates, whose origin is its first row.
   */
  readonly bottomUp?: boolean;
}

// A bitmap larger than this is a malformed header, not a picture: refuse it
// rather than reserve the memory it asks for.
const MAX_PIXELS = 30_000_000;

/** What a caller knows about the bitmap that its own header does not say. */
export interface DibOptions {
  /**
   * Where the pixels start. EMF states this; a WMF stores the two parts back to
   * back, so leaving it out means "after the palette".
   */
  readonly bitsAt?: number;
  /** How many pixel bytes there are, when the record states it. */
  readonly cbBits?: number;
  /**
   * Whether a 32-bit bitmap's fourth byte is ALPHA. It is reserved in an
   * ordinary blit — GDI ignores it and so must this, or a picture whose
   * producer left it zero comes out invisible — and it is the alpha channel
   * in an EMR_ALPHABLEND that says so.
   */
  readonly alpha?: boolean;
}

/**
 * Decode a packed DIB into RGBA.
 *
 * @param bytes The metafile.
 * @param bmiAt Where the bitmap's header starts.
 * @param o     What the record says about the bitmap; see {@link DibOptions}.
 * @returns The bitmap, or `undefined` when the header is malformed or its
 *          compression is one this decoder does not do.
 */
export function readDib(
  bytes: Uint8Array,
  bmiAt: number,
  o: DibOptions = {},
): DibImage | undefined {
  const { bitsAt, cbBits } = o;
  if (bmiAt < 0 || bmiAt + 12 > bytes.length) return undefined;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerSize = v.getUint32(bmiAt, true);
  // §2.2.2.9 BitmapCoreHeader is the OS/2 spelling: 16-bit extents and a
  // three-byte palette entry. Anything else is a BitmapInfoHeader or one of
  // its extensions, which only add fields past the ones read here.
  const core = headerSize === 12;
  if (!core && (headerSize < 40 || bmiAt + 40 > bytes.length)) return undefined;
  const width = core ? v.getInt16(bmiAt + 4, true) : v.getInt32(bmiAt + 4, true);
  const rawHeight = core ? v.getInt16(bmiAt + 6, true) : v.getInt32(bmiAt + 8, true);
  const bpp = core ? v.getUint16(bmiAt + 10, true) : v.getUint16(bmiAt + 14, true);
  const compression = core ? 0 : v.getUint32(bmiAt + 16, true);
  // A NEGATIVE height means the rows are stored top-down, the way everything
  // else in this file thinks of them.
  const topDown = rawHeight < 0;
  const height = Math.abs(rawHeight);
  if (width <= 0 || height <= 0 || width * height > MAX_PIXELS) return undefined;
  if (compression !== 0 && compression !== 3) return undefined;
  if (bpp !== 1 && bpp !== 4 && bpp !== 8 && bpp !== 16 && bpp !== 24 && bpp !== 32) {
    return undefined;
  }

  const paletteAt = bmiAt + headerSize;
  const entry = core ? 3 : 4;
  // §2.2.2.9 `ColorUsed` — how many of the table's entries the bitmap uses; 0
  // means all of them. BI_BITFIELDS spends the same space on channel masks.
  const declared = core || headerSize < 40 ? 0 : v.getUint32(bmiAt + 32, true);
  const paletteCount = bpp <= 8 ? (declared > 0 ? Math.min(declared, 1 << bpp) : 1 << bpp) : 0;
  const masksAt = compression === 3 ? paletteAt : undefined;
  const paletteBytes = compression === 3 ? 12 : paletteCount * entry;
  const start = bitsAt ?? paletteAt + paletteBytes;
  const stride = (((width * bpp + 7) >> 3) + 3) & ~3;
  const need = stride * height;
  if (start < 0 || start + Math.min(need, cbBits ?? need) > bytes.length) return undefined;

  const palette = readPalette(v, paletteAt, paletteCount, entry);
  const channels =
    masksAt !== undefined
      ? [
          maskChannel(v.getUint32(masksAt, true)),
          maskChannel(v.getUint32(masksAt + 4, true)),
          maskChannel(v.getUint32(masksAt + 8, true)),
        ]
      : undefined;

  const rgba = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    const row = start + (topDown ? y : height - 1 - y) * stride;
    let dst = y * width * 4;
    for (let x = 0; x < width; x++) {
      let r = 0;
      let g = 0;
      let b = 0;
      let a = 255;
      if (bpp <= 8) {
        const bitsPer = bpp;
        const bit = x * bitsPer;
        const byte = bytes[row + (bit >> 3)] ?? 0;
        const shift = 8 - bitsPer - (bit & 7);
        const index = (byte >> shift) & ((1 << bitsPer) - 1);
        const rgb = palette[index] ?? 0;
        r = (rgb >>> 16) & 0xff;
        g = (rgb >>> 8) & 0xff;
        b = rgb & 0xff;
      } else if (bpp === 16) {
        const word = (bytes[row + x * 2] ?? 0) | ((bytes[row + x * 2 + 1] ?? 0) << 8);
        if (channels) {
          [r, g, b] = [channels[0]!(word), channels[1]!(word), channels[2]!(word)];
        } else {
          // BI_RGB at 16 bits is X1R5G5B5 — five bits a channel, scaled up so
          // that a full channel comes out white rather than 248/255 grey.
          r = expand5((word >> 10) & 31);
          g = expand5((word >> 5) & 31);
          b = expand5(word & 31);
        }
      } else {
        const at = row + x * (bpp >> 3);
        b = bytes[at] ?? 0;
        g = bytes[at + 1] ?? 0;
        r = bytes[at + 2] ?? 0;
        if (bpp === 32) {
          const word = ((b | (g << 8) | (r << 16) | ((bytes[at + 3] ?? 0) << 24)) >>> 0) >>> 0;
          if (channels) [r, g, b] = [channels[0]!(word), channels[1]!(word), channels[2]!(word)];
          // The fourth byte is the alpha channel only when the record that
          // carries the bitmap says it is (see DibOptions).
          if (o.alpha === true) {
            a = bytes[at + 3] ?? 0;
            // AC_SRC_ALPHA states PREMULTIPLIED colour; PNG wants it straight.
            if (a > 0 && a < 255) {
              r = Math.min(255, Math.round((r * 255) / a));
              g = Math.min(255, Math.round((g * 255) / a));
              b = Math.min(255, Math.round((b * 255) / a));
            }
          }
        }
      }
      rgba[dst] = r;
      rgba[dst + 1] = g;
      rgba[dst + 2] = b;
      rgba[dst + 3] = a;
      dst += 4;
    }
  }
  return { width, height, rgba, ...(topDown ? {} : { bottomUp: true }) };
}

/**
 * The colour bitmap seen through a monochrome one — the two-blit idiom every
 * clipart of this age is drawn with: an AND of the mask, which knocks the
 * picture's ground to white, and then an OR of the picture, which is black
 * exactly where the ground was. Kept apart they draw a black box; put together
 * they are one bitmap with an alpha channel.
 *
 * @param color The bitmap the second blit paints.
 * @param mask  The first blit's mask: WHITE where the picture is to show
 *              through, black where it is opaque.
 */
export function maskedDib(color: DibImage, mask: DibImage): DibImage {
  const rgba = new Uint8Array(color.rgba);
  for (let y = 0; y < color.height; y++) {
    // The two are the same size in every file that draws this way, but a
    // stretched blit need not be: sample the mask by relative position.
    const my = Math.min(mask.height - 1, Math.floor((y * mask.height) / color.height));
    for (let x = 0; x < color.width; x++) {
      const mx = Math.min(mask.width - 1, Math.floor((x * mask.width) / color.width));
      const m = my * mask.width * 4 + mx * 4;
      const lum = (mask.rgba[m]! + mask.rgba[m + 1]! + mask.rgba[m + 2]!) / 3;
      if (lum >= 128) rgba[(y * color.width + x) * 4 + 3] = 0;
    }
  }
  return { width: color.width, height: color.height, rgba };
}

/**
 * The bitmap as a PNG file, ready for the resource store. An opaque one is
 * written without an alpha channel — a quarter less to deflate, and the
 * writers that flatten transparency then have nothing to do.
 */
export function dibToPng(img: DibImage): Uint8Array {
  let opaque = true;
  for (let i = 3; i < img.rgba.length; i += 4) {
    if (img.rgba[i] !== 255) {
      opaque = false;
      break;
    }
  }
  if (!opaque) return encodePng(img.width, img.height, 'rgba', img.rgba);
  const rgb = new Uint8Array(img.width * img.height * 3);
  for (let i = 0, j = 0; i < img.rgba.length; i += 4, j += 3) {
    rgb[j] = img.rgba[i]!;
    rgb[j + 1] = img.rgba[i + 1]!;
    rgb[j + 2] = img.rgba[i + 2]!;
  }
  return encodePng(img.width, img.height, 'rgb', rgb);
}

/** The bitmap at a constant transparency — what a blend function's `SrcConstantAlpha` asks for. */
export function fadedDib(img: DibImage, alpha: number): DibImage {
  const rgba = new Uint8Array(img.rgba);
  for (let i = 3; i < rgba.length; i += 4) rgba[i] = Math.round(rgba[i]! * alpha);
  return { width: img.width, height: img.height, rgba };
}

/** The part of the bitmap a blit's source rectangle names, or all of it. */
export function cropDib(img: DibImage, x: number, y: number, w: number, h: number): DibImage {
  const left = Math.max(0, Math.min(img.width - 1, x));
  const top = Math.max(0, Math.min(img.height - 1, y));
  const width = Math.max(1, Math.min(img.width - left, w));
  const height = Math.max(1, Math.min(img.height - top, h));
  if (left === 0 && top === 0 && width === img.width && height === img.height) return img;
  const rgba = new Uint8Array(width * height * 4);
  for (let row = 0; row < height; row++) {
    const src = ((top + row) * img.width + left) * 4;
    rgba.set(img.rgba.subarray(src, src + width * 4), row * width * 4);
  }
  return { width, height, rgba };
}

// The colour table as packed 0xRRGGBB, in the order the bitmap indexes it.
function readPalette(v: DataView, at: number, count: number, entry: number): Uint32Array {
  const out = new Uint32Array(Math.max(count, 2));
  // With no table of its own a 1-bit bitmap is black and white, which is what
  // a mask always is.
  out[1] = 0xffffff;
  for (let i = 0; i < count; i++) {
    const o = at + i * entry;
    if (o + 3 > v.byteLength) break;
    out[i] = (v.getUint8(o + 2) << 16) | (v.getUint8(o + 1) << 8) | v.getUint8(o);
  }
  return out;
}

// One BI_BITFIELDS channel: pull the masked bits down and scale them to 0..255.
function maskChannel(mask: number): (word: number) => number {
  if (mask === 0) return () => 0;
  let shift = 0;
  while (((mask >>> shift) & 1) === 0) shift++;
  const span = (mask >>> shift) + 1;
  return (word) => Math.round((((word & mask) >>> shift) * 255) / (span - 1));
}

// Five bits of a channel as eight — 31 has to come out 255, not 248.
function expand5(v: number): number {
  return (v << 3) | (v >> 2);
}
