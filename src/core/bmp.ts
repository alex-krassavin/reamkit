// Windows bitmaps (BMP / DIB) → 8-bit RGB samples.
//
// A BMP file is a 14-byte BITMAPFILEHEADER, a DIB header, an optional palette
// and then the rows — bottom-up unless the height is negative. Everything
// interesting is in the DIB header, and there are five of them: the 12-byte
// OS/2 core header, the 40-byte BITMAPINFOHEADER every writer still emits, and
// the V2/V3/V4/V5 extensions that only add fields on the end. They are read as
// one, by length.
//
// The reason this module exists is the `.ppt` and `.xls` picture stores: a
// `BlipDIB` record (MS-ODRAW §2.2.28) holds a bitmap with the FILE header cut
// off, because the record already says how long it is. Nothing downstream can
// sniff such a payload — it starts with a length field, not a signature — so
// {@link dibToBmp} puts the missing header back and the ordinary raster path
// takes it from there.

/** One decoded bitmap: chunky 8-bit samples, plus alpha when the file carries it. */
export interface DecodedBmp {
  readonly width: number;
  readonly height: number;
  /** Row-major RGB samples, `width * height * 3` long. */
  readonly data: Uint8Array;
  /** One byte per pixel, present only when the file states real transparency. */
  readonly alpha?: Uint8Array;
  /** `biXPelsPerMeter`/`biYPelsPerMeter` turned into pixels per inch. */
  readonly dpiX?: number;
  readonly dpiY?: number;
}

const FILE_HEADER_BYTES = 14;
const CORE_HEADER_BYTES = 12;
const INFO_HEADER_BYTES = 40;
const MAX_PIXELS = 40_000_000; // DoS guard (~40 MP), as the other raster paths use

// §BITMAPINFOHEADER `biCompression`.
const BI_RGB = 0;
const BI_RLE8 = 1;
const BI_RLE4 = 2;
const BI_BITFIELDS = 3;
const BI_ALPHABITFIELDS = 6;

/** Whether these bytes open with the `BM` signature of a bitmap FILE. */
export function isBmp(bytes: Uint8Array): boolean {
  return bytes.length >= FILE_HEADER_BYTES && bytes[0] === 0x42 && bytes[1] === 0x4d;
}

/**
 * MS-ODRAW §2.2.28 — a `BlipDIB`'s payload is a bitmap without its 14-byte
 * BITMAPFILEHEADER, which the Escher record makes redundant. Put it back.
 *
 * @param dib The bytes from the DIB header onwards.
 * @returns A complete BMP file, or `undefined` when the header does not read
 *          as a bitmap at all.
 */
export function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  const head = readDibHeader(dib, 0);
  if (!head) return undefined;
  const offBits = FILE_HEADER_BYTES + head.headerBytes + head.maskBytes + head.paletteBytes;
  const out = new Uint8Array(FILE_HEADER_BYTES + dib.length);
  const view = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  view.setUint32(2, out.length, true);
  view.setUint32(10, offBits, true);
  out.set(dib, FILE_HEADER_BYTES);
  return out;
}

/**
 * Decode a Windows bitmap into 8-bit RGB samples.
 *
 * @param bytes A complete BMP file (`BM` signature included).
 * @returns The samples, their size and any alpha channel.
 * @throws Error when the file is malformed, or carries a bitmap this does not
 *         read — a JPEG or PNG smuggled inside a DIB (`BI_JPEG`/`BI_PNG`),
 *         which is a whole other file wearing a bitmap header.
 */
export function decodeBmp(bytes: Uint8Array): DecodedBmp {
  if (!isBmp(bytes)) throw new Error('BMP: not a bitmap');
  const file = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const head = readDibHeader(bytes, FILE_HEADER_BYTES);
  if (!head) throw new Error('BMP: unreadable header');
  const { width, height, topDown, bitCount, compression } = head;
  if (width <= 0 || height <= 0) throw new Error('BMP: empty bitmap');
  if (width * height > MAX_PIXELS) throw new Error('BMP: image too large');
  if (compression === 4 || compression === 5) {
    throw new Error('BMP: embedded JPEG/PNG bitmaps are not supported');
  }
  // §BITMAPFILEHEADER `bfOffBits` names where the rows start. A writer that
  // leaves it at zero means "right after the palette", which is where the
  // header arithmetic puts them anyway.
  const stated = file.getUint32(10, true);
  const computed = FILE_HEADER_BYTES + head.headerBytes + head.maskBytes + head.paletteBytes;
  const pixelsAt = stated >= computed && stated < bytes.length ? stated : computed;
  const palette = readPalette(bytes, FILE_HEADER_BYTES + head.headerBytes, head);

  const count = width * height;
  const data = new Uint8Array(count * 3);
  const rowOf = (y: number): number => (topDown ? y : height - 1 - y);

  if (compression === BI_RLE8 || compression === BI_RLE4) {
    const indices = decodeRle(bytes.subarray(pixelsAt), width, height, compression === BI_RLE4);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        // An RLE bitmap is stored bottom-up like any other, and the runs that
        // never reached a pixel leave it at palette entry 0.
        const at = (rowOf(y) * width + x) * 3;
        const p = (indices[y * width + x] ?? 0) * 3;
        data[at] = palette[p] ?? 0;
        data[at + 1] = palette[p + 1] ?? 0;
        data[at + 2] = palette[p + 2] ?? 0;
      }
    }
    return { width, height, data, ...density(head) };
  }

  // Rows are padded to a four-byte boundary, always.
  const stride = (((width * bitCount + 31) / 32) | 0) * 4;
  const alpha = head.alphaMask !== 0 ? new Uint8Array(count) : undefined;
  const shifts = {
    r: maskShift(head.redMask),
    g: maskShift(head.greenMask),
    b: maskShift(head.blueMask),
    a: maskShift(head.alphaMask),
  };
  for (let y = 0; y < height; y++) {
    const row = pixelsAt + y * stride;
    if (row + stride > bytes.length + stride) break;
    for (let x = 0; x < width; x++) {
      const at = (rowOf(y) * width + x) * 3;
      if (bitCount <= 8) {
        const p = paletteIndex(bytes, row, x, bitCount) * 3;
        data[at] = palette[p] ?? 0;
        data[at + 1] = palette[p + 1] ?? 0;
        data[at + 2] = palette[p + 2] ?? 0;
        continue;
      }
      if (bitCount === 24) {
        // §RGBTRIPLE — blue, green, red, in that order.
        const o = row + x * 3;
        data[at] = bytes[o + 2] ?? 0;
        data[at + 1] = bytes[o + 1] ?? 0;
        data[at + 2] = bytes[o] ?? 0;
        continue;
      }
      const raw =
        bitCount === 16
          ? (bytes[row + x * 2] ?? 0) | ((bytes[row + x * 2 + 1] ?? 0) << 8)
          : ((bytes[row + x * 4] ?? 0) |
              ((bytes[row + x * 4 + 1] ?? 0) << 8) |
              ((bytes[row + x * 4 + 2] ?? 0) << 16) |
              ((bytes[row + x * 4 + 3] ?? 0) << 24)) >>>
            0;
      data[at] = channel(raw, head.redMask, shifts.r);
      data[at + 1] = channel(raw, head.greenMask, shifts.g);
      data[at + 2] = channel(raw, head.blueMask, shifts.b);
      if (alpha) alpha[rowOf(y) * width + x] = channel(raw, head.alphaMask, shifts.a);
    }
  }
  // A 32-bit bitmap whose alpha is zero everywhere is one whose writer left the
  // fourth byte alone — the field is reserved under `BI_RGB` — and honouring it
  // would paint the whole picture away.
  const opaque = alpha?.every((v) => v === 0) ?? true;
  return {
    width,
    height,
    data,
    ...(alpha && !opaque ? { alpha } : {}),
    ...density(head),
  };
}

/** What the header says about the bitmap, whichever of the five it is. */
interface DibHeader {
  readonly headerBytes: number;
  readonly maskBytes: number;
  readonly paletteBytes: number;
  readonly paletteEntryBytes: 3 | 4;
  readonly width: number;
  readonly height: number;
  readonly topDown: boolean;
  readonly bitCount: number;
  readonly compression: number;
  readonly redMask: number;
  readonly greenMask: number;
  readonly blueMask: number;
  readonly alphaMask: number;
  readonly pxPerMeterX: number;
  readonly pxPerMeterY: number;
}

function readDibHeader(bytes: Uint8Array, at: number): DibHeader | undefined {
  if (bytes.length < at + 4) return undefined;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerBytes = v.getUint32(at, true);
  const core = headerBytes === CORE_HEADER_BYTES;
  if (!core && headerBytes < INFO_HEADER_BYTES) return undefined;
  if (bytes.length < at + headerBytes) return undefined;
  const width = core ? v.getInt16(at + 4, true) : v.getInt32(at + 4, true);
  const rawHeight = core ? v.getInt16(at + 6, true) : v.getInt32(at + 8, true);
  const planes = core ? v.getUint16(at + 6, true) : v.getUint16(at + 12, true);
  const bitCount = core ? v.getUint16(at + 8, true) : v.getUint16(at + 14, true);
  const compression = core ? BI_RGB : v.getUint32(at + 16, true);
  // Enough of the header to be sure this IS one: a `BlipDIB` is probed at two
  // offsets because the tag byte before it is not always written, and a header
  // that reads as nonsense at the first is how the second gets its turn.
  if (planes !== 1) return undefined;
  if (![1, 2, 4, 8, 16, 24, 32].includes(bitCount)) return undefined;
  if (compression > 6) return undefined;
  if (width <= 0 || rawHeight === 0) return undefined;
  const clrUsed = core ? 0 : v.getUint32(at + 32, true);
  // §BI_BITFIELDS on a 40-byte header puts the masks AFTER it, in the space the
  // palette would occupy; a longer header carries them as fields of its own.
  const inHeader = headerBytes >= 52;
  const maskBytes =
    !inHeader && (compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS)
      ? compression === BI_ALPHABITFIELDS
        ? 16
        : 12
      : 0;
  const maskAt = inHeader ? at + 40 : at + headerBytes;
  const stated = compression === BI_BITFIELDS || compression === BI_ALPHABITFIELDS || inHeader;
  const readMask = (i: number): number =>
    stated && bytes.length >= maskAt + (i + 1) * 4 ? v.getUint32(maskAt + i * 4, true) : 0;
  const defaults = bitCount === 16 ? [0x7c00, 0x03e0, 0x001f] : [0xff0000, 0xff00, 0x00ff];
  const red = readMask(0) || defaults[0]!;
  const green = readMask(1) || defaults[1]!;
  const blue = readMask(2) || defaults[2]!;
  // The alpha mask is a V3-and-later field; a V2 or a plain BITMAPINFOHEADER
  // has no way to say a bitmap is translucent, whatever its fourth byte holds.
  const alphaMask =
    headerBytes >= 56 && bytes.length >= maskAt + 16 ? v.getUint32(maskAt + 12, true) : 0;
  const paletteEntryBytes = core ? 3 : 4;
  const entries = bitCount <= 8 ? clrUsed || 1 << bitCount : 0;
  return {
    headerBytes,
    maskBytes,
    paletteBytes: entries * paletteEntryBytes,
    paletteEntryBytes,
    width,
    height: Math.abs(rawHeight),
    topDown: rawHeight < 0,
    bitCount,
    compression,
    redMask: red,
    greenMask: green,
    blueMask: blue,
    alphaMask,
    pxPerMeterX: core ? 0 : v.getInt32(at + 24, true),
    pxPerMeterY: core ? 0 : v.getInt32(at + 28, true),
  };
}

/** The palette as flat RGB triples, whichever entry width the header uses. */
function readPalette(bytes: Uint8Array, at: number, head: DibHeader): Uint8Array {
  const entries = head.paletteBytes / head.paletteEntryBytes;
  const out = new Uint8Array(entries * 3);
  for (let i = 0; i < entries; i++) {
    const o = at + i * head.paletteEntryBytes;
    // §RGBQUAD — blue, green, red (and a reserved byte on the four-byte form).
    out[i * 3] = bytes[o + 2] ?? 0;
    out[i * 3 + 1] = bytes[o + 1] ?? 0;
    out[i * 3 + 2] = bytes[o] ?? 0;
  }
  return out;
}

/** The palette index of one pixel of a 1/2/4/8-bit row. */
function paletteIndex(bytes: Uint8Array, row: number, x: number, bitCount: number): number {
  if (bitCount === 8) return bytes[row + x] ?? 0;
  const perByte = 8 / bitCount;
  const byte = bytes[row + ((x / perByte) | 0)] ?? 0;
  const shift = 8 - bitCount * ((x % perByte) + 1);
  return (byte >> shift) & ((1 << bitCount) - 1);
}

/** How far right a mask's field sits, so its bits can be read off. */
function maskShift(mask: number): number {
  if (mask === 0) return 0;
  let shift = 0;
  while (((mask >>> shift) & 1) === 0) shift++;
  return shift;
}

/** One channel of a packed pixel, stretched to the full byte range. */
function channel(raw: number, mask: number, shift: number): number {
  if (mask === 0) return 0;
  const width = (mask >>> shift).toString(2).length;
  const value = (raw & mask) >>> shift;
  const max = (1 << width) - 1;
  return max === 0 ? 0 : Math.round((value * 255) / max);
}

/** §biXPelsPerMeter — the resolution the bitmap claims, in pixels per inch. */
function density(head: DibHeader): { dpiX?: number; dpiY?: number } {
  if (head.pxPerMeterX <= 0 || head.pxPerMeterY <= 0) return {};
  return {
    dpiX: head.pxPerMeterX * 0.0254,
    dpiY: head.pxPerMeterY * 0.0254,
  };
}

/**
 * §BI_RLE8 / §BI_RLE4 — run-length rows of palette indices.
 *
 * Each pair is either a run (a non-zero count and the index to repeat) or an
 * escape (a zero count and a code): 0 ends the row, 1 ends the bitmap, 2 is a
 * delta that skips forward, and anything else is that many literal pixels,
 * padded to a two-byte boundary.
 *
 * @param data   The bytes from the first run onwards.
 * @param width  The bitmap's width in pixels.
 * @param height Its height in rows.
 * @param four   Whether this is the four-bit form, which packs two indices per byte.
 * @returns One palette index per pixel, top row first.
 */
function decodeRle(data: Uint8Array, width: number, height: number, four: boolean): Uint8Array {
  const out = new Uint8Array(width * height);
  let x = 0;
  let y = 0;
  let at = 0;
  const put = (index: number): void => {
    if (x < width && y < height) out[y * width + x] = index;
    x++;
  };
  while (at + 1 < data.length && y < height) {
    const count = data[at]!;
    const value = data[at + 1]!;
    at += 2;
    if (count > 0) {
      for (let i = 0; i < count; i++) {
        put(four ? (i % 2 === 0 ? value >> 4 : value & 0x0f) : value);
      }
      continue;
    }
    if (value === 0) {
      x = 0;
      y++;
      continue;
    }
    if (value === 1) break;
    if (value === 2) {
      x += data[at] ?? 0;
      y += data[at + 1] ?? 0;
      at += 2;
      continue;
    }
    // An absolute run: `value` literal pixels, then padding to an even length.
    const bytesUsed = four ? ((value + 1) / 2) | 0 : value;
    for (let i = 0; i < value; i++) {
      const byte = data[at + (four ? (i / 2) | 0 : i)] ?? 0;
      put(four ? (i % 2 === 0 ? byte >> 4 : byte & 0x0f) : byte);
    }
    at += bytesUsed + (bytesUsed % 2);
  }
  return out;
}
