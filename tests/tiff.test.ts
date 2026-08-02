import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { prepareImage } from '@/core/images';
import { decodeTiff, isTiff } from '@/core/tiff';

// A TIFF is a header, an image directory and the strips the directory points
// at (TIFF 6.0 §2). This builder writes exactly that: the tags, sorted as the
// specification requires, then the image data after them.

const TAG = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
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
} as const;

interface Tag {
  readonly tag: number;
  readonly type: 3 | 4;
  readonly values: ReadonlyArray<number>;
}

const short = (tag: number, ...values: Array<number>): Tag => ({ tag, type: 3, values });
const long = (tag: number, ...values: Array<number>): Tag => ({ tag, type: 4, values });

/** Assemble a TIFF from its tags and one blob of image data. */
function tiff(
  tags: ReadonlyArray<Tag>,
  data: Uint8Array,
  options: { be?: boolean } = {},
): Uint8Array {
  const be = options.be ?? false;
  // Layout: header (8) | IFD | out-of-line tag values | image data.
  const sorted = [...tags].sort((a, b) => a.tag - b.tag);
  const ifdSize = 2 + sorted.length * 12 + 4;
  const inline = (t: Tag): boolean => (t.type === 3 ? t.values.length <= 2 : t.values.length <= 1);
  let extra = 0;
  for (const t of sorted) if (!inline(t)) extra += t.values.length * (t.type === 3 ? 2 : 4);
  const dataAt = 8 + ifdSize + extra;
  const buf = new Uint8Array(dataAt + data.length);
  const view = new DataView(buf.buffer);
  const u16 = (o: number, v: number): void => view.setUint16(o, v, !be);
  const u32 = (o: number, v: number): void => view.setUint32(o, v, !be);

  buf[0] = be ? 0x4d : 0x49;
  buf[1] = be ? 0x4d : 0x49;
  u16(2, 42);
  u32(4, 8);
  u16(8, sorted.length);
  let entry = 10;
  let extraAt = 8 + ifdSize;
  for (const t of sorted) {
    u16(entry, t.tag);
    u16(entry + 2, t.type);
    u32(entry + 4, t.values.length);
    if (inline(t)) {
      // A value that fits in the four value bytes sits LEFT-aligned in them.
      t.values.forEach((v, i) => {
        if (t.type === 3) u16(entry + 8 + i * 2, v);
        else u32(entry + 8, v);
      });
    } else {
      u32(entry + 8, extraAt);
      t.values.forEach((v, i) => {
        if (t.type === 3) u16(extraAt + i * 2, v);
        else u32(extraAt + i * 4, v);
      });
      extraAt += t.values.length * (t.type === 3 ? 2 : 4);
    }
    entry += 12;
  }
  buf.set(data, dataAt);
  return buf;
}

/** The offset the image data lands at for a given tag list — for stripOffsets. */
function dataOffset(tags: ReadonlyArray<Tag>): number {
  const ifdSize = 2 + tags.length * 12 + 4;
  let extra = 0;
  for (const t of tags) {
    const inline = t.type === 3 ? t.values.length <= 2 : t.values.length <= 1;
    if (!inline) extra += t.values.length * (t.type === 3 ? 2 : 4);
  }
  return 8 + ifdSize + extra;
}

/**
 * LZW (§13) that only ever emits literals: Clear, one code per byte, EOD. Under
 * 253 bytes the codes stay nine bits wide, which keeps the encoder honest and
 * still drives the decoder's table, widening and KwKwK-free path.
 */
function lzwLiterals(bytes: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let acc = 0;
  let bits = 0;
  const push = (code: number): void => {
    acc = (acc << 9) | code;
    bits += 9;
    while (bits >= 8) {
      bits -= 8;
      out.push((acc >> bits) & 0xff);
    }
  };
  push(256);
  for (const b of bytes) push(b);
  push(257);
  if (bits > 0) out.push((acc << (8 - bits)) & 0xff);
  return Uint8Array.from(out);
}

/** PackBits (§9) as pure literal runs — the encoding a flat encoder produces. */
function packBitsLiterals(bytes: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  for (let i = 0; i < bytes.length; i += 128) {
    const chunk = bytes.subarray(i, i + 128);
    out.push(chunk.length - 1, ...chunk);
  }
  return Uint8Array.from(out);
}

// A 2×2 RGB picture: red, green / blue, white.
const RGB_2X2 = Uint8Array.of(255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255);

function rgbTags(compression: number, byteCount: number, extra: Array<Tag> = []): Array<Tag> {
  const tags = [
    short(TAG.width, 2),
    short(TAG.height, 2),
    short(TAG.bitsPerSample, 8, 8, 8),
    short(TAG.compression, compression),
    short(TAG.photometric, 2),
    short(TAG.samplesPerPixel, 3),
    short(TAG.rowsPerStrip, 2),
    long(TAG.stripByteCounts, byteCount),
    short(TAG.planarConfig, 1),
    ...extra,
    long(TAG.stripOffsets, 0),
  ];
  // stripOffsets must name where the data actually lands, which depends on the
  // tag list it is part of — so it is filled in once the list is complete.
  return tags.map((t) =>
    t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, dataOffset(tags)) : t,
  );
}

describe('baseline TIFF (TIFF 6.0)', () => {
  it('knows a TIFF by its header, either byte order', () => {
    expect(isTiff(Uint8Array.of(0x49, 0x49, 42, 0, 8, 0, 0, 0))).toBe(true);
    expect(isTiff(Uint8Array.of(0x4d, 0x4d, 0, 42, 0, 0, 0, 8))).toBe(true);
    expect(isTiff(Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 13, 10, 26, 10))).toBe(false);
  });

  it('reads an uncompressed RGB strip', () => {
    const decoded = decodeTiff(tiff(rgbTags(1, RGB_2X2.length), RGB_2X2));
    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(2);
    expect(decoded.channels).toBe(3);
    expect([...decoded.data]).toEqual([...RGB_2X2]);
  });

  it('reads the same picture written big-endian', () => {
    const decoded = decodeTiff(tiff(rgbTags(1, RGB_2X2.length), RGB_2X2, { be: true }));
    expect([...decoded.data]).toEqual([...RGB_2X2]);
  });

  it('inflates an LZW strip (§13)', () => {
    const lzw = lzwLiterals(RGB_2X2);
    expect([...decodeTiff(tiff(rgbTags(5, lzw.length), lzw)).data]).toEqual([...RGB_2X2]);
  });

  it('undoes horizontal differencing (§14 Predictor 2)', () => {
    // Each row stores the first pixel then the DIFFERENCE to the one before it,
    // per channel — tdf115094.docx's logos are written this way.
    const diffed = Uint8Array.of(255, 0, 0, 1, 255, 0, 0, 0, 255, 255, 255, 1);
    const lzw = lzwLiterals(diffed);
    const decoded = decodeTiff(tiff(rgbTags(5, lzw.length, [short(TAG.predictor, 2)]), lzw));
    expect([...decoded.data]).toEqual([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0]);
  });

  it('unpacks a PackBits strip (§9)', () => {
    const packed = packBitsLiterals(RGB_2X2);
    expect([...decodeTiff(tiff(rgbTags(32773, packed.length), packed)).data]).toEqual([...RGB_2X2]);
  });

  it('assembles the strips of a multi-strip picture in order', () => {
    // Two strips of one row each: the second must land BELOW the first.
    const base = [
      short(TAG.width, 2),
      short(TAG.height, 2),
      short(TAG.bitsPerSample, 8, 8, 8),
      short(TAG.compression, 1),
      short(TAG.photometric, 2),
      short(TAG.samplesPerPixel, 3),
      short(TAG.rowsPerStrip, 1),
      long(TAG.stripByteCounts, 6, 6),
      short(TAG.planarConfig, 1),
      long(TAG.stripOffsets, 0, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) =>
      t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, at, at + 6) : t,
    );
    expect([...decodeTiff(tiff(tags, RGB_2X2)).data]).toEqual([...RGB_2X2]);
  });

  it('crops the padding a tile carries past the image edge (§15)', () => {
    // One 2×2 tile over a 1×1 image: only the top-left pixel is the picture.
    const base = [
      short(TAG.width, 1),
      short(TAG.height, 1),
      short(TAG.bitsPerSample, 8, 8, 8),
      short(TAG.compression, 1),
      short(TAG.photometric, 2),
      short(TAG.samplesPerPixel, 3),
      short(TAG.planarConfig, 1),
      short(TAG.tileWidth, 2),
      short(TAG.tileLength, 2),
      long(TAG.tileByteCounts, 12),
      long(TAG.tileOffsets, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) => (t.tag === TAG.tileOffsets ? long(TAG.tileOffsets, at) : t));
    expect([...decodeTiff(tiff(tags, RGB_2X2)).data]).toEqual([255, 0, 0]);
  });

  it('expands a palette against its colour map (§8 photometric 3)', () => {
    // Four-bit indices, two pixels per byte; the map holds all reds, then all
    // greens, then all blues, each 0…65535.
    const map = [0xffff, 0, 0, 0, 0, 0xffff, 0, 0, 0, 0, 0, 0]; // all reds, all greens, all blues
    const base = [
      short(TAG.width, 2),
      short(TAG.height, 1),
      short(TAG.bitsPerSample, 4),
      short(TAG.compression, 1),
      short(TAG.photometric, 3),
      short(TAG.samplesPerPixel, 1),
      short(TAG.rowsPerStrip, 1),
      long(TAG.stripByteCounts, 1),
      short(TAG.planarConfig, 1),
      short(TAG.colorMap, ...map),
      long(TAG.stripOffsets, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) => (t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, at) : t));
    // Index 0 then index 1 — the map above makes those red and green.
    const decoded = decodeTiff(tiff(tags, Uint8Array.of(0x01)));
    expect([...decoded.data]).toEqual([255, 0, 0, 0, 255, 0]);
  });

  it('inverts WhiteIsZero grey into the DeviceGray a PDF paints (§8 photometric 0)', () => {
    const base = [
      short(TAG.width, 2),
      short(TAG.height, 1),
      short(TAG.bitsPerSample, 1),
      short(TAG.compression, 1),
      short(TAG.photometric, 0),
      short(TAG.samplesPerPixel, 1),
      short(TAG.rowsPerStrip, 1),
      long(TAG.stripByteCounts, 1),
      short(TAG.planarConfig, 1),
      long(TAG.stripOffsets, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) => (t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, at) : t));
    // Bits 0 then 1: white then black, which DeviceGray states the other way.
    const decoded = decodeTiff(tiff(tags, Uint8Array.of(0b01000000)));
    expect(decoded.channels).toBe(1);
    expect([...decoded.data]).toEqual([255, 0]);
  });

  it('splits an ExtraSamples alpha channel off the colour (§8)', () => {
    const rgba = Uint8Array.of(255, 0, 0, 128, 0, 255, 0, 255);
    const base = [
      short(TAG.width, 2),
      short(TAG.height, 1),
      short(TAG.bitsPerSample, 8, 8, 8, 8),
      short(TAG.compression, 1),
      short(TAG.photometric, 2),
      short(TAG.samplesPerPixel, 4),
      short(TAG.rowsPerStrip, 1),
      long(TAG.stripByteCounts, rgba.length),
      short(TAG.planarConfig, 1),
      short(TAG.extraSamples, 2),
      long(TAG.stripOffsets, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) => (t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, at) : t));
    const decoded = decodeTiff(tiff(tags, rgba));
    expect([...decoded.data]).toEqual([255, 0, 0, 0, 255, 0]);
    expect([...(decoded.alpha ?? [])]).toEqual([128, 255]);
  });

  it('refuses a planar-separated file rather than drawing it wrong', () => {
    const tags = rgbTags(1, RGB_2X2.length).map((t) =>
      t.tag === TAG.planarConfig ? short(TAG.planarConfig, 2) : t,
    );
    expect(() => decodeTiff(tiff(tags, RGB_2X2))).toThrow(/planar/iu);
  });

  it('refuses a compression it cannot read', () => {
    expect(() => decodeTiff(tiff(rgbTags(7, RGB_2X2.length), RGB_2X2))).toThrow(/compression/iu);
  });
});

describe('a TIFF prepared for embedding', () => {
  it('arrives as Flate-compressed samples — PDF has no TIFF filter', () => {
    const prepared = prepareImage(tiff(rgbTags(1, RGB_2X2.length), RGB_2X2));
    expect(prepared.format).toBe('tiff');
    expect(prepared.mimeType).toBe('image/tiff');
    expect(prepared.filter).toBe('FlateDecode');
    expect(prepared.colorSpace).toBe('DeviceRGB');
    expect(prepared.bitsPerComponent).toBe(8);
    expect(prepared.widthPx).toBe(2);
    expect(prepared.heightPx).toBe(2);
    expect([...unzlibSync(prepared.data)]).toEqual([...RGB_2X2]);
  });

  it('turns CMYK into the RGB an sRGB output intent can hold', () => {
    const cmyk = Uint8Array.of(0, 255, 255, 0, 0, 0, 0, 255); // red, then black
    const base = [
      short(TAG.width, 2),
      short(TAG.height, 1),
      short(TAG.bitsPerSample, 8, 8, 8, 8),
      short(TAG.compression, 1),
      short(TAG.photometric, 5),
      short(TAG.samplesPerPixel, 4),
      short(TAG.rowsPerStrip, 1),
      long(TAG.stripByteCounts, cmyk.length),
      short(TAG.planarConfig, 1),
      long(TAG.stripOffsets, 0),
    ];
    const at = dataOffset(base);
    const tags = base.map((t) => (t.tag === TAG.stripOffsets ? long(TAG.stripOffsets, at) : t));
    const prepared = prepareImage(tiff(tags, cmyk));
    expect(prepared.colorSpace).toBe('DeviceRGB');
    expect([...unzlibSync(prepared.data)]).toEqual([255, 0, 0, 0, 0, 0]);
  });
});
