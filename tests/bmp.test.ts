import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { decodeBmp, dibToBmp, isBmp } from '@/core/bmp';
import { detectImageFormat, prepareImage } from '@/core/images';
import { readEscherBlip } from '@/core/ole/escher-blip';

// A BMP is a 14-byte file header, a DIB header, an optional palette and then
// the rows — bottom-up unless the height is negative, each padded to four
// bytes. This builder writes exactly that, so a test can state a bitmap by its
// pixels and get back the file a writer would have produced.

interface BmpSpec {
  readonly width: number;
  readonly height: number;
  readonly bitCount: number;
  /** Flat BGRA palette entries, for the ≤ 8-bit forms. */
  readonly palette?: ReadonlyArray<number>;
  /** The pixel rows, BOTTOM row first; the builder pads each to four bytes. */
  readonly rows: ReadonlyArray<ReadonlyArray<number>>;
  readonly compression?: number;
  /** A negative height, which means the rows run top-down. */
  readonly topDown?: boolean;
  readonly headerBytes?: number;
  /** The four bit masks a `BI_BITFIELDS` bitmap carries. */
  readonly masks?: readonly [number, number, number, number];
  readonly pxPerMeter?: number;
}

function buildBmp(spec: BmpSpec): Uint8Array {
  const headerBytes = spec.headerBytes ?? 40;
  const palette = spec.palette ?? [];
  const masks = spec.masks;
  const maskBytes = masks && headerBytes === 40 ? 16 : 0;
  // Every row of a bitmap is padded to a four-byte boundary — the trap a
  // hand-written fixture falls into, so the builder does it.
  const stride =
    spec.compression !== undefined && spec.compression > 0
      ? 0
      : (((spec.width * spec.bitCount + 31) / 32) | 0) * 4;
  const pixels = spec.rows.flatMap((row) => [
    ...row,
    ...new Array<number>(Math.max(0, stride - row.length)).fill(0),
  ]);
  const offBits = 14 + headerBytes + maskBytes + palette.length;
  const out = new Uint8Array(offBits + pixels.length);
  const v = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  v.setUint32(2, out.length, true);
  v.setUint32(10, offBits, true);
  v.setUint32(14, headerBytes, true);
  v.setInt32(18, spec.width, true);
  v.setInt32(22, spec.topDown === true ? -spec.height : spec.height, true);
  v.setUint16(26, 1, true); // biPlanes
  v.setUint16(28, spec.bitCount, true);
  v.setUint32(30, spec.compression ?? 0, true);
  if (spec.pxPerMeter !== undefined) {
    v.setInt32(38, spec.pxPerMeter, true);
    v.setInt32(42, spec.pxPerMeter, true);
  }
  v.setUint32(46, spec.bitCount <= 8 ? palette.length / 4 : 0, true); // biClrUsed
  if (masks) {
    const at = headerBytes === 40 ? 14 + 40 : 14 + 40;
    masks.forEach((m, i) => v.setUint32(at + i * 4, m, true));
  }
  out.set(palette, 14 + headerBytes + maskBytes);
  out.set(pixels, offBits);
  return out;
}

/** The bitmap's pixels as `#`/`.` rows, top row first, by their red channel. */
function ink(bytes: Uint8Array): Array<string> {
  const img = decodeBmp(bytes);
  const rows: Array<string> = [];
  for (let y = 0; y < img.height; y++) {
    let row = '';
    for (let x = 0; x < img.width; x++) row += img.data[(y * img.width + x) * 3]! > 127 ? '#' : '.';
    rows.push(row);
  }
  return rows;
}

// Black and white, in the BGRA order a palette entry uses.
const MONO = [0, 0, 0, 0, 255, 255, 255, 0];

describe('Windows bitmaps (BMP / DIB)', () => {
  it('is sniffed by its signature, and only there', () => {
    expect(isBmp(new Uint8Array([0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]))).toBe(true);
    expect(isBmp(new Uint8Array([0x42, 0x4d]))).toBe(false); // too short to be a file
    expect(isBmp(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBe(false);
    expect(
      detectImageFormat(buildBmp({ width: 1, height: 1, bitCount: 24, rows: [[0, 0, 0, 0]] })),
    ).toBe('bmp');
  });

  it('reads the rows bottom-up, which is the way a bitmap stores them', () => {
    // Two rows of two pixels, the FIRST stored row being the BOTTOM one.
    const bmp = buildBmp({
      width: 2,
      height: 2,
      bitCount: 24,
      rows: [
        [0, 0, 0, 0, 0, 0], // bottom: black black
        [255, 255, 255, 255, 255, 255], // top: white white
      ],
    });
    expect(ink(bmp)).toEqual(['##', '..']);
    // …and a NEGATIVE height says the writer stored them the other way up.
    const flipped = buildBmp({
      width: 2,
      height: 2,
      bitCount: 24,
      topDown: true,
      rows: [
        [0, 0, 0, 0, 0, 0],
        [255, 255, 255, 255, 255, 255],
      ],
    });
    expect(ink(flipped)).toEqual(['..', '##']);
  });

  it('puts a 24-bit pixel back in RGB order', () => {
    // §RGBTRIPLE stores blue, green, red.
    const bmp = buildBmp({ width: 1, height: 1, bitCount: 24, rows: [[0x11, 0x22, 0x33, 0]] });
    expect([...decodeBmp(bmp).data]).toEqual([0x33, 0x22, 0x11]);
  });

  it('unpacks the palette forms — one, four and eight bits to a pixel', () => {
    // 1-bit: the high bit of the byte is the leftmost pixel.
    const one = buildBmp({
      width: 4,
      height: 1,
      bitCount: 1,
      palette: MONO,
      rows: [[0b1010_0000]],
    });
    // Palette entry 0 is black and entry 1 white, so the set bits are the ink.
    expect(ink(one)).toEqual(['#.#.']);
    // 4-bit: two pixels to a byte, high nibble first.
    const four = buildBmp({
      width: 4,
      height: 1,
      bitCount: 4,
      palette: MONO,
      rows: [[0x01, 0x10]],
    });
    expect(ink(four)).toEqual(['.##.']);
    // 8-bit: one index per byte.
    const eight = buildBmp({
      width: 4,
      height: 1,
      bitCount: 8,
      palette: MONO,
      rows: [[1, 0, 0, 1]],
    });
    expect(ink(eight)).toEqual(['#..#']);
  });

  it('reads a packed pixel through the masks its header states', () => {
    // §BI_BITFIELDS with 5-6-5, which is not a byte boundary anywhere.
    const bmp = buildBmp({
      width: 1,
      height: 1,
      bitCount: 16,
      compression: 3,
      masks: [0xf800, 0x07e0, 0x001f, 0],
      rows: [[0x00, 0xf8]], // red all set, green and blue clear
    });
    expect([...decodeBmp(bmp).data]).toEqual([255, 0, 0]);
  });

  it('leaves a 32-bit bitmap opaque when its fourth byte is reserved', () => {
    // Under `BI_RGB` the fourth byte has no meaning, and every writer that
    // leaves it at zero would otherwise paint the whole picture away.
    const bmp = buildBmp({
      width: 1,
      height: 1,
      bitCount: 32,
      rows: [[0x33, 0x22, 0x11, 0x00]],
    });
    const img = decodeBmp(bmp);
    expect([...img.data]).toEqual([0x11, 0x22, 0x33]);
    expect(img.alpha).toBeUndefined();
  });

  it('unrolls a run-length bitmap, rows and all', () => {
    // §BI_RLE8: a run is a count and an index; a zero count is an escape —
    // 0 ends the row and 1 ends the bitmap.
    const bmp = buildBmp({
      width: 4,
      height: 2,
      bitCount: 8,
      compression: 1,
      palette: MONO,
      rows: [[4, 1, 0, 0, 2, 1, 2, 0, 0, 1]],
    });
    // The first stored row is the bottom one, as always.
    expect(ink(bmp)).toEqual(['##..', '####']);
  });

  it('states the resolution the header claims', () => {
    // §biXPelsPerMeter — 3780 px/m is the 96 dpi every Office writer assumes.
    const bmp = buildBmp({
      width: 1,
      height: 1,
      bitCount: 24,
      pxPerMeter: 3780,
      rows: [[0, 0, 0, 0]],
    });
    expect(decodeBmp(bmp).dpiX).toBeCloseTo(96, 0);
  });

  it('refuses a bitmap that is a whole other file in disguise', () => {
    // §BI_JPEG / §BI_PNG — the payload is not pixels at all.
    const bmp = buildBmp({
      width: 1,
      height: 1,
      bitCount: 24,
      compression: 4,
      rows: [[0, 0, 0, 0]],
    });
    expect(() => decodeBmp(bmp)).toThrow(/JPEG\/PNG/u);
  });

  it('embeds as Flate-compressed RGB samples', () => {
    const bmp = buildBmp({ width: 1, height: 1, bitCount: 24, rows: [[0x33, 0x22, 0x11, 0]] });
    const prepared = prepareImage(bmp);
    expect(prepared.format).toBe('bmp');
    expect(prepared.mimeType).toBe('image/bmp');
    expect(prepared.filter).toBe('FlateDecode');
    expect(prepared.colorSpace).toBe('DeviceRGB');
    expect([...unzlibSync(prepared.data)]).toEqual([0x11, 0x22, 0x33]);
  });
});

describe('a DIB inside an Escher picture store (MS-ODRAW §2.2.28)', () => {
  // The record holds a bitmap with its FILE header cut off, the record's own
  // length having made it redundant. Nothing downstream can sniff that — it
  // opens with a length field, not a signature.
  const headerless = (): Uint8Array => {
    const file = buildBmp({
      width: 2,
      height: 1,
      bitCount: 24,
      rows: [[0x33, 0x22, 0x11, 0x33, 0x22, 0x11]],
    });
    return file.subarray(14);
  };

  it('puts the missing file header back', () => {
    const bmp = dibToBmp(headerless());
    expect(bmp).toBeDefined();
    expect(isBmp(bmp!)).toBe(true);
    // The pixels survive the round trip intact.
    expect([...decodeBmp(bmp!).data.slice(0, 3)]).toEqual([0x11, 0x22, 0x33]);
  });

  it('finds the bitmap after the record’s UID and tag byte', () => {
    // §2.2.28 — one UID (16 bytes) then a one-byte tag, for the one-UID
    // recInstance; the blip reader has to skip both.
    const dib = headerless();
    const body = new Uint8Array(17 + dib.length);
    body.set(dib, 17);
    const out = readEscherBlip(0xf01f, 0x7a8, body);
    expect(out).toBeDefined();
    expect(isBmp(out!)).toBe(true);
    expect(decodeBmp(out!).width).toBe(2);
  });

  it('says nothing when the payload is not a bitmap at all', () => {
    expect(dibToBmp(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).toBeUndefined();
    // A header whose planes and depth are nonsense is not one of ours.
    const junk = new Uint8Array(60);
    new DataView(junk.buffer).setUint32(0, 40, true);
    expect(dibToBmp(junk)).toBeUndefined();
  });
});
