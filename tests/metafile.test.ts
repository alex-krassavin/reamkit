import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { MetaPicture, PictureImage, PicturePath, PictureText } from '@/core/metafile/picture';
import { FontRegistry } from '@/core/font';
import { ResourceStore, pt } from '@/core/ir';
import { paintPlan } from '@/layout/page-doc';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { isEmf, readEmf } from '@/core/metafile/emf';
import { readDib } from '@/core/metafile/dib';
import { isWmf, readWmf } from '@/core/metafile/wmf';

// A metafile is a list of records over a device context, so the tests build the
// records by hand: a header, the state, the drawing, an end.

class Bytes {
  private readonly parts: Array<Uint8Array> = [];
  u16(v: number): this {
    const b = new Uint8Array(2);
    new DataView(b.buffer).setUint16(0, v, true);
    this.parts.push(b);
    return this;
  }
  i16(v: number): this {
    const b = new Uint8Array(2);
    new DataView(b.buffer).setInt16(0, v, true);
    this.parts.push(b);
    return this;
  }
  u32(v: number): this {
    const b = new Uint8Array(4);
    new DataView(b.buffer).setUint32(0, v >>> 0, true);
    this.parts.push(b);
    return this;
  }
  i32(v: number): this {
    const b = new Uint8Array(4);
    new DataView(b.buffer).setInt32(0, v, true);
    this.parts.push(b);
    return this;
  }
  f32(v: number): this {
    const b = new Uint8Array(4);
    new DataView(b.buffer).setFloat32(0, v, true);
    this.parts.push(b);
    return this;
  }
  ascii(s: string, pad = 0): this {
    const b = new Uint8Array(Math.max(pad, s.length));
    for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i);
    this.parts.push(b);
    return this;
  }
  utf16(s: string, padUnits = 0): this {
    const n = Math.max(padUnits, s.length);
    const b = new Uint8Array(n * 2);
    const v = new DataView(b.buffer);
    for (let i = 0; i < s.length; i++) v.setUint16(i * 2, s.charCodeAt(i), true);
    this.parts.push(b);
    return this;
  }
  zeros(n: number): this {
    this.parts.push(new Uint8Array(n));
    return this;
  }
  build(): Uint8Array {
    const total = this.parts.reduce((n, p) => n + p.length, 0);
    const out = new Uint8Array(total);
    let at = 0;
    for (const p of this.parts) {
      out.set(p, at);
      at += p.length;
    }
    return out;
  }
}

/** An EMF record: type, its own size, then the body. */
const emr = (type: number, body: Uint8Array): Uint8Array =>
  concat(
    new Bytes()
      .u32(type)
      .u32(8 + body.length)
      .build(),
    body,
  );

function concat(...parts: Array<Uint8Array>): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

/** An EMF header over the given bounds, followed by `records` and an EOF. */
function emf(
  bounds: { l: number; t: number; r: number; b: number },
  records: Array<Uint8Array>,
): Uint8Array {
  const header = concat(
    new Bytes()
      .u32(1) // EMR_HEADER
      .u32(88)
      .i32(bounds.l)
      .i32(bounds.t)
      .i32(bounds.r)
      .i32(bounds.b)
      .zeros(16) // rclFrame
      .ascii(' EMF') // dSignature at offset 40
      .u32(0x10000)
      .u32(0) // nBytes, filled by nobody: the reader walks records
      .zeros(88 - 52)
      .build(),
  );
  return concat(header, ...records, emr(14, new Bytes().zeros(12).build()));
}

const rect = (l: number, t: number, r: number, b: number) =>
  emr(43, new Bytes().i32(l).i32(t).i32(r).i32(b).build());
const createBrush = (ih: number, colorBgr: number) =>
  emr(39, new Bytes().u32(ih).u32(0).u32(colorBgr).u32(0).build());
const createPen = (ih: number, style: number, width: number, colorBgr: number) =>
  emr(38, new Bytes().u32(ih).u32(style).i32(width).i32(0).u32(colorBgr).build());
const select = (ih: number) => emr(37, new Bytes().u32(ih).build());

const paths = (prims: MetaPicture['prims']): Array<PicturePath> =>
  prims.filter((p): p is PicturePath => p.kind === 'path');
const texts = (prims: MetaPicture['prims']): Array<PictureText> =>
  prims.filter((p): p is PictureText => p.kind === 'text');

describe('EMF (MS-EMF)', () => {
  it('knows an enhanced metafile by its header', () => {
    expect(isEmf(emf({ l: 0, t: 0, r: 10, b: 10 }, []))).toBe(true);
    expect(isEmf(new Uint8Array(64))).toBe(false);
  });

  it('reads the bounds as the picture’s box', () => {
    const pic = readEmf(emf({ l: 5, t: 7, r: 105, b: 57 }, []));
    expect([pic.left, pic.top, pic.width, pic.height]).toEqual([5, 7, 100, 50]);
  });

  it('fills a rectangle with the selected brush', () => {
    // COLORREF is 0x00bbggrr: 0x0000FF is pure red.
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 100 }, [
        createBrush(1, 0x0000ff),
        select(1),
        rect(10, 20, 30, 40),
      ]),
    );
    const [p] = paths(pic.prims);
    expect(p?.fillColorHex).toBe('FF0000');
    expect(p?.paths[0]!.segments).toEqual([
      { op: 'move', x: 10, y: 20 },
      { op: 'line', x: 30, y: 20 },
      { op: 'line', x: 30, y: 40 },
      { op: 'line', x: 10, y: 40 },
      { op: 'close' },
    ]);
  });

  it('draws no outline through a NULL pen (§2.1.25 PS_NULL)', () => {
    const withPen = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 100 }, [createPen(1, 0, 2, 0), select(1), rect(0, 0, 10, 10)]),
    );
    expect(paths(withPen.prims)[0]?.stroke?.widthPt).toBe(2);
    const none = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 100 }, [createPen(1, 5, 2, 0), select(1), rect(0, 0, 10, 10)]),
    );
    expect(paths(none.prims)[0]?.stroke).toBeUndefined();
  });

  it('maps logical units through the window and viewport (§2.3.11)', () => {
    // A 100-unit window shown in a 200-unit viewport doubles every coordinate.
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 200, b: 200 }, [
        emr(9, new Bytes().i32(100).i32(100).build()), // SETWINDOWEXTEX
        emr(11, new Bytes().i32(200).i32(200).build()), // SETVIEWPORTEXTEX
        createBrush(1, 0),
        select(1),
        rect(10, 10, 20, 20),
      ]),
    );
    const seg = paths(pic.prims)[0]!.paths[0]!.segments[0]!;
    expect(seg).toEqual({ op: 'move', x: 20, y: 20 });
  });

  it('restores the state a SAVEDC put away', () => {
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 100 }, [
        createBrush(1, 0x0000ff),
        select(1),
        emr(33, new Uint8Array()), // SAVEDC
        createBrush(2, 0x00ff00),
        select(2),
        emr(34, new Bytes().i32(-1).build()), // RESTOREDC
        rect(0, 0, 10, 10),
      ]),
    );
    expect(paths(pic.prims)[0]?.fillColorHex).toBe('FF0000');
  });

  it('reads the text a EMR_EXTTEXTOUTW places, with its font and colour', () => {
    // The string offset is measured from the START of the record: the record's
    // own fields are 76 bytes before it.
    const body = new Bytes()
      .zeros(16) // rclBounds
      .u32(1) // iGraphicsMode
      .f32(1)
      .f32(1)
      .i32(30) // ptlReference.x
      .i32(40) // …y
      .u32(2) // nChars
      .u32(76) // offString, from the record start
      .u32(0)
      .zeros(16) // rcl
      .u32(0) // offDx
      .utf16('hi')
      .build();
    const font = new Bytes()
      .u32(1) // ihFont
      .i32(-24) // lfHeight
      .i32(0)
      .i32(0) // lfEscapement
      .i32(0)
      .i32(700) // lfWeight — bold
      .zeros(8)
      .utf16('Arial', 32)
      .build();
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 100 }, [
        emr(82, font),
        select(1),
        emr(24, new Bytes().u32(0x0000ff).build()), // SETTEXTCOLOR — red
        emr(84, body),
      ]),
    );
    const [t] = texts(pic.prims);
    expect(t?.text).toBe('hi');
    expect([t?.x, t?.y]).toEqual([30, 40]);
    expect(t?.colorHex).toBe('FF0000');
    expect(t?.sizeLu).toBe(24);
    expect(t?.bold).toBe(true);
    expect(t?.fontFamily).toBe('Arial');
  });

  it('names what it could not draw instead of drawing it wrong', () => {
    // A blit WITH a source bitmap: 100 bytes of BITMAPINFO it cannot decode.
    const blit = new Bytes()
      .zeros(16)
      .i32(0)
      .i32(0)
      .i32(10)
      .i32(10) // dest x, y, w, h
      .u32(0x00cc0020) // SRCCOPY
      .zeros(44)
      .u32(100) // cbBmiSrc, at offset 88
      .zeros(8)
      .build();
    expect(readEmf(emf({ l: 0, t: 0, r: 10, b: 10 }, [emr(76, blit)])).skipped).toContain('bitmap');
  });
});

// A WMF record: its size in WORDS (itself included), then the function.
const meta = (fn: number, params: Uint8Array): Uint8Array =>
  concat(
    new Bytes()
      .u32(3 + params.length / 2)
      .u16(fn)
      .build(),
    params,
  );

function wmf(records: Array<Uint8Array>, placeableBox?: [number, number, number, number]) {
  const head = placeableBox
    ? new Bytes()
        .u32(0x9ac6cdd7)
        .u16(0)
        .i16(placeableBox[0])
        .i16(placeableBox[1])
        .i16(placeableBox[2])
        .i16(placeableBox[3])
        .u16(1440)
        .u32(0)
        .u16(0)
        .build()
    : new Uint8Array();
  const header = new Bytes()
    .u16(1) // type: memory metafile
    .u16(9) // header size in words
    .u16(0x0300)
    .u32(0)
    .u16(0)
    .u32(0)
    .u16(0)
    .build();
  return concat(head, header, ...records, meta(0x0000, new Uint8Array()));
}

describe('WMF (MS-WMF)', () => {
  it('knows a metafile placeable or bare', () => {
    expect(isWmf(wmf([], [0, 0, 100, 50]))).toBe(true);
    expect(isWmf(wmf([]))).toBe(true);
    expect(isWmf(new Uint8Array(32))).toBe(false);
  });

  it('takes its box from the placeable header', () => {
    const pic = readWmf(wmf([], [10, 20, 110, 70]));
    expect([pic.left, pic.top, pic.width, pic.height]).toEqual([10, 20, 100, 50]);
  });

  it('…and from the window it sets when there is no placeable header', () => {
    const pic = readWmf(
      wmf([
        meta(0x020b, new Bytes().i16(5).i16(3).build()), // SETWINDOWORG: y, x
        meta(0x020c, new Bytes().i16(40).i16(80).build()), // SETWINDOWEXT: cy, cx
      ]),
    );
    expect([pic.left, pic.top, pic.width, pic.height]).toEqual([3, 5, 80, 40]);
  });

  it('fills the rectangle a PATBLT names with the current brush', () => {
    // The record most used in the corpus: a bar, a rule or a panel.
    const pic = readWmf(
      wmf(
        [
          meta(0x02fc, new Bytes().u16(0).u32(0x0000ff).u16(0).build()), // CREATEBRUSHINDIRECT
          meta(0x012d, new Bytes().u16(0).build()), // SELECTOBJECT 0
          meta(
            0x061d,
            new Bytes().u32(0x00f00021).i16(4).i16(30).i16(3).i16(2).build(), // rop, h, w, y, x
          ),
        ],
        [0, 0, 100, 50],
      ),
    );
    const [p] = paths(pic.prims);
    expect(p?.fillColorHex).toBe('FF0000');
    expect(p?.paths[0]!.segments[0]).toEqual({ op: 'move', x: 2, y: 3 });
    expect(p?.paths[0]!.segments[2]).toEqual({ op: 'line', x: 32, y: 7 });
  });

  it('reuses the slot a deleted object leaves (§3.1.4.1)', () => {
    const pic = readWmf(
      wmf(
        [
          meta(0x02fc, new Bytes().u16(0).u32(0x0000ff).u16(0).build()), // slot 0: red
          meta(0x01f0, new Bytes().u16(0).build()), // delete it
          meta(0x02fc, new Bytes().u16(0).u32(0x00ff00).u16(0).build()), // slot 0 again: green
          meta(0x012d, new Bytes().u16(0).build()),
          meta(0x041b, new Bytes().i16(10).i16(10).i16(0).i16(0).build()), // RECTANGLE
        ],
        [0, 0, 100, 50],
      ),
    );
    expect(paths(pic.prims)[0]?.fillColorHex).toBe('00FF00');
  });

  it('reads a TEXTOUT with the point it stands at', () => {
    const pic = readWmf(
      wmf(
        [
          meta(0x0209, new Bytes().u32(0x0000ff).build()), // SETTEXTCOLOR
          meta(0x0521, new Bytes().u16(2).ascii('hi').i16(9).i16(4).build()), // count, string, y, x
        ],
        [0, 0, 100, 50],
      ),
    );
    const [t] = texts(pic.prims);
    expect(t?.text).toBe('hi');
    expect([t?.x, t?.y]).toEqual([4, 9]);
    expect(t?.colorHex).toBe('FF0000');
  });

  it('draws a symbol font\'s circle rather than spelling it "n"', () => {
    // The bullets in an embedded diagram are Webdings `n`, which is a filled
    // circle — and no substitute font has one either, so it is DRAWN.
    // MS-WMF §2.2.1.2 LogFont: height, width, escapement, orientation, weight,
    // then eight bytes of flags before the 32-byte face name.
    const font = new Bytes()
      .i16(-20) // height
      .i16(0)
      .i16(0) // escapement
      .i16(0)
      .i16(400) // weight
      .ascii('\u0000'.repeat(8))
      .ascii('Webdings', 32)
      .build();
    const pic = readWmf(
      wmf(
        [
          meta(0x02fb, font), // CREATEFONTINDIRECT
          meta(0x012d, new Bytes().u16(0).build()), // SELECTOBJECT
          meta(0x0209, new Bytes().u32(0x800000).build()), // SETTEXTCOLOR — blue
          meta(0x0521, new Bytes().u16(1).ascii('n\u0000').i16(40).i16(10).build()), // TEXTOUT
        ],
        [0, 0, 100, 50],
      ),
    );
    expect(texts(pic.prims)).toEqual([]); // nothing spelled
    const [shape] = paths(pic.prims);
    expect(shape?.fillColorHex).toBe('000080');
    // One em wide, and round: four quadrant curves.
    const segs = shape?.paths[0]?.segments ?? [];
    expect(segs.filter((sg) => sg.op === 'cubic').length).toBe(4);
  });

  it('drops what a picture declares transparent (a:clrChange)', () => {
    // §20.1.8.16 — tdf113163's whole slide is a metafile whose white ground is
    // declared away, so the black slide shows through. Drawn as stored, the
    // slide is a white sheet.
    const white = new Bytes().i16(20).i16(80).i16(0).i16(0).build(); // b,r,t,l
    const red = new Bytes().i16(20).i16(40).i16(0).i16(0).build();
    const bytes = wmf(
      [
        meta(0x02fc, new Bytes().u16(0).u32(0xffffff).u16(0).build()), // white brush
        meta(0x012d, new Bytes().u16(0).build()),
        meta(0x041b, white),
        meta(0x02fc, new Bytes().u16(0).u32(0x0000ff).u16(0).build()), // red brush
        meta(0x012d, new Bytes().u16(1).build()),
        meta(0x041b, red),
      ],
      [0, 0, 100, 50],
    );
    const store = new ResourceStore();
    const resource = store.put(bytes);
    const lay = (colorChange?: { fromHex: string; toHex: string; transparent: boolean }) =>
      layoutStyledDocument(
        [
          {
            kind: 'image',
            image: {
              resource,
              width: pt(100),
              height: pt(50),
              paragraphProperties: {},
              ...(colorChange ? { colorChange } : {}),
            },
          },
        ],
        {
          registry: FontRegistry.fromBytes({
            regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          }),
          resources: store,
          styles: { defaultRunProperties: {}, defaultParagraphProperties: {}, styles: new Map() },
        },
      );
    const fills = (laid: ReturnType<typeof lay>): Array<string | undefined> =>
      laid.pages[0]!.commands.flatMap((c) => (c.type === 'shape' ? [c.shape.fillColorHex] : []));
    expect(fills(lay())).toEqual(['FFFFFF', 'FF0000']);
    // The white FILL goes; the rectangle's black outline is another colour and
    // stays. The corpus deck draws its ground with a null pen, so there the
    // primitive disappears completely.
    expect(fills(lay({ fromHex: 'FFFFFF', toHex: 'FFFFFF', transparent: true }))).toEqual([
      undefined,
      'FF0000',
    ]);
    // …and a change that names another colour repaints instead of dropping.
    expect(fills(lay({ fromHex: 'FFFFFF', toHex: '00FF00', transparent: false }))).toEqual([
      '00FF00',
      'FF0000',
    ]);
  });

  it('paints as ONE picture, in the order the metafile draws', () => {
    // A metafile writes a label, lays a panel over it and writes it again as a
    // drop shadow; painted as "every shape, then every line" the buried copy
    // would show through — 45541_Footer's rotated headers came out doubled.
    const panel = new Bytes().i16(30).i16(60).i16(0).i16(0).build(); // b,r,t,l
    const bytes = wmf(
      [
        meta(0x0521, new Bytes().u16(2).ascii('hi').i16(20).i16(4).build()), // TEXTOUT
        meta(0x02fc, new Bytes().u16(0).u32(0x00ff00).u16(0).build()), // CREATEBRUSH — green
        meta(0x012d, new Bytes().u16(0).build()), // SELECTOBJECT
        meta(0x041b, panel), // RECTANGLE over it
      ],
      [0, 0, 100, 50],
    );
    const store = new ResourceStore();
    const resource = store.put(bytes);
    const laid = layoutStyledDocument(
      [
        {
          kind: 'image',
          image: { resource, width: pt(100), height: pt(50), paragraphProperties: {} },
        },
      ],
      {
        registry: FontRegistry.fromBytes({
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        }),
        resources: store,
        styles: { defaultRunProperties: {}, defaultParagraphProperties: {}, styles: new Map() },
      },
    );
    const items = laid.pages[0]!.commands.filter((c) => c.pictureId !== undefined);
    expect(items.map((c) => c.type)).toEqual(['line', 'shape']); // the text FIRST
    expect(new Set(items.map((c) => c.pictureId)).size).toBe(1);
    // …and the plan keeps them out of the passes that would reorder them.
    const plan = paintPlan(laid.pages[0]!.commands);
    expect(plan.pictures).toHaveLength(1);
    expect(plan.shapes).toHaveLength(0);
    expect(plan.lines).toHaveLength(0);
  });

  it('draws a polygon closed and a polyline open', () => {
    const pts = new Bytes().u16(3).i16(0).i16(0).i16(10).i16(0).i16(10).i16(10).build();
    const poly = readWmf(wmf([meta(0x0324, pts)], [0, 0, 20, 20]));
    const line = readWmf(wmf([meta(0x0325, pts)], [0, 0, 20, 20]));
    expect(paths(poly.prims)[0]?.paths[0]!.segments.at(-1)).toEqual({ op: 'close' });
    expect(paths(line.prims)[0]?.paths[0]!.segments.at(-1)).not.toEqual({ op: 'close' });
  });
});

// MS-WMF §2.2.2.9 — a packed DIB: a BITMAPINFOHEADER, a colour table, and the
// rows, stored BOTTOM first and padded out to a four-byte boundary.
const bmi = (w: number, h: number, bpp: number, clrUsed = 0): Uint8Array =>
  new Bytes()
    .u32(40)
    .i32(w)
    .i32(h)
    .u16(1)
    .u16(bpp)
    .u32(0) // BI_RGB
    .u32(0)
    .i32(0)
    .i32(0)
    .u32(clrUsed)
    .u32(0)
    .build();

/** A 24-bit DIB over `pixels`, given TOP row first as `[r, g, b]`. */
function dib24(w: number, h: number, pixels: Array<[number, number, number]>): Uint8Array {
  const stride = (w * 3 + 3) & ~3;
  const bits = new Uint8Array(stride * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const [r, g, b] = pixels[(h - 1 - y) * w + x]!;
      bits.set([b, g, r], y * stride + x * 3);
    }
  }
  return concat(bmi(w, h, 24), bits);
}

/** A monochrome DIB over `on` bits, TOP row first: 0 is black, 1 is white. */
function dib1(w: number, h: number, on: Array<number>): Uint8Array {
  const stride = (((w + 7) >> 3) + 3) & ~3;
  const bits = new Uint8Array(stride * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (on[(h - 1 - y) * w + x] === 1) bits[y * stride + (x >> 3)]! |= 0x80 >> (x & 7);
    }
  }
  const palette = new Bytes().u32(0x000000).u32(0xffffff).build();
  return concat(bmi(w, h, 1, 2), palette, bits);
}

/** EMR_STRETCHDIBITS (§2.3.1.7) — the whole bitmap onto the destination box. */
const stretchDibits = (
  dest: [number, number, number, number],
  bitmap: Uint8Array,
  src: [number, number, number, number],
  rop = 0x00cc0020,
  bmiSize = 40,
): Uint8Array =>
  emr(
    81,
    concat(
      new Bytes()
        .zeros(16) // Bounds
        .i32(dest[0])
        .i32(dest[1])
        .i32(src[0])
        .i32(src[1])
        .i32(src[2])
        .i32(src[3])
        .u32(80) // offBmiSrc, from the record's start
        .u32(bmiSize)
        .u32(80 + bmiSize) // offBitsSrc
        .u32(bitmap.length - bmiSize)
        .u32(0) // UsageSrc
        .u32(rop)
        .i32(dest[2])
        .i32(dest[3])
        .build(),
      bitmap,
    ),
  );

const images = (prims: MetaPicture['prims']): Array<PictureImage> =>
  prims.filter((p): p is PictureImage => p.kind === 'image');

// A PNG's IHDR: its extent, and colour type 2 (RGB) or 6 (RGBA).
const ihdr = (png: Uint8Array): { w: number; h: number; color: number } => {
  const v = new DataView(png.buffer, png.byteOffset, png.byteLength);
  return { w: v.getUint32(16), h: v.getUint32(20), color: png[25]! };
};

describe('a blit (MS-EMF §2.3.1)', () => {
  const red: [number, number, number] = [255, 0, 0];
  const blue: [number, number, number] = [0, 0, 255];

  it('reads a DIB bottom-up, through its padding, into the box it lands in', () => {
    const bitmap = dib24(2, 2, [red, red, blue, blue]);
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 50 }, [stretchDibits([10, 20, 30, 40], bitmap, [0, 0, 2, 2])]),
    );
    expect(images(pic.prims)).toHaveLength(1);
    const img = images(pic.prims)[0]!;
    expect([img.x, img.y, img.width, img.height]).toEqual([10, 20, 30, 40]);
    expect(ihdr(img.png)).toEqual({ w: 2, h: 2, color: 2 });
    // The first row of the PICTURE is the LAST row of the file, and it is red.
    const back = readDib(bitmap, 0)!;
    expect([...back.rgba.subarray(0, 4)]).toEqual([255, 0, 0, 255]);
    expect([...back.rgba.subarray(8, 12)]).toEqual([0, 0, 255, 255]);
  });

  it('puts an AND mask and the OR that follows it back into one picture', () => {
    // The idiom every clipart of this age is drawn with: a monochrome mask
    // knocks the ground out, and the picture is ORed into the hole.
    const dest: [number, number, number, number] = [0, 0, 20, 20];
    const src: [number, number, number, number] = [0, 0, 2, 2];
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 50 }, [
        stretchDibits(dest, dib1(2, 2, [1, 1, 0, 0]), src, 0x008800c6, 48),
        stretchDibits(dest, dib24(2, 2, [red, red, red, red]), src, 0x00ee0086),
      ]),
    );
    // ONE picture, not two, and — the colour bitmap having no black of its own
    // to key out — an alpha channel it can only have got from the mask.
    expect(images(pic.prims)).toHaveLength(1);
    expect(ihdr(images(pic.prims)[0]!.png).color).toBe(6);
  });

  it('keeps an OR blit’s own black transparent when no mask came with it', () => {
    const black: [number, number, number] = [0, 0, 0];
    const half = dib24(2, 2, [black, black, red, red]);
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 50 }, [
        stretchDibits([0, 0, 20, 20], half, [0, 0, 2, 2], 0x00ee0086),
      ]),
    );
    expect(ihdr(images(pic.prims)[0]!.png).color).toBe(6);
    // …while a plain copy of the same bitmap keeps its black.
    const copied = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 50 }, [stretchDibits([0, 0, 20, 20], half, [0, 0, 2, 2])]),
    );
    expect(ihdr(images(copied.prims)[0]!.png).color).toBe(2);
  });

  it('draws nothing for a mask with no picture behind it', () => {
    const pic = readEmf(
      emf({ l: 0, t: 0, r: 100, b: 50 }, [
        stretchDibits([0, 0, 20, 20], dib1(2, 2, [1, 1, 0, 0]), [0, 0, 2, 2], 0x008800c6, 48),
      ]),
    );
    expect(images(pic.prims)).toHaveLength(0);
  });

  it('fills the rectangle with the brush when a WMF blit carries no bitmap', () => {
    // 288 records of the corpus are exactly this: META_DIBBITBLT with PATCOPY
    // and no source at all, which is the brush painting a rule or a panel. The
    // record is then TWELVE words — one more than its fields account for — and
    // read without that reserved word the rectangle comes out of no height.
    const pic = readWmf(
      wmf(
        [
          meta(0x02fc, new Bytes().u16(0).u32(0x00ff00).u16(0).build()), // CREATEBRUSH — green
          meta(0x012d, new Bytes().u16(0).build()), // SELECTOBJECT
          meta(
            0x0940,
            new Bytes()
              .u32(0x00f00021) // PATCOPY
              .i16(0) // reserved
              .i16(0) // YSrc
              .i16(0) // XSrc
              .i16(6) // Height
              .i16(8) // Width
              .i16(2) // YDest
              .i16(3) // XDest
              .build(),
          ),
        ],
        [0, 0, 100, 50],
      ),
    );
    expect(paths(pic.prims)[0]?.fillColorHex).toBe('00FF00');
    expect(paths(pic.prims)[0]?.paths[0]!.segments).toEqual([
      { op: 'move', x: 3, y: 2 },
      { op: 'line', x: 11, y: 2 },
      { op: 'line', x: 11, y: 8 },
      { op: 'line', x: 3, y: 8 },
      { op: 'close' },
    ]);
  });

  it('reads the bitmap a WMF stretches into its destination rectangle', () => {
    const bitmap = dib24(2, 1, [red, blue]);
    const pic = readWmf(
      wmf(
        [
          meta(
            0x0b41,
            concat(
              new Bytes()
                .u32(0x00cc0020)
                .i16(1) // SrcHeight
                .i16(2) // SrcWidth
                .i16(0) // YSrc
                .i16(0) // XSrc
                .i16(30) // DestHeight
                .i16(40) // DestWidth
                .i16(5) // YDest
                .i16(7) // XDest
                .build(),
              bitmap,
            ),
          ),
        ],
        [0, 0, 100, 50],
      ),
    );
    expect(images(pic.prims)).toHaveLength(1);
    const img = images(pic.prims)[0]!;
    expect([img.x, img.y, img.width, img.height]).toEqual([7, 5, 40, 30]);
  });

  it('carries the bitmap on as a resource the page draws', () => {
    const bytes = emf({ l: 0, t: 0, r: 100, b: 50 }, [
      stretchDibits([0, 0, 100, 50], dib24(2, 2, [red, red, blue, blue]), [0, 0, 2, 2]),
    ]);
    const store = new ResourceStore();
    const resource = store.put(bytes);
    const laid = layoutStyledDocument(
      [
        {
          kind: 'image',
          image: { resource, width: pt(100), height: pt(50), paragraphProperties: {} },
        },
      ],
      {
        registry: FontRegistry.fromBytes({
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        }),
        resources: store,
        styles: { defaultRunProperties: {}, defaultParagraphProperties: {}, styles: new Map() },
      },
    );
    const item = laid.pages[0]!.commands.find((c) => c.type === 'image');
    expect(item?.imageResourceName).not.toBe('');
    // …and the picture it names is embedded, not left dangling.
    expect(
      [...laid.imageResources.values()].some(
        (r) => r.resourceName === item?.imageResourceName && r.prepared !== undefined,
      ),
    ).toBe(true);
  });
});
