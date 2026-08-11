// E-PDF EP6 — image extraction. Three layers: the PNG encoder produces a
// decodable file; the XObject decoder handles the colour spaces / filters /
// passthroughs; and end-to-end, an image embedded in a docx survives the
// docx → pdf → parse round-trip back into the FlowDoc (and onward to HTML).

import { readFileSync } from 'node:fs';

import { unzlibSync, zlibSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { Ream } from '@/core/converter/ream';
import { detectImageFormat, prepareImage } from '@/core/images';
import { PdfFile } from '@/pdf-reader/document';
import { collectPageImages } from '@/pdf-reader/images';
import { decodePdfImage } from '@/pdf-reader/image-decode';
import { encodePng } from '@/core/png-encode';
import { PdfHexString, dict, name, stream } from '@/pdf/objects';

/** The 8x8 flat-tone baseline JPEG the decoder's own test assembles. */
function flatGrayJpeg(): Uint8Array {
  const seg = (marker: number, body: ReadonlyArray<number>): Array<number> => [
    0xff,
    marker,
    (body.length + 2) >> 8,
    (body.length + 2) & 0xff,
    ...body,
  ];
  const table = (tcth: number, symbol: number): Array<number> => [
    tcth,
    1,
    ...new Array<number>(15).fill(0),
    symbol,
  ];
  return new Uint8Array([
    0xff,
    0xd8,
    ...seg(0xdb, [0x00, ...new Array<number>(64).fill(8)]),
    ...seg(0xc0, [8, 0, 8, 0, 8, 1, 1, 0x11, 0]),
    ...seg(0xc4, table(0x00, 0x04)),
    ...seg(0xc4, table(0x10, 0x00)),
    ...seg(0xda, [1, 1, 0x00, 0x00, 0x3f, 0x00]),
    0b0111_1011,
    0xff,
    0xd9,
  ]);
}

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// A lone inline picture referencing rId20 (the EMU extent is 72pt × 54pt).
const drawingXml = (rId: string, cxEmu: number, cyEmu: number): string =>
  `<w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="${cxEmu}" cy="${cyEmu}"/>
      <wp:docPr id="1" name="Picture 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r>`;

const docxWithImage = (): Uint8Array => {
  const png = buildTinyPng(8, 6, [200, 60, 60, 255]);
  return buildDocxFromBody(`<w:p>${drawingXml('rId20', 914400, 685800)}</w:p>`, {
    images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
  });
};

describe('PNG encoder (E-PDF EP6)', () => {
  it('encodes RGB samples into a decodable PNG', () => {
    const samples = Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0]); // 2×2
    const png = encodePng(2, 2, 'rgb', samples);
    expect(detectImageFormat(png)).toBe('png');
    const back = prepareImage(png);
    expect([back.widthPx, back.heightPx]).toEqual([2, 2]);
    expect(unzlibSync(back.data)).toEqual(samples);
  });

  it('encodes gray + alpha (split back into an /SMask)', () => {
    const png = encodePng(1, 2, 'gray-alpha', Uint8Array.from([10, 255, 20, 128]));
    const back = prepareImage(png);
    expect(back.colorSpace).toBe('DeviceGray');
    expect(back.smaskData).toBeDefined();
  });
});

describe('PDF image XObject decode (E-PDF EP6)', () => {
  let file: PdfFile;
  beforeAll(async () => {
    const pdf = await Ream.parse(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>')).convert(
      'pdf',
      { fonts: FONTS },
    );
    file = PdfFile.parse(pdf);
  });

  it('decodes a Flate DeviceRGB image to PNG', () => {
    const raw = Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0]); // 2×2
    const xobj = stream(
      {
        Width: 2,
        Height: 2,
        ColorSpace: name('DeviceRGB'),
        BitsPerComponent: 8,
        Filter: name('FlateDecode'),
      },
      zlibSync(raw),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(decoded.format).toBe('png');
    expect([decoded.widthPx, decoded.heightPx]).toEqual([2, 2]);
    expect(detectImageFormat(decoded.bytes)).toBe('png');
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(raw);
  });

  it('takes a JPEG apart to attach the /SMask a JPEG cannot carry', () => {
    // ISO 32000-1 §8.9.5.4 — the mask is a second image supplying the alpha,
    // and a JPEG has nowhere to put it. 22060_A1_01_Plans.pdf stores each of
    // its floor plans this way: the drawing as a JPEG, its line work as a grey
    // JPEG mask. Carried through unmasked, they render as dark rectangles.
    const jpeg = flatGrayJpeg();
    const mask = stream({ Width: 8, Height: 8, Filter: name('DCTDecode') }, jpeg);
    const xobj = stream(
      {
        Width: 8,
        Height: 8,
        ColorSpace: name('DeviceGray'),
        Filter: name('DCTDecode'),
        SMask: mask,
      },
      jpeg,
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    // A PNG now, not the original bytes, and with no loss to report.
    expect(decoded.format).toBe('png');
    expect(decoded.degraded).toBeUndefined();
    const back = prepareImage(decoded.bytes);
    expect(back.smaskData).toBeDefined();
  });

  it('passes a DCTDecode (JPEG) image through verbatim', () => {
    const jpeg = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 16, 1, 2, 3, 4]);
    const xobj = stream(
      { Width: 12, Height: 9, ColorSpace: name('DeviceRGB'), Filter: name('DCTDecode') },
      jpeg,
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(decoded.format).toBe('jpeg');
    expect(decoded.widthPx).toBe(12);
    expect(decoded.bytes).toEqual(jpeg);
  });

  it('expands a 1-bit Indexed image against its palette', () => {
    // palette 0→red, 1→green; two pixels [0, 1] packed into one byte (row-aligned).
    const palette = Uint8Array.from([255, 0, 0, 0, 255, 0]);
    const indices = Uint8Array.from([0b01000000]);
    const xobj = stream(
      {
        Width: 2,
        Height: 1,
        BitsPerComponent: 1,
        ColorSpace: [name('Indexed'), name('DeviceRGB'), 1, new PdfHexString(palette)],
        Filter: name('FlateDecode'),
      },
      zlibSync(indices),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(
      Uint8Array.from([255, 0, 0, 0, 255, 0]),
    );
  });

  it('reports a stencil image mask as a loss', () => {
    const xobj = stream(
      { Width: 2, Height: 2, ImageMask: true, Filter: name('FlateDecode') },
      zlibSync(Uint8Array.from([0])),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(false);
    if (!decoded.ok) expect(decoded.detail).toMatch(/mask/i);
  });
});

// A PDF/TIFF LZW encoder (the inverse of the reader's decoder) — 9→12-bit
// variable-width codes, a leading clear (256) and a trailing end-of-data (257),
// with the same `nextCode + earlyChange === 2^width` width-bump rule.
function lzwEncode(data: Uint8Array, earlyChange = 1): Uint8Array {
  const out: Array<number> = [];
  let bitBuffer = 0;
  let bitCount = 0;
  let codeWidth = 9;
  const writeCode = (code: number): void => {
    bitBuffer = (bitBuffer << codeWidth) | code;
    bitCount += codeWidth;
    while (bitCount >= 8) {
      bitCount -= 8;
      out.push((bitBuffer >>> bitCount) & 0xff);
    }
  };
  let dictionary = new Map<string, number>();
  let nextCode = 258;
  const reset = (): void => {
    dictionary = new Map();
    for (let i = 0; i < 256; i++) dictionary.set(String.fromCharCode(i), i);
    nextCode = 258;
    codeWidth = 9;
  };
  reset();
  writeCode(256);
  let w = '';
  for (const b of data) {
    const c = String.fromCharCode(b);
    const wc = w + c;
    if (dictionary.has(wc)) {
      w = wc;
    } else {
      writeCode(dictionary.get(w)!);
      if (nextCode < 4096) {
        dictionary.set(wc, nextCode++);
        // The decoder's table lags the encoder's by one entry, so the encoder
        // widens one code later than the decoder's `nextCode + earlyChange`.
        if (nextCode + earlyChange === 513) codeWidth = 10;
        else if (nextCode + earlyChange === 1025) codeWidth = 11;
        else if (nextCode + earlyChange === 2049) codeWidth = 12;
      }
      w = c;
    }
  }
  if (w !== '') writeCode(dictionary.get(w)!);
  writeCode(257);
  if (bitCount > 0) out.push((bitBuffer << (8 - bitCount)) & 0xff);
  return Uint8Array.from(out);
}

// Forward TIFF Predictor 2 (horizontal differencing, 8-bit) — the encoder side
// of the reader's reversePredictor.
function tiffPredictor2(raw: Uint8Array, width: number, colors: number): Uint8Array {
  const rowBytes = width * colors;
  const out = new Uint8Array(raw.length);
  for (let off = 0; off < raw.length; off += rowBytes) {
    for (let i = 0; i < rowBytes; i++) {
      const left = i >= colors ? raw[off + i - colors]! : 0;
      out[off + i] = (raw[off + i]! - left) & 0xff;
    }
  }
  return out;
}

describe('PDF LZW image decode (E-PDF EP12)', () => {
  let file: PdfFile;
  beforeAll(async () => {
    const pdf = await Ream.parse(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>')).convert(
      'pdf',
      { fonts: FONTS },
    );
    file = PdfFile.parse(pdf);
  });

  it('decodes an LZW DeviceGray image across 9→11-bit code widths', () => {
    const W = 64;
    const H = 40; // 2560 varied bytes — crosses the 512 and 1024 code boundaries
    const raw = new Uint8Array(W * H);
    for (let i = 0; i < raw.length; i++) raw[i] = (i * 31 + (i >> 4) * 7) & 0xff;
    const xobj = stream(
      {
        Width: W,
        Height: H,
        ColorSpace: name('DeviceGray'),
        BitsPerComponent: 8,
        Filter: name('LZWDecode'),
      },
      lzwEncode(raw),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(decoded.format).toBe('png');
    expect([decoded.widthPx, decoded.heightPx]).toEqual([W, H]);
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(raw);
  });

  it('handles the KwKwK case (a run of repeated bytes)', () => {
    const raw = new Uint8Array(20).fill(65); // "AAAA…" forces code === nextCode
    const xobj = stream(
      {
        Width: 4,
        Height: 5,
        ColorSpace: name('DeviceGray'),
        BitsPerComponent: 8,
        Filter: name('LZWDecode'),
      },
      lzwEncode(raw),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(raw);
  });

  it('reverses a TIFF Predictor 2 layered over LZW', () => {
    const W = 6;
    const H = 4;
    const raw = new Uint8Array(W * H * 3);
    for (let i = 0; i < raw.length; i++) raw[i] = (i * 5 + 3) & 0xff;
    const xobj = stream(
      {
        Width: W,
        Height: H,
        ColorSpace: name('DeviceRGB'),
        BitsPerComponent: 8,
        Filter: name('LZWDecode'),
        DecodeParms: dict({ Predictor: 2, Colors: 3, BitsPerComponent: 8, Columns: W }),
      },
      lzwEncode(tiffPredictor2(raw, W, 3)),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(raw);
  });

  it('honours /EarlyChange 0', () => {
    const W = 64;
    const H = 24;
    const raw = new Uint8Array(W * H);
    for (let i = 0; i < raw.length; i++) raw[i] = (i * 13 + 5) & 0xff;
    const xobj = stream(
      {
        Width: W,
        Height: H,
        ColorSpace: name('DeviceGray'),
        BitsPerComponent: 8,
        Filter: name('LZWDecode'),
        DecodeParms: dict({ EarlyChange: 0 }),
      },
      lzwEncode(raw, 0),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(unzlibSync(prepareImage(decoded.bytes).data)).toEqual(raw);
  });
});

describe('image reconstruction end-to-end (E-PDF EP6)', () => {
  it('lifts a raster image back out of an untagged PDF, sized from the CTM', async () => {
    const pdf = await Ream.parse(docxWithImage()).convert('pdf', { fonts: FONTS });
    const doc = Ream.parse(pdf);
    expect(doc.format).toBe('pdf');
    const img = doc.flow.body.find((b) => b.kind === 'image');
    expect(img).toBeDefined();
    if (img?.kind !== 'image') return;
    const bytes = doc.flow.resources.get(img.image.resource!);
    expect(bytes).toBeDefined();
    expect(detectImageFormat(bytes!)).not.toBeNull();
    // 914400 × 685800 EMU = 72 × 54 pt.
    expect(img.image.width).toBeGreaterThan(60);
    expect(img.image.width).toBeLessThan(84);
    expect(img.image.height).toBeGreaterThan(44);
    expect(img.image.height).toBeLessThan(64);
  });

  it('lifts the image from a tagged PDF (a /Figure)', async () => {
    const pdf = await Ream.parse(docxWithImage()).convert('pdf', { fonts: FONTS, tagged: true });
    const doc = Ream.parse(pdf);
    expect(doc.flow.body.some((b) => b.kind === 'image')).toBe(true);
  });

  it('carries the lifted image into HTML output', async () => {
    const pdf = await Ream.parse(docxWithImage()).convert('pdf', { fonts: FONTS });
    const html = new TextDecoder().decode(await Ream.parse(pdf).convert('html'));
    expect(html).toContain('<img');
    expect(html).toContain('data:image/');
  });

  it('reports a bare `sh` region where there is one, and NOT where there is none', async () => {
    // §8.7.4.3 — `sh` paints the clip rather than filling a path, and nothing
    // here lifts it. Reported unconditionally it fired on all four hundred
    // files of the pdf.js corpus, most of which contain no `sh` at all — and a
    // loss report that cries wolf on every document tells a reader nothing.
    const plain = await Ream.parse(docxWithImage()).convert('pdf', { fonts: FONTS });
    expect(Ream.parse(plain).losses.some((l) => /bare-shading/u.test(l.detail))).toBe(false);

    const content = '/Sh0 sh';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R ' +
        '/Resources << /Shading << /Sh0 << /ShadingType 2 /ColorSpace /DeviceRGB ' +
        '/Coords [0 0 100 0] /Function << /FunctionType 2 /Domain [0 1] ' +
        '/C0 [1 0 0] /C1 [0 0 1] /N 1 >> >> >> >> >>',
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ];
    let pdf = '%PDF-1.7\n';
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(pdf.length);
      pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
    });
    const xref = pdf.length;
    pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
    for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
    pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
    const shaded = Ream.parse(new TextEncoder().encode(pdf));
    expect(shaded.losses.some((l) => /bare-shading/u.test(l.detail))).toBe(true);
  });
});

/** Read back a PNG this suite made: its size and its pixels as RGBA. */
function decodePngPixels(png: Uint8Array): { width: number; height: number; rgba: Uint8Array } {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  let at = 8;
  let width = 0;
  let height = 0;
  let channels = 4;
  const parts: Array<Uint8Array> = [];
  while (at + 8 <= png.length) {
    const len = view.getUint32(at);
    const type = String.fromCharCode(...png.subarray(at + 4, at + 8));
    const body = png.subarray(at + 8, at + 8 + len);
    if (type === 'IHDR') {
      width = view.getUint32(at + 8);
      height = view.getUint32(at + 12);
      channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[body[9] as 0 | 2 | 4 | 6];
    } else if (type === 'IDAT') {
      parts.push(body);
    } else if (type === 'IEND') {
      break;
    }
    at += 12 + len;
  }
  const joined = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let off = 0;
  for (const p of parts) {
    joined.set(p, off);
    off += p.length;
  }
  const raw = unzlibSync(joined);
  // Every scanline this suite writes carries filter 0 (None).
  const rgba = new Uint8Array(width * height * 4);
  const stride = width * channels;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const src = y * (stride + 1) + 1 + x * channels;
      const dst = (y * width + x) * 4;
      if (channels >= 3) {
        rgba[dst] = raw[src]!;
        rgba[dst + 1] = raw[src + 1]!;
        rgba[dst + 2] = raw[src + 2]!;
        rgba[dst + 3] = channels === 4 ? raw[src + 3]! : 255;
      } else {
        rgba[dst] = raw[src]!;
        rgba[dst + 1] = raw[src]!;
        rgba[dst + 2] = raw[src]!;
        rgba[dst + 3] = channels === 2 ? raw[src + 1]! : 255;
      }
    }
  }
  return { width, height, rgba };
}

describe('§8.9.6.2 — a stencil mask, and §8.9.7 — an image written into the stream', () => {
  /** A one-page file: `content` its stream, extra objects appended after it. */
  const page = (
    content: string,
    extra: ReadonlyArray<string>,
    streams: ReadonlyMap<number, Uint8Array>,
  ): Uint8Array => {
    const bytes: Array<number> = [];
    const push = (t: string): void => {
      for (const ch of t) bytes.push(ch.charCodeAt(0) & 0xff);
    };
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R ' +
        '/Resources << /XObject << /Im 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>`,
      ...extra,
    ];
    push('%PDF-1.7\n');
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(bytes.length);
      push(`${String(i + 1)} 0 obj\n${body}\n`);
      const data =
        i === 3 ? Uint8Array.from([...content].map((c) => c.charCodeAt(0))) : streams.get(i + 1);
      if (data) {
        push('stream\n');
        for (const b of data) bytes.push(b);
        push('\nendstream\n');
      }
      push('endobj\n');
    });
    const xref = bytes.length;
    push(`xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`);
    for (const off of offsets) push(`${String(off).padStart(10, '0')} 00000 n \n`);
    push(
      `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`,
    );
    return Uint8Array.from(bytes);
  };

  /** The one image such a page carries, decoded to its PNG pixels. */
  const pixels = (pdf: Uint8Array): { width: number; height: number; rgba: Uint8Array } => {
    const file = PdfFile.parse(pdf);
    const { images } = collectPageImages(file, file.pages()[0]!);
    expect(images).toHaveLength(1);
    return decodePngPixels(images[0]!.bytes);
  };

  it('paints a stencil in the page’s own fill colour, clear elsewhere', () => {
    // §8.9.6.2 — a sample of 0 paints, 1 leaves the page alone. Two pixels
    // across, one row: paint, then leave.
    const mask = Uint8Array.from([0b0100_0000]);
    const pdf = page(
      '1 0 0 rg 0 0 100 100 cm /Im Do',
      [
        '<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true ' +
          '/BitsPerComponent 1 /Length 1 >>',
      ],
      new Map([[5, mask]]),
    );
    const { width, rgba } = pixels(pdf);
    expect(width).toBe(2);
    // Painted: the fill colour, opaque. Not painted: clear.
    expect([...rgba.slice(0, 4)]).toEqual([255, 0, 0, 255]);
    expect(rgba[7]).toBe(0);
  });

  it('lets /Decode [1 0] say the OTHER bit paints', () => {
    const mask = Uint8Array.from([0b0100_0000]);
    const pdf = page(
      '0 0 1 rg 0 0 100 100 cm /Im Do',
      [
        '<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true ' +
          '/Decode [1 0] /BitsPerComponent 1 /Length 1 >>',
      ],
      new Map([[5, mask]]),
    );
    const { rgba } = pixels(pdf);
    expect(rgba[3]).toBe(0);
    expect([...rgba.slice(4, 8)]).toEqual([0, 0, 255, 255]);
  });

  it('MEASURES an inline image rather than hunting for EI in its bytes', () => {
    // §8.9.7 — the data is binary and may hold `EI` itself. Three grey pixels
    // whose bytes are exactly 'E', 'I' and a space: searching for the keyword
    // would cut this image off at its first pixel.
    const pdf = page('0 0 100 100 cm BI /W 3 /H 1 /BPC 8 /CS /G ID EI  EI', [], new Map());
    const { width, rgba } = pixels(pdf);
    expect(width).toBe(3);
    expect([rgba[0], rgba[4], rgba[8]]).toEqual([0x45, 0x49, 0x20]);
  });

  it('takes an inline image’s abbreviated filter name', () => {
    // /AHx is ASCIIHexDecode: three bytes written as six hex digits.
    const real = page(
      '0 0 100 100 cm BI /W 3 /H 1 /BPC 8 /CS /G /F /AHx ID 0080ff> EI',
      [],
      new Map(),
    );
    const { rgba } = pixels(real);
    expect([rgba[0], rgba[4], rgba[8]]).toEqual([0x00, 0x80, 0xff]);
  });
  it('reads /BlackIs1, which says which bit the fax filter hands on', () => {
    // §7.4.6 — the decoder codes black as 1; false (the default) means the
    // black pixel LEAVES as a 0, and true means it leaves as a 1, which
    // DeviceGray then reads as white. images_1bit_grayscale.pdf carries the
    // same picture twice, once encoded each way, and read without the flag the
    // second came back white on black.
    //
    // One all-white 8-pixel row, Group 4: the white-run code for 8 is 10011,
    // then the EOFB. Under BlackIs1 the same bits mean the row is black.
    const g4 = Uint8Array.from([0b1001_1000, 0b0000_0001, 0b0000_0000]);
    const shot = (blackIs1: boolean): number => {
      const pdf = page(
        '0 0 100 100 cm /Im Do',
        [
          '<< /Type /XObject /Subtype /Image /Width 8 /Height 1 /ColorSpace /DeviceGray ' +
            '/BitsPerComponent 1 /Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 ' +
            `/Rows 1 /BlackIs1 ${blackIs1 ? 'true' : 'false'} >> /Length ${String(g4.length)} >>`,
        ],
        new Map([[5, g4]]),
      );
      return pixels(pdf).rgba[0]!;
    };
    expect(shot(false)).toBe(255);
    expect(shot(true)).toBe(0);
  });
  it('samples a function-based shading a bare `sh` paints', () => {
    // §8.7.4.5.3 — type 1 is a function of TWO variables over a rectangle, and
    // no gradient stands for one; a picture does exactly.
    // function_based_shading.pdf is nine such squares and 43% of the page's
    // ink, and reconstructed to a blank sheet.
    const ps = '{ pop }'; // grey = x, whatever y is
    const content = 'q 10 10 100 100 re W n /Sh0 sh Q';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        '/Resources << /Shading << /Sh0 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>`,
      '<< /ShadingType 1 /ColorSpace /DeviceGray /Domain [0 1 0 1] ' +
        '/Matrix [100 0 0 100 10 10] /Function 6 0 R >>',
      `<< /FunctionType 4 /Domain [0 1 0 1] /Range [0 1] /Length ${String(ps.length)} >>`,
    ];
    const enc = (t: string): Uint8Array => Uint8Array.from([...t].map((c) => c.charCodeAt(0)));
    const bytes: Array<number> = [];
    const push = (t: string): void => {
      for (const ch of t) bytes.push(ch.charCodeAt(0) & 0xff);
    };
    push('%PDF-1.7\n');
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(bytes.length);
      push(`${String(i + 1)} 0 obj\n${body}\n`);
      const data = i === 3 ? enc(content) : i === 5 ? enc(ps) : undefined;
      if (data) {
        push('stream\n');
        for (const b of data) bytes.push(b);
        push('\nendstream\n');
      }
      push('endobj\n');
    });
    const xref = bytes.length;
    push(`xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`);
    for (const off of offsets) push(`${String(off).padStart(10, '0')} 00000 n \n`);
    push(
      `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`,
    );
    const file = PdfFile.parse(Uint8Array.from(bytes));
    const { images } = collectPageImages(file, file.pages()[0]!);
    expect(images).toHaveLength(1);
    // The `/Matrix` maps the unit domain onto 100pt at (10,10).
    expect([images[0]!.x, images[0]!.y, images[0]!.widthPt, images[0]!.heightPt]).toEqual([
      10, 10, 100, 100,
    ]);
    // …and the picture runs black on the left to white on the right.
    const { rgba, width } = decodePngPixels(images[0]!.bytes);
    expect(rgba[0]).toBeLessThan(8);
    expect(rgba[(width - 1) * 4]).toBeGreaterThan(247);
  });

  it('lifts the picture a Type 3 glyph PAINTS', () => {
    // §9.6.5 — a Type 3 glyph is a content stream, and a bitmap font's glyph is
    // a picture. This pass interpreted with an empty font map, so it never knew
    // a face was Type 3 and never saw a glyph call at all:
    // french_diacritics.pdf draws each accented letter as a 40×59 inline
    // stencil inside its glyph, and the page came back with nothing on it.
    const glyph = '10 0 0 0 10 10 d1 q 10 0 0 10 0 0 cm BI /W 2 /H 1 /IM true /BPC 1 ID \x40 EI Q';
    const content = 'BT /T3 10 Tf 20 100 Td (A) Tj ET';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        '/Resources << /Font << /T3 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>`,
      '<< /Type /Font /Subtype /Type3 /FontBBox [0 0 10 10] /FontMatrix [0.1 0 0 0.1 0 0] ' +
        '/CharProcs << /square 6 0 R >> /Encoding << /Differences [65 /square] >> ' +
        '/FirstChar 65 /LastChar 65 /Widths [10] >>',
      `<< /Length ${String(glyph.length)} >>`,
    ];
    const enc = (t: string): Uint8Array => Uint8Array.from([...t].map((c) => c.charCodeAt(0)));
    // Assembled by hand: `page` above wires its own fixed object list.
    const bytes: Array<number> = [];
    const push = (t: string): void => {
      for (const ch of t) bytes.push(ch.charCodeAt(0) & 0xff);
    };
    push('%PDF-1.7\n');
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(bytes.length);
      push(`${String(i + 1)} 0 obj\n${body}\n`);
      const data = i === 3 ? enc(content) : i === 5 ? enc(glyph) : undefined;
      if (data) {
        push('stream\n');
        for (const b of data) bytes.push(b);
        push('\nendstream\n');
      }
      push('endobj\n');
    });
    const xref = bytes.length;
    push(`xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`);
    for (const off of offsets) push(`${String(off).padStart(10, '0')} 00000 n \n`);
    push(
      `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`,
    );
    const file = PdfFile.parse(Uint8Array.from(bytes));
    expect(collectPageImages(file, file.pages()[0]!).images).toHaveLength(1);
  });

  it('runs a DeviceN image through its own tint transform', () => {
    // §8.6.6.4 — the space names its colorants, an alternate space and the
    // transform between them. Without running it the whole image is
    // unreadable and the page comes back BLANK: colorspace_sin.pdf is one
    // 256×256 picture in a `/DeviceN [/X /Y /Z]` whose transform is a
    // PostScript program, and that picture was the entire page.
    const ps = '{ 3 1 roll exch }'; // hands the components back reversed
    const pdf = page(
      '0 0 100 100 cm /Im Do',
      [
        '<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 ' +
          '/ColorSpace [/DeviceN [/X /Y /Z] /DeviceRGB 6 0 R] /Length 6 >>',
        `<< /FunctionType 4 /Domain [0 1 0 1 0 1] /Range [0 1 0 1 0 1] /Length ${String(ps.length)} >>`,
      ],
      new Map([
        [5, Uint8Array.from([255, 0, 0, 0, 0, 255])],
        [6, Uint8Array.from([...ps].map((c) => c.charCodeAt(0)))],
      ]),
    );
    const { rgba } = pixels(pdf);
    // (1,0,0) reversed is blue; (0,0,1) reversed is red.
    expect([...rgba.slice(0, 3)]).toEqual([0, 0, 255]);
    expect([...rgba.slice(4, 7)]).toEqual([255, 0, 0]);
  });

  it('reads a Lab image, and an Indexed palette whose base is one', () => {
    // §8.6.5.8 — `L*` 0..100 and `a*`/`b*` over the stated range, against the
    // stated white. issue10339_reduced.pdf paints two grids of blue swatches
    // through an Indexed palette whose base is one, and read as anything else
    // the page came back blank.
    //
    // Two 8-bit pixels: pure white (L*=100, a*=b*=0) and mid grey (L*≈54).
    const lab = Uint8Array.from([255, 128, 128, 138, 128, 128]);
    const pdf = page(
      '0 0 100 100 cm /Im Do',
      [
        '<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 ' +
          '/ColorSpace [/Lab << /WhitePoint [0.9505 1 1.089] /Range [-128 127 -128 127] >>] ' +
          `/Length ${String(lab.length)} >>`,
      ],
      new Map([[5, lab]]),
    );
    const { rgba } = pixels(pdf);
    // White stays white; the mid grey is neutral and close to sRGB's middle.
    expect([...rgba.slice(0, 3)]).toEqual([255, 255, 255]);
    expect(rgba[4]).toBe(rgba[5]);
    expect(rgba[5]).toBe(rgba[6]);
    expect(rgba[4]).toBeGreaterThan(110);
    expect(rgba[4]).toBeLessThan(150);
  });

  it('maps an Indexed image through /Decode, which gives INDEX values', () => {
    // §8.9.5.2 — an Indexed image's decode default is `[0 2^bpc − 1]`, so the
    // sample IS the index unless the file says otherwise.
    // issue10339_reduced.pdf draws its two grids from one palette, the second
    // through `/Decode [255 0]`, and read without it the two came back
    // identical instead of mirrored.
    const shot = (decode: string): Array<number> => {
      const pdf = page(
        '0 0 100 100 cm /Im Do',
        [
          '<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 ' +
            '/ColorSpace [/Indexed /DeviceRGB 1 <FF000000FF00>] ' +
            `${decode} /Length 2 >>`,
        ],
        new Map([[5, Uint8Array.from([0, 255])]]),
      );
      const { rgba } = pixels(pdf);
      return [rgba[0]!, rgba[1]!, rgba[4]!, rgba[5]!];
    };
    // The samples run the width of the 8-bit range, as an Indexed image's do:
    // 0 is the palette's first entry (red) and 255 its last (green).
    expect(shot('')).toEqual([255, 0, 0, 255]);
    // …and a decode that runs the other way swaps them.
    expect(shot('/Decode [255 0]')).toEqual([0, 255, 255, 0]);
  });
});
