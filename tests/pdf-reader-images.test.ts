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
import { decodePdfImage } from '@/pdf-reader/image-decode';
import { encodePng } from '@/pdf-reader/png-encode';
import { PdfHexString, dict, name, stream } from '@/pdf/objects';

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

  it('still reports vector graphics as an unreconstructed loss', async () => {
    const pdf = await Ream.parse(docxWithImage()).convert('pdf', { fonts: FONTS });
    const doc = Ream.parse(pdf);
    expect(doc.losses.some((l) => /vector/i.test(l.detail))).toBe(true);
  });
});
