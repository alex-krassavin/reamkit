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
import { PdfHexString, name, stream } from '@/pdf/objects';

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
