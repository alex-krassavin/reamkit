import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { ResourceStore, eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';
import { convertDocxToPdfSync } from '@/core/converter';
import { parseTtf } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { parseDocument } from '@/word';
import { readDocx } from '@/word/docx-reader';
import { detectImageFormat, embedImage } from '@/pdf';
import { PdfDocument } from '@/pdf/writer';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

function drawingXml(rId: string, cxEmu: number, cyEmu: number): string {
  return `<w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="${cxEmu}" cy="${cyEmu}"/>
      <wp:docPr id="1" name="Picture 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:blipFill>
              <a:blip r:embed="${rId}"/>
            </pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r>`;
}

describe('detectImageFormat', () => {
  it('recognises PNG magic bytes', () => {
    expect(
      detectImageFormat(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])),
    ).toBe('png');
  });

  it('recognises JPEG magic bytes', () => {
    expect(detectImageFormat(new Uint8Array([0xff, 0xd8, 0xff, 0xe0]))).toBe('jpeg');
  });

  it('returns null for non-image bytes', () => {
    expect(detectImageFormat(new Uint8Array([0x00, 0x01, 0x02, 0x03]))).toBeNull();
  });

  it('recognises a JPEG 2000 JP2 box signature and a raw codestream', () => {
    const jp2Sig = new Uint8Array([
      0x00, 0x00, 0x00, 0x0c, 0x6a, 0x50, 0x20, 0x20, 0x0d, 0x0a, 0x87, 0x0a,
    ]);
    expect(detectImageFormat(jp2Sig)).toBe('jpeg2000');
    expect(detectImageFormat(new Uint8Array([0xff, 0x4f, 0xff, 0x51]))).toBe('jpeg2000');
  });
});

// Minimal JP2: signature box + jp2h{ ihdr(height=64, width=128) }. No codestream
// (we pass the bytes through to /JPXDecode and only read the dimensions).
const JP2_64x128 = new Uint8Array([
  0x00,
  0x00,
  0x00,
  0x0c,
  0x6a,
  0x50,
  0x20,
  0x20,
  0x0d,
  0x0a,
  0x87,
  0x0a, // signature
  0x00,
  0x00,
  0x00,
  0x1e,
  0x6a,
  0x70,
  0x32,
  0x68, // jp2h box (len 30)
  0x00,
  0x00,
  0x00,
  0x16,
  0x69,
  0x68,
  0x64,
  0x72, // ihdr box (len 22)
  0x00,
  0x00,
  0x00,
  0x40, // HEIGHT = 64
  0x00,
  0x00,
  0x00,
  0x80, // WIDTH = 128
  0x00,
  0x01,
  0x07,
  0x07,
  0x00,
  0x00, // NC=1, BPC, C, UnkC, IPR
]);

describe('JPEG 2000 embedding (/JPXDecode pass-through)', () => {
  it('reads ihdr dimensions and emits a JPXDecode image XObject', () => {
    const doc = new PdfDocument();
    const img = embedImage(doc, JP2_64x128);
    expect(img.widthPx).toBe(128);
    expect(img.heightPx).toBe(64);
    const pdf = new TextDecoder('latin1').decode(doc.build(img.ref));
    expect(pdf).toContain('/Filter /JPXDecode');
    expect(pdf).toContain('/Width 128');
    expect(pdf).toContain('/Height 64');
  });
});

describe('Drawing parser', () => {
  it('produces a BodyElement of kind=image when a w:drawing is present', () => {
    const body = `<w:p>${drawingXml('rId20', 914400, 685800)}</w:p>`;
    const docx = buildDocxFromBody(body);
    const pkg = OpcPackage.open(docx);
    // The parser resolves drawing relationship ids through the supplied
    // ImageResolver into content-addressed ResourceIds.
    const store = new ResourceStore();
    const expectedId = store.put(new Uint8Array([1, 2, 3]));
    const parsed = parseDocument(pkg.getMainDocument().data, {
      resolveColor: defaultColorResolver,
      resolveImage: (relId) => (relId === 'rId20' ? expectedId : undefined),
    });
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('image');
    if (parsed[0]!.kind !== 'image') throw new Error('unreachable');
    expect(parsed[0]!.image.resource).toBe(expectedId);
    expect(parsed[0]!.image.width).toBe(emuToPt(914400));
    expect(parsed[0]!.image.height).toBe(emuToPt(685800));
  });
});

describe('Image rendering end-to-end', () => {
  it('emits an XObject Image and Do operator for an embedded PNG', () => {
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]); // 2×2 red
    const body = `
      <w:p><w:r><w:t>Before image</w:t></w:r></w:p>
      <w:p>${drawingXml('rId20', 914400, 914400)}</w:p>
      <w:p><w:r><w:t>After image</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toContain('/Type /XObject');
    expect(text).toContain('/Subtype /Image');
    expect(text).toContain('/Filter /FlateDecode');
    expect(text).toMatch(/\/Im\d+ Do/);
    expect(text).toMatch(/\d+(\.\d+)? 0 0 \d+(\.\d+)? \d+(\.\d+)? \d+(\.\d+)? cm/);
  });

  it('preserves text and renders image when a paragraph mixes both', () => {
    const png = buildTinyPng(4, 4, [0, 200, 0, 255]); // small green square
    // Paragraph with text + inline image + more text in the same w:p.
    const body = `
      <w:p>
        <w:r><w:t xml:space="preserve">Before </w:t></w:r>
        ${drawingXml('rId20', 304800, 304800)}
        <w:r><w:t xml:space="preserve"> After</w:t></w:r>
      </w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hexOf = (s: string) =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');

    // Both text segments must be present (not lost to image-block collapse).
    expect(text).toContain(`<${hexOf('Before')}> Tj`);
    expect(text).toContain(`<${hexOf('After')}> Tj`);
    // Image XObject is also drawn.
    expect(text).toMatch(/\/Im\d+ Do/);
    // ET / BT sequence indicates we left text mode for the inline image.
    expect(text).toMatch(/ET\nq\n[^\n]+cm\n\/Im\d+ Do\nQ\nBT/);
  });

  it('image XObject lists the right Width and Height', () => {
    const png = buildTinyPng(3, 5, [0, 128, 255, 255]);
    const body = `<w:p>${drawingXml('rId20', 914400, 1828800)}</w:p>`;
    const docx = buildDocxFromBody(body, {
      images: { rId20: { contentType: 'image/png', bytes: png, extension: 'png' } },
    });
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toMatch(/\/Width 3/);
    expect(text).toMatch(/\/Height 5/);
  });
});

describe('image robustness', () => {
  it('skips an unsupported/corrupt image instead of crashing the document', () => {
    // Garbage bytes labelled as PNG — embedImage throws; the document must still
    // render (the bad image is simply omitted, no dangling XObject reference).
    const body =
      `<w:p><w:r><w:t>Before image.</w:t></w:r></w:p>` +
      `<w:p>${drawingXml('rId99', 914400, 914400)}</w:p>` +
      `<w:p><w:r><w:t>After image.</w:t></w:r></w:p>`;
    const opts = {
      images: {
        rId99: {
          contentType: 'image/png' as const,
          bytes: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]),
          extension: 'png' as const,
        },
      },
    };
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body, opts), { fonts: FONTS });
    const text = asLatin1(pdf);
    // Conversion succeeded and produced a valid PDF...
    expect(text.startsWith('%PDF')).toBe(true);
    // ...with no image drawn (no `/Im… Do`, and crucially no dangling `/ Do`).
    expect(text).not.toMatch(/\/Im\d* Do/);
    expect(text).not.toMatch(/\/ Do/);
    // The surrounding text still rendered (font subset embedded).
    expect(text).toMatch(/\/BaseFont \/[A-Za-z]/);
  });
});

describe('per-part image resolution (C5)', () => {
  it("resolves a header image through the header's own rels, not the main part's", () => {
    // Same rId in both parts, pointing at DIFFERENT images: blue in the
    // header's rels, red in the main document's. The header must get blue.
    const bluePng = buildTinyPng(2, 2, [0, 0, 255, 255]);
    const redPng = buildTinyPng(2, 2, [255, 0, 0, 255]);
    const body =
      '<w:p><w:r><w:t>body</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/></w:sectPr>';
    const docx = buildDocxFromBody(body, {
      headerXml: `<w:p>${drawingXml('rId20', 190500, 190500)}</w:p>`,
      headerImages: { rId20: { contentType: 'image/png', bytes: bluePng, extension: 'png' } },
      images: { rId20: { contentType: 'image/png', bytes: redPng, extension: 'png' } },
    });
    const { doc } = readDocx(docx);
    const header = [...(doc.headersFooters?.values() ?? [])][0];
    expect(header).toBeDefined();
    const img = header!.find((el) => el.kind === 'image');
    expect(img).toBeDefined();
    if (img?.kind !== 'image') throw new Error('unreachable');
    const bytes = doc.resources.get(img.image.resource!);
    expect(bytes && Buffer.from(bytes).equals(Buffer.from(bluePng))).toBe(true);
  });
});
