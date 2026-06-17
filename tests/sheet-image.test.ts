// E-SHEET W1 — floating pictures on a worksheet. The sheet drawing part used to
// yield charts only; `xdr:pic` anchors now resolve their media bytes into the
// SheetDoc resource store and project to an ImageBlock after the grid (like
// chart frames). PDF embeds an image XObject; HTML emits a data URI.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

// A 1×1 transparent PNG — the smallest valid image the decoder accepts.
const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
);

const imageXlsx = () => buildXlsx({ rows: [['cell']], sheetImage: { pngBytes: PNG_1x1 } });

describe('sheet pictures — resolve (E-SHEET W1)', () => {
  it('resolves an anchored picture into Sheet.images with its bytes', () => {
    const doc = readXlsxToSheetDoc(imageXlsx());
    const sheet = doc.sheets[0]!;
    expect(sheet.images).toHaveLength(1);
    const img = sheet.images![0]!;
    expect(img.widthPt).toBeGreaterThan(0);
    expect(img.heightPt).toBeGreaterThan(0);
    // The bytes are content-addressed into the SheetDoc resource store.
    expect(doc.resources.get(img.resourceId)).toEqual(PNG_1x1);
  });

  it('leaves a sheet with no drawing without an images field', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]] })).sheets[0]!;
    expect(sheet.images).toBeUndefined();
  });
});

describe('sheet pictures — projection (E-SHEET W1)', () => {
  it('projects the picture as an image block carrying the resource', () => {
    const flow = Ream.parse(imageXlsx()).flow;
    const image = flow.body.find((el) => el.kind === 'image');
    if (image?.kind !== 'image') throw new Error('expected an image block');
    expect(image.image.resource).toBeDefined();
    expect(flow.resources.get(image.image.resource!)).toEqual(PNG_1x1);
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('sheet pictures — render (E-SHEET W1)', () => {
  it('emits the picture as a PNG data URI in HTML', async () => {
    const html = new TextDecoder().decode(await Ream.parse(imageXlsx()).convert('html'));
    expect(html).toContain('data:image/png;base64,');
  });

  it('renders a sheet with a picture to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(imageXlsx(), { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
