// XLS-5 — embedded images in a legacy `.xls`. Pictures live in the Office Drawing
// (Escher) layer: the workbook globals hold the BLIP store (MSODrawingGroup) and
// each sheet's MSODrawing references it by index. The reader pulls the image
// bytes into the shared resource store and renders them like xlsx pictures.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXls, msoDrawingGroupRec, msoDrawingRec, numberRec } from './fixtures/build-xls';
import { parseBlipStore, parseSheetPictures } from '@/excel/xls/escher';
import { readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { Ream } from '@/core/converter/ream';

// A 1×1 transparent PNG — the smallest image the decoder accepts.
const PNG_1x1 = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
);

// Unwrap a BIFF record (4-byte header) → its Escher payload.
const escherOf = (biffRec: Uint8Array): Uint8Array => biffRec.subarray(4);

describe('Escher drawing parser (XLS-5)', () => {
  it('extracts image bytes from the BLIP store, index-aligned', () => {
    const blips = parseBlipStore(escherOf(msoDrawingGroupRec([PNG_1x1])));
    expect(blips).toEqual([PNG_1x1]);
  });

  it('reads picture shapes (BLIP index + anchor) from a sheet drawing', () => {
    const pics = parseSheetPictures(escherOf(msoDrawingRec([{ blipIndex: 1, col2: 4, row2: 6 }])));
    expect(pics).toHaveLength(1);
    expect(pics[0]?.blipIndex).toBe(1);
    expect(pics[0]?.anchor).toMatchObject({ col1: 0, row1: 0, col2: 4, row2: 6 });
  });
});

describe('xls embedded images — end to end (XLS-5)', () => {
  const imageXls = (): Uint8Array =>
    buildXls({
      styleRecords: [msoDrawingGroupRec([PNG_1x1])],
      sheets: [{ name: 'S', records: [numberRec(0, 0, 1), msoDrawingRec([{ blipIndex: 1 }])] }],
    });

  it('resolves the picture into the resource store and sizes it from the anchor', () => {
    const doc = readXlsToSheetDoc(imageXls());
    const img = doc.sheets[0]?.images?.[0];
    expect(img).toBeDefined();
    expect(doc.resources.get(img!.resourceId)).toEqual(PNG_1x1);
    expect(img!.widthPt).toBeGreaterThan(0);
    expect(img!.heightPt).toBeGreaterThan(0);
  });

  it('projects the picture as an image block after the grid', () => {
    const flow = projectSheetDoc(readXlsToSheetDoc(imageXls()));
    expect(flow.body.some((el) => el.kind === 'image')).toBe(true);
  });

  it('adds no images to a sheet without a drawing (byte-zero)', () => {
    const doc = readXlsToSheetDoc(
      buildXls({ sheets: [{ name: 'S', records: [numberRec(0, 0, 1)] }] }),
    );
    expect(doc.sheets[0]?.images).toBeUndefined();
  });

  it('renders an .xls with an embedded image to a valid PDF', async () => {
    const pdf = await Ream.parse(imageXls()).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
