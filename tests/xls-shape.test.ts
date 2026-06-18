// XLS-7 — drawing shapes (autoshapes + text boxes) in a legacy `.xls`. Non-picture
// Escher shapes become ShapeBlocks: the preset geometry comes from the MSOSPT
// shape type, the fill/line from the Escher OPT, and a text box's text from its
// TXO record (associated by order). They render through the same shape pipeline
// as xlsx drawings.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXls, msoDrawingShapesRec, numberRec, txoRecs } from './fixtures/build-xls';
import { readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';
import { Ream } from '@/core/converter/ream';

const shapeXls = (): Uint8Array =>
  buildXls({
    sheets: [
      {
        name: 'S',
        records: [
          numberRec(0, 0, 1),
          msoDrawingShapesRec([
            { shapeType: 3, fillRgb: [255, 0, 0] }, // ellipse, red fill
            { shapeType: 202, hasText: true }, // text box
          ]),
          ...txoRecs('Hello box'),
        ],
      },
    ],
  });

describe('xls drawing shapes (XLS-7)', () => {
  it('reads autoshapes with preset geometry + fill, and text boxes with their text', () => {
    const shapes = readXlsToSheetDoc(shapeXls()).sheets[0]?.shapes;
    expect(shapes).toHaveLength(2);

    expect(shapes![0]?.geometry).toMatchObject({ kind: 'preset', preset: 'ellipse' });
    expect(shapes![0]?.fill).toMatchObject({ kind: 'solid', colorHex: 'FF0000' });

    expect(shapes![1]?.geometry.preset).toBe('rect'); // text box → rectangle
    const runs = shapes![1]?.text?.content.flatMap((el) =>
      el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text) : [],
    );
    expect(runs).toEqual(['Hello box']);
  });

  it('projects the shapes as shape blocks after the grid', () => {
    const flow = projectSheetDoc(readXlsToSheetDoc(shapeXls()));
    expect(flow.body.filter((el) => el.kind === 'shape')).toHaveLength(2);
  });

  it('adds no shapes to a sheet without a drawing (byte-zero)', () => {
    const doc = readXlsToSheetDoc(
      buildXls({ sheets: [{ name: 'S', records: [numberRec(0, 0, 1)] }] }),
    );
    expect(doc.sheets[0]?.shapes).toBeUndefined();
  });

  it('renders an .xls with shapes to a valid PDF', async () => {
    const pdf = await Ream.parse(shapeXls()).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
