// E-SHEET W2 — floating shapes on a worksheet. An xdr:sp anchor renders as a
// ShapeBlock (geometry / fill / line / text), reusing the DrawingML readers the
// pptx + SmartArt paths use. The shape's box comes from its sheet anchor; runs
// use their direct a:rPr formatting (no placeholder cascade). It projects to a
// shape block after the grid — PDF/HTML draw it through the existing shape path.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const shapeXlsx = () =>
  buildXlsx({ rows: [['cell']], sheetShape: { text: 'Shape text', fillHex: '4472C4' } });

describe('sheet shapes — resolve (E-SHEET W2)', () => {
  it('resolves an xdr:sp anchor into a ShapeBlock with geometry, fill, line and text', () => {
    const sheet = readXlsxToSheetDoc(shapeXlsx()).sheets[0]!;
    expect(sheet.shapes).toHaveLength(1);
    const shape = sheet.shapes![0]!;
    expect(shape.width).toBeGreaterThan(0);
    expect(shape.height).toBeGreaterThan(0);
    expect(shape.geometry).toMatchObject({ kind: 'preset', preset: 'roundRect' });
    expect(shape.fill).toMatchObject({ kind: 'solid', colorHex: '4472C4' });
    expect(shape.line).toBeDefined();
    const para = shape.text?.content[0];
    expect(para?.kind).toBe('paragraph');
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(para.paragraph.runs[0]?.text).toBe('Shape text');
  });

  it('leaves a sheet with no drawing without a shapes field', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]] })).sheets[0]!;
    expect(sheet.shapes).toBeUndefined();
  });

  it('skips a chart-only drawing (no shapes field)', () => {
    const chartXml = `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]], sheetChart: { chartXml } }))
      .sheets[0]!;
    expect(sheet.shapes).toBeUndefined();
  });
});

describe('sheet shapes — projection (E-SHEET W2)', () => {
  it('projects the shape as a shape block after the grid', () => {
    const body = Ream.parse(shapeXlsx()).flow.body;
    const shape = body.find((el) => el.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('expected a shape block');
    expect(shape.shape.fill).toMatchObject({ kind: 'solid', colorHex: '4472C4' });
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('sheet shapes — render (E-SHEET W2)', () => {
  it('draws the shape fill and text into HTML', async () => {
    const html = new TextDecoder().decode(await Ream.parse(shapeXlsx()).convert('html'));
    expect(html).toContain('#4472C4'); // the shape fill
    expect(html).toContain('Shape text'); // the shape's text body
  });

  it('renders a sheet with a shape to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(shapeXlsx(), { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
