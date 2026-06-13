// E-PDF EP16 — gradient fills. A docx a:gradFill is parsed into a real gradient
// (not averaged to a solid) and rendered faithfully: SVG and HTML emit a
// <linearGradient>/<radialGradient>, docx round-trips the a:gradFill, and the
// PDF/layout path falls back to the gradient's solid average (byte-stable).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { OpcPackage } from '@/core/opc';
import { PdfFile } from '@/pdf-reader/document';
import { parseDocument } from '@/word/document-parser';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const shapeRun = (spPrInner: string): string =>
  `<w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Shape 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>${spPrInner}</wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r>`;

const GRAD =
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:gradFill><a:gsLst>' +
  '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
  '</a:gsLst><a:lin ang="0"/></a:gradFill>';

const gradientDocx = (): Uint8Array => buildDocxFromBody(`<w:p>${shapeRun(GRAD)}</w:p>`);

describe('gradient fills (E-PDF EP16)', () => {
  it('renders a gradient shape as an SVG linearGradient with its stops', async () => {
    const svg = new TextDecoder().decode(
      await Ream.parse(gradientDocx()).convert('svg', { fonts: FONTS }),
    );
    expect(svg).toContain('<linearGradient');
    expect(svg).toContain('stop-color="#FF0000"');
    expect(svg).toContain('stop-color="#0000FF"');
    expect(svg).toContain('fill="url(#grad0)"');
  });

  it('renders a gradient shape as an HTML inline-SVG gradient', async () => {
    const html = new TextDecoder().decode(await Ream.parse(gradientDocx()).convert('html'));
    expect(html).toContain('<linearGradient');
    expect(html).toContain('stop-color="#FF0000"');
    expect(html).toContain('stop-color="#0000FF"');
  });

  it('round-trips the gradient through docx', async () => {
    const out = await Ream.parse(gradientDocx()).convert('docx');
    const reparsed = parseDocument(OpcPackage.open(out).getMainDocument().data);
    const shape = reparsed.find((el) => el.kind === 'shape');
    expect(shape?.kind).toBe('shape');
    if (shape?.kind !== 'shape') return;
    expect(shape.shape.fill.kind).toBe('gradient');
    expect(shape.shape.fill.gradient?.stops).toEqual([
      { offset: 0, colorHex: 'FF0000' },
      { offset: 1, colorHex: '0000FF' },
    ]);
  });

  it('emits a gradient as a PDF axial shading pattern (E-PDF EP16b)', async () => {
    const pdf = await Ream.parse(gradientDocx()).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    const page = file.pages()[0]!;
    const patterns = file.get(page.resources!, 'Pattern');
    expect(patterns instanceof Map).toBe(true);
    const pattern = file.resolve((patterns as PdfDict).get('Sh0') ?? null);
    expect((pattern as PdfDict).get('PatternType')).toBe(2);
    const shading = file.resolve((pattern as PdfDict).get('Shading') ?? null);
    expect((shading as PdfDict).get('ShadingType')).toBe(2); // axial (a:lin)
    const fn = file.resolve((shading as PdfDict).get('Function') ?? null);
    expect((fn as PdfDict).get('C0')).toEqual([1, 0, 0]); // FF0000
    expect((fn as PdfDict).get('C1')).toEqual([0, 0, 1]); // 0000FF
    const content = new TextDecoder('latin1').decode(file.pageContent(page));
    expect(content).toContain('/Pattern cs');
    expect(content).toContain('/Sh0 scn');
  });

  it('keeps the solid fallback under PDF/A (no shading pattern)', async () => {
    const pdf = await Ream.parse(gradientDocx()).convert('pdf', {
      fonts: FONTS,
      pdfA: 'PDF/A-2b',
    });
    const file = PdfFile.parse(pdf);
    const page = file.pages()[0]!;
    const patterns: PdfValue = file.get(page.resources!, 'Pattern');
    expect(patterns instanceof Map).toBe(false); // no /Pattern resource
  });
});
