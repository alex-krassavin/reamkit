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
    const pattern = file.resolve((patterns as PdfDict).get('Sh0')!);
    expect((pattern as PdfDict).get('PatternType')).toBe(2);
    const shading = file.resolve((pattern as PdfDict).get('Shading')!);
    expect((shading as PdfDict).get('ShadingType')).toBe(2); // axial (a:lin)
    const fn = file.resolve((shading as PdfDict).get('Function')!);
    expect((fn as PdfDict).get('C0')).toEqual([1, 0, 0]); // FF0000
    expect((fn as PdfDict).get('C1')).toEqual([0, 0, 1]); // 0000FF
    const content = new TextDecoder('latin1').decode(file.pageContent(page));
    expect(content).toContain('/Pattern cs');
    expect(content).toContain('/Sh0 scn');
  });

  it('holds a stop that does not start at the edge where the file puts it', async () => {
    // §20.1.8.36 — a ramp whose first stop sits a fifth of the way along holds
    // its first colour flat for that fifth and only then begins. Dragged back
    // to the edge, the whole ramp stretches over it: themes.pptx's last slide
    // states exactly that and its background came out a bucket too dark across
    // the middle of the page.
    const offset =
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst>' +
      '<a:gs pos="20000"><a:srgbClr val="FF0000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
      '</a:gsLst><a:lin ang="0"/></a:gradFill>';
    const pdf = await Ream.parse(buildDocxFromBody(`<w:p>${shapeRun(offset)}</w:p>`)).convert(
      'pdf',
      { fonts: FONTS },
    );
    const file = PdfFile.parse(pdf);
    const page = file.pages()[0]!;
    const patterns = file.get(page.resources!, 'Pattern') as PdfDict;
    const pattern = file.resolve(patterns.get('Sh0')!) as PdfDict;
    const shading = file.resolve(pattern.get('Shading')!) as PdfDict;
    const fn = file.resolve(shading.get('Function')!) as PdfDict;
    // A stitching function: red held to 0.2, then red → blue over the rest.
    expect(fn.get('FunctionType')).toBe(3);
    expect(fn.get('Bounds')).toEqual([0.2]);
    const parts = (fn.get('Functions') as Array<PdfValue>).map((f) => file.resolve(f) as PdfDict);
    expect(parts[0]?.get('C0')).toEqual([1, 0, 0]);
    expect(parts[0]?.get('C1')).toEqual([1, 0, 0]);
    expect(parts[1]?.get('C1')).toEqual([0, 0, 1]);
  });

  it('reads a PDF shading pattern back into a gradient fill (E-PDF EP16c)', async () => {
    const pdf = await Ream.parse(gradientDocx()).convert('pdf', { fonts: FONTS });
    const back = Ream.parse(pdf);
    const shape = back.flow.body.find(
      (el) => el.kind === 'shape' && el.shape.fill.kind === 'gradient',
    );
    expect(shape?.kind).toBe('shape');
    if (shape?.kind !== 'shape') return;
    const g = shape.shape.fill.gradient;
    expect(g?.kind).toBe('linear');
    const stops = g?.stops ?? [];
    expect(stops[0]?.colorHex).toBe('FF0000');
    expect(stops[stops.length - 1]?.colorHex).toBe('0000FF');
  });

  // §20.1.2.3.1 — a stop's own transparency, read as transparency rather than as
  // a colour washed toward white. The glow on tdf123684's master is three stops
  // at 7%, 6% and nothing, and composited over the paper it was an opaque white
  // disc on a dark slide.
  it('reads a stop’s alpha as the fill’s transparency, colour unwashed', () => {
    const faint =
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst>' +
      '<a:gs pos="0"><a:srgbClr val="80C0FF"><a:alpha val="7000"/></a:srgbClr></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="80C0FF"><a:alpha val="0"/></a:srgbClr></a:gs>' +
      '</a:gsLst><a:path path="circle"/></a:gradFill>';
    const parsed = parseDocument(
      OpcPackage.open(buildDocxFromBody(`<w:p>${shapeRun(faint)}</w:p>`)).getMainDocument().data,
    );
    const el = parsed.find((e) => e.kind === 'shape');
    expect(el?.kind).toBe('shape');
    if (el?.kind !== 'shape') return;
    const fill = el.shape.fill;
    expect(fill.kind).toBe('gradient');
    // The strongest stop decides whether the shape is there at all.
    expect(fill.alpha).toBeCloseTo(0.07, 5);
    expect(fill.gradient?.stops[0]?.alpha).toBeCloseTo(0.07, 5);
    expect(fill.gradient?.stops[1]?.alpha).toBe(0);
    // …and the colour is the stop's own, not a 7% wash of it over white. The
    // recovery is off by a couple of levels — a 7% composite lands every colour
    // within four levels of white, and coming back out multiplies the rounding
    // by fourteen — which is invisible at the opacity it is drawn with.
    const stop = fill.gradient?.stops[0]?.colorHex ?? '';
    for (const [i, want] of [0x80, 0xc0, 0xff].entries()) {
      expect(Math.abs(parseInt(stop.slice(i * 2, i * 2 + 2), 16) - want)).toBeLessThan(10);
    }
  });

  // §11.6.5.2 — a PDF shading has one colour per point and no alpha, so a
  // gradient that FADES OUT is painted through a luminosity mask of the same
  // sweep: white where the stop is opaque, black where it is clear. Painted
  // flat, tdf123684's 7%-and-fading glow was an opaque disc on a dark slide.
  it('paints a fading gradient through a luminosity soft mask', async () => {
    const faint =
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst>' +
      '<a:gs pos="0"><a:srgbClr val="80C0FF"><a:alpha val="7000"/></a:srgbClr></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="80C0FF"><a:alpha val="0"/></a:srgbClr></a:gs>' +
      '</a:gsLst><a:path path="circle"/></a:gradFill>';
    const pdf = await Ream.parse(buildDocxFromBody(`<w:p>${shapeRun(faint)}</w:p>`)).convert(
      'pdf',
      { fonts: FONTS },
    );
    const text = new TextDecoder('latin1').decode(pdf);
    expect(text).toContain('/Luminosity');
    expect(text).toMatch(/\/SMask/u);
    // …and an opaque gradient needs none.
    const plain = await Ream.parse(gradientDocx()).convert('pdf', { fonts: FONTS });
    expect(new TextDecoder('latin1').decode(plain)).not.toContain('/Luminosity');
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
