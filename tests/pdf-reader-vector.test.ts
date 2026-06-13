// E-PDF EP10 — filled vector paths. The interpreter captures a `re … f` fill
// with its colour; end-to-end, a filled docx shape comes back as a solid-fill
// shape when the (untagged) PDF is read.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { interpretContent } from '@/pdf-reader/content';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const NO_FONTS = new Map();

// A wps:wsp rectangle filled solid C0504D.
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
const RECT =
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill>';
// An outline-only rectangle: no fill, a 1.5pt blue stroke.
const STROKED =
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
  '<a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>';

describe('filled vector paths (E-PDF EP10)', () => {
  it('captures a filled rectangle and its colour from the content stream', () => {
    const { vectors } = interpretContent(
      new TextEncoder().encode('1 0 0 rg 10 20 30 40 re f'),
      NO_FONTS,
    );
    expect(vectors).toHaveLength(1);
    expect(vectors[0]!.fillHex).toBe('FF0000');
    expect(vectors[0]!.strokeHex).toBeUndefined();
    expect(vectors[0]!.segs[0]).toMatchObject({ op: 'move', x: 10, y: 20 });
  });

  it('lifts a filled docx shape back out of an untagged PDF', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(`<w:p>${shapeRun(RECT)}</w:p>`)).convert('pdf', {
      fonts: FONTS,
    });
    const shape = Ream.parse(pdf).flow.body.find((el) => el.kind === 'shape');
    expect(shape).toBeDefined();
    if (shape?.kind !== 'shape') return;
    expect(shape.shape.fill.kind).toBe('solid');
    expect(shape.shape.fill.colorHex).toMatch(/^[0-9A-F]{6}$/);
    expect(shape.shape.geometry.kind).toBe('custom');
  });
});

describe('stroked vector paths (E-PDF EP11)', () => {
  it('captures a stroke-only path with its colour and CTM-scaled width', () => {
    const { vectors } = interpretContent(
      new TextEncoder().encode('1 0 0 RG 3 w 10 20 m 40 60 l S'),
      NO_FONTS,
    );
    expect(vectors).toHaveLength(1);
    expect(vectors[0]!.fillHex).toBeUndefined();
    expect(vectors[0]!.strokeHex).toBe('FF0000');
    expect(vectors[0]!.lineWidth).toBe(3);
  });

  it('captures fill and stroke together for the B operator', () => {
    const { vectors } = interpretContent(
      new TextEncoder().encode('0 1 0 rg 1 0 0 RG 2 w 10 20 30 40 re B'),
      NO_FONTS,
    );
    expect(vectors).toHaveLength(1);
    expect(vectors[0]!.fillHex).toBe('00FF00');
    expect(vectors[0]!.strokeHex).toBe('FF0000');
  });

  it('scales the line width by the CTM', () => {
    const { vectors } = interpretContent(
      new TextEncoder().encode('2 0 0 2 0 0 cm 0 0 0 RG 1 w 10 20 m 40 60 l S'),
      NO_FONTS,
    );
    expect(vectors[0]!.lineWidth).toBe(2); // 1 user-space unit × scale 2
  });

  it('lifts a stroked docx shape back out of an untagged PDF as a line', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(`<w:p>${shapeRun(STROKED)}</w:p>`)).convert(
      'pdf',
      { fonts: FONTS },
    );
    const lined = Ream.parse(pdf).flow.body.find(
      (el) => el.kind === 'shape' && el.shape.line !== undefined,
    );
    expect(lined?.kind).toBe('shape');
    if (lined?.kind !== 'shape') return;
    expect(lined.shape.line?.colorHex).toMatch(/^[0-9A-F]{6}$/);
    expect(lined.shape.fill.kind).toBe('none');
  });
});
