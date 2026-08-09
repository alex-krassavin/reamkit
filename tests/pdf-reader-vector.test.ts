// E-PDF EP10 — filled vector paths. The interpreter captures a `re … f` fill
// with its colour; end-to-end, a filled docx shape comes back as a solid-fill
// shape when the (untagged) PDF is read.

import { readFileSync } from 'node:fs';

import { zlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { interpretContent } from '@/pdf-reader/content';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';

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

/**
 * A one-page PDF that fills a black square and then draws a picture over it —
 * the shape a legend swatch with an icon on it takes.
 */
function fillThenImage(): Uint8Array {
  const gray = zlibSync(Uint8Array.from([0, 255, 255, 0])); // 2x2 DeviceGray
  const content = 'q 0 0 0 rg 100 100 200 200 re f 50 0 0 50 150 150 cm /Im Do Q';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 400] ' +
      '/Resources << /XObject << /Im 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    `<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray ` +
      `/BitsPerComponent 8 /Filter /FlateDecode /Length ${String(gray.length)} >>`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\n`;
    if (i === 4) pdf += `stream\n${String.fromCharCode(...gray)}\nendstream\n`;
    pdf += 'endobj\n';
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return Uint8Array.from([...pdf].map((c) => c.charCodeAt(0)));
}

describe('painting order (§8.5.3)', () => {
  it('lays a picture over the path it was drawn after', () => {
    // Pictures under paths loses an icon sitting on its swatch; paths under
    // pictures loses a white box backing a legend. Only the page's own order
    // gets both, and 22060_A1_01_Plans.pdf has one of each.
    const { doc } = reconstructByLayout(PdfFile.parse(fillThenImage()));
    const shape = doc.body.find((b) => b.kind === 'shape');
    const image = doc.body.find((b) => b.kind === 'image');
    expect(shape?.kind).toBe('shape');
    expect(image?.kind).toBe('image');
    if (shape?.kind !== 'shape' || image?.kind !== 'image') return;
    expect(image.image.float?.zOrder).toBeGreaterThan(shape.shape.float?.zOrder ?? 0);
  });

  it('numbers paths and XObject calls in one sequence', () => {
    // Later marks cover earlier ones, and a form is drawn where its `Do`
    // stands. Collected apart, every form ends up on top: 22060_A1_01_Plans.pdf
    // backs its legend with a white box inside a form, and hoisted to the end
    // that box covered the legend's own words.
    const { vectors, images } = interpretContent(
      new TextEncoder().encode('0 0 0 rg 0 0 10 10 re f /Fm Do 0 0 20 20 re f'),
      NO_FONTS,
    );
    expect(vectors).toHaveLength(2);
    expect(images).toHaveLength(1);
    // The form's call falls BETWEEN the two fills, which is where it paints.
    expect(vectors[0]!.order).toBeLessThan(images[0]!.order);
    expect(images[0]!.order).toBeLessThan(vectors[1]!.order);
  });
});

describe('tiling patterns (§8.7.3)', () => {
  const fills = (stream: string) =>
    interpretContent(new TextEncoder().encode(stream), NO_FONTS).vectors;

  it('names the tiling pattern a path is filled with instead of inventing a colour', () => {
    // §8.6.6.2 — in a Pattern colour space `scn` takes a NAME. Read as a
    // colour it left whatever was set before it standing: 22060_A1_01_Plans.pdf
    // filled four floor plans with a pattern and we painted four black squares.
    const [painted] = fills('0 0 0 rg /Pat cs /P1 scn 0 0 100 100 re f');
    expect(painted?.patternName).toBe('P1');
  });

  it('forgets the pattern when a real colour is set after it', () => {
    const [painted] = fills('/Pat cs /P1 scn 1 0 0 rg 0 0 100 100 re f');
    expect(painted?.patternName).toBeUndefined();
    expect(painted?.fillHex).toBe('FF0000');
  });

  it('names no pattern for an ordinary colour fill', () => {
    const [painted] = fills('0 0 1 rg 0 0 100 100 re f');
    expect(painted?.patternName).toBeUndefined();
  });
});

describe('clipping paths (§8.5.4)', () => {
  const clipOf = (stream: string) =>
    interpretContent(new TextEncoder().encode(stream), NO_FONTS).vectors;

  it('installs the clip at the painting operator that ends its path', () => {
    // `W` names the clip; `n` ends the path and installs it. The fill that
    // follows is painted through it.
    const [painted] = clipOf('q 100 100 50 50 re W n 0 0 0 rg 0 0 500 500 re f Q');
    expect(painted?.clip).toBeDefined();
    expect(painted?.clip?.minX).toBe(100);
    expect(painted?.clip?.maxX).toBe(150);
  });

  it('leaves a path painted under no clip unclipped', () => {
    const [painted] = clipOf('0 0 0 rg 10 20 30 40 re f');
    expect(painted?.clip).toBeUndefined();
  });

  it('restores the clip a Q pops', () => {
    // §8.4.2 — the clip belongs to the graphics state, so `Q` takes it back.
    const painted = clipOf('q 0 0 10 10 re W n Q 0 0 0 rg 0 0 500 500 re f');
    expect(painted[painted.length - 1]?.clip).toBeUndefined();
  });

  it('keeps the smaller of two nested clips', () => {
    const [painted] = clipOf('q 0 0 400 400 re W n 10 10 20 20 re W n 0 0 0 rg 0 0 500 500 re f Q');
    expect(painted?.clip?.maxX).toBe(30);
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
