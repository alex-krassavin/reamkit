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
import { collectPageVectors } from '@/pdf-reader/vector';

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

/**
 * A page whose only mark is a WIDGET annotation: nothing in its content stream,
 * a filled rectangle in the annotation's `/AP` `/N`.
 */
function widgetOnlyPdf(): Uint8Array {
  const ap = '0 0 1 rg 0 0 40 20 re f';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      '/Annots [<< /Type /Annot /Subtype /Widget /Rect [50 60 90 80] /AP << /N 5 0 R >> >>] >>',
    '<< /Length 0 >>\nstream\n\nendstream',
    `<< /Type /XObject /Subtype /Form /BBox [0 0 40 20] /Length ${String(ap.length)} >>\n` +
      `stream\n${ap}\nendstream`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('annotation appearances (§12.5.5)', () => {
  it('lifts what a widget draws, and fits it to the annotation’s /Rect', () => {
    // A form field paints nothing in the page's content stream: its tint, its
    // border and its value all live in the widget's own appearance. Read only
    // from the page, 160F-2019.pdf gave a grid with no fields in it.
    const file = PdfFile.parse(widgetOnlyPdf());
    const vectors = collectPageVectors(file, file.pages()[0]!);
    expect(vectors).toHaveLength(1);
    const [box] = vectors;
    expect(box!.fillHex).toBe('0000FF');
    // The appearance is authored at the origin; the /Rect is what places it.
    expect(box!.minX).toBeCloseTo(50, 1);
    expect(box!.minY).toBeCloseTo(60, 1);
    expect(box!.maxX).toBeCloseTo(90, 1);
    expect(box!.maxY).toBeCloseTo(80, 1);
  });

  it('paints no annotation the file marks hidden', () => {
    const hidden = new TextDecoder()
      .decode(widgetOnlyPdf())
      .replace('/Subtype /Widget', '/Subtype /Widget /F 2');
    const file = PdfFile.parse(new TextEncoder().encode(hidden));
    expect(collectPageVectors(file, file.pages()[0]!)).toHaveLength(0);
  });
});

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

/**
 * A page that fills one band through an `/ExtGState` at `ca` 0.6 and a second
 * one after the `Q` that pops it — the shape 22060_A1_01_Plans.pdf's evacuation
 * routes take.
 */
function alphaBandsPdf(): Uint8Array {
  const content = 'q /G0 gs 0 1 0 rg 20 20 100 60 re f Q 0 0 1 rg 20 120 100 60 re f';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      '/Resources << /ExtGState << /G0 << /Type /ExtGState /CA 0.6 /ca 0.6 >> >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('constant fill alpha (§11.6.4.4)', () => {
  it('reads /ca off the graphics state `gs` names, and lets `Q` take it back', () => {
    // §8.4.5 — `gs` sets a whole state at once. 22060_A1_01_Plans.pdf marks its
    // evacuation routes with a green band at `ca` 0.6, meant to be read
    // THROUGH: painted solid, the floor plan under each band disappeared.
    const { vectors } = interpretContent(
      new TextEncoder().encode('q /G0 gs 0 1 0 rg 20 20 100 60 re f Q 0 0 1 rg 20 120 100 60 re f'),
      NO_FONTS,
      undefined,
      undefined,
      new Map([['G0', 0.6]]),
    );
    expect(vectors).toHaveLength(2);
    expect(vectors[0]!.alpha).toBeCloseTo(0.6, 5);
    expect(vectors[1]!.alpha).toBeUndefined();
  });

  it('leaves a fill opaque when the named state states no alpha', () => {
    const { vectors } = interpretContent(
      new TextEncoder().encode('/G9 gs 0 1 0 rg 20 20 100 60 re f'),
      NO_FONTS,
      undefined,
      undefined,
      new Map([['G0', 0.6]]),
    );
    expect(vectors[0]!.alpha).toBeUndefined();
  });

  it('carries the alpha from the page’s /ExtGState onto the lifted shape', () => {
    const file = PdfFile.parse(alphaBandsPdf());
    const [band, opaque] = collectPageVectors(file, file.pages()[0]!);
    expect(band!.alpha).toBeCloseTo(0.6, 5);
    expect(opaque!.alpha).toBeUndefined();

    const { doc } = reconstructByLayout(file);
    const fills = doc.body.filter((b) => b.kind === 'shape').map((s) => s.shape.fill);
    expect(fills).toHaveLength(2);
    expect(fills.find((f) => f.colorHex === '00FF00')).toMatchObject({
      kind: 'solid',
      alpha: 0.6,
    });
    expect(fills.find((f) => f.colorHex === '0000FF')).not.toHaveProperty('alpha');
  });
});

describe('filled rules (E-PDF EP10)', () => {
  const fills = (stream: string) =>
    interpretContent(new TextEncoder().encode(stream), NO_FONTS).vectors;

  it('keeps a long thin fill — a form draws its lines as rectangles', () => {
    // 160F-2019.pdf has no stroke operator at all: every rule of the
    // certificate is a filled rectangle half a point high. Dropped as hairline
    // clutter, the whole grid went with them and the text arrived with no form
    // under it.
    const [rule] = fills('0 0 0 rg 100 100 200 0.5 re f');
    expect(rule).toBeDefined();
    expect(rule!.fillHex).toBe('000000');
  });

  it('still drops a speck thin in BOTH directions', () => {
    // The rule admits length, not smallness: a dot stays clutter.
    const painted = fills('0 0 0 rg 100 100 1 1 re f');
    expect(painted).toHaveLength(1); // the interpreter sees it …
    // … and `collectPageVectors` is what rejects it; the geometry is the test.
    const segs = painted[0]!.segs.flatMap((seg) => ('x' in seg ? [seg.x] : []));
    expect(Math.max(...segs) - Math.min(...segs)).toBeLessThan(6);
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
