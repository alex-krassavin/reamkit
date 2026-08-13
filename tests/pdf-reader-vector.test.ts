// E-PDF EP10 — filled vector paths. The interpreter captures a `re … f` fill
// with its colour; end-to-end, a filled docx shape comes back as a solid-fill
// shape when the (untagged) PDF is read.

import { readFileSync } from 'node:fs';

import { zlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PdfImage } from '@/pdf-reader/images';
import { Ream } from '@/core/converter/ream';
import { interpretContent } from '@/pdf-reader/content';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { shapeBlock } from '@/pdf-reader/flow-build';
import { collectPageVectors } from '@/pdf-reader/vector';
import { collectPageImages } from '@/pdf-reader/images';

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
 * A page whose only mark is a check box whose `/AP` `/N` is a SET of states —
 * one stream, named `/1` — while `/AS` names the state actually in force.
 */
function checkBoxPdf(state: string): Uint8Array {
  const ap = '0 0 1 rg 0 0 10 10 re f';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      `/Annots [<< /Type /Annot /Subtype /Widget /Rect [50 60 60 70] /AS /${state} ` +
      '/AP << /N << /1 5 0 R >> >> >>] >>',
    '<< /Length 0 >>\nstream\n\nendstream',
    `<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length ${String(ap.length)} >>\n` +
      `stream\n${ap}\nendstream`,
  ];
  return assemble(objects);
}

/**
 * A page that writes a word and lays a band over it through a graphics state
 * whose `/BM` is `/Multiply` — a highlighter, and nothing else.
 */
function highlighterPdf(): Uint8Array {
  const content = 'q /Hi gs 1 1 0 rg 20 20 100 20 re f Q';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      '/Resources << /ExtGState << /Hi << /Type /ExtGState /BM /Multiply >> >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
  ];
  return assemble(objects);
}

/** The objects as a file: header, bodies, cross-reference table, trailer. */
function assemble(objects: ReadonlyArray<string>): Uint8Array {
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

/**
 * A page with one markup annotation and NO `/AP` — the case §12.5.5 leaves to
 * the reader. `extra` states the geometry the subtype is drawn from.
 */
function noAppearancePdf(subtype: string, extra: string): Uint8Array {
  return assemble([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      `/Annots [<< /Type /Annot /Subtype /${subtype} /C [1 0 0] ${extra} >>] >>`,
    '<< /Length 0 >>\nstream\n\nendstream',
  ]);
}

describe('annotation appearances (§12.5.5)', () => {
  it('draws a markup annotation that carries no appearance at all', () => {
    // "If the annotation does not contain an appearance stream, the conforming
    // reader shall generate one." Nine files of the pdf.js corpus are that
    // case, and every one came back a blank page while poppler and
    // LibreOffice both drew it.
    const line = PdfFile.parse(noAppearancePdf('Line', '/Rect [0 0 0 0] /L [20 30 180 30]'));
    const [rule] = collectPageVectors(line, line.pages()[0]!).vectors;
    expect(rule?.strokeHex).toBe('FF0000');
    expect([rule?.minX, rule?.maxX]).toEqual([20, 180]);

    const ink = PdfFile.parse(
      noAppearancePdf('Ink', '/Rect [0 0 200 200] /InkList [[20 20 40 60] [80 20 100 60]]'),
    );
    // Two strokes, one path: a scrawl is stroked once, not once per stroke.
    expect(collectPageVectors(ink, ink.pages()[0]!).vectors).toHaveLength(1);

    const circle = PdfFile.parse(
      noAppearancePdf('Circle', '/Rect [20 20 120 80] /IC [0 0 1] /BS << /W 2 >>'),
    );
    const [disc] = collectPageVectors(circle, circle.pages()[0]!).vectors;
    expect(disc?.fillHex).toBe('0000FF');
    expect(disc?.strokeHex).toBe('FF0000');
  });

  it('leaves a WIDGET with no appearance for its state undrawn', () => {
    // The same rule must not invent a look for a form field: a check box whose
    // clear state has no appearance draws nothing, and that is the point.
    const file = PdfFile.parse(noAppearancePdf('Widget', '/Rect [20 20 40 40] /FT /Btn'));
    expect(collectPageVectors(file, file.pages()[0]!).vectors).toHaveLength(0);
  });

  it('paints the state /AS names, and nothing when the set has no such state', () => {
    // A check box is drawn by its ON appearance and cleared by having none:
    // annotation-button-widget.pdf carries `/N << /1 … >>` and `/AS /Off` for
    // the boxes that are clear. Taking "the only stream in the set" for those
    // ticked every box and filled every radio button on the form.
    const on = PdfFile.parse(checkBoxPdf('1'));
    expect(collectPageVectors(on, on.pages()[0]!).vectors).toHaveLength(1);
    const off = PdfFile.parse(checkBoxPdf('Off'));
    expect(collectPageVectors(off, off.pages()[0]!).vectors).toHaveLength(0);
  });

  it('lifts what a widget draws, and fits it to the annotation’s /Rect', () => {
    // A form field paints nothing in the page's content stream: its tint, its
    // border and its value all live in the widget's own appearance. Read only
    // from the page, 160F-2019.pdf gave a grid with no fields in it.
    const file = PdfFile.parse(widgetOnlyPdf());
    const { vectors } = collectPageVectors(file, file.pages()[0]!);
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
    expect(collectPageVectors(file, file.pages()[0]!).vectors).toHaveLength(0);
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

/** A page that paints `count` little squares, each its own path. */
function manyPathsPdf(count: number): Uint8Array {
  let content = '0 0 1 rg ';
  for (let i = 0; i < count; i++) {
    content += `${String(i % 500) + ' ' + String(Math.floor(i / 500))} 6 6 re f `;
  }
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 600] /Contents 4 0 R >>',
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

describe('the guard on how much a page may paint', () => {
  // The cap was two thousand and it was silent, so a drawing that ran past it
  // simply arrived without its tail: Brotli-Prototype-FileA.pdf lost its whole
  // title block, the vegetation of its perspective and every hatch in its
  // legend, with nothing in the report to say a thing was missing.
  it('reads a page well inside the guard with nothing to report', () => {
    const file = PdfFile.parse(manyPathsPdf(50));
    const lifted = collectPageVectors(file, file.pages()[0]!);
    expect(lifted.vectors.length).toBe(50);
    expect(lifted.losses).toHaveLength(0);
  });

  it('reports the cut once the page runs past it', () => {
    const file = PdfFile.parse(manyPathsPdf(20_001));
    const lifted = collectPageVectors(file, file.pages()[0]!);
    expect(lifted.losses).toHaveLength(1);
    expect(lifted.losses[0]!.severity).toBe('dropped');
    expect(lifted.losses[0]!.detail).toContain('20000');
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
  it('marks a fill that only DARKENS, so what it covers reads through', () => {
    // §11.3.5 — annotation-highlight.pdf lays its yellow band over the words
    // with `/BM /Multiply`, which is how a highlighter works. Nothing anchored
    // in a .docx blends, so a band read as paint buried the line it marked.
    const file = PdfFile.parse(highlighterPdf());
    const [band] = collectPageVectors(file, file.pages()[0]!).vectors;
    expect(band?.darkens).toBe(true);
    const block = shapeBlock(band!, { left: 0, top: 200 }, 0);
    expect(block.kind === 'shape' && block.shape.float?.behind).toBe(true);
  });

  it('reads /ca off the graphics state `gs` names, and lets `Q` take it back', () => {
    // §8.4.5 — `gs` sets a whole state at once. 22060_A1_01_Plans.pdf marks its
    // evacuation routes with a green band at `ca` 0.6, meant to be read
    // THROUGH: painted solid, the floor plan under each band disappeared.
    const { vectors } = interpretContent(
      new TextEncoder().encode('q /G0 gs 0 1 0 rg 20 20 100 60 re f Q 0 0 1 rg 20 120 100 60 re f'),
      NO_FONTS,
      undefined,
      undefined,
      new Map([['G0', { alpha: 0.6 }]]),
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
      new Map([['G0', { alpha: 0.6 }]]),
    );
    expect(vectors[0]!.alpha).toBeUndefined();
  });

  it('carries the alpha from the page’s /ExtGState onto the lifted shape', () => {
    const file = PdfFile.parse(alphaBandsPdf());
    const [band, opaque] = collectPageVectors(file, file.pages()[0]!).vectors;
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

/** A page that paints `content`, for asking what survives the de-cluttering. */
function contentPdf(content: string): Uint8Array {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>',
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

/**
 * A page set in a Type 3 font whose one glyph strokes a square, and — from
 * inside that glyph — shows a second Type 3 font whose glyph strokes a
 * triangle. The shape ContentStreamCycleType3insideType3.pdf takes.
 */
function type3Pdf(inner: boolean): Uint8Array {
  const outerProc =
    '1000 0 d0 20 w 1 0 0 RG 0 0 750 750 re s' + (inner ? ' BT /FB 50 Tf (c) Tj ET' : '');
  const innerProc = '1000 0 d0 20 w 0 1 0 RG 0 0 m 375 750 l 750 0 l s';
  const content = 'BT /FA 200 Tf 50 50 Td (a) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 400] ' +
      '/Resources << /Font << /FA 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type3 /FontBBox [0 0 750 750] ' +
      '/FontMatrix [0.001 0 0 0.001 0 0] /Widths [1000] /FirstChar 97 /LastChar 97 ' +
      '/Encoding << /Differences [97 /square] >> /CharProcs << /square 6 0 R >> ' +
      '/Resources << /Font << /FB 7 0 R >> >> >>',
    `<< /Length ${String(outerProc.length)} >>\nstream\n${outerProc}\nendstream`,
    '<< /Type /Font /Subtype /Type3 /FontBBox [0 0 750 750] ' +
      '/FontMatrix [0.01 0 0 0.01 0 0] /Widths [1000] /FirstChar 99 /LastChar 99 ' +
      '/Encoding << /Differences [99 /tri] >> /CharProcs << /tri 8 0 R >> >>',
    `<< /Length ${String(innerProc.length)} >>\nstream\n${innerProc}\nendstream`,
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

describe('a Type 3 glyph is a drawing (§9.6.5)', () => {
  const strokesOf = (bytes: Uint8Array) => {
    const file = PdfFile.parse(bytes);
    return collectPageVectors(file, file.pages()[0]!).vectors;
  };

  it('paints what the glyph procedure paints, where the character stands', () => {
    // ContentStreamCycleType3insideType3.pdf is a stroked square and a stroked
    // triangle drawn as GLYPHS. With the procedures unread the page came back
    // as two letters of substituted type an eighth of an inch tall.
    const vectors = strokesOf(type3Pdf(false));
    expect(vectors).toHaveLength(1);
    expect(vectors[0]!.strokeHex).toBe('FF0000');
    // 750 glyph units × 0.001 × 200pt = 150pt, at the pen's 50, 50.
    expect(vectors[0]!.minX).toBeCloseTo(50, 1);
    expect(vectors[0]!.minY).toBeCloseTo(50, 1);
    expect(vectors[0]!.maxX).toBeCloseTo(200, 1);
  });

  it('follows a glyph that shows a second Type 3 font from inside itself', () => {
    const vectors = strokesOf(type3Pdf(true));
    expect(vectors.map((v) => v.strokeHex).sort()).toEqual(['00FF00', 'FF0000']);
  });
});

describe('white paint is invisible only over white', () => {
  const kept = (content: string) => {
    const file = PdfFile.parse(contentPdf(content));
    return collectPageVectors(file, file.pages()[0]!).vectors;
  };

  it('keeps a white stroke drawn over something, as it keeps a white fill', () => {
    // 160F-2019.pdf's every form field is a tinted box with a WHITE one-point
    // border stroked inside it. Dropped, each field came back a point wider on
    // every side than the file draws it — seventy-six times over.
    const vectors = kept('0 0 1 rg 50 50 100 50 re f 1 G 55 55 90 40 re s');
    expect(vectors).toHaveLength(2);
    expect(vectors[0]!.fillHex).toBe('0000FF');
    expect(vectors[1]!.strokeHex).toBe('FFFFFF');
  });

  it('drops a white stroke over bare paper, which shows nothing', () => {
    expect(kept('1 G 55 55 90 40 re s')).toHaveLength(0);
  });
});

describe('the pen is as wide as the page says (§8.4.3.2)', () => {
  const strokeOf = (stream: string) => {
    const [v] = interpretContent(new TextEncoder().encode(stream), NO_FONTS).vectors;
    const el = shapeBlock({
      orderKey: [0],
      segs: v!.segs,
      strokeHex: '000000',
      ...(v?.lineWidth !== undefined ? { lineWidth: v.lineWidth } : {}),
      minX: 0,
      minY: 0,
      maxX: 100,
      maxY: 0,
    });
    return el.kind === 'shape' ? el.shape : undefined;
  };

  it('draws a hairline pen at the width the page set, not at a floor', () => {
    // Brotli-Prototype-FileA.pdf draws its elevations with a 0.12pt pen, and
    // raised to half a point every clapboard line came out four times too
    // heavy: a drawing that reads grey in every viewer arrived black.
    expect(strokeOf('0.12 w 0 0 m 100 0 l S')?.line?.width).toBeCloseTo(0.12, 5);
    expect(strokeOf('2 w 0 0 m 100 0 l S')?.line?.width).toBeCloseTo(2, 5);
  });

  it('takes a width of zero for the thinnest line there is', () => {
    const width = strokeOf('0 w 0 0 m 100 0 l S')?.line?.width ?? 0;
    expect(width).toBeGreaterThan(0);
    expect(width).toBeLessThan(0.25);
  });

  it('still gives a flat line a box to draw in', () => {
    // The floor was there for a reason: a horizontal rule has no height, and a
    // shape of no height has nowhere to put a stroke.
    expect(strokeOf('0.12 w 0 0 m 100 0 l S')?.height).toBeCloseTo(0.5, 5);
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

describe('a blend the page asks for and nothing here performs (§11.3.5)', () => {
  /** A page that paints an image through a graphics state naming `mode`. */
  const blended = (mode: string): Uint8Array => {
    const gray = zlibSync(Uint8Array.from([0, 255, 255, 0]));
    const content = `q /BM0 gs 50 0 0 50 20 20 cm /Im Do Q`;
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        '/Resources << /XObject << /Im 5 0 R >> ' +
        `/ExtGState << /BM0 << /Type /ExtGState /BM /${mode} >> >> >> >>`,
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
      '<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray ' +
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
  };

  it('says so, rather than drawing one picture over the other in silence', () => {
    // No anchored picture blends, in any format this writes. blendmode.pdf
    // lays a second photograph over a dog in each of sixteen cells, one blend
    // to a cell, and every one came back as the dog alone with nothing said.
    const file = PdfFile.parse(blended('Difference'));
    const { losses, images } = collectPageImages(file, file.pages()[0]!);
    expect(images).toHaveLength(1);
    expect(losses.some((l) => l.detail.includes('/Difference'))).toBe(true);
  });

  it('says nothing where the blend is Normal, which is no blend', () => {
    const file = PdfFile.parse(blended('Normal'));
    expect(collectPageImages(file, file.pages()[0]!).losses).toHaveLength(0);
  });

  it('says nothing where the mark only DARKENS, which it approximates', () => {
    // Multiply and Darken go behind the text instead, and come to the same
    // picture — a highlighter with its words on top.
    const file = PdfFile.parse(blended('Multiply'));
    expect(collectPageImages(file, file.pages()[0]!).losses).toHaveLength(0);
  });
});

describe('a picture the CTM turns and a clip cuts (§8.9.5, §8.5.4)', () => {
  /** A page whose only mark is one image drawn through `cm`, under `clip`. */
  const placed = (cm: string, clip: string): Uint8Array => {
    const gray = zlibSync(Uint8Array.from([0, 255, 255, 0]));
    const content = `q ${clip}${cm} cm /Im Do Q`;
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 400] /Contents 4 0 R ' +
        '/Resources << /XObject << /Im 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
      '<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray ' +
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
  };

  /** The one image such a page carries. */
  const only = (cm: string, clip = ''): PdfImage => {
    const file = PdfFile.parse(placed(cm, clip));
    const { images } = collectPageImages(file, file.pages()[0]!);
    expect(images).toHaveLength(1);
    return images[0]!;
  };

  it('leaves an upright picture upright, and where it was put', () => {
    const img = only('100 0 0 60 40 30');
    expect(img.rotationDeg).toBeUndefined();
    expect(img.crop).toBeUndefined();
    expect([img.x, img.y, img.widthPt, img.heightPt]).toEqual([40, 30, 100, 60]);
  });

  it('turns it as the CTM turns it, about its own centre', () => {
    // 90° counter-clockwise: the columns swap and the second one flips.
    // image-rotated-black-white-ratio.pdf sets its picture at thirty-one
    // degrees and it came back upright in the corner, because taking only the
    // column LENGTHS threw the turn away.
    const img = only('0 100 -60 0 40 30');
    expect(Math.round(img.rotationDeg ?? 0)).toBe(90);
    expect(Math.round(img.widthPt)).toBe(100);
    expect(Math.round(img.heightPt)).toBe(60);
    // The unturned box is centred where the turned one is: (40,30) is the
    // unit square's origin, so its middle lands at (40-30, 30+50).
    expect([Math.round(img.x + img.widthPt / 2), Math.round(img.y + img.heightPt / 2)]).toEqual([
      10, 80,
    ]);
  });

  it('cuts it to a clip, on the picture\u2019s own edges', () => {
    // The left half and the top three quarters, of a picture placed square on.
    const img = only('100 0 0 100 0 0', '0 25 50 75 re W n ');
    expect(img.crop).toEqual({ left: 0, right: 0.5, top: 0, bottom: 0.25 });
    expect([img.x, img.y, img.widthPt, img.heightPt]).toEqual([0, 25, 50, 75]);
  });

  it('cuts along the picture\u2019s axes when clip and picture turn together', () => {
    // A clip square on in the page's axes but drawn at the picture's own
    // angle: in the picture's space it is square on, and the crop is exact.
    // image-rotated-black-white-ratio.pdf turns both by thirty-one degrees.
    const cm = '0 100 -100 0 100 0';
    const img = only(cm, '100 25 m 25 25 l 25 75 l 100 75 l h W n ');
    expect(img.crop?.left).toBeCloseTo(0.25, 6);
    expect(img.crop?.right).toBeCloseTo(0.25, 6);
    expect(img.crop?.top).toBeCloseTo(0.25, 6);
    expect(img.crop?.bottom).toBeCloseTo(0, 6);
    expect(Math.round(img.widthPt)).toBe(50);
    expect(Math.round(img.heightPt)).toBe(75);
  });

  it('draws the whole picture where the clip does not reach it', () => {
    // A clip that leaves nothing has been read wrong; a hairline is worse
    // than the picture.
    const img = only('100 0 0 100 0 0', '300 300 20 20 re W n ');
    expect(img.crop).toBeUndefined();
    expect(img.widthPt).toBe(100);
  });
});

describe('a fill the size of the page', () => {
  /** A page whose only mark is one filled rectangle covering `fraction` of it. */
  const backgroundPdf = (fraction: number, color: string): Uint8Array => {
    const side = Math.sqrt(fraction) * 200;
    const content = `${color} rg 0 0 ${String(side)} ${String(side)} re f`;
    return assemble([
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>',
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ]);
  };

  it('is a page background, not clutter', () => {
    // Capped at 0.85 of the sheet, filled-background.pdf — which is nothing but
    // that fill — came back a blank page, and bug1755507.pdf's card floated on
    // white where the file has pale blue. Eight files of the corpus improved
    // when the cap went and not one worsened.
    const whole = PdfFile.parse(backgroundPdf(1, '0.7 0.85 0.95'));
    const [bg] = collectPageVectors(whole, whole.pages()[0]!).vectors;
    expect(bg?.fillHex).toBe('B3D9F2');
    expect(bg!.maxX - bg!.minX).toBeCloseTo(200, 3);
  });

  it('still drops white paint that covers nothing', () => {
    // White over white paper marks nothing, whatever its size.
    const white = PdfFile.parse(backgroundPdf(1, '1 1 1'));
    expect(collectPageVectors(white, white.pages()[0]!).vectors).toHaveLength(0);
  });
});

describe('the CIE-based grey space (§8.6.5.6)', () => {
  /** A page filling a box in a CalGray space with the given parameters. */
  const calGrayPdf = (params: string, value: string): Uint8Array => {
    const content = `/Cs cs ${value} sc 20 20 100 100 re f`;
    return assemble([
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        `/Resources << /ColorSpace << /Cs [/CalGray << ${params} >>] >> >> >>`,
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ]);
  };

  it('puts the value through the gamma the space states', () => {
    // A CalGray's number looks like a DeviceGray's and is not: calgray.pdf
    // reads 0.258 against the source as one and 0.044 as the other.
    const plain = PdfFile.parse(calGrayPdf('/WhitePoint [0.9505 1 1.089]', '0.5'));
    const [flat] = collectPageVectors(plain, plain.pages()[0]!).vectors;
    // Gamma 1: the sRGB transfer alone, which lifts a mid grey well above half.
    expect(flat?.fillHex).toBe('BCBCBC');
    const dark = PdfFile.parse(calGrayPdf('/WhitePoint [0.9505 1 1.089] /Gamma 2.2', '0.5'));
    const [gammaed] = collectPageVectors(dark, dark.pages()[0]!).vectors;
    expect(gammaed?.fillHex).toBe('808080');
  });

  it('keeps the device reading where the space states no white point', () => {
    // §8.6.5.6 makes /WhitePoint required; a space without one states nothing
    // this can transform, and the number is read as the device grey it looks
    // like — which at gamma 1 is a different grey from the transform's.
    const odd = PdfFile.parse(calGrayPdf('/Gamma 1', '0.5'));
    expect(collectPageVectors(odd, odd.pages()[0]!).vectors[0]?.fillHex).toBe('808080');
  });
});

describe('the CIE-based RGB space (§8.6.5.7)', () => {
  /** A page filling a box in a CalRGB space with the given parameters. */
  const calRgbPdf = (params: string, value: string): Uint8Array => {
    const content = `/Cs cs ${value} sc 20 20 100 100 re f`;
    return assemble([
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
        `/Resources << /ColorSpace << /Cs [/CalRGB << ${params} >>] >> >> >>`,
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ]);
  };

  it('decodes through the gamma and the matrix, and adapts the white it states', () => {
    // Three numbers that look like a DeviceRGB's and are not. Under the sRGB
    // white and an identity matrix the space IS sRGB's own primaries scaled,
    // so a half-red comes back the sRGB transfer of what the matrix gives.
    const d65 =
      '/WhitePoint [0.9505 1 1.089] /Matrix [0.4124 0.2126 0.0193 0.3576 0.7152 0.1192 0.1805 0.0722 0.9505]';
    const half = PdfFile.parse(calRgbPdf(d65, '0.5 0 0'));
    expect(collectPageVectors(half, half.pages()[0]!).vectors[0]?.fillHex).toBe('BC0000');
    // …and the gamma is the space's own: 0.5^2.2 is a quarter of the light.
    const gammaed = PdfFile.parse(calRgbPdf(`${d65} /Gamma [2.2 2.2 2.2]`, '0.5 0 0'));
    expect(collectPageVectors(gammaed, gammaed.pages()[0]!).vectors[0]?.fillHex).toBe('800000');
  });

  it('keeps the device reading where the space states no white point', () => {
    // §8.6.5.7 makes /WhitePoint required. Without one there is nothing to
    // adapt from, and the three numbers are read as the device colour they
    // look like.
    const odd = PdfFile.parse(calRgbPdf('/Gamma [1 1 1]', '0.5 0 0'));
    expect(collectPageVectors(odd, odd.pages()[0]!).vectors[0]?.fillHex).toBe('800000');
  });
});

describe('a shading stitched out of shadings (§7.10.4)', () => {
  /** A page filling one square with pattern `/P1`, whose function is `fn`. */
  const shaded = (fn: string): Uint8Array => {
    const content = '/Pattern cs /P1 scn 10 10 80 80 re f';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R ' +
        '/Resources << /Pattern << /P1 5 0 R >> >> >>',
      `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
      '<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB ' +
        `/Coords [0 0 0 100] /Function ${fn} >> >>`,
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
    return Uint8Array.from([...pdf].map((c) => c.charCodeAt(0)));
  };

  const stopsOf = (fn: string): Array<string> => {
    const shape = Ream.parse(shaded(fn)).flow.body.find((el) => el.kind === 'shape');
    if (shape?.kind !== 'shape' || shape.shape.fill.kind !== 'gradient') return [];
    const gradient = shape.shape.fill.gradient;
    if (!gradient) return [];
    return gradient.stops.map((s) => `${s.offset.toFixed(3)}:${s.colorHex}`);
  };

  const RAMP = '<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >>';
  /** Green, then a zero-width fade, then blue — which is a hard EDGE. */
  const STEP =
    '<< /FunctionType 3 /Domain [0 1] /Bounds [0.5 0.5] /Encode [0 1 0 1 0 1] /Functions [' +
    '<< /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 1 0] /N 1 >> ' +
    '<< /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >> ' +
    '<< /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 0 1] /N 1 >>] >>';

  it('keeps a plain ramp as its two ends', () => {
    expect(stopsOf(RAMP)).toEqual(['0.000:FF0000', '1.000:0000FF']);
  });

  it('keeps a subfunction’s OWN steps, not just its ends', () => {
    // issue10572.pdf stitches twelve copies of a green/blue pair whose
    // `/Bounds [0.5 0.5]` makes a hard edge; reduced to first-and-last, each
    // pair came back a smooth fade instead of two flat bands.
    expect(stopsOf(STEP)).toEqual(['0.000:00FF00', '0.500:00FF00', '0.500:0000FF', '1.000:0000FF']);
  });

  it('lays a nested stitch onto the piece of the domain it holds', () => {
    const outer = `<< /FunctionType 3 /Domain [0 1] /Bounds [0.5] /Encode [0 1 0 1] /Functions [${STEP} ${STEP}] >>`;
    // Two bands, each stepping green → blue halfway through its own half.
    expect(stopsOf(outer)).toEqual([
      '0.000:00FF00',
      '0.250:00FF00',
      '0.250:0000FF',
      '0.500:0000FF',
      '0.500:00FF00',
      '0.750:00FF00',
      '0.750:0000FF',
      '1.000:0000FF',
    ]);
  });
});
