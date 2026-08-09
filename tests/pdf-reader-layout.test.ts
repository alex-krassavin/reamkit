// E-PDF EP4 — heuristic reconstruction for untagged PDFs. Convert a docx to a
// plain (untagged) PDF, then rebuild a FlowDoc from the positioned text alone:
// lines clustered by baseline, paragraphs by vertical spacing, headings by a
// font size above the median.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const spaced = (text: string): string =>
  `<w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function layoutFlow(body: string) {
  const pdf = await Ream.parse(buildDocxFromBody(body)).convert('pdf', { fonts: FONTS });
  return reconstructByLayout(PdfFile.parse(pdf)).doc;
}

const paragraphs = (flow: { body: ReadonlyArray<{ kind: string }> }) =>
  flow.body.filter(
    (
      b,
    ): b is {
      kind: 'paragraph';
      paragraph: { properties: { outlineLevel?: number }; runs: ReadonlyArray<{ text: string }> };
    } => b.kind === 'paragraph',
  );

describe('a heuristic line keeps how it looked (E-PDF EP4)', () => {
  it('carries each run’s size and colour, not just its letters', async () => {
    // The tagged path has carried these since it learned to; this one never
    // did, so every line came back at the 11pt default in black. Placed, that
    // is not a wrong shade but a wrong SHAPE: 160F-2019.pdf's footnotes are set
    // in 7pt nine and a half apart, and drawn at eleven they climbed over each
    // other.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:rPr><w:sz w:val="14"/><w:color w:val="FF0000"/></w:rPr>' +
        '<w:t>SmallRedText</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const flow = reconstructByLayout(PdfFile.parse(pdf)).doc;
    const runs = flow.body.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
    const small = runs.find((r) => r.text.includes('SmallRed'));
    expect(small).toBeDefined();
    expect(small!.properties.fontSizePt).toBeCloseTo(7, 0);
    expect(small!.properties.colorHex).toBe('FF0000');
  });
});

describe('a multi-page PDF keeps its pages (E-PDF EP4)', () => {
  it('opens an output page for each source page after the first', async () => {
    // Flowed, the layout repaginates and this hardly shows. PLACED, every mark
    // is anchored to "the page", so without a break all twenty-five pages of
    // Brotli-Prototype-FileA.pdf stacked onto one.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>PageOne</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>PageTwo</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    expect(file.pages().length).toBe(2);
    const placed = reconstructByLayout(file, 'positional').doc;
    const breaks = placed.body.filter(
      (b) => b.kind === 'paragraph' && b.paragraph.properties.pageBreakBefore === true,
    );
    expect(breaks).toHaveLength(file.pages().length - 1);
  });
});

/** A one-page PDF of hand-written objects; `page` is the page dict's body. */
function onePagePdf(page: string, content: string): Uint8Array {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /Contents 4 0 R ${page} >>`,
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

/** Two words on one baseline, three hundred points apart — a table row. */
const twoColumnLinePdf = (): Uint8Array =>
  onePagePdf('/MediaBox [0 0 400 800]', 'BT /F1 10 Tf 20 700 Td (Left) Tj 300 0 Td (Right) Tj ET');

/** A page whose `/MediaBox` starts twelve points off the origin, as 160F's does. */
const offsetBoxPdf = (): Uint8Array =>
  onePagePdf(
    '/MediaBox [-12 12 388 812]',
    'BT /F1 10 Tf 20 700 Td (Hello) Tj ET 0 0 1 rg 20 100 50 30 re f',
  );

describe('placed reconstruction (E-PDF EP4)', () => {
  it('anchors every line where its glyphs stand, instead of flowing them', async () => {
    // A form is a grid of ruled boxes with a label in each: flowed, the labels
    // land an inch from the boxes they label, because the artwork is placed and
    // the words are not. 160F-2019.pdf is that document.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>FirstLine</w:t></w:r></w:p><w:p><w:r><w:t>SecondLine</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const placed = reconstructByLayout(PdfFile.parse(pdf), 'positional');

    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes.length).toBeGreaterThanOrEqual(2);
    // Each carries its words and an anchor of its own on the page.
    for (const shape of shapes) {
      expect(shape.float?.posV?.relativeFrom).toBe('page');
      expect(shape.text?.content.length).toBeGreaterThan(0);
    }
    // The first line sits higher on the page, so its offset from the top is less.
    const offsets = shapes.map((shape) => shape.float?.posV?.offsetPt ?? 0);
    expect(offsets[0]!).toBeLessThan(offsets[1]!);
  });

  it('splits a baseline where a column gap opens, so each piece keeps its x', () => {
    // 160F-2019.pdf sets a line number, a label and a right-hand column on one
    // baseline. Read as a single line, the right-hand text was dragged left
    // against its neighbour with one space between: "sous-total: n° dossier:".
    const placed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()), 'positional');
    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes).toHaveLength(2);
    const offsets = shapes.map((s) => s.float?.posH?.offsetPt ?? 0).sort((a, b) => a - b);
    expect(offsets[0]).toBeCloseTo(20, 0);
    expect(offsets[1]).toBeCloseTo(320, 0);
    // Both stand on the same baseline, so neither moved vertically.
    const tops = shapes.map((s) => s.float?.posV?.offsetPt ?? 0);
    expect(tops[0]).toBeCloseTo(tops[1]!, 5);
  });

  it('reads one flowing line across the same gap', () => {
    // A paragraph is meant to be read across: only the placed reading splits.
    const flowed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()));
    const paras = paragraphs(flowed.doc);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.paragraph.runs.map((r) => r.text).join('')).toBe('Left Right');
  });

  it('measures a placed mark off the page’s own corner, not off the origin', () => {
    // §14.11.2 — a /MediaBox need not start at (0, 0), and 160F-2019.pdf's is
    // [-11.96 11.99 583.24 853.67]. Taking the corner for the origin put every
    // line and every rule twelve points up and to the left of the page's own.
    const placed = reconstructByLayout(PdfFile.parse(offsetBoxPdf()), 'positional');
    const text = placed.doc.body.find((b) => b.kind === 'shape' && b.shape.text !== undefined);
    const rule = placed.doc.body.find((b) => b.kind === 'shape' && b.shape.text === undefined);
    expect(text?.kind).toBe('shape');
    expect(rule?.kind).toBe('shape');
    if (text?.kind !== 'shape' || rule?.kind !== 'shape') return;
    // Text at x 20 on a box whose left edge is −12 stands 32pt in from the page.
    expect(text.shape.float?.posH?.offsetPt).toBeCloseTo(32, 5);
    // Baseline 700, box top 812: 812 − (700 − 2.5) − 12.5.
    expect(text.shape.float?.posV?.offsetPt).toBeCloseTo(102, 5);
    expect(rule.shape.float?.posH?.offsetPt).toBeCloseTo(32, 5);
    expect(rule.shape.float?.posV?.offsetPt).toBeCloseTo(682, 5); // 812 − 130
  });

  it('still flows by default, so a PDF reads back as a document', async () => {
    // The placed reading is opt-in: it has no reading order, no paragraphs and
    // no tables — which is exactly what a docx or a markdown conversion needs,
    // so the default must stay the flowed one.
    const docx = buildDocxFromBody('<w:p><w:r><w:t>FirstLine</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    expect(flowed.doc.body.some((b) => b.kind === 'paragraph')).toBe(true);
    expect(flowed.doc.body.some((b) => b.kind === 'shape')).toBe(false);
  });
});

describe('heuristic layout reconstruction (E-PDF EP4)', () => {
  it('groups untagged text into paragraphs in reading order', async () => {
    const flow = await layoutFlow(
      spaced('AlphaLine') + spaced('BravoLine') + spaced('CharlieLine'),
    );
    const paras = paragraphs(flow);
    expect(paras.length).toBeGreaterThanOrEqual(2); // the paragraphs separated
    const joined = paras.map((p) => p.paragraph.runs.map((r) => r.text).join('')).join(' | ');
    expect(joined).toContain('AlphaLine');
    expect(joined).toContain('BravoLine');
    expect(joined).toContain('CharlieLine');
    expect(joined.indexOf('Alpha')).toBeLessThan(joined.indexOf('Bravo'));
    expect(joined.indexOf('Bravo')).toBeLessThan(joined.indexOf('Charlie'));
  });

  it('carries the source page size + orientation into the FlowDoc section (F1)', async () => {
    // Force a non-A4 landscape MediaBox on the generated PDF, then confirm the
    // reader reflects it back — so a re-render keeps the size/orientation
    // instead of falling back to the layout engine's A4 default.
    const pdf = await Ream.parse(buildDocxFromBody(spaced('OnlyLine'))).convert('pdf', {
      fonts: FONTS,
      pageWidth: 1000,
      pageHeight: 600,
    });
    const section = reconstructByLayout(PdfFile.parse(pdf)).doc.section;
    expect(section?.pageSize?.width).toBe(1000);
    expect(section?.pageSize?.height).toBe(600);
    expect(section?.pageSize?.orientation).toBe('landscape');
    // A PDF has no margin model — the page box is the content box.
    expect(section?.margins?.left).toBe(0);
    expect(section?.margins?.top).toBe(0);
  });

  it('marks a line far larger than the median as a heading', async () => {
    const big = '<w:p><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>BigTitle</w:t></w:r></w:p>';
    const flow = await layoutFlow(
      big + spaced('body one') + spaced('body two') + spaced('body three'),
    );
    const title = paragraphs(flow).find((p) =>
      p.paragraph.runs
        .map((r) => r.text)
        .join('')
        .includes('BigTitle'),
    );
    expect(title?.paragraph.properties.outlineLevel).toBe(0);
  });
});
