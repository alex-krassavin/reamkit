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
