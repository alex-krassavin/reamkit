// E-PDF EP3 — tagged fast-path. Convert a docx to a TAGGED PDF, then rebuild a
// FlowDoc from its /StructTreeRoot and confirm the headings, paragraphs and
// reading order come back — the honest inverse of the tagged PDF Ream writes.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructTaggedPdf } from '@/pdf-reader/tagged';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const para = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const heading = (text: string, level: number): string =>
  `<w:p><w:pPr><w:outlineLvl w:val="${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function taggedFlow(body: string) {
  const pdf = await Ream.parse(buildDocxFromBody(body)).convert('pdf', {
    fonts: FONTS,
    tagged: true,
  });
  const flow = reconstructTaggedPdf(PdfFile.parse(pdf));
  if (!flow) throw new Error('reconstruction returned no FlowDoc');
  return flow.doc;
}

const paragraphTexts = (flow: { body: ReadonlyArray<{ kind: string }> }): Array<string> =>
  flow.body
    .filter(
      (b): b is { kind: 'paragraph'; paragraph: { runs: ReadonlyArray<{ text: string }> } } =>
        b.kind === 'paragraph',
    )
    .map((p) =>
      p.paragraph.runs
        .map((r) => r.text)
        .join('')
        .replace(/\s+/g, ' ')
        .trim(),
    );

describe('tagged-PDF reconstruction (E-PDF EP3)', () => {
  it('recovers headings, paragraphs and reading order', async () => {
    const flow = await taggedFlow(
      heading('Chapter One', 0) + para('First paragraph body.') + para('Second paragraph body.'),
    );
    const texts = paragraphTexts(flow);
    expect(texts.some((t) => t.includes('Chapter One'))).toBe(true);
    expect(texts.some((t) => t.includes('First paragraph body'))).toBe(true);
    expect(texts.some((t) => t.includes('Second paragraph body'))).toBe(true);
    // reading order: the heading precedes the first body paragraph.
    const joined = texts.join(' | ');
    expect(joined.indexOf('Chapter One')).toBeLessThan(joined.indexOf('First paragraph'));
  });

  it('maps an H1 structure element back to outline level 0', async () => {
    const flow = await taggedFlow(heading('A Heading', 0) + para('Body text here.'));
    const headingPara = flow.body.find(
      (b) =>
        b.kind === 'paragraph' &&
        b.paragraph.runs
          .map((r) => r.text)
          .join('')
          .includes('A Heading'),
    );
    expect(headingPara?.kind).toBe('paragraph');
    expect(
      headingPara?.kind === 'paragraph' ? headingPara.paragraph.properties.outlineLevel : undefined,
    ).toBe(0);
  });

  it('returns undefined for an untagged PDF', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(para('Plain untagged text.'))).convert('pdf', {
      fonts: FONTS,
    });
    expect(reconstructTaggedPdf(PdfFile.parse(pdf))).toBeUndefined();
  });

  it('reconstructs a table as a Table element with its cell text (EP3b)', async () => {
    const cell = (t: string): string => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr>${cell('NameCol')}${cell('ScoreCol')}</w:tr>` +
      `<w:tr>${cell('AlphaRow')}${cell('NinetyNine')}</w:tr></w:tbl>`;
    const flow = await taggedFlow(tbl);
    const table = flow.body.find((b) => b.kind === 'table');
    expect(table).toBeDefined();
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows).toHaveLength(2);
    expect(table.table.rows[0]!.cells).toHaveLength(2);
    const dump = JSON.stringify(table);
    for (const word of ['NameCol', 'ScoreCol', 'AlphaRow', 'NinetyNine']) {
      expect(dump).toContain(word);
    }
  });

  it('reads the column grid from where the cells sit, not in equal shares', async () => {
    // A structure tree states no widths, so an equal share was the old answer.
    // It is right only for a table whose columns really are equal: here the
    // first is four times the second, and a form's are further apart still.
    const cell = (t: string): string => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const tbl =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="6400"/><w:gridCol w:w="1600"/></w:tblGrid>' +
      `<w:tr>${cell('WideColumnHere')}${cell('Nx')}</w:tr>` +
      `<w:tr>${cell('AlsoWideHere')}${cell('Ny')}</w:tr></w:tbl>`;
    const flow = await taggedFlow(tbl);
    const table = flow.body.find((b) => b.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const [first, second] = table.table.grid;
    expect(first).toBeDefined();
    expect(second).toBeDefined();
    expect(first!).toBeGreaterThan(second!);
  });

  it('keeps the colour the glyphs were painted in', async () => {
    // §8.6.8. Dropped, 160F-2019.pdf's blue field labels and its red warning
    // all came back plain black.
    const flow = await taggedFlow(
      '<w:p><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>BlueText</w:t></w:r></w:p>',
    );
    const colours = flow.body.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.properties.colorHex) : [],
    );
    expect(colours).toContain('0000FF');
  });

  it('keeps the size the glyphs were shown at', async () => {
    // §9.3.1 Tf. Dropped, every run came back at the 11pt default, and a form
    // set in 7pt grew by half again — 160F-2019.pdf rebuilt one page as five.
    const flow = await taggedFlow(
      '<w:p><w:r><w:rPr><w:sz w:val="14"/></w:rPr><w:t>SevenPointText</w:t></w:r></w:p>',
    );
    const sizes = flow.body.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.properties.fontSizePt) : [],
    );
    expect(sizes.some((s) => s !== undefined && Math.abs(s - 7) < 0.5)).toBe(true);
  });

  it('reconstructs list items as paragraphs carrying their text (EP3b)', async () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
      '<w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const item = (t: string): string =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${t}</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(item('ItemOne') + item('ItemTwo'), { numberingXml: numbering });
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS, tagged: true });
    const flow = reconstructTaggedPdf(PdfFile.parse(pdf));
    if (!flow) throw new Error('reconstruction returned no FlowDoc');
    const texts = paragraphTexts(flow.doc);
    expect(texts.some((t) => t.includes('ItemOne'))).toBe(true);
    expect(texts.some((t) => t.includes('ItemTwo'))).toBe(true);
  });
});
