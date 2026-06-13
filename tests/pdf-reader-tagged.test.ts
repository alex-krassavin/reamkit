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
      (
        b,
      ): b is {
        kind: 'paragraph';
        paragraph: { properties: { outlineLevel?: number }; runs: ReadonlyArray<{ text: string }> };
      } =>
        b.kind === 'paragraph' &&
        (b as { paragraph: { runs: ReadonlyArray<{ text: string }> } }).paragraph.runs
          .map((r) => r.text)
          .join('')
          .includes('A Heading'),
    );
    expect(headingPara?.paragraph.properties.outlineLevel).toBe(0);
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
