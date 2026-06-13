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
  return reconstructByLayout(PdfFile.parse(pdf));
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
