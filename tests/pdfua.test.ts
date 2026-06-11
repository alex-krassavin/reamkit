import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b);

const BODY = '<w:p><w:r><w:t>accessible text</w:t></w:r></w:p>';

describe('PDF/UA-1 (ISO 14289-1)', () => {
  it('emits the pdfuaid XMP identifier and a tagged structure', async () => {
    const pdf = latin1(
      await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', {
        fonts: FONTS,
        pdfUA: true,
        info: { title: 'T' },
      }),
    );
    expect(pdf).toContain('<pdfuaid:part>1</pdfuaid:part>');
    expect(pdf).not.toContain('pdfaid:part'); // UA alone is not PDF/A
    expect(pdf).toContain('/StructTreeRoot');
    expect(pdf).toContain('/DisplayDocTitle true');
  });

  it('synthesizes a title when the source has none (Matterhorn 06/07)', async () => {
    const pdf = latin1(
      await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', { fonts: FONTS, pdfUA: true }),
    );
    expect(pdf).toContain('Untitled document');
    expect(pdf).toContain('/DisplayDocTitle true');
  });

  it('combines with PDF/A: both identifiers in one XMP', async () => {
    const pdf = latin1(
      await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', {
        fonts: FONTS,
        pdfA: 'PDF/A-2a',
        pdfUA: true,
        info: { title: 'T' },
      }),
    );
    expect(pdf).toContain('<pdfaid:part>2</pdfaid:part>');
    expect(pdf).toContain('<pdfuaid:part>1</pdfuaid:part>');
  });

  it('Note structure elements carry unique IDs (§7.9)', async () => {
    const body = '<w:p><w:r><w:t>text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const pdf = latin1(
      await Ream.parse(
        buildDocxFromBody(body, {
          footnotesXml:
            '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> note</w:t></w:r></w:p></w:footnote>',
        }),
      ).convert('pdf', { fonts: FONTS, pdfUA: true }),
    );
    expect(pdf).toMatch(/\/S \/Note[^>]*\/ID \(note-\d+\)|\/ID \(note-\d+\)[^>]*\/S \/Note/);
  });
});
