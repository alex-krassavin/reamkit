import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const pdfText = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');

describe('PDF link annotations (ISO 32000 §12.5.6.5)', () => {
  it('emits a /Link annotation with a /URI action for an external hyperlink', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>before </w:t></w:r>' +
        '<w:hyperlink r:id="rId30"><w:r><w:t>visit reamkit</w:t></w:r></w:hyperlink>' +
        '<w:r><w:t> after</w:t></w:r></w:p>',
      { hyperlinks: { rId30: 'https://reamkit.dev' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).toContain('/Subtype /Link');
    expect(pdf).toContain('/URI (https://reamkit.dev)');
    expect(pdf).toContain('/Annots');
    expect(pdf.match(/\/Subtype \/Link/g)!.length).toBe(1);
  });

  it('one rect per line: a wrapping link produces an annotation per line', async () => {
    const words = Array.from({ length: 12 }, (_, i) => `linkword${i}`).join(' ');
    const docx = buildDocxFromBody(
      `<w:p><w:hyperlink r:id="rId30"><w:r><w:t>${words}</w:t></w:r></w:hyperlink></w:p>`,
      { hyperlinks: { rId30: 'https://reamkit.dev/long' } },
    );
    // A narrow page forces the link text across several lines.
    const pdf = pdfText(
      await Ream.parse(docx).convert('pdf', { fonts: FONTS, pageWidth: 220, pageHeight: 400 }),
    );
    expect(pdf.match(/\/Subtype \/Link/g)!.length).toBeGreaterThan(1);
  });

  it('disallowed schemes produce no annotation (text still renders)', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:hyperlink r:id="rId31"><w:r><w:t>not clickable</w:t></w:r></w:hyperlink></w:p>',
      { hyperlinks: { rId31: 'javascript:alert(1)' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).not.toContain('/Subtype /Link');
    expect(pdf).not.toContain('javascript:');
  });

  it('tagged mode encloses the annotation in a Link StructElem (OBJR + StructParent)', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:hyperlink r:id="rId30"><w:r><w:t>tagged link</w:t></w:r></w:hyperlink></w:p>',
      { hyperlinks: { rId30: 'https://reamkit.dev' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS, tagged: true }));
    expect(pdf).toContain('/S /Link');
    expect(pdf).toContain('/Type /OBJR');
    expect(pdf).toContain('/StructParent ');
    expect(pdf).toContain('/F 4');
  });

  it('documents without links carry no /Annots at all', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>plain</w:t></w:r></w:p>');
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).not.toContain('/Annots');
  });
});
