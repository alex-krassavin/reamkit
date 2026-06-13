// E-PDF EP8 — link annotations. A docx hyperlink becomes a /Link annotation in
// the PDF Ream writes; reading that PDF back must re-attach the URI to the run
// over which the annotation's /Rect sits, so the link survives onward to HTML.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement } from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import { Ream } from '@/core/converter/ream';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const LINK = 'https://example.com/page';

const linkDocx = (): Uint8Array =>
  buildDocxFromBody(
    `<w:p><w:r><w:t>Visit </w:t></w:r>` +
      `<w:hyperlink r:id="rId50"><w:r><w:t>our website</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t> today.</w:t></w:r></w:p>`,
    { hyperlinks: { rId50: LINK } },
  );

function hrefs(flow: FlowDoc): Array<string> {
  const out: Array<string> = [];
  const visit = (el: BodyElement): void => {
    if (el.kind === 'paragraph') {
      for (const r of el.paragraph.runs) if (r.href) out.push(r.href);
    } else if (el.kind === 'table') {
      for (const row of el.table.rows)
        for (const cell of row.cells) for (const child of cell.content) visit(child);
    }
  };
  for (const el of flow.body) visit(el);
  return out;
}

describe('PDF link annotations → hrefs (E-PDF EP8)', () => {
  it('recovers a hyperlink from an untagged PDF', async () => {
    const pdf = await Ream.parse(linkDocx()).convert('pdf', { fonts: FONTS });
    expect(hrefs(Ream.parse(pdf).flow)).toContain(LINK);
  });

  it('recovers a hyperlink from a tagged PDF', async () => {
    const pdf = await Ream.parse(linkDocx()).convert('pdf', { fonts: FONTS, tagged: true });
    expect(hrefs(Ream.parse(pdf).flow)).toContain(LINK);
  });

  it('carries the recovered link into HTML output', async () => {
    const pdf = await Ream.parse(linkDocx()).convert('pdf', { fonts: FONTS });
    const html = new TextDecoder().decode(await Ream.parse(pdf).convert('html'));
    expect(html).toContain(`href="${LINK}"`);
  });
});
