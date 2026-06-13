// E-DOCX WT2 — footnote/endnote write-back. A docx footnote survives the
// round-trip docx → FlowDoc → docx: the body keeps its w:footnoteReference and a
// re-emitted word/footnotes.xml carries the note content.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';

const BODY =
  '<w:p><w:r><w:t>Body text</w:t></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
  '<w:r><w:t> continues.</w:t></w:r></w:p>';

const FOOTNOTES =
  '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
  '<w:footnoteRef/></w:r><w:r><w:t>The footnote body.</w:t></w:r></w:p></w:footnote>';

const noteText = (blocks: ReadonlyArray<BodyElement> | undefined): string =>
  (blocks ?? [])
    .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
    .join('')
    .replace(/\s+/g, ' ')
    .trim();

const hasFootnoteRef = (body: ReadonlyArray<BodyElement>, id: string): boolean =>
  body.some((el) => el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.footnoteRef === id));

describe('docx footnote write-back (WT2)', () => {
  it('round-trips a footnote through docx → FlowDoc → docx', async () => {
    const docx = buildDocxFromBody(BODY, { footnotesXml: FOOTNOTES });
    // Sanity: the reader picked the footnote up.
    expect(Ream.parse(docx).flow.footnotes?.size).toBe(1);

    const written = await Ream.parse(docx).convert('docx');
    const flow = Ream.parse(written).flow;

    expect(flow.footnotes?.size).toBe(1);
    expect(noteText(flow.footnotes?.get('1'))).toContain('The footnote body.');
    expect(hasFootnoteRef(flow.body, '1')).toBe(true);
  });
});
