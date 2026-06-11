import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { flowRenderOptions } from '@/core/converter/project';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// Body that paginates to three pages via forced page breaks.
const THREE_PAGES =
  '<w:p><w:r><w:t>one</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>three</w:t></w:r></w:p>' +
  '<w:sectPr><w:footerReference w:type="default" r:id="rId11"/></w:sectPr>';

// "Page <PAGE> of <NUMPAGES>" — PAGE as a complex field (fldChar/instrText),
// NUMPAGES as a simple field, both with a stale cached result of "1".
const FOOTER =
  '<w:p><w:r><w:t>Page </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> PAGE \\* MERGEFORMAT </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t> of </w:t></w:r>' +
  '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';

function pageText(commands: ReadonlyArray<{ type: string }>): string {
  let out = '';
  for (const c of commands) {
    if (c.type !== 'line') continue;
    const line = (c as { line: { tokens: ReadonlyArray<{ kind: string; text?: string }> } }).line;
    for (const t of line.tokens) if (t.kind === 'text') out += t.text ?? '';
    out += '\n';
  }
  return out;
}

describe('PAGE / NUMPAGES fields (§17.16.5.33/.35)', () => {
  it('parses both field syntaxes into field runs', () => {
    const { doc } = readDocx(buildDocxFromBody(FOOTER));
    const para = doc.body[0]!;
    if (para.kind !== 'paragraph') throw new Error('expected paragraph');
    const fields = para.paragraph.runs.filter((r) => r.field !== undefined);
    expect(fields.map((r) => r.field)).toEqual(['PAGE', 'NUMPAGES']);
    expect(fields[0]!.text).toBe('1'); // cached result preserved
  });

  it('substitutes real numbers per page in the footer band', () => {
    const docx = buildDocxFromBody(THREE_PAGES, { footerXml: FOOTER });
    const flow = Ream.parse(docx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    expect(laid.pages.length).toBe(3);
    expect(pageText(laid.pages[0]!.commands)).toContain('Page 1 of 3');
    expect(pageText(laid.pages[1]!.commands)).toContain('Page 2 of 3');
    expect(pageText(laid.pages[2]!.commands)).toContain('Page 3 of 3');
    // The stale cache must not leak through anywhere.
    expect(pageText(laid.pages[1]!.commands)).not.toContain('Page 1 of 1');
  });

  it('renders the substituted numbers in the PDF (smoke)', async () => {
    const docx = buildDocxFromBody(THREE_PAGES, { footerXml: FOOTER });
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000); // converts; glyph identity is covered above
  });

  it('body fields keep their cached text (no page context in the flow)', () => {
    const body = '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple></w:p>';
    const { doc } = readDocx(buildDocxFromBody(body));
    const para = doc.body[0]!;
    if (para.kind !== 'paragraph') throw new Error('expected paragraph');
    expect(para.paragraph.runs[0]!.field).toBe('PAGE');
    expect(para.paragraph.runs[0]!.text).toBe('7');
  });

  it('unrecognized instructions keep their cached result runs', () => {
    const body =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> REF _Toc123 </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>Chapter 3</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const { doc } = readDocx(buildDocxFromBody(body));
    const para = doc.body[0]!;
    if (para.kind !== 'paragraph') throw new Error('expected paragraph');
    expect(para.paragraph.runs.map((r) => r.text).join('')).toBe('Chapter 3');
    expect(para.paragraph.runs.every((r) => r.field === undefined)).toBe(true);
  });
});
