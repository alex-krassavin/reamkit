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

const decode = (b: Uint8Array) => new TextDecoder().decode(b);
const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b);

// "See chapter 2" links to a bookmark wrapping a heading on page 2.
const BODY =
  '<w:p><w:r><w:t>See </w:t></w:r>' +
  '<w:hyperlink w:anchor="_Toc42"><w:r><w:t>chapter two</w:t></w:r></w:hyperlink>' +
  '<w:r><w:t> for details.</w:t></w:r></w:p>' +
  '<w:bookmarkStart w:id="0" w:name="orphan"/><w:bookmarkEnd w:id="0"/>' +
  '<w:p><w:pPr><w:pageBreakBefore/></w:pPr>' +
  '<w:bookmarkStart w:id="1" w:name="_Toc42"/>' +
  '<w:r><w:t>Chapter Two</w:t></w:r>' +
  '<w:bookmarkEnd w:id="1"/>' +
  '<w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>' +
  '</w:p>';

describe('bookmarks + internal links (§17.13.6.2 / §17.16.22 @anchor)', () => {
  it('parses bookmark names onto paragraphs and anchors onto runs', () => {
    const { doc } = readDocx(buildDocxFromBody(BODY));
    const first = doc.body[0]!;
    if (first.kind !== 'paragraph') throw new Error('expected paragraph');
    const linked = first.paragraph.runs.find((r) => r.anchor !== undefined);
    expect(linked?.anchor).toBe('_Toc42');
    expect(linked?.text).toBe('chapter two');

    const second = doc.body[1]!;
    if (second.kind !== 'paragraph') throw new Error('expected paragraph');
    // Own bookmark + the body-level one carried over; _GoBack filtered.
    expect(second.paragraph.bookmarks).toEqual(['orphan', '_Toc42']);
  });

  it('layout records the destination on the right page', () => {
    const flow = Ream.parse(buildDocxFromBody(BODY)).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    expect(laid.pages.length).toBe(2);
    const pos = laid.pdf.bookmarks.get('_Toc42');
    expect(pos?.pageIdx).toBe(1);
    expect(pos!.yTopPt).toBeGreaterThan(0);
  });

  it('PDF carries a GoTo annotation and a /Names /Dests entry (referenced only)', async () => {
    const pdf = latin1(await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', { fonts: FONTS }));
    expect(pdf).toContain('/S /GoTo');
    expect(pdf).toContain('(_Toc42)');
    expect(pdf).toContain('/Dests');
    expect(pdf).toContain('/FitH');
    // The unreferenced bookmark gets no destination entry.
    expect(pdf).not.toContain('(orphan)');
  });

  it('HTML emits fragment links and ids for referenced bookmarks only', async () => {
    const html = decode(await Ream.parse(buildDocxFromBody(BODY)).convert('html'));
    expect(html).toContain('<a href="#bm-_Toc42">');
    expect(html).toContain('id="bm-_Toc42"');
    expect(html).not.toContain('id="bm-orphan"');
  });

  it('documents without internal links are unchanged (no Dests, no GoTo)', async () => {
    const plain = '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>';
    const pdf = latin1(await Ream.parse(buildDocxFromBody(plain)).convert('pdf', { fonts: FONTS }));
    expect(pdf).not.toContain('/Dests');
    expect(pdf).not.toContain('/GoTo');
  });
});
