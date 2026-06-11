import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { readDocx } from '@/word/docx-reader';

const decode = (b: Uint8Array) => new TextDecoder().decode(b);

const FOOTNOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> second note</w:t></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> first note</w:t></w:r></w:p></w:footnote>';

const BODY =
  '<w:p><w:r><w:t>alpha</w:t></w:r>' +
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
  '<w:r><w:t> beta</w:t></w:r>' +
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' +
  '</w:p>';

describe('footnotes / endnotes (§17.11)', () => {
  it('parses references and the notes part (separators skipped)', () => {
    const { doc } = readDocx(buildDocxFromBody(BODY, { footnotesXml: FOOTNOTES }));
    const para = doc.body[0]!;
    if (para.kind !== 'paragraph') throw new Error('expected paragraph');
    const refs = para.paragraph.runs.filter((r) => r.footnoteRef !== undefined);
    expect(refs.map((r) => r.footnoteRef)).toEqual(['1', '2']);
    expect(doc.footnotes?.size).toBe(2); // separator stubs excluded
    const note1 = doc.footnotes?.get('1')?.[0];
    expect(note1?.kind === 'paragraph' && note1.paragraph.runs.some((r) => r.noteNumber)).toBe(
      true,
    );
  });

  it('HTML: superscript anchors, ordered notes section, backlinks', async () => {
    const html = decode(
      await Ream.parse(buildDocxFromBody(BODY, { footnotesXml: FOOTNOTES })).convert('html'),
    );
    expect(html).toContain('<sup><a href="#fn-1" id="fnref-1">1</a></sup>');
    expect(html).toContain('<sup><a href="#fn-2" id="fnref-2">2</a></sup>');
    // Numbering follows reference order, not part order: id=1 → note 1.
    const fn1 = html.indexOf('<li id="fn-1">');
    const fn2 = html.indexOf('<li id="fn-2">');
    expect(fn1).toBeGreaterThan(-1);
    expect(fn2).toBeGreaterThan(fn1);
    expect(html.slice(fn1, fn2)).toContain('first note');
    expect(html).toContain('href="#fnref-1"'); // backlink
  });

  it('HTML: endnotes render in their own section', async () => {
    const endnotes =
      '<w:endnote w:id="5"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t> the endnote</w:t></w:r></w:p></w:endnote>';
    const body = '<w:p><w:r><w:t>text</w:t></w:r><w:r><w:endnoteReference w:id="5"/></w:r></w:p>';
    const html = decode(
      await Ream.parse(buildDocxFromBody(body, { endnotesXml: endnotes })).convert('html'),
    );
    expect(html).toContain('<sup><a href="#en-1" id="enref-1">1</a></sup>');
    expect(html).toContain('<li id="en-1">');
    expect(html).toContain('the endnote');
  });
});
