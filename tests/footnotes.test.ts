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

function layoutOf(docx: Uint8Array, extra: Record<string, unknown> = {}) {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
    ...extra,
  });
}

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

  it('layout: notes land at the bottom of the referencing page with a separator', () => {
    const body =
      BODY + '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>page two</w:t></w:r></w:p>';
    const laid = layoutOf(buildDocxFromBody(body, { footnotesXml: FOOTNOTES }));
    expect(laid.pages.length).toBe(2);
    const p1 = pageText(laid.pages[0]!.commands);
    expect(p1).toContain('first note');
    expect(p1).toContain('second note');
    // Reference numbers render in the body text (1 then 2, reading order).
    expect(p1).toContain('alpha1');
    expect(p1).toContain('beta2');
    // The separator rule: a short 0.75pt fill.
    const sep = laid.pages[0]!.commands.find(
      (c) => c.type === 'fill' && Math.abs((c as { width: number }).width - 144) < 0.01,
    );
    expect(sep).toBeDefined();
    // Page 2 carries neither the notes nor the separator.
    expect(pageText(laid.pages[1]!.commands)).not.toContain('first note');
  });

  it('layout: a reference near the page bottom moves to the next page WITH its note', () => {
    const filler = Array.from(
      { length: 12 },
      (_, i) => `<w:p><w:r><w:t>filler line ${i}</w:t></w:r></w:p>`,
    ).join('');
    const refPara =
      '<w:p><w:r><w:t>ref here</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>';
    const notes =
      '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> tall note line one</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>tall note line two</w:t></w:r></w:p></w:footnote>';
    const laid = layoutOf(buildDocxFromBody(filler + refPara, { footnotesXml: notes }), {
      pageHeight: 260,
      marginTop: 36,
      marginBottom: 36,
    });
    expect(laid.pages.length).toBeGreaterThan(1);
    // The page holding the reference also holds the note text.
    const refPage = laid.pages.findIndex((p) => pageText(p.commands).includes('ref here'));
    expect(refPage).toBeGreaterThanOrEqual(0);
    const text = pageText(laid.pages[refPage]!.commands);
    expect(text).toContain('tall note line one');
    expect(text).toContain('tall note line two');
  });

  it('layout: endnotes flow after the body', () => {
    const endnotes =
      '<w:endnote w:id="3"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t> closing remark</w:t></w:r></w:p></w:endnote>';
    const body =
      '<w:p><w:r><w:t>main text</w:t></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p>';
    const laid = layoutOf(buildDocxFromBody(body, { endnotesXml: endnotes }));
    const text = pageText(laid.pages[laid.pages.length - 1]!.commands);
    expect(text).toContain('main text');
    expect(text).toContain('closing remark');
    expect(text.indexOf('closing remark')).toBeGreaterThan(text.indexOf('main text'));
  });
});
