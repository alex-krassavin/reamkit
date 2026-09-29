// ECMA-376 Part 1 §17.3.3.1 — a paragraph that is nothing but a page break
// keeps its mark with the break, on the page the break ends; what follows
// starts the next page at its first line. Measured in Word for Mac, in every
// compatibility mode, by page counts and paragraph positions.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>';
const para = (text: string): string =>
  `<w:p><w:pPr><w:spacing w:after="0"/>${RPR}</w:pPr><w:r>${RPR}<w:t>${text}</w:t></w:r></w:p>`;
const BREAK = `<w:p><w:pPr><w:spacing w:after="0"/>${RPR}</w:pPr><w:r>${RPR}<w:br w:type="page"/></w:r></w:p>`;
const EMPTY = `<w:p><w:pPr><w:spacing w:after="0"/>${RPR}</w:pPr></w:p>`;
const SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

function pages(body: string): ReadonlyArray<ReadonlyArray<PageItem>> {
  const flow = Ream.parse(buildDocxFromBody(body + SECTION)).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.map((p) => p.commands);
}

const lines = (items: ReadonlyArray<PageItem>): ReadonlyArray<TextLineItem> =>
  items.filter((c): c is TextLineItem => c.type === 'line');

/** The first line's top on a page, from its baseline: 11pt Calibri's ascent above it. */
const firstLineTop = (items: ReadonlyArray<PageItem>): number =>
  lines(items)[0]!.baselineY - (11 * 1950) / 2048;

const texts = (items: ReadonlyArray<PageItem>): Array<string> =>
  lines(items).map((l) => l.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join(''));

describe('a paragraph that is only a page break (§17.3.3.1)', () => {
  it('stays on the page it ends, and the next paragraph starts the next page', () => {
    // Word: "three" at 72.00 on page 2, the break's mark on page 1.
    const laid = pages(para('one') + BREAK + para('three'));
    expect(laid).toHaveLength(2);
    expect(texts(laid[1]!)).toEqual(['three']);
    expect(firstLineTop(laid[1]!)).toBeCloseTo(72, 3);
  });

  it('makes no page of its own at the end of the document', () => {
    // Word: one page; and two with an empty paragraph after the break.
    expect(pages(para('one') + BREAK)).toHaveLength(1);
    expect(pages(para('one') + BREAK + EMPTY)).toHaveLength(2);
  });

  it('leaves a second one alone at the top of the page between', () => {
    // Word: three pages, "three" at 72.00 on the third.
    const laid = pages(para('one') + BREAK + BREAK + para('three'));
    expect(laid).toHaveLength(3);
    expect(texts(laid[2]!)).toEqual(['three']);
    expect(firstLineTop(laid[2]!)).toBeCloseTo(72, 3);
  });

  it('still takes the paragraph along when the break comes before its words', () => {
    const leading = `<w:p><w:r>${RPR}<w:br w:type="page"/></w:r><w:r>${RPR}<w:t>two</w:t></w:r></w:p>`;
    const laid = pages(para('one') + leading);
    expect(laid).toHaveLength(2);
    expect(texts(laid[1]!)).toEqual(['two']);
  });
});
