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

describe('a page break among the words of a paragraph (§17.3.3.1)', () => {
  // Measured in Word for Mac by page counts and word positions: the break
  // happens where it stands in its run, and what follows it starts the next
  // page at its first line.
  const PAGE = '<w:br w:type="page"/>';
  const t = (text: string): string => `<w:t>${text}</w:t>`;
  const run = (inner: string, pPr = ''): string =>
    `<w:p><w:pPr>${pPr}<w:spacing w:after="0"/>${RPR}</w:pPr><w:r>${RPR}${inner}</w:r></w:p>`;

  it('takes the words after it in its own run to the next page', () => {
    // Word: "Second" at 72.00 on page 2.
    const laid = pages(para('Latin') + run(PAGE + t('Second')));
    expect(laid).toHaveLength(2);
    expect(texts(laid[0]!)).toEqual(['Latin']);
    expect(texts(laid[1]!)).toEqual(['Second']);
    expect(firstLineTop(laid[1]!)).toBeCloseTo(72, 3);
  });

  it('breaks the page between two words of one run', () => {
    // Word: "Latin" on page 1, "Second" at 72.00 on page 2.
    const laid = pages(run(t('Latin') + PAGE + t('Second')));
    expect(laid).toHaveLength(2);
    expect(texts(laid[0]!)).toEqual(['Latin']);
    expect(texts(laid[1]!)).toEqual(['Second']);
    expect(firstLineTop(laid[1]!)).toBeCloseTo(72, 3);
  });

  it('breaks it the same from a run of its own between two others', () => {
    const between =
      `<w:p><w:pPr><w:spacing w:after="0"/>${RPR}</w:pPr><w:r>${RPR}${t('Latin')}</w:r>` +
      `<w:r>${RPR}${PAGE}</w:r><w:r>${RPR}${t('Second')}</w:r></w:p>`;
    const laid = pages(between);
    expect(laid.map(texts)).toEqual([['Latin'], ['Second']]);
  });

  it('leaves the line before it on its page under widow control', () => {
    // Word: "one" and "Latin" on page 1, "Second" on page 2. Read as a line
    // the page could not hold, the break left "Latin" alone at its foot and
    // widow control moved it on.
    const widow = '<w:widowControl/>';
    const laid = pages(run(t('one'), widow) + run(t('Latin') + PAGE + t('Second'), widow));
    expect(laid.map(texts)).toEqual([['one', 'Latin'], ['Second']]);
  });

  it('leaves an empty page between two in a row', () => {
    // Word: three pages either way, "Second" at 72.00 on the third.
    for (const body of [
      run(t('Latin') + PAGE + PAGE + t('Second')),
      para('one') + run(PAGE + PAGE + t('Second')),
    ]) {
      const laid = pages(body);
      expect(laid).toHaveLength(3);
      expect(texts(laid[1]!)).toEqual(['']);
      expect(texts(laid[2]!)).toEqual(['Second']);
      expect(firstLineTop(laid[2]!)).toBeCloseTo(72, 3);
    }
  });

  it('sends what follows two pages on when two end the paragraph', () => {
    // Word: three pages, "next" on the third — after words or after nothing.
    for (const body of [
      run(t('Latin') + PAGE + PAGE) + para('next'),
      para('one') + run(PAGE + PAGE) + para('next'),
    ]) {
      const laid = pages(body);
      expect(laid).toHaveLength(3);
      expect(texts(laid[2]!)).toEqual(['next']);
    }
  });

  it('breaks no page and no line in a table cell', () => {
    // Word sets "Second" on the line of "Latin", where the break ends: there
    // is no page in a cell to break.
    const cell =
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
      run(t('Latin') + PAGE + t('Second')) +
      '</w:tc></w:tr></w:tbl>' +
      para('after');
    const laid = pages(cell);
    expect(laid).toHaveLength(1);
    expect(texts(laid[0]!)).toEqual(['LatinSecond', 'after']);
  });
});
