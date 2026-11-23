// ECMA-376 Part 1 §17.3.3.1 `w:br`, §17.3.3.4 `w:cr` — a soft break at the
// very end of a paragraph still ends a line, and the paragraph mark stands on
// the next one by itself (issue #144). Every expectation was measured in Word
// for Mac on Courier New probes (compatibility mode 15, positions by
// AppleScript); LibreOffice sets them the same. Word's Courier New line is
// 2320/2048 of the size, 13.59pt at 12pt.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Line, PageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** Word's Courier New line and ascent at a size. */
const L = (sizePt: number): number => (sizePt * (1705 + 615)) / 2048;
const ASCENT = (sizePt: number): number => (sizePt * 1705) / 2048;

const rPr = (sizePt = 12): string =>
  `<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="${sizePt * 2}"/></w:rPr>`;
const t = (text: string, sizePt = 12): string =>
  `<w:r>${rPr(sizePt)}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const br = (sizePt = 12, type?: string): string =>
  `<w:r>${rPr(sizePt)}<w:br${type === undefined ? '' : ` w:type="${type}"`}/></w:r>`;
const tab = (sizePt = 12): string => `<w:r>${rPr(sizePt)}<w:tab/></w:r>`;

interface ParaOptions {
  /** pPr children that come before `w:spacing` in schema order. */
  readonly pre?: string;
  /** The `w:spacing` attributes. */
  readonly spacing?: string;
  /** pPr children that come after it. */
  readonly post?: string;
  /** The paragraph mark's size. */
  readonly mark?: number;
}

const para = (content: string, o: ParaOptions = {}): string =>
  `<w:p><w:pPr>${o.pre ?? ''}<w:spacing w:after="0" ${o.spacing ?? ''}/>${o.post ?? ''}${rPr(o.mark)}</w:pPr>${content}</w:p>`;

const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';

function pages(body: string, headerXml?: string): ReadonlyArray<ReadonlyArray<PageItem>> {
  const section =
    headerXml === undefined
      ? `<w:sectPr>${PAGE}</w:sectPr>`
      : `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>${PAGE}</w:sectPr>`;
  const flow = Ream.parse(
    buildDocxFromBody(body + section, headerXml === undefined ? {} : { headerXml }),
  ).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.map((p) => p.commands);
}

const textOf = (line: Line): string =>
  line.tokens.map((tk) => (tk.kind === 'text' ? tk.text : '')).join('');

/** Every drawn line, with the page it stands on. */
function lineItems(body: string, headerXml?: string): Array<TextLineItem & { page: number }> {
  return pages(body, headerXml).flatMap((items, page) =>
    items.filter((c): c is TextLineItem => c.type === 'line').map((c) => ({ ...c, page })),
  );
}

/** Where the 12pt line holding `text` starts: its baseline less the ascent. */
function top(items: ReadonlyArray<TextLineItem & { page: number }>, text: string): number {
  const item = items.find((c) => textOf(c.line).includes(text));
  if (!item) throw new Error(`no line holds ${text}`);
  return item.baselineY - ASCENT(12);
}

/** How far the paragraph holding `content` reaches: from "A" above it to "B" below. */
function height(content: string, o: ParaOptions = {}): number {
  const items = lineItems(para(t('A')) + para(content, o) + para(t('B')));
  return top(items, 'B') - top(items, 'A') - L(12);
}

describe('a soft break at the end of a paragraph (§17.3.3.1, issue #144)', () => {
  it('leaves the paragraph mark a line of its own', () => {
    // Word: "P01" and a `w:br` stand 27.7pt from the next paragraph, two lines.
    expect(height(t('P01') + br())).toBeCloseTo(2 * L(12), 3);
  });

  it("does so after the issue's heading, a tab and a break", () => {
    const heading = t('Scanned Assets', 13) + tab(13) + br(13);
    expect(height(heading, { mark: 13 })).toBeCloseTo(2 * L(13), 3);
  });

  it('makes a paragraph of nothing but a break two lines tall', () => {
    // Word: 27.7pt — the issue's cover page opens with one at 22pt.
    expect(height(br())).toBeCloseTo(2 * L(12), 3);
    expect(height(br(22), { mark: 22 })).toBeCloseTo(2 * L(22), 3);
  });

  it("sets the break's line in the break's size and the last in the mark's", () => {
    // Word: a 22pt break over a 12pt mark is 24.5 + 13.45pt, a 12pt break over
    // a 22pt mark 13.45 + 24.55pt.
    expect(height(br(22))).toBeCloseTo(L(22) + L(12), 3);
    expect(height(br(12), { mark: 22 })).toBeCloseTo(L(12) + L(22), 3);
  });

  it('counts the break for nothing on a line with text, and the mark for the last line', () => {
    // Word: 12pt text and a 24pt break stand 13.45 + 13.45pt; with a 24pt mark
    // the empty line is 27.65pt.
    expect(height(t('P06') + br(24))).toBeCloseTo(2 * L(12), 3);
    expect(height(t('P07') + br(), { mark: 24 })).toBeCloseTo(L(12) + L(24), 3);
  });

  it('measures the last line by the mark, whatever spaces or tabs follow the break', () => {
    // Word: a 24pt space or a 24pt tab after the break leaves the line 12pt.
    expect(height(t('P08') + br() + t(' '))).toBeCloseTo(2 * L(12), 3);
    expect(height(t('P16') + br(24) + t(' ', 24))).toBeCloseTo(2 * L(12), 3);
    expect(height(t('P09') + br() + tab(24))).toBeCloseTo(2 * L(12), 3);
    // …while a letter there is a line like any other: 13.45 + 27.7pt.
    expect(height(t('P07') + br() + t('Z', 24))).toBeCloseTo(L(12) + L(24), 3);
  });

  it('leaves a line for every break, the middle ones in their own size', () => {
    // Word: two breaks, three lines; a 12pt and then a 24pt break stand
    // 13.45, 27.2 and 13.45pt.
    expect(height(t('P10') + br() + br())).toBeCloseTo(3 * L(12), 3);
    expect(height(t('P17') + br() + br(24))).toBeCloseTo(2 * L(12) + L(24), 3);
  });

  it("spaces the mark's line like the paragraph's other lines", () => {
    // Word: 54.6pt at double spacing, 40.35 at exactly 20pt, 40 at least 20pt.
    expect(height(t('P12') + br(), { spacing: 'w:line="480" w:lineRule="auto"' })).toBeCloseTo(
      4 * L(12),
      3,
    );
    expect(height(t('P13') + br(), { spacing: 'w:line="400" w:lineRule="exact"' })).toBeCloseTo(
      40,
      3,
    );
    expect(height(t('P14') + br(), { spacing: 'w:line="400" w:lineRule="atLeast"' })).toBeCloseTo(
      40,
      3,
    );
  });

  it('reads a carriage return as a break, mid-paragraph and at the end (§17.3.3.4)', () => {
    const cr = `<w:r>${rPr()}<w:cr/></w:r>`;
    expect(height(t('P11') + cr)).toBeCloseTo(2 * L(12), 3);
    const items = lineItems(para(t('left') + cr + t('right')));
    expect(items.map((c) => textOf(c.line))).toEqual(['left', 'right']);
  });

  it('justifies the line before the break, which is no longer the last', () => {
    // Word stretches "short one" across the whole measure when a break ends it,
    // as it does any line a soft break ends in a justified paragraph.
    const items = lineItems(para(t('short one') + br(), { post: '<w:jc w:val="both"/>' }));
    const line = items.find((c) => textOf(c.line) === 'short one')!.line;
    expect(line.isLastInParagraph).toBe(false);
  });

  it('does so in a table cell and in a header', () => {
    // Word: a cell of "C1" and a break makes its row two lines tall; a header
    // paragraph so ended pushes the next one down a line.
    const cell = (content: string): string =>
      `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>' +
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr>${cell(para(t('C1') + br()))}</w:tr><w:tr>${cell(para(t('C2')))}</w:tr></w:tbl>`;
    const rows = lineItems(table + para(t('after')));
    expect(top(rows, 'C2') - top(rows, 'C1')).toBeCloseTo(2 * L(12), 3);

    // Word: the header, 36pt from the paper's top, holds three lines and so
    // pushes the body below the 72pt margin.
    const body = lineItems(para(t('body')), para(t('H1') + br()) + para(t('H2')));
    expect(top(body, 'body')).toBeCloseTo(36 + 3 * L(12), 3);
  });
});

describe('a page or column break at the end of a paragraph (§17.3.3.1)', () => {
  it('keeps the mark with a page break: the next paragraph opens the next page', () => {
    // Word: "A2" and a page break on page 1, "A3" at 72.00 on page 2.
    const items = lineItems(para(t('A1')) + para(t('A2') + br(12, 'page')) + para(t('A3')));
    const a3 = items.find((c) => textOf(c.line) === 'A3')!;
    expect(a3.page).toBe(1);
    expect(top(items, 'A3')).toBeCloseTo(72, 3);
  });

  it('sends the mark after a column break on to the next column, a line of its own', () => {
    // Word, one column: "C1" and the break on page 1, the mark alone at 72.00
    // on page 2 and "C2" a line below it.
    const items = lineItems(para(t('C1') + br(12, 'column')) + para(t('C2')));
    const c2 = items.find((c) => textOf(c.line) === 'C2')!;
    expect(c2.page).toBe(1);
    expect(top(items, 'C2')).toBeCloseTo(72 + L(12), 3);
  });

  it('gives a paragraph of nothing but a column break a line in each column', () => {
    // Word: the break on the line under "M1", its mark at the top of page 2,
    // "M3" a line below.
    const items = lineItems(para(t('M1')) + para(br(12, 'column')) + para(t('M3')));
    const m3 = items.find((c) => textOf(c.line) === 'M3')!;
    expect(m3.page).toBe(1);
    expect(top(items, 'M3')).toBeCloseTo(72 + L(12), 3);
  });

  it("does not balance a continuous section's columns around a column break (§17.6.4)", () => {
    // Word, three columns: "D1", a column break and eight lines stand one,
    // eight and none, and the next section starts under the eight; eight lines
    // with no break stand four and four. (LibreOffice balances past the break.)
    const cols = (n: number): string =>
      `<w:type w:val="continuous"/>${PAGE}<w:cols w:num="${n}" w:space="360"/>`;
    const lines = [2, 3, 4, 5, 6, 7, 8].map((i) => para(t(`D${i}`))).join('');
    const items = lineItems(
      para(t('TOP'), { post: `<w:sectPr>${PAGE}</w:sectPr>` }) +
        para(t('D1') + br(12, 'column')) +
        lines +
        para(t('D9'), { post: `<w:sectPr>${cols(3)}</w:sectPr>` }) +
        para(t('AFTER'), { post: `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>` }) +
        para(t('END')),
    );
    const d1 = items.find((c) => textOf(c.line) === 'D1')!;
    const column = (text: string): number => items.find((c) => textOf(c.line) === text)!.originX;
    expect(column('D2')).toBeGreaterThan(d1.originX);
    expect(column('D9')).toBe(column('D2'));
    // The mark after the break opens the second column; "D9" is its ninth line.
    expect(top(items, 'AFTER') - top(items, 'D1')).toBeCloseTo(9 * L(12), 3);
  });

  it('starts the next continuous section under the longest column', () => {
    // Word: "A1"…"A6" and a column break end the section; "AFTER" stands a
    // line under "A6", not under the mark alone in the second column.
    const cols2 = `<w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/>`;
    const six = [1, 2, 3, 4, 5].map((i) => para(t(`A${i}`))).join('');
    const items = lineItems(
      para(t('TOP'), { post: `<w:sectPr>${PAGE}</w:sectPr>` }) +
        six +
        para(t('A6') + br(12, 'column'), { post: `<w:sectPr>${cols2}</w:sectPr>` }) +
        para(t('AFTER'), { post: `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>` }) +
        para(t('END')),
    );
    expect(top(items, 'AFTER') - top(items, 'A6')).toBeCloseTo(L(12), 3);
  });

  it('leaves a column of its own to the second of two column breaks', () => {
    // Word: "N1" on page 1, the second break alone on page 2, "N2" on page 3.
    const items = lineItems(
      para(t('N1') + br(12, 'column') + br(12, 'column') + t('N2')) + para(t('end')),
    );
    expect(items.find((c) => textOf(c.line) === 'N2')?.page).toBe(2);
  });
});
