import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { cjkBreakBetween, isCjkWide, splitCjkSegment } from '@/core/line-breaker/cjk';

const cp = (s: string): number => s.codePointAt(0)!;

describe('CJK line breaking helpers (UAX #14 subset)', () => {
  it('splits a pure-Chinese run into per-character pieces', () => {
    expect(splitCjkSegment('很长的中文')).toEqual(['很', '长', '的', '中', '文']);
  });

  it('keeps Latin runs grouped, splitting only at CJK boundaries', () => {
    expect(splitCjkSegment('订单ABC123编号')).toEqual(['订', '单', 'ABC123', '编', '号']);
  });

  it('returns a non-CJK segment unchanged (Latin tokenization is byte-stable)', () => {
    expect(splitCjkSegment('hello-world')).toEqual(['hello-world']);
    expect(splitCjkSegment('')).toEqual(['']);
  });

  it('recognises wide ranges incl. hiragana and astral Ext B', () => {
    expect(isCjkWide(cp('中'))).toBe(true);
    expect(isCjkWide(cp('あ'))).toBe(true);
    expect(isCjkWide('𠀀'.codePointAt(0)!)).toBe(true); // CJK Ext B (astral)
    expect(isCjkWide(cp('A'))).toBe(false);
    expect(isCjkWide(cp(' '))).toBe(false);
  });

  it('allows a break between two ideographs', () => {
    expect(cjkBreakBetween(cp('中'), cp('文'))).toBe(true);
  });

  it('forbids a break before closing punctuation (it clings to its char)', () => {
    expect(cjkBreakBetween(cp('子'), cp('。'))).toBe(false);
    expect(cjkBreakBetween(cp('话'), cp('，'))).toBe(false);
    expect(cjkBreakBetween(cp('好'), cp('）'))).toBe(false);
  });

  it('forbids a break after opening punctuation', () => {
    expect(cjkBreakBetween(cp('（'), cp('中'))).toBe(false);
  });

  it('allows a break after closing punctuation (the next line starts fresh)', () => {
    expect(cjkBreakBetween(cp('。'), cp('下'))).toBe(true);
  });

  it('allows a break at a CJK↔Latin boundary', () => {
    expect(cjkBreakBetween(cp('文'), cp('A'))).toBe(true);
    expect(cjkBreakBetween(cp('A'), cp('文'))).toBe(true);
  });

  it('never breaks between two Latin characters', () => {
    expect(cjkBreakBetween(cp('a'), cp('b'))).toBe(false);
  });
});

// End-to-end: the reported bug — long CJK content in a narrow container (table
// cell / narrow page) wrapped in Word but overflowed in the PDF because the
// whitespace-only tokenizer never opened a break opportunity in a space-less run.
const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

function lineCount(docx: Uint8Array): number {
  const flow = Ream.parse(docx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.flatMap((p) => p.commands).filter((c) => c.type === 'line').length;
}

// A single paragraph on a page whose text column is `pageTwips/20 − 72` pt wide.
const doc = (text: string, pageTwips: number): Uint8Array =>
  buildDocxFromBody(
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="${pageTwips}" w:h="16838"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`,
  );

const LONG_CJK = '这是一个非常长的中文句子它应该在狭窄的单元格里自动换行而不是溢出边界';
const NARROW = 3200; // ≈ 88pt text column
const WIDE = 16000; // ≈ 728pt text column

describe('CJK wrapping in a narrow container (regression for the reported issue)', () => {
  it('wraps a long CJK run into multiple lines when it is wider than the column', () => {
    expect(lineCount(doc(LONG_CJK, NARROW))).toBeGreaterThan(1);
  });

  it('keeps the same CJK run on one line when the column is wide enough', () => {
    expect(lineCount(doc(LONG_CJK, WIDE))).toBe(1);
  });

  it('leaves a short CJK run on a single line', () => {
    expect(lineCount(doc('短句', NARROW))).toBe(1);
  });

  it('still wraps ordinary English (sanity — whitespace breaking unchanged)', () => {
    const english = 'this is a fairly long english sentence that must wrap in a narrow column';
    expect(lineCount(doc(english, NARROW))).toBeGreaterThan(1);
  });
});
