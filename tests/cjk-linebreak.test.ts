import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildPptx } from './fixtures/build-pptx';
import { buildXlsx } from './fixtures/build-xlsx';
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

// Issue #131. Total-fit rates a line by how far its glue stretches, and a line
// of ideographs has none: its break opportunities are penalties between glyph
// boxes. When every way to break the paragraph holds such a line, every answer
// scored ∞ and the DP's tie-break chose: it filled the lines from the END and
// left the remainder on the FIRST line — PowerPoint sets a 33-character
// sentence in an eight-character box 8/8/8/8/1, and we set it 1/8/8/8/8.
// Such a paragraph now breaks first-fit, as PowerPoint and Word set it.
describe('a paragraph total-fit cannot rate breaks first-fit (issue #131)', () => {
  type Profile = 'ream' | 'word';
  const textLines = (bytes: Uint8Array, profile: Profile, fonts = FONTS): Array<string> => {
    const flow = Ream.parse(bytes).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(fonts),
      ...flowRenderOptions(flow),
      ...(profile === 'word' ? { layoutProfile: 'word' as const } : {}),
    });
    return laid.pages
      .flatMap((p) => p.commands)
      .flatMap((c) =>
        c.type === 'line'
          ? [c.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('')]
          : [],
      )
      .map((t) => t.replace(/\u200b/g, ''))
      .filter((t) => t.trim().length > 0);
  };

  // The issue's minimal case: one wrapping text box, 18pt zh-CN.
  const deck = (text: string, cx: number): Uint8Array =>
    buildPptx([
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="${cx}" cy="2000000"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>` +
        `<a:rPr lang="zh-CN" sz="1800"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`,
    ]);
  const SENTENCE = '这是一个非常长的中文句子它应该在狭窄的文本框里自动换行而不是溢出边';

  it('fills the first line of a CJK text box, leaving the remainder last', () => {
    const lines = textLines(deck(SENTENCE, 1016000), 'ream');
    expect(lines.length).toBeGreaterThan(2);
    expect(lines.join('')).toBe(SENTENCE);
    // Every line but the last as full as the first; the remainder at the end.
    const full = [...lines[0]!].length;
    expect(full).toBeGreaterThan(1);
    for (const l of lines.slice(0, -1)) expect([...l].length).toBe(full);
    expect([...lines.at(-1)!].length).toBeLessThanOrEqual(full);
    // …which is the 'word' profile's first-fit, line for line.
    expect(lines).toEqual(textLines(deck(SENTENCE, 1016000), 'word'));
  });

  it('breaks Japanese first-fit too, keeping its non-starters off the line start', () => {
    // っ ょ and 。 must not begin a line (UAX #14 NS/CL): the first-fit answer
    // works around them the way the 'word' profile does.
    const ja =
      'これは非常に長い日本語の文章で狭いテキストボックスの中で自動的に折り返されるべきです。' +
      'ちょっとだけ待ってください。';
    const bytes = doc(ja, NARROW);
    const lines = textLines(bytes, 'ream');
    expect(lines.length).toBeGreaterThan(2);
    expect(lines).toEqual(textLines(bytes, 'word'));
    for (const l of lines.slice(1)) expect('っょ。、').not.toContain(l[0]);
  });

  it('breaks a narrow column of words first-fit when one word must stand alone', () => {
    // A word that fills the column alone is a line with no space to stretch,
    // so this paragraph had only infinite answers as well — and set one "ab"
    // on a line of its own under it, above four lines of six.
    const words = ['internationalization', ...Array<string>(25).fill('ab')].join(' ');
    const bytes = doc(words, 3400);
    const lines = textLines(bytes, 'ream');
    expect(lines[0]).toBe('internationalization');
    expect(lines[1]!.split(' ').length).toBeGreaterThan(1);
    expect(lines.at(-1)).toBe('ab');
    expect(lines).toEqual(textLines(bytes, 'word'));
  });

  it('wraps a CJK cell first-fit in a spreadsheet', () => {
    const stylesXml =
      `<fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill/></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf/><xf applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>`;
    const lines = textLines(
      buildXlsx({
        rows: [[{ value: SENTENCE, styleIndex: 1 }]],
        columns: [{ min: 1, max: 1, widthChars: 12 }],
        stylesXml,
      }),
      'ream',
    );
    expect(lines.length).toBeGreaterThan(2);
    expect(lines.join('')).toBe(SENTENCE);
    const full = [...lines[0]!].length;
    for (const l of lines.slice(0, -1)) expect([...l].length).toBe(full);
  });
});
