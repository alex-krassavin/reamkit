// A legacy Word paragraph's outline level ([MS-DOC] §2.6.2 sprmPOutLvl, and the
// level Word's built-in "heading N" styles fix) — what the PDF outline (#138),
// tagged-PDF headings, HTML h1–h6 and markdown `#` are all read from. Every
// document here is synthetic; the expectations come from Word 16 for Mac,
// reading back two .doc files it saved itself (`outline level of paragraph
// format`, 2026-11-20).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDoc } from './fixtures/build-doc';
import type { DocParaFormat, DocStyleDef } from './fixtures/build-doc';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { readPdfOutline } from '@/pdf-reader/outline';
import { readDoc } from '@/word/doc/doc-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// §2.9.271 STSH keeps istd 1–9 for sti 1–9, Word's "heading 1" … "heading 9".
const STYLES: ReadonlyArray<DocStyleDef> = [
  { istd: 0, kind: 'paragraph' },
  // Saved by Word with no sprmPOutLvl of its own, as it saves a "heading 4"
  // whose .docx stated none: the level is the sti's all the same.
  { istd: 1, sti: 1, kind: 'paragraph', base: 0, para: { keepNext: true } },
  { istd: 2, sti: 2, kind: 'paragraph', base: 0, para: { outlineLevel: 1 } },
  // A "heading 5" whose own sprmPOutLvl says level 2: Word reads level 5.
  { istd: 5, sti: 5, kind: 'paragraph', base: 0, para: { outlineLevel: 1 } },
  // A style of the user's own at level 3.
  { istd: 15, kind: 'paragraph', base: 0, para: { outlineLevel: 2 } },
  // Based on that "heading 5", stating nothing: level 5, not 2.
  { istd: 16, kind: 'paragraph', base: 5 },
  // Based on "heading 1", stating body text: body text.
  { istd: 17, kind: 'paragraph', base: 1, para: { outlineLevel: 9 } },
  // Based on the user's level-3 style: level 3.
  { istd: 18, kind: 'paragraph', base: 15 },
];

const PARAGRAPHS: ReadonlyArray<{ readonly text: string } & Omit<DocParaFormat, 'length'>> = [
  { text: 'Introduction', istd: 1 },
  // A built-in heading keeps its level over the paragraph's own: §2.6.2 says
  // sprmPOutLvl "MUST be ignored" at istd 1–9, and Word reads a Heading 2
  // paragraph with a direct body-text level as level 2.
  { text: 'Background', istd: 2, outlineLevel: 9 },
  { text: 'Chapter', istd: 15 },
  { text: 'Body', istd: 0 },
  // On a style that is no built-in heading, the paragraph's level stands.
  { text: 'Stated body', istd: 15, outlineLevel: 9 },
  { text: 'Direct', istd: 0, outlineLevel: 4 },
  { text: 'Five', istd: 5 },
  { text: 'Under five', istd: 16 },
  { text: 'Not a heading', istd: 17 },
  { text: 'Deeper chapter', istd: 18 },
];

const DOC = buildDoc([{ text: PARAGRAPHS.map((p) => `${p.text}\r`).join(''), compressed: false }], {
  styles: STYLES,
  paraRuns: PARAGRAPHS.map(({ text, ...format }) => ({ length: text.length + 1, ...format })),
});

describe('.doc outline levels', () => {
  it('reads each paragraph’s level as Word does', () => {
    const levels = readDoc(DOC).doc.body.map((el) =>
      el.kind === 'paragraph'
        ? [el.paragraph.runs.map((r) => r.text).join(''), el.paragraph.properties.outlineLevel]
        : el.kind,
    );
    expect(levels).toEqual([
      ['Introduction', 0],
      ['Background', 1],
      ['Chapter', 2],
      ['Body', undefined],
      ['Stated body', 9],
      ['Direct', 4],
      ['Five', 4],
      ['Under five', 4],
      ['Not a heading', 9],
      ['Deeper chapter', 2],
    ]);
  });

  it('builds the PDF outline from them, nested by level', async () => {
    const pdf = await Ream.parse(DOC).convert('pdf', { fonts: FONTS });
    // A deeper item hangs off the nearest shallower one before it: level 5
    // under level 3, and the second level 3 under level 2.
    expect(readPdfOutline(PdfFile.parse(pdf)).map((i) => [i.level, i.title])).toEqual([
      [1, 'Introduction'],
      [2, 'Background'],
      [3, 'Chapter'],
      [4, 'Direct'],
      [4, 'Five'],
      [4, 'Under five'],
      [3, 'Deeper chapter'],
    ]);
  });

  it('gives markdown the same headings', async () => {
    const md = new TextDecoder().decode(await Ream.parse(DOC).convert('md'));
    expect(md.split('\n').filter((line) => line.startsWith('#'))).toEqual([
      '# Introduction',
      '## Background',
      '### Chapter',
      '##### Direct',
      '##### Five',
      '##### Under five',
      '### Deeper chapter',
    ]);
  });
});
