// §17.3.2.43 `w:w` — a run set at a share of its face's width: laid out that
// narrow, drawn through the PDF's horizontal scaling (ISO 32000-1 §9.3.3), and
// written back as `w:w`. And §17.3.2.35 `w:spacing`, which the writer read and
// never wrote.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement, RunProperties } from '@/core/document-model';
import type { TextLineItem, TextToken } from '@/layout/page-doc';
import { FontRegistry } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { renderStyledPdf } from '@/pdf/styled-page-renderer';
import { readDocx } from '@/word/docx-reader';
import { writeDocx } from '@/word/docx-writer';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** "Hi Hi", its space set with `space`'s properties. */
const words = (space: RunProperties): ReadonlyArray<BodyElement> => [
  {
    kind: 'paragraph',
    paragraph: {
      properties: {},
      runs: [
        { text: 'Hi', properties: {} },
        { text: ' ', properties: space },
        { text: 'Hi', properties: {} },
      ],
    },
  },
];

const spaceWidth = (body: ReadonlyArray<BodyElement>): number => {
  const laid = layoutStyledDocument(body, {
    registry: FontRegistry.fromBytes(FONTS),
    styles: EMPTY_STYLE_SHEET,
  });
  const line = laid.pages[0]!.commands.find((c): c is TextLineItem => c.type === 'line')!;
  const space = line.line.tokens.find(
    (t): t is TextToken => t.kind === 'text' && t.isSpace && t.text === ' ',
  );
  return space!.widthPt;
};

describe('a run set at a share of its width (§17.3.2.43)', () => {
  it('is laid out that narrow', () => {
    expect(spaceWidth(words({ widthScale: 0.5 }))).toBeCloseTo(spaceWidth(words({})) / 2, 6);
  });

  it('is drawn through the horizontal scaling, and the next run at its own width', () => {
    const pdf = renderStyledPdf(words({ widthScale: 0.5 }), {
      registry: FontRegistry.fromBytes(FONTS),
      styles: EMPTY_STYLE_SHEET,
    });
    const text = Buffer.from(pdf).toString('latin1');
    expect(text).toContain('50 Tz');
    expect(text).toContain('100 Tz');
  });

  it('is written as `w:w`, after the spacing and before the kerning', () => {
    // The spacing is read from a .docx; the share is the model's own.
    const { doc } = readDocx(
      buildDocxFromBody(
        '<w:p><w:r><w:rPr><w:spacing w:val="-20"/><w:kern w:val="2"/></w:rPr><w:t>Hi</w:t></w:r></w:p>',
      ),
    );
    const para = doc.body[0]!;
    if (para.kind !== 'paragraph') throw new Error('a paragraph');
    const run = para.paragraph.runs[0]!;
    const body: ReadonlyArray<BodyElement> = [
      {
        ...para,
        paragraph: {
          ...para.paragraph,
          runs: [{ ...run, properties: { ...run.properties, widthScale: 0.8 } }],
        },
      },
    ];
    const xml = new TextDecoder().decode(
      OpcPackage.open(writeDocx({ ...doc, body }).bytes).getMainDocument().data,
    );
    expect(xml).toContain('<w:spacing w:val="-20"/><w:w w:val="80"/><w:kern w:val="2"/>');
  });
});
