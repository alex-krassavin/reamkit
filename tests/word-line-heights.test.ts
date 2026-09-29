// ECMA-376 Part 1 §17.3.1.33 — a line of a Word document stands as tall as
// Word sets it: the line of the faces on it, not a flat 1.2× of the size. Every
// expectation below was measured in Word for Mac, averaged over 10–20 lines
// (Word reports positions in steps of 0.05pt or coarser): the figure in each
// comment is Word's, the one tested is the rule's.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildDoc } from './fixtures/build-doc';
import { buildTinyPng } from './fixtures/build-png';
import type { TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry, parseTtf } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDoc } from '@/word/doc/doc-reader';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** A face's line at a size, from its ascent, gap and descent in 2048ths. */
const line = (sizePt: number, ascent: number, gap: number, descent: number): number =>
  (sizePt * (ascent + gap + descent)) / 2048;

const CALIBRI_11 = line(11, 1950, 0, 550);

const rPr = (font: string, sizePt = 11): string =>
  `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="${sizePt * 2}"/></w:rPr>`;
const run = (font: string, text: string, sizePt = 11): string =>
  `<w:r>${rPr(font, sizePt)}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A paragraph single-spaced (or at `line` 240ths) with nothing after it. */
const para = (content: string, pPr = '', spacing = 240): string =>
  `<w:p><w:pPr>${pPr}<w:spacing w:after="0" w:line="${spacing}" w:lineRule="auto"/></w:pPr>${content}</w:p>`;

const SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

function lines(docx: Uint8Array): ReadonlyArray<TextLineItem> {
  const flow = Ream.parse(docx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line');
}

/** How far apart the baselines of the second and third lines stand. */
function lineHeight(docx: Uint8Array): number {
  const [, second, third] = lines(docx);
  return third!.baselineY - second!.baselineY;
}

/** Three paragraphs of the same content. */
const three = (content: string, pPr = '', spacing = 240): Uint8Array =>
  buildDocxFromBody(
    [0, 1, 2].map((i) => para(content.replace('#', String(i)), pPr, spacing)).join('') + SECTION,
  );

describe('a line of text (§17.3.1.33)', () => {
  it('stands on the line Word sets its family at, whatever face draws it', () => {
    // Word: 13.44 for 11pt Calibri, 12.90 for Cambria, 13.80 for 12pt Times
    // New Roman — none of them the 1.2× of the size, and the test's only face
    // is Roboto.
    expect(lineHeight(three(run('Calibri', 'line #')))).toBeCloseTo(CALIBRI_11, 3);
    expect(lineHeight(three(run('Cambria', 'line #')))).toBeCloseTo(line(11, 1946, 0, 455), 3);
    expect(lineHeight(three(run('Times New Roman', 'line #', 12)))).toBeCloseTo(
      line(12, 1825, 87, 443),
      3,
    );
  });

  it('stands a family Word’s line is not known for on the line of the face that draws it', () => {
    const roboto = parseTtf(FONTS.regular);
    const own = (11 * (roboto.ascender - roboto.descender + roboto.lineGap)) / roboto.unitsPerEm;
    expect(lineHeight(three(run('Nobody Sans', 'line #')))).toBeCloseTo(own, 3);
  });

  it('looks East Asian text up by the family the run gives for it, not the Latin one', () => {
    // §17.3.2.26: Word sets Han in the run's eastAsia font — here one whose
    // line is not known, so the face that draws it answers, not Calibri.
    const roboto = parseTtf(FONTS.regular);
    const own = (11 * (roboto.ascender - roboto.descender + roboto.lineGap)) / roboto.unitsPerEm;
    const han =
      '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Nobody Mincho"/>' +
      '<w:sz w:val="22"/></w:rPr><w:t>中文#</w:t></w:r>';
    expect(lineHeight(three(han))).toBeCloseTo(own, 3);
    expect(own).not.toBeCloseTo(CALIBRI_11, 1);
  });

  it('takes the tallest ascent and the deepest descent of the faces on it, each apart', () => {
    // Word: 14.36 for Verdana beside Courier New — Verdana's ascent over
    // Courier's descent, where either face alone stands 13.4.
    expect(lineHeight(three(run('Verdana', 'line #') + run('Courier New', ' cour')))).toBeCloseTo(
      line(11, 2059, 0, 615),
      3,
    );
  });

  it('sets a face’s gap above its text, where a taller ascent swallows it', () => {
    // Word: 14.11 for Candara beside Verdana — not 16.5, as it would stand
    // were Candara's 452-unit gap added on top of Verdana's ascent.
    expect(lineHeight(three(run('Candara', 'line #') + run('Verdana', ' verd')))).toBeCloseTo(
      line(11, 2059, 0, 564),
      3,
    );
  });

  it('grows by lines of itself at a spacing of more lines', () => {
    expect(lineHeight(three(run('Calibri', 'line #'), '', 276))).toBeCloseTo(1.15 * CALIBRI_11, 3);
    expect(lineHeight(three(run('Calibri', 'line #'), '', 480))).toBeCloseTo(2 * CALIBRI_11, 3);
  });
});

describe('a list marker on the line', () => {
  const bullet = (font: string, text: string, sizePt?: number): string =>
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    `<w:numFmt w:val="bullet"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
    `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>` +
    `${sizePt ? `<w:sz w:val="${sizePt * 2}"/>` : ''}</w:rPr></w:lvl></w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const list = (numbering: string): Uint8Array =>
    buildDocxFromBody(
      [0, 1, 2]
        .map((i) =>
          para(
            run('Calibri', `item ${String(i)}`),
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>',
          ),
        )
        .join('') + SECTION,
      { numberingXml: numbering },
    );

  it('raises it by the marker’s ascent, and adds nothing below', () => {
    // Word: 14.02 under a Symbol bullet over 11pt Calibri — Symbol's ascent
    // over Calibri's descent; 19.59 under a 20pt Courier New "o", whose own
    // descent (6pt) would make it 22.7.
    expect(lineHeight(list(bullet('Symbol', '')))).toBeCloseTo(line(11, 2059, 0, 550), 3);
    expect(lineHeight(list(bullet('Courier New', 'o', 20)))).toBeCloseTo(
      (20 * 1705 + 11 * 550) / 2048,
      3,
    );
  });

  it('leaves a line alone that its marker does not reach above', () => {
    // Word: 13.44 under an 11pt Courier New "o", Calibri's line to the point.
    expect(lineHeight(list(bullet('Courier New', 'o')))).toBeCloseTo(CALIBRI_11, 3);
  });
});

describe('a picture on a line of text', () => {
  const PNG = buildTinyPng(4, 4, [0, 120, 200, 255]);
  const picture = (sizePt: number): string => {
    const emu = Math.round(sizePt * 12700);
    return (
      `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu}" cy="${emu}"/>` +
      '<wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rId20"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu}" cy="${emu}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    );
  };
  const withPicture = (spacing: number): Uint8Array =>
    buildDocxFromBody(
      [0, 1, 2]
        .map((i) =>
          para(
            run('Calibri', `line ${String(i)} x `) + picture(30) + run('Calibri', ' y'),
            '',
            spacing,
          ),
        )
        .join('') + SECTION,
      { images: { rId20: { contentType: 'image/png', bytes: PNG, extension: 'png' } } },
    );
  const descent = (11 * 550) / 2048;

  it('stands it on the baseline, with the text’s descent under it', () => {
    // Word: 32.99 for a 30pt picture among 11pt Calibri.
    expect(lineHeight(withPicture(240))).toBeCloseTo(30 + descent, 3);
  });

  it('grows by lines of the text at a spacing of more lines, not of the picture', () => {
    // Word: 35.00 at 1.15 and 46.38 at double — 2.01 and 13.4 more.
    expect(lineHeight(withPicture(276))).toBeCloseTo(30 + descent + 0.15 * CALIBRI_11, 3);
    expect(lineHeight(withPicture(480))).toBeCloseTo(30 + descent + CALIBRI_11, 3);
  });
});

describe('a blank paragraph', () => {
  it('stands as tall as its mark, formatted by the document’s defaults and styles', () => {
    // A 12pt Calibri default and a 20pt style: the two blank paragraphs
    // between the lines stand at their own marks' lines.
    const docx = buildDocxFromBody(
      para(run('Calibri', 'before', 12)) +
        para('') +
        para('', '<w:pStyle w:val="Big"/>') +
        para(run('Calibri', 'after', 12)) +
        SECTION,
      {
        stylesXml:
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>' +
          '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
          '<w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/>' +
          '<w:rPr><w:sz w:val="40"/></w:rPr></w:style>',
      },
    );
    const text = lines(docx).filter((l) =>
      l.line.tokens.some((t) => t.kind === 'text' && !t.isSpace),
    );
    const [before, after] = text;
    // Baseline to baseline: the first line, then the two blank ones.
    const calibri = (size: number): number => line(size, 1950, 0, 550);
    expect(after!.baselineY - before!.baselineY).toBeCloseTo(2 * calibri(12) + calibri(20), 3);
  });
});

describe('the documents set as Word sets them', () => {
  it('are the ones Word writes: .docx and .doc', () => {
    expect(readDocx(buildDocxFromBody(para(run('Calibri', 'x')))).doc.typesetBy).toBe('word');
    expect(readDoc(buildDoc([{ text: 'x\r', compressed: false }])).doc.typesetBy).toBe('word');
  });
});
