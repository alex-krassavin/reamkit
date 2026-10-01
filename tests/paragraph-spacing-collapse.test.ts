// ECMA-376 Part 1 §17.3.1.33 — in a document Word sets, two paragraphs stand
// the larger of the first's space after and the second's space before apart,
// not the two together. Measured in Word for Mac (modes 14 and 15 alike) by
// the positions of 11pt Calibri paragraphs, averaged over ten pairs.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import type { PageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** 11pt Calibri's single line, as Word sets it. */
const LINE = (11 * 2500) / 2048;

const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>';

/** A paragraph with the given space before and after, in twips. */
function para(text: string, before: number, after: number, pPr = ''): string {
  return (
    `<w:p><w:pPr>${pPr}<w:spacing w:before="${String(before)}" w:after="${String(after)}" ` +
    `w:line="240" w:lineRule="auto"/>${RPR}</w:pPr>` +
    (text ? `<w:r>${RPR}<w:t>${text}</w:t></w:r>` : '') +
    '</w:p>'
  );
}

const SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"/></w:sectPr>';

function pages(docx: Uint8Array, asWord = true): ReadonlyArray<ReadonlyArray<PageItem>> {
  const flow = Ream.parse(docx).flow;
  const { typesetBy: _word, ...plain } = flowRenderOptions(flow);
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...(asWord ? flowRenderOptions(flow) : plain),
  });
  return laid.pages.map((p) => p.commands);
}

function baseline(items: ReadonlyArray<PageItem>, text: string): number {
  const hit = items.find(
    (c): c is TextLineItem =>
      c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === text),
  );
  if (!hit) throw new Error(`no line "${text}"`);
  return hit.baselineY;
}

/** The space between paragraph A (`after`) and B (`before`), past their lines. */
function gap(after: number, before: number, asWord = true): number {
  const [page] = pages(
    buildDocxFromBody(para('A', 0, after) + para('B', before, 0) + SECTION),
    asWord,
  );
  return baseline(page!, 'B') - baseline(page!, 'A') - LINE;
}

describe('two paragraphs in a document Word sets (§17.3.1.33)', () => {
  it('stand the larger of the space after and the space before apart', () => {
    // Word: 9.98, 24.03, 24.03, 12.03, 12.03.
    expect(gap(200, 200)).toBeCloseTo(10, 3);
    expect(gap(200, 480)).toBeCloseTo(24, 3);
    expect(gap(480, 200)).toBeCloseTo(24, 3);
    expect(gap(120, 240)).toBeCloseTo(12, 3);
    expect(gap(240, 120)).toBeCloseTo(12, 3);
  });

  it('share each side with an empty paragraph between them', () => {
    // Word: 34.05 — 10 above the empty one and 24 below it.
    const [page] = pages(
      buildDocxFromBody(para('A', 0, 200) + para('', 0, 200) + para('B', 480, 0) + SECTION),
    );
    expect(baseline(page!, 'B') - baseline(page!, 'A') - 2 * LINE).toBeCloseTo(34, 3);
  });

  it('add up where the document is not Word’s', () => {
    expect(gap(200, 200, false) - (11 * 1.2 - LINE)).toBeCloseTo(20, 3);
  });

  it('stack the same way in a table cell and in a header', () => {
    const cell =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>' +
      `<w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${para('A', 0, 200)}${para('B', 200, 0)}` +
      '</w:tc></w:tr></w:tbl>';
    const [inCell] = pages(buildDocxFromBody(cell + SECTION));
    expect(baseline(inCell!, 'B') - baseline(inCell!, 'A') - LINE).toBeCloseTo(10, 3);

    const docx = buildDocxFromBody(
      para('body', 0, 0) +
        '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>' +
        '<w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"/></w:sectPr>',
      { headerXml: para('A', 0, 200) + para('B', 200, 0) },
    );
    const [withHeader] = pages(docx);
    expect(baseline(withHeader!, 'B') - baseline(withHeader!, 'A') - LINE).toBeCloseTo(10, 3);
  });
});

describe('a paragraph at the top of a page, in a document Word sets', () => {
  const ASCENT = (11 * 1950) / 2048;
  const WORD_2013 =
    '<w:compat><w:compatSetting w:name="compatibilityMode" ' +
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';

  /** Where "B" stands down its page, in a document of the given body. */
  function topOfB(body: string, settingsXml?: string): number {
    const laid = pages(
      buildDocxFromBody(body + SECTION, settingsXml === undefined ? {} : { settingsXml }),
    );
    const page = laid.find((items) =>
      items.some(
        (c) => c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === 'B'),
      ),
    );
    return baseline(page!, 'B') - ASCENT;
  }

  it('loses its space before where it ran over onto the page', () => {
    // Word: 71.15 after 51 and after 50 lines of a 52-line page, in every mode.
    const fill = Array.from({ length: 51 }, (_, i) => para(`fill ${String(i)}`, 0, 0)).join('');
    expect(topOfB(fill + para('B', 480, 0))).toBeCloseTo(72, 3);
  });

  it('loses it after a page break, and keeps it as the first paragraph of the document', () => {
    // Word: 71.15 after a w:br w:type="page", 95.45 as the very first paragraph.
    const brokenA = `<w:p><w:r>${RPR}<w:t>A</w:t><w:br w:type="page"/></w:r></w:p>`;
    expect(topOfB(brokenA + para('B', 480, 0))).toBeCloseTo(72, 3);
    expect(topOfB(para('B', 480, 0))).toBeCloseTo(72 + 24, 3);
  });

  it('keeps it where its own page break brings it there, but not in Word 2013’s mode', () => {
    // Word: 95.45 in modes 14 and none, 71.15 in mode 15.
    const body = para('A', 0, 200) + para('B', 480, 0, '<w:pageBreakBefore/>');
    expect(topOfB(body)).toBeCloseTo(72 + 24, 3);
    expect(topOfB(body, WORD_2013)).toBeCloseTo(72, 3);
  });

  it('keeps it as the first paragraph with a drawing anchored off it, which opens nothing', () => {
    // tdf60351.docx hangs its cover picture off its first paragraph.
    const anchored =
      '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ' +
      'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="635000" cy="635000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="P"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      '<pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rId20"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="635000" cy="635000"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>' +
      '</wp:anchor></w:drawing></w:r>';
    const withDrawing = para('B', 480, 0).replace('</w:p>', `${anchored}</w:p>`);
    const png = {
      contentType: 'image/png',
      bytes: buildTinyPng(4, 4, [0, 0, 0, 255]),
      extension: 'png',
    };
    const laid = pages(buildDocxFromBody(withDrawing + SECTION, { images: { rId20: png } }));
    expect(laid[0]!.some((c) => c.type === 'image')).toBe(true);
    expect(baseline(laid[0]!, 'B') - (11 * 1950) / 2048).toBeCloseTo(72 + 24, 3);
  });

  it('keeps it as the first paragraph of a section on a page of its own', () => {
    // Word: 95.45.
    const sectionBreak =
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>';
    expect(topOfB(para('A', 0, 0) + sectionBreak + para('B', 480, 0), WORD_2013)).toBeCloseTo(
      72 + 24,
      3,
    );
  });
});
