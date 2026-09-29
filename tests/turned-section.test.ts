// §17.6.20 `w:textDirection` — a section whose lines run DOWN the sheet.
//
// Word turns the TEXT of such a section a quarter clockwise and nothing else:
// the header stays across the top of the sheet, and a drawing anchored to the
// page stands where its offsets put it, upright. Laid out flat, a page a viewer
// shows turned came back with its words lying across the sheet.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import type { ImageItem, TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { convertDocxToPdfSync } from '@/core/converter';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** US Letter turned landscape, an inch of margin all round, its lines running down it. */
const TURNED =
  '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>' +
  '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>' +
  '<w:textDirection w:val="tbRl"/></w:sectPr>';

const SHEET_WIDTH = 792;

const para = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** A 40×20pt picture anchored to the page at (100pt, 50pt). */
const anchoredPicture = (): string =>
  `<w:p><w:r><w:drawing>
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
               distT="0" distB="0" distL="0" distR="0" behindDoc="0" locked="0"
               layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>635000</wp:posOffset></wp:positionV>
      <wp:extent cx="508000" cy="254000"/>
      <wp:wrapNone/>
      <wp:docPr id="1" name="Picture 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="508000" cy="254000"/></a:xfrm></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing></w:r></w:p>`;

const docx = (body: string): Uint8Array =>
  buildDocxFromBody(body + TURNED, {
    headerXml: para('HEADER'),
    images: {
      rId20: {
        contentType: 'image/png',
        bytes: buildTinyPng(2, 2, [255, 0, 0, 255]),
        extension: 'png',
      },
    },
  });

function firstPage(body: string) {
  const flow = Ream.parse(docx(body)).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  }).pages[0]!;
}

const textOf = (item: TextLineItem): string =>
  item.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('');

const lineWith = (page: ReturnType<typeof firstPage>, text: string): TextLineItem => {
  const found = page.commands.find(
    (c): c is TextLineItem => c.type === 'line' && textOf(c).includes(text),
  );
  if (!found) throw new Error(`no line reads "${text}"`);
  return found;
};

describe('a section whose lines run down the sheet (§17.6.20)', () => {
  it('prints on the sheet it states, its first line down the right margin', () => {
    const page = firstPage(para('First line') + para('Second line'));
    expect([page.width, page.height]).toEqual([792, 612]);
    const first = lineWith(page, 'First line');
    const second = lineWith(page, 'Second line');
    // A quarter clockwise: the words run down the sheet from its top margin…
    expect(first.rotationDeg).toBe(-90);
    expect(first.baselineY).toBeCloseTo(72, 5);
    // …the first line stands inside the right margin, its glyphs' tops
    // toward the edge…
    expect(first.originX).toBeLessThan(SHEET_WIDTH - 72);
    expect(first.originX).toBeGreaterThan(SHEET_WIDTH - 72 - 24);
    // …and the next line stands to the LEFT of it, starting where it starts.
    expect(second.originX).toBeLessThan(first.originX);
    expect(second.baselineY).toBeCloseTo(first.baselineY, 5);
  });

  it('keeps the header across the top of the sheet', () => {
    const header = lineWith(firstPage(para('Body')), 'HEADER');
    expect(header.rotationDeg).toBeUndefined();
    expect(header.originX).toBeCloseTo(72, 5);
    expect(header.baselineY).toBeGreaterThan(36);
    expect(header.baselineY).toBeLessThan(72);
  });

  it('stands a picture anchored to the page where its offsets put it, upright', () => {
    const page = firstPage(anchoredPicture() + para('Body'));
    const picture = page.commands.find((c): c is ImageItem => c.type === 'image');
    expect(picture).toBeDefined();
    expect(picture!.x).toBeCloseTo(100, 3);
    expect(picture!.y).toBeCloseTo(50, 3);
    expect([picture!.width, picture!.height]).toEqual([40, 20]);
    expect(picture!.rotationDeg).toBeUndefined();
  });

  it('writes the turn into the PDF, and a tabbed line keeps its tab turned', () => {
    // A line one text matrix can set is turned by that matrix; a line with a
    // tab in it is set flat inside a turned CTM, so the tab's distance turns
    // with it rather than being lost to the font's own advances.
    const pdf = new TextDecoder('latin1').decode(
      convertDocxToPdfSync(
        docx(
          para('Plain') +
            '<w:p><w:r><w:t>Left</w:t></w:r><w:r><w:tab/><w:t>Right</w:t></w:r></w:p>',
        ),
        { fonts: FONTS },
      ),
    );
    expect(pdf).toMatch(/\n0 -1 1 0 [\d.]+ [\d.]+ Tm\n/u);
    expect(pdf).toMatch(/\n0 -1 1 0 [\d.-]+ [\d.-]+ cm\n/u);
  });
});
