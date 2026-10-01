// ECMA-376 Part 1 §17.10 — a header or footer band taller than the room its
// margin leaves moves the body out of its way, on each page by the band that
// page shows; and a picture on a paragraph of its own is as tall, and spaced,
// as that paragraph makes it (§17.3.1.33). Bug51170.docx's header logo was not
// drawn at all (its picture is offered in a markup-compatibility block under a
// `ve:` prefix), its body stood 10pt too high under it, and ran into a footer
// taller than its bottom margin.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import type { FlowDoc } from '@/core/ir/flow';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const PNG = buildTinyPng(4, 4, [0, 120, 200, 255]);
const PDF = new TextEncoder().encode('%PDF-1.4\n%%EOF\n');
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** An inline picture, `sizePt` square, of relationship `rId`. */
function inlinePicture(sizePt: number, blipFill: string): string {
  const emu = Math.round(sizePt * 12700);
  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu}" cy="${emu}"/>` +
    '<wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="P"/><pic:cNvPicPr/></pic:nvPicPr>${blipFill}` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu}" cy="${emu}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

const blip = (rId: string): string =>
  `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`;

/** A page 400pt tall with 72pt margins and its bands 36pt from its edges. */
const SECTION =
  '<w:pgSz w:w="8000" w:h="8000"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';

const para = (text: string): string =>
  `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

function layoutOf(docx: Uint8Array): ReturnType<typeof layoutStyledDocument> {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
}

/** The y (from the paper's top) of the first body line on each page. */
function bodyTops(docx: Uint8Array): Array<number> {
  return layoutOf(docx).pages.map((page) => {
    const line = page.commands.find(
      (c) =>
        c.type === 'line' &&
        c.line.tokens.some((t) => t.kind === 'text' && t.text.includes('body')),
    );
    return line?.type === 'line' ? line.baselineY : NaN;
  });
}

describe('a picture offered in a markup-compatibility block', () => {
  it('is taken from the branch a reader like us is to take, whatever the prefix', () => {
    // Word 2008 for Mac bound the namespace to `ve:` and offered a PDF for
    // itself (Requires="ma") and a PNG for everyone else.
    const fill =
      `<ve:AlternateContent xmlns:ve="${MC}">` +
      `<ve:Choice Requires="ma">${blip('rId1')}</ve:Choice>` +
      `<ve:Fallback>${blip('rId2')}</ve:Fallback></ve:AlternateContent>`;
    const docx = buildDocxFromBody(
      `${para('body')}<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>${SECTION}</w:sectPr>`,
      {
        headerXml: `<w:p><w:r>${inlinePicture(50, fill)}</w:r></w:p>`,
        headerImages: {
          rId1: { contentType: 'application/pdf', bytes: PDF, extension: 'pdf' },
          rId2: { contentType: 'image/png', bytes: PNG, extension: 'png' },
        },
      },
    );
    const flow: FlowDoc = Ream.parse(docx).flow;
    const [picture] = flow.headersFooters?.get('rId10') ?? [];
    expect(picture?.kind).toBe('image');
    const resource = picture?.kind === 'image' ? picture.image.resource : undefined;
    expect(resource).toBeDefined();
    expect([...(flow.resources.get(resource!)?.subarray(1, 4) ?? [])]).toEqual([0x50, 0x4e, 0x47]); // "PNG"
  });
});

describe('a picture on a paragraph of its own (§17.3.1.33)', () => {
  const pictureThen = (line: number, style = ''): Uint8Array =>
    buildDocxFromBody(
      `<w:p><w:pPr>${style}<w:spacing w:line="${line}" w:lineRule="auto"/></w:pPr>` +
        `<w:r>${inlinePicture(100, blip('rId20'))}</w:r></w:p>${para('body')}` +
        `<w:sectPr>${SECTION}</w:sectPr>`,
      {
        images: { rId20: { contentType: 'image/png', bytes: PNG, extension: 'png' } },
        stylesXml:
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>' +
          '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
          '<w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/>' +
          '<w:pPr><w:spacing w:after="400"/></w:pPr></w:style>',
      },
    );

  it('is spaced by its paragraph’s style', () => {
    const plain = bodyTops(pictureThen(240))[0]!;
    const spaced = bodyTops(pictureThen(240, '<w:pStyle w:val="Spaced"/>'))[0]!;
    expect(spaced - plain).toBeCloseTo(20, 1);
  });

  it('gains what a spacing of more lines adds to the mark’s font, not a share of itself', () => {
    // Word for Mac: 100, 102.5 and 113 points at 1, 1.15 and 2 lines of an
    // 11pt Calibri mark — the line of the mark's font (13.43pt, see
    // word-line-heights.test.ts), not a picture 15% or twice as tall.
    const line = (11 * 2500) / 2048;
    const single = bodyTops(pictureThen(240))[0]!;
    expect(bodyTops(pictureThen(276))[0]! - single).toBeCloseTo(0.15 * line, 1);
    expect(bodyTops(pictureThen(480))[0]! - single).toBeCloseTo(line, 1);
  });
});

describe('a band taller than its margin (§17.10)', () => {
  const lines = (n: number): string =>
    Array.from({ length: n }, (_, i) => para(`body ${String(i)}`)).join('');

  it('moves the body down under its header, on the pages that show it', () => {
    // A 100pt picture 36pt from the top, 10pt after it: the body starts 146pt
    // down, not at the 72pt margin — but the title page shows no header.
    const docx = buildDocxFromBody(
      `${lines(60)}<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>` +
        `<w:titlePg/>${SECTION}</w:sectPr>`,
      {
        headerXml:
          '<w:p><w:pPr><w:spacing w:after="200"/></w:pPr>' +
          `<w:r>${inlinePicture(100, blip('rId30'))}</w:r></w:p>`,
        headerImages: { rId30: { contentType: 'image/png', bytes: PNG, extension: 'png' } },
      },
    );
    const [first, second] = bodyTops(docx);
    expect(first!).toBeLessThan(72 + 12 + 1);
    expect(second!).toBeGreaterThan(146);
    expect(second!).toBeLessThan(146 + 12 + 1);
  });

  it('lifts the body above a footer taller than the bottom margin', () => {
    // Word: 10 lines to a 200pt page with 36pt margins, 8 under a four-line
    // footer 18pt up from the edge (66pt), 10 under a one-line one (30pt).
    const page =
      '<w:pgSz w:w="8400" w:h="4000"/>' +
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/>';
    const perPage = (footerLines: number): Array<number> => {
      const docx = buildDocxFromBody(
        `${lines(40)}<w:sectPr><w:footerReference w:type="default" r:id="rId11"/>${page}</w:sectPr>`,
        {
          footerXml: Array.from({ length: footerLines }, (_, i) => para(`foot ${String(i)}`)).join(
            '',
          ),
        },
      );
      return layoutOf(docx).pages.map(
        (p) =>
          p.commands.filter(
            (c) =>
              c.type === 'line' &&
              c.line.tokens.some((t) => t.kind === 'text' && t.text.includes('body')),
          ).length,
      );
    };
    expect(perPage(1)).toEqual([10, 10, 10, 10]);
    expect(perPage(4)).toEqual([8, 8, 8, 8, 8]);
  });
});
