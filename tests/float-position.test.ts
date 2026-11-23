// ECMA-376 Part 1 §20.4.2.10/§20.4.2.11 — where an anchored drawing stands on
// its page, measured against Word (Word for Mac 16, PDF export of synthetic
// probes: every `relativeFrom` × every alignment and an offset, in a header, a
// footer and the body, on an odd and an even page). Word places a header's or
// a footer's drawing exactly as it places the body's, and every expected
// number below is one Word printed.
//
// #143: a first-page header picture as wide as the page, aligned right
// against the PAGE, was aligned against the text area instead — a white strip
// as wide as the right margin down the side of the sheet, and the picture's
// left side cut by as much.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import type { FixtureImage } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const EMU = 12700;
const PNG: FixtureImage = {
  contentType: 'image/png',
  bytes: buildTinyPng(4, 4, [48, 112, 208, 255]),
  extension: 'png',
};

/** One `wp:positionH`/`wp:positionV`: a keyword, or an offset in points. */
const pos = (axis: 'H' | 'V', rel: string, at: string | number): string =>
  `<wp:position${axis} relativeFrom="${rel}">` +
  (typeof at === 'number'
    ? `<wp:posOffset>${String(Math.round(at * EMU))}</wp:posOffset>`
    : `<wp:align>${at}</wp:align>`) +
  `</wp:position${axis}>`;

let docPrId = 1;
/** A floating picture `wPt` × `hPt`, wrapped in no text. */
function picture(wPt: number, hPt: number, posH: string, posV: string, behind = false): string {
  const id = docPrId++;
  const cx = Math.round(wPt * EMU);
  const cy = Math.round(hPt * EMU);
  return (
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${String(251658240 + id)}" ` +
    `behindDoc="${behind ? '1' : '0'}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${String(cx)}" cy="${String(cy)}"/><wp:wrapNone/>` +
    `<wp:docPr id="${String(id)}" name="P${String(id)}"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    `<pic:nvPicPr><pic:cNvPr id="${String(id)}" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rIdP"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>' +
    '</wp:anchor></w:drawing></w:r>'
  );
}

/** Paragraphs of exactly 12pt lines, so every renderer agrees where they are. */
const P = (inner: string): string =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${inner}</w:p>`;
const T = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BR = '<w:r><w:br/></w:r>';
const PAGE_BREAK = P('<w:r><w:br w:type="page"/></w:r>');

/**
 * #143's page: A4, its margins 72pt left, 49.25pt right, 57.6pt top, 43.15pt
 * bottom; the header 14.2pt from the top edge and the footer 28.35pt up from
 * the bottom one.
 */
const pgMar =
  '<w:pgSz w:w="11900" w:h="16840"/>' +
  '<w:pgMar w:top="1152" w:right="985" w:bottom="863" w:left="1440" w:header="284" w:footer="567" w:gutter="0"/>';

type Where = 'header' | 'footer' | 'body';

/** A two-page document whose header, footer or body (on each page) holds `pictures`. */
function twoPages(where: Where, pictures: string, band = P(T('band'))): Uint8Array {
  const refs =
    (where === 'header' ? '<w:headerReference w:type="default" r:id="rId10"/>' : '') +
    (where === 'footer' ? '<w:footerReference w:type="default" r:id="rId11"/>' : '');
  const content = where === 'body' ? P(pictures) : P(T('one'));
  const body = `${content}${PAGE_BREAK}${where === 'body' ? P(pictures) : P(T('two'))}`;
  const bandXml = band.replace('</w:p>', `${pictures}</w:p>`);
  return buildDocxFromBody(`${body}<w:sectPr>${refs}${pgMar}</w:sectPr>`, {
    ...(where === 'body' ? { images: { rIdP: PNG } } : {}),
    ...(where === 'header' ? { headerXml: bandXml, headerImages: { rIdP: PNG } } : {}),
    ...(where === 'footer' ? { footerXml: bandXml, footerImages: { rIdP: PNG } } : {}),
  });
}

/** Each picture on each page, by its width: its left edge and its top (from the page's top). */
function placed(docx: Uint8Array): Array<Map<number, { x: number; y: number }>> {
  const flow = Ream.parse(docx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.map((page) => {
    const out = new Map<number, { x: number; y: number }>();
    for (const c of page.commands) {
      if (c.type === 'image') out.set(Math.round(c.width * 100) / 100, { x: c.x, y: c.y });
    }
    return out;
  });
}

describe('#143 — a header picture aligned against the PAGE', () => {
  // The picture as reported: 7588250 × 11242149 EMU, right-aligned to the page,
  // 14.2pt above its paragraph — which is the header's, 14.2pt down the page.
  const issuePicture = picture(
    7588250 / EMU,
    11242149 / EMU,
    pos('H', 'page', 'right'),
    pos('V', 'paragraph', -180340 / EMU),
    true,
  );

  it('ends at the page’s right edge, not at the text’s (first-page header)', () => {
    const docx = buildDocxFromBody(
      `${P(T('Body text on the first page.'))}<w:sectPr>` +
        '<w:headerReference w:type="first" r:id="rId12"/>' +
        `${pgMar}<w:titlePg/></w:sectPr>`,
      { firstHeaderXml: P(issuePicture), firstHeaderImages: { rIdP: PNG } },
    );
    const [page] = placed(docx);
    const pic = page!.get(597.5);
    // Word: the picture's left edge at −2.5pt (595 − 597.5) and its top on the
    // paper's own top edge. Aligned to the text, it stood at −51.75pt.
    expect(pic?.x).toBeCloseTo(-2.5, 1);
    expect(pic?.y).toBeCloseTo(0, 1);
  });

  it('does the same in the default header', () => {
    const [page] = placed(twoPages('header', issuePicture, P('')));
    expect(page!.get(597.5)?.x).toBeCloseTo(-2.5, 1);
  });
});

// Word's left edges for each `relativeFrom` × alignment, odd page then even
// page. A picture's width is 30pt + its place in this table (row × 6 +
// column), as in the probe that measured it; `null` is one Word does not
// print at all (it hangs half off the page) and is not asserted.
const ALIGNS = ['left', 'center', 'right', 'inside', 'outside', 36] as const;
const H_MATRIX: ReadonlyArray<
  readonly [
    rel: string,
    row: number,
    odd: ReadonlyArray<number | null>,
    even: ReadonlyArray<number | null>,
  ]
> = [
  ['page', 0, [0, 282, 563, 0, 561, 36], [0, 282, 563, 562, 0, 36]],
  ['margin', 1, [72, 290.375, 507.75, 72, 505.75, 108], [72, 290.375, 507.75, 506.75, 72, 108]],
  ['column', 2, [72, 287.375, 501.75, 72, 499.75, 108], [72, 287.375, 501.75, 500.75, 72, 108]],
  ['leftMargin', 4, [0, 8.5, 16, 0, 14, 36], [0, 8.5, 16, 15, 0, 36]],
  [
    'rightMargin',
    5,
    [545.75, 539.875, 533, 545.75, 531, null],
    [545.75, 539.875, 533, 532, 545.75, null],
  ],
  ['insideMargin', 6, [0, 2.5, 4, 0, 2, 36], [545.75, 536.875, 527, 526, 545.75, null]],
  ['outsideMargin', 7, [545.75, 533.875, 521, 545.75, 519, null], [0, -0.5, -2, -3, 0, 36]],
];

describe('across the page (§20.4.2.10), as Word places it', () => {
  const matrix = H_MATRIX.flatMap(([rel, row]) =>
    ALIGNS.map((align, a) => {
      const k = row * 6 + a;
      return picture(30 + k, 10, pos('H', rel, align), pos('V', 'page', 60 + k * 15));
    }),
  ).join('');

  for (const where of ['header', 'footer', 'body'] as const) {
    it(`puts every anchor in the ${where} where Word does, on odd and even pages`, () => {
      const pages = placed(twoPages(where, matrix));
      for (const [rel, row, odd, even] of H_MATRIX) {
        ALIGNS.forEach((align, a) => {
          const k = row * 6 + a;
          [odd, even].forEach((expected, p) => {
            const want = expected[a];
            if (want === null || want === undefined) return;
            const got = pages[p]!.get(30 + k);
            expect(got?.x, `${where} p${String(p + 1)} ${rel} ${String(align)}`).toBeCloseTo(
              want,
              1,
            );
            expect(got?.y).toBeCloseTo(60 + k * 15, 1);
          });
        });
      }
    });
  }

  it('reads `character` as a box of no width where the anchor stands', () => {
    // Word: centred on the anchor, or ending at it — an anchor at the start of
    // an unindented paragraph stands at the left margin.
    const pics =
      picture(40, 10, pos('H', 'character', 'left'), pos('V', 'page', 100)) +
      picture(41, 10, pos('H', 'character', 'center'), pos('V', 'page', 120)) +
      picture(42, 10, pos('H', 'character', 'right'), pos('V', 'page', 140)) +
      picture(43, 10, pos('H', 'character', 36), pos('V', 'page', 160));
    const [page] = placed(twoPages('body', pics));
    expect(page!.get(40)?.x).toBeCloseTo(72, 1);
    expect(page!.get(41)?.x).toBeCloseTo(72 - 20.5, 1);
    expect(page!.get(42)?.x).toBeCloseTo(72 - 42, 1);
    expect(page!.get(43)?.x).toBeCloseTo(108, 1);
  });
});

describe('down the page (§20.4.2.11), as Word places it', () => {
  /** Pictures 10pt tall, 20pt + n wide, each anchored down the page as `specs` say. */
  const column = (specs: ReadonlyArray<readonly [rel: string, at: string | number]>): string =>
    specs
      .map(([rel, at], n) => picture(20 + n, 10, pos('H', 'page', 8 + n * 25), pos('V', rel, at)))
      .join('');
  const tops = (pages: ReturnType<typeof placed>, count: number): Array<Array<number>> =>
    pages.map((page) => Array.from({ length: count }, (_, n) => page.get(20 + n)?.y ?? NaN));
  // Two 12pt lines: the anchor hangs off the band's (or the body's) paragraph.
  const twoLines = P(`${T('BAND1')}${BR}${T('BAND2')}`);

  const SPECS = [
    ['page', 'top'],
    ['page', 'center'],
    ['page', 'bottom'],
    ['page', 'inside'],
    ['page', 'outside'],
    ['margin', 'top'],
    ['margin', 'center'],
    ['margin', 'bottom'],
    ['margin', 'inside'],
    ['paragraph', 'bottom'],
    ['paragraph', 'center'],
    ['paragraph', 36],
    ['topMargin', 'center'],
    ['topMargin', 'bottom'],
    ['bottomMargin', 'top'],
    ['bottomMargin', 'center'],
    ['insideMargin', 'top'],
    ['insideMargin', 'bottom'],
    ['outsideMargin', 'top'],
    ['outsideMargin', 'inside'],
  ] as const;

  it('in a header: the page, the text area, the paragraph and the margin bands', () => {
    const [odd, even] = tops(placed(twoPages('header', column(SPECS), twoLines)), SPECS.length);
    // Word, odd page: `inside`/`outside` of the page stop 18pt short of its
    // edges in a band; the header's paragraph is 14.2pt down, and an
    // alignment against it changes nothing.
    expect(odd).toEqual(
      [
        0, 416, 832, 18, 814, 57.6, 423.225, 788.85, 57.6, 14.2, 14.2, 50.2, 23.8, 47.6, 798.85,
        815.425, 0, 47.6, 798.85, 798.85,
      ].map((v) => expect.closeTo(v, 1)),
    );
    // …and on an even page `inside` is the bottom, `insideMargin` the bottom
    // band and `outsideMargin` the top one.
    expect(even).toEqual(
      [
        0, 416, 832, 814, 18, 57.6, 423.225, 788.85, 788.85, 14.2, 14.2, 50.2, 23.8, 47.6, 798.85,
        815.425, 798.85, 832, 0, 47.6,
      ].map((v) => expect.closeTo(v, 1)),
    );
  });

  it('in a footer taller than its margin: the text area ends where the footer begins', () => {
    // The footer is two 12pt lines 28.35pt up from the edge, so it begins
    // 52.35pt up the page — above the 43.15pt margin — and Word's `margin`
    // and `bottomMargin` boxes both turn at its top (789.65pt down).
    const [odd] = tops(placed(twoPages('footer', column(SPECS), twoLines)), SPECS.length);
    expect(odd).toEqual(
      [
        0, 416, 832, 18, 814, 57.6, 418.625, 779.65, 57.6, 789.65, 789.65, 825.65, 23.8, 47.6,
        789.65, 810.825, 0, 47.6, 789.65, 789.65,
      ].map((v) => expect.closeTo(v, 1)),
    );
  });

  it('in the body: inside and outside the page stop halfway to the header and footer', () => {
    // Word: 7.1pt down (half the header's 14.2pt) and 14.175pt up (half the
    // footer's 28.35pt) — in a band it is 18pt, whatever the distances.
    const specs = [
      ['page', 'inside'],
      ['page', 'outside'],
      ['line', 'top'],
    ] as const;
    const [odd, even] = tops(placed(twoPages('body', column(specs))), specs.length);
    expect(odd!.slice(0, 2)).toEqual([7.1, 817.825].map((v) => expect.closeTo(v, 1)));
    expect(even!.slice(0, 2)).toEqual([817.825, 7.1].map((v) => expect.closeTo(v, 1)));
    // A `line` box hangs from the anchoring paragraph, here the first on its page.
    expect(odd![2]).toBeCloseTo(57.6, 1);
  });
});

describe('a frame (§17.3.1.11 w:framePr) keyed into the page', () => {
  /** A 100 × 50pt frame holding `label`, after ten lines of text. */
  const frameDoc = (frames: ReadonlyArray<readonly [label: string, attrs: string]>): Uint8Array =>
    buildDocxFromBody(
      Array.from({ length: 10 }, (_, i) => P(T(`filler ${String(i)}`))).join('') +
        frames
          .map(
            ([label, attrs]) =>
              `<w:p><w:pPr><w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around" w:hAnchor="page" ${attrs}/>` +
              `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${T(label)}</w:p>`,
          )
          .join('') +
        `<w:sectPr>${pgMar}</w:sectPr>`,
    );
  const baselines = (docx: Uint8Array): Map<string, number> => {
    const flow = Ream.parse(docx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const out = new Map<string, number>();
    for (const c of laid.pages[0]!.commands) {
      if (c.type !== 'line') continue;
      const text = c.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('');
      out.set(text.trim(), c.baselineY);
    }
    return out;
  };

  it('hangs off the MARGIN when it names no vertical anchor', () => {
    // Word: `w:yAlign="bottom"` with no `w:vAnchor` ends the frame on the
    // margin's foot, as `w:vAnchor="margin"` does; `w:y="1440"` stands 72pt
    // below the margin's top, not 72pt below its paragraph.
    const lines = baselines(
      frameDoc([
        ['NOANCHOR', 'w:x="1440" w:yAlign="bottom"'],
        ['MARGINANCHOR', 'w:x="5040" w:vAnchor="margin" w:yAlign="bottom"'],
        ['TOP', 'w:x="8640" w:yAlign="top"'],
      ]),
    );
    const lines2 = baselines(frameDoc([['OFFSET', 'w:x="1440" w:y="1440"']]));
    expect(lines.get('NOANCHOR')).toBeCloseTo(lines.get('MARGINANCHOR')!, 1);
    // A 50pt frame on the margin's foot (798.85pt down) begins 748.85pt down.
    expect(lines.get('NOANCHOR')! - lines.get('TOP')!).toBeCloseTo(748.85 - 57.6, 1);
    expect(lines2.get('OFFSET')! - lines.get('TOP')!).toBeCloseTo(72, 1);
  });

  it('ignores an alignment against its own paragraph, as a drawing does', () => {
    // Word: `w:vAnchor="text" w:yAlign="bottom"` stands at the paragraph.
    const lines = baselines(
      frameDoc([['TEXTANCHOR', 'w:x="1440" w:vAnchor="text" w:yAlign="bottom"']]),
    );
    const filler = lines.get('filler 9')!;
    expect(lines.get('TEXTANCHOR')! - filler).toBeCloseTo(12, 0);
  });

  it('turns `inside` and `outside` with the page', () => {
    const doc = (attrs: string): Uint8Array =>
      buildDocxFromBody(
        P(T('one')) +
          `<w:p><w:pPr><w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around" ${attrs}/></w:pPr>${T('F1')}</w:p>` +
          PAGE_BREAK +
          `<w:p><w:pPr><w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around" ${attrs}/></w:pPr>${T('F2')}</w:p>` +
          `<w:sectPr>${pgMar}</w:sectPr>`,
      );
    const flow = Ream.parse(
      doc('w:hAnchor="margin" w:xAlign="inside" w:vAnchor="margin" w:y="0"'),
    ).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const x = (page: number, label: string): number | undefined => {
      for (const c of laid.pages[page]!.commands) {
        if (c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === label))
          return c.originX;
      }
      return undefined;
    };
    // Word: on the margin's left edge on page 1, flush with its right edge
    // (545.75 − 100pt) on page 2.
    expect(x(0, 'F1')).toBeCloseTo(72, 0);
    expect(x(1, 'F2')).toBeCloseTo(445.75, 0);
  });
});

describe('reading the anchor (§20.4.3.1–§20.4.3.5)', () => {
  it('keeps the bases and keywords that turn with the page', () => {
    const docx = buildDocxFromBody(
      P(
        picture(10, 10, pos('H', 'insideMargin', 'outside'), pos('V', 'outsideMargin', 'inside')) +
          picture(11, 10, pos('H', 'character', 12), pos('V', 'line', 'center')),
      ) + `<w:sectPr>${pgMar}</w:sectPr>`,
      { images: { rIdP: PNG } },
    );
    const floats = readDocx(docx)
      .doc.body.map((el) => (el.kind === 'image' ? el.image.float : undefined))
      .filter((f) => f !== undefined);
    expect(floats[0]?.posH).toEqual({ relativeFrom: 'insideMargin', align: 'outside' });
    expect(floats[0]?.posV).toEqual({ relativeFrom: 'outsideMargin', align: 'inside' });
    expect(floats[1]?.posH).toEqual({ relativeFrom: 'character', offsetPt: 12 });
    expect(floats[1]?.posV).toEqual({ relativeFrom: 'line', align: 'center' });
  });
});
