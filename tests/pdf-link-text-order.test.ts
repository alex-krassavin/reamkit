// A link annotation's /Contents is the link's alternate description (ISO 14289-1
// §7.18.5), the text a screen reader speaks for it, so it has to read in the
// order the text was written in, not the order it is drawn in. The emitter
// walks a right-to-left line in visual order to place its glyphs, and it used
// to gather the link's text on that same walk: a bidi paragraph's link
// "العنوان الرئيسي" reached macOS PDFKit (probe, 2026-11-20) as
// "الرئيسي العنوان", its words swapped. Every document here is synthetic;
// Roboto has no Arabic, which /Contents does not need.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PdfDict, PdfName, PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { PdfHexString } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { pdfTextString } from '@/pdf-reader/encodings';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const URL = 'https://reamkit.dev/ar';
const URL2 = 'https://reamkit.dev/ar/2';

/** A run of text, right-to-left when `rtl` is set (§17.3.2.30). */
const run = (text: string, rtl = false): string =>
  `<w:r>${rtl ? '<w:rPr><w:rtl/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A hyperlink to {@link URL} around the given runs. */
const link = (...runs: ReadonlyArray<string>): string =>
  `<w:hyperlink r:id="rIdL">${runs.join('')}</w:hyperlink>`;

/** A hyperlink to {@link URL2}. */
const link2 = (...runs: ReadonlyArray<string>): string =>
  `<w:hyperlink r:id="rIdM">${runs.join('')}</w:hyperlink>`;

/** A paragraph, right-to-left when `bidi` is set (§17.3.1.6). */
const para = (bidi: boolean, ...content: ReadonlyArray<string>): string =>
  `<w:p>${bidi ? '<w:pPr><w:bidi/></w:pPr>' : ''}${content.join('')}</w:p>`;

/** A page `widthPt` wide with half-inch margins, so a long link wraps. */
const narrow = (widthPt: number): string =>
  `<w:sectPr><w:pgSz w:w="${String(widthPt * 20)}" w:h="16838"/>` +
  `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;

interface LinkAnnot {
  readonly uri: string;
  readonly contents: string;
  readonly rect: readonly [number, number, number, number];
}

/** A string value as a reader decodes it: the bytes as written, then §7.9.2.2. */
function decoded(raw: PdfValue | undefined): string {
  if (raw instanceof PdfHexString) return pdfTextString(raw.bytes);
  if (typeof raw === 'string') {
    // The reader hands a literal back as one char per byte.
    return pdfTextString(Uint8Array.from(raw, (c) => c.charCodeAt(0) & 0xff));
  }
  throw new Error(`not a string: ${String(raw)}`);
}

/** Every Link annotation on the first page, in the order the page lists them. */
async function linksOf(body: string): Promise<Array<LinkAnnot>> {
  const docx = buildDocxFromBody(body, { hyperlinks: { rIdL: URL, rIdM: URL2 } });
  const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
  const file = PdfFile.parse(pdf);
  const annots = file.get(file.pages()[0]!.dict, 'Annots');
  if (!Array.isArray(annots)) return [];
  return annots
    .map((a) => file.resolve(a))
    .filter((a): a is PdfDict => a instanceof Map)
    .filter((a) => (a.get('Subtype') as PdfName | undefined)?.value === 'Link')
    .map((a) => {
      const rect = (file.get(a, 'Rect') as Array<PdfValue>).map((v) => Number(v));
      const action = file.get(a, 'A') as PdfDict;
      return {
        uri: decoded(file.get(action, 'URI')),
        contents: decoded(file.get(a, 'Contents')),
        rect: [rect[0]!, rect[1]!, rect[2]!, rect[3]!],
      };
    });
}

describe('link /Contents reads in logical order', () => {
  it('keeps the words of a right-to-left link in the order they were written', async () => {
    const links = await linksOf(para(true, link(run('العنوان الرئيسي', true))));
    expect(links.map((l) => l.contents)).toEqual(['العنوان الرئيسي']);
  });

  it('keeps a left-to-right word inside a right-to-left link where it was written', async () => {
    const text = 'العنوان ABC 2024 الرئيسي';
    const links = await linksOf(
      para(true, link(run('العنوان ', true), run('ABC 2024'), run(' الرئيسي', true))),
    );
    expect(links.map((l) => l.contents)).toEqual([text]);
  });

  it('keeps an Arabic phrase inside a left-to-right link where it was written', async () => {
    const links = await linksOf(
      para(false, run('See '), link(run('the العنوان الرئيسي page')), run(' now.')),
    );
    expect(links.map((l) => l.contents)).toEqual(['the العنوان الرئيسي page']);
  });

  it('gives a wrapped right-to-left link one rectangle per line, each in logical order', async () => {
    const words = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس', 'السابع', 'الثامن'];
    const links = await linksOf(para(true, link(run(words.join(' '), true))) + narrow(200));
    expect(links.length).toBeGreaterThan(1);
    // One rectangle per line, top line first: each sits wholly below the last.
    for (let i = 1; i < links.length; i++) {
      expect(links[i]!.rect[3]).toBeLessThanOrEqual(links[i - 1]!.rect[1] + 0.01);
    }
    // Read line by line, the pieces are the link as written.
    expect(links.map((l) => l.contents).join(' ')).toBe(words.join(' '));
  });

  it('lists two links of a right-to-left line in the order they were written', async () => {
    const links = await linksOf(
      para(true, link(run('الأول', true)), run(' و ', true), link2(run('الثاني', true))),
    );
    // Tagged output hangs each annotation's Link element off the line in this
    // order, and /Tabs /S has Tab follow it: the first link written comes
    // first though it is drawn on the right.
    expect(links.map((l) => [l.uri, l.contents])).toEqual([
      [URL, 'الأول'],
      [URL2, 'الثاني'],
    ]);
    expect(links[0]!.rect[0]).toBeGreaterThan(links[1]!.rect[2]);
  });

  it('orders the pieces of a link that unlinked text splits on its line as written', async () => {
    // Base direction right to left, no run marked: "عربي ABC" is drawn
    // "ABC DEF عربي", so the link stands either side of the unlinked "DEF".
    const links = await linksOf(para(true, link(run('عربي ABC')), run(' DEF')));
    expect(links.map((l) => l.contents)).toEqual(['عربي', 'ABC']);
    // Two rectangles, neither over "DEF": the first written is the one on the right.
    expect(links[0]!.rect[0]).toBeGreaterThan(links[1]!.rect[2]);
  });
});
