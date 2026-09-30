// §17.6.12 — the page numbers a PDF prints in its running head or foot, read
// back as the numbering they are: which numeral is the page's, in what
// numerals, and where the count starts again.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { numeralsIn, pageNumberingOf } from '@/pdf-reader/page-numbers';
import { reconstructTaggedPdf } from '@/pdf-reader/tagged';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

describe('the page numbers a band prints', () => {
  it('reads front matter in roman numerals and a body that starts again at 1', () => {
    // bug793632.pdf: four pages, footed i, ii, iii and 1.
    const numbering = pageNumberingOf(['i', 'ii', 'iii', '1']);
    expect(numbering?.runs).toEqual([
      { from: 0, format: 'lowerRoman', start: 1 },
      { from: 3, format: 'decimal', start: 1 },
    ]);
  });

  it('takes the numeral that changes from page to page for the page’s', () => {
    expect(
      pageNumberingOf(['Chapter I — 47', 'Chapter I — 48'])?.numbers.map((n) => n?.text),
    ).toEqual(['47', '48']);
    expect(pageNumberingOf(['Page 1 of 3', 'Page 2 of 3', 'Page 3 of 3'])?.runs).toEqual([
      { from: 0, format: 'decimal', start: 1 },
    ]);
  });

  it('counts an unnumbered title page with the body it opens', () => {
    expect(pageNumberingOf([undefined, '2', '3'])?.runs).toEqual([
      { from: 0, format: 'decimal', start: 1 },
    ]);
  });

  it('reads nothing where the band says the same on every page', () => {
    // ZapfDingbats.pdf signs every sheet "© RenderX 2000".
    expect(pageNumberingOf(['© RenderX 2000', '© RenderX 2000'])).toBeUndefined();
  });

  it('reads a book that leaves its blank pages out, its count skipping their numbers', () => {
    // freeculture.pdf leaves out the blank page before five of its chapters,
    // and its count jumps two at each: seven sequences in all, where a band
    // counting something else starts one on nearly every page.
    const count = (from: number): Array<string> =>
      Array.from({ length: 14 }, (_, k) => String(from + k));
    const numbering = pageNumberingOf([
      'xiii',
      'xiv',
      'xv',
      'xvi',
      ...count(1),
      ...count(17),
      ...count(33),
      ...count(49),
      ...count(65),
    ]);
    expect(numbering?.runs).toEqual([
      { from: 0, format: 'lowerRoman', start: 13 },
      { from: 4, format: 'decimal', start: 1 },
      { from: 18, format: 'decimal', start: 17 },
      { from: 32, format: 'decimal', start: 33 },
      { from: 46, format: 'decimal', start: 49 },
      { from: 60, format: 'decimal', start: 65 },
    ]);
    expect(
      pageNumberingOf(Array.from({ length: 40 }, (_, k) => `Figure ${String((k * 7) % 31)}`)),
    ).toBeUndefined();
  });

  it('takes only letters written as a numeral for one', () => {
    expect(numeralsIn('ic vx iv XIV 12345').map((n) => [n.text, n.value])).toEqual([
      ['iv', 4],
      ['XIV', 14],
    ]);
  });
});

describe('a document read back numbered as its pages are (§17.6.12)', () => {
  // Two pages of front matter numbered i, ii and a body that starts again at
  // 1, each footed "Page <n> of <total>" — written by Ream, then read back.
  const FOOTER =
    '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
    '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
    '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' +
    '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
  const BODY =
    '<w:p><w:r><w:t>Front one</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Front two</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rId11"/>' +
    '<w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr></w:p>' +
    '<w:p><w:r><w:t>Body one</w:t></w:r></w:p>' +
    '<w:sectPr><w:footerReference w:type="default" r:id="rId11"/>' +
    '<w:pgNumType w:start="1"/></w:sectPr>';

  const pdf = (tagged: boolean): Promise<Uint8Array> =>
    Ream.parse(buildDocxFromBody(BODY, { footerXml: FOOTER })).convert('pdf', {
      fonts: FONTS,
      tagged,
    });

  const footText = (bands: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined): string =>
    [...(bands?.values() ?? [])]
      .flat()
      .flatMap((el) =>
        el.kind === 'paragraph'
          ? el.paragraph.runs.map((r) => (r.field ? `{${r.field}}` : r.text))
          : [],
      )
      .join('');

  it('reads the untagged document’s numbering into its sections', async () => {
    const flow = reconstructByLayout(PdfFile.parse(await pdf(false))).doc;
    expect(
      flow.sections.map((s) => [s.properties.pageNumberFormat, s.properties.pageNumberStart]),
    ).toEqual([
      ['lowerRoman', undefined],
      [undefined, 1],
    ]);
    // The foot's number is the page's own, in the numerals each section names.
    expect(footText(flow.headersFooters)).toBe('Page {PAGE} of {NUMPAGES}');
  });

  it('reads a tagged document’s foot, which its tree leaves out as an artifact', async () => {
    const flow = reconstructTaggedPdf(PdfFile.parse(await pdf(true)))?.doc;
    expect(
      flow?.sections.map((s) => [s.properties.pageNumberFormat, s.properties.pageNumberStart]),
    ).toEqual([
      ['lowerRoman', undefined],
      [undefined, 1],
    ]);
    expect(footText(flow?.headersFooters)).toBe('Page {PAGE} of {NUMPAGES}');
  });
});
