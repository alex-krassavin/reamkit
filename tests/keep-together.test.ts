// ECMA-376 Part 1 §17.3.1.14 `w:keepNext` and §17.3.1.15 `w:keepLines`: a
// paragraph kept with the start of the next one, or kept whole, goes to the
// next page together where the rest of this one cannot hold it.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildDoc } from './fixtures/build-doc';
import type { FlowDoc } from '@/core/ir/flow';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDoc } from '@/word/doc/doc-reader';
import { readDocx } from '@/word/docx-reader';
import { writeDocx } from '@/word/docx-writer';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** A paragraph of `lines` soft-broken lines, with `pPr` added to its properties. */
interface Para {
  readonly lines: number;
  readonly pPr?: string;
}

/** The paragraphs on pages of exactly five 12pt lines. */
function document(paragraphs: ReadonlyArray<Para>): Uint8Array {
  const body = paragraphs.map((p, n) => {
    const lines = Array.from({ length: p.lines }, (_, i) => `<w:t>p${n} line ${i}</w:t>`);
    return (
      `<w:p><w:pPr>${p.pPr ?? ''}<w:spacing w:line="240" w:lineRule="exact"/></w:pPr>` +
      `<w:r>${lines.join('<w:br/>')}</w:r></w:p>`
    );
  });
  return buildDocxFromBody(
    body.join('') +
      '<w:sectPr><w:pgSz w:w="8400" w:h="2640"/>' +
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>',
  );
}

/** The lines each page holds. */
function linesPerPage(docx: Uint8Array): Array<number> {
  const flow = Ream.parse(docx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages.map((page) => page.commands.filter((c) => c.type === 'line').length);
}

const KEEP_NEXT = '<w:keepNext/>';
const KEEP_LINES = '<w:keepLines/>';
const NO_WIDOWS = '<w:widowControl w:val="0"/>';

describe('a paragraph kept with the next (§17.3.1.14)', () => {
  it('goes to the next page with the paragraph it is kept with', () => {
    // A heading after three lines, then three lines its widow control keeps
    // together: they go to the next page, and the heading goes with them.
    expect(linesPerPage(document([{ lines: 3 }, { lines: 1 }, { lines: 3 }]))).toEqual([4, 3]);
    expect(
      linesPerPage(document([{ lines: 3 }, { lines: 1, pPr: KEEP_NEXT }, { lines: 3 }])),
    ).toEqual([3, 4]);
  });

  it('takes a chain of them along, and as much of the next as its widow control keeps', () => {
    const heading = { lines: 1, pPr: KEEP_NEXT };
    expect(linesPerPage(document([{ lines: 3 }, heading, heading, { lines: 2 }]))).toEqual([3, 4]);
  });

  it('decides a chain at its head, and moves one no page holds, as Word does', () => {
    // Word and LibreOffice alike: six headings kept with a two-line paragraph
    // stand where they begin a page and break where it runs out; after a
    // line of text they move to the next page first.
    const heading = { lines: 1, pPr: KEEP_NEXT };
    const chain = Array.from({ length: 6 }, () => heading);
    expect(linesPerPage(document([...chain, { lines: 2 }]))).toEqual([5, 3]);
    expect(linesPerPage(document([{ lines: 1 }, ...chain, { lines: 2 }]))).toEqual([1, 5, 3]);
    // …and a paragraph kept with one its widow control keeps whole leaves a
    // page where it would fit, though the two cannot share the next.
    expect(
      linesPerPage(document([{ lines: 2 }, { lines: 3, pPr: KEEP_NEXT }, { lines: 3 }])),
    ).toEqual([2, 3, 3]);
  });
});

describe('a paragraph that keeps its lines together (§17.3.1.15)', () => {
  it('goes whole to the next page rather than break', () => {
    expect(linesPerPage(document([{ lines: 3 }, { lines: 3, pPr: NO_WIDOWS }]))).toEqual([5, 1]);
    expect(
      linesPerPage(document([{ lines: 3 }, { lines: 3, pPr: NO_WIDOWS + KEEP_LINES }])),
    ).toEqual([3, 3]);
  });

  it('moves to a page of its own before breaking where no page holds it', () => {
    // Word and LibreOffice both start it on the next page, then break it.
    expect(
      linesPerPage(document([{ lines: 1 }, { lines: 7, pPr: NO_WIDOWS + KEEP_LINES }])),
    ).toEqual([1, 5, 2]);
  });
});

describe('keeping, read and written', () => {
  it('travels through a .docx from the paragraph or its style', () => {
    const source = buildDocxFromBody(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Kept</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>',
      {
        stylesXml:
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
          '<w:pPr><w:keepNext/><w:keepLines/></w:pPr></w:style>',
      },
    );
    const keeps = (doc: FlowDoc): Array<[boolean | undefined, boolean | undefined]> =>
      doc.body.map((el) =>
        el.kind === 'paragraph'
          ? [el.paragraph.properties.keepNext, el.paragraph.properties.keepLines]
          : [undefined, undefined],
      );
    const flow = readDocx(source).doc;
    expect(keeps(flow)).toEqual([
      [true, true],
      [false, true],
      [false, false],
    ]);
    const written = writeDocx(flow).bytes;
    const body = new TextDecoder().decode(OpcPackage.open(written).getMainDocument().data);
    expect(body).toContain('<w:pPr><w:keepNext/><w:keepLines/>');
    expect(keeps(readDocx(written).doc)).toEqual(keeps(flow));
  });

  it('is read from a .doc paragraph’s sprms', () => {
    const doc = readDoc(
      buildDoc([{ text: 'Head\rBody\r', compressed: false }], {
        paraRuns: [{ length: 5, keepNext: true, keepLines: true }, { length: 5 }],
      }),
    ).doc;
    const [head, body] = doc.body.flatMap((el) => (el.kind === 'paragraph' ? [el] : []));
    expect(head?.paragraph.properties.keepNext).toBe(true);
    expect(head?.paragraph.properties.keepLines).toBe(true);
    expect(body?.paragraph.properties.keepNext).toBe(false);
  });
});
