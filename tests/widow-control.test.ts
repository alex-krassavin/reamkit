// ECMA-376 Part 1 §17.3.1.44 — widow control: a page break that would leave a
// paragraph's first line alone at the foot of a page, or its last line alone
// at the head of the next, is moved so that neither stands alone. Word applies
// it wherever nothing says otherwise.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/**
 * Paragraphs of the given line counts on pages of exactly five 12pt lines:
 * each line its own soft break, so nothing but the page decides where a line
 * falls.
 */
function document(lineCounts: ReadonlyArray<number>, widowControl?: boolean): Uint8Array {
  const control =
    widowControl === undefined ? '' : `<w:widowControl w:val="${widowControl ? '1' : '0'}"/>`;
  const paragraphs = lineCounts.map((count, p) => {
    const lines = Array.from({ length: count }, (_, i) => `<w:t>p${p} line ${i}</w:t>`);
    return (
      `<w:p><w:pPr>${control}<w:spacing w:line="240" w:lineRule="exact"/></w:pPr>` +
      `<w:r>${lines.join('<w:br/>')}</w:r></w:p>`
    );
  });
  return buildDocxFromBody(
    paragraphs.join('') +
      // 132pt tall, 36pt margins: 60pt of text, five lines.
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

describe('widow control (§17.3.1.44)', () => {
  it('takes a first line that would stand alone at the foot of a page to the next', () => {
    expect(linesPerPage(document([4, 3]))).toEqual([4, 3]);
  });

  it('takes a line along with a last line that would stand alone at the head of a page', () => {
    expect(linesPerPage(document([2, 4]))).toEqual([4, 2]);
  });

  it('moves a three-line paragraph whole when two lines of it would fit', () => {
    // The last line alone is a widow; taking one more along leaves the first
    // alone, an orphan — so all three go.
    expect(linesPerPage(document([3, 3]))).toEqual([3, 3]);
  });

  it('keeps two lines at the head of the last page of a paragraph longer than a page', () => {
    expect(linesPerPage(document([11]))).toEqual([5, 4, 2]);
    expect(linesPerPage(document([12]))).toEqual([5, 5, 2]);
  });

  it('breaks where the page is full when a paragraph turns it off', () => {
    expect(linesPerPage(document([4, 3], false))).toEqual([5, 2]);
    expect(linesPerPage(document([2, 4], false))).toEqual([5, 1]);
    expect(linesPerPage(document([11], false))).toEqual([5, 5, 1]);
  });

  it('is on where nothing says otherwise', () => {
    expect(linesPerPage(document([4, 3], true))).toEqual(linesPerPage(document([4, 3])));
  });
});
