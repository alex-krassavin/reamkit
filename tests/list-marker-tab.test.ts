import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Line, TextLineItem } from '@/layout/page-doc';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** Every line of the first page, as laid out. */
function lines(docx: Uint8Array): Array<Line> {
  const flow = readDocx(docx).doc;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  return laid.pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line').map(
    (c) => c.line,
  );
}

describe('the tab a list marker ends in is spent before the first line breaks', () => {
  // A bullet at a first-line indent, its tab out to the next default stop some
  // 30pt on: the first line has that much less room for words than the rest.
  const numberingXml =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>' +
    '<w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
    '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const words = Array.from({ length: 80 }, (_, i) => `w${i}`).join(' ');
  const item = (jc: string): string =>
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' +
    `<w:ind w:firstLine="709"/><w:jc w:val="${jc}"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${words}</w:t></w:r></w:p>`;

  it('keeps the first line of an item inside the measure', () => {
    const laid = lines(buildDocxFromBody(item('left'), { numberingXml }));
    const first = laid[0]!;
    const tab = first.tokens.find((t) => t.kind === 'text' && t.tab === true);
    expect(tab?.widthPt).toBeGreaterThan(20);
    expect(first.contentWidthPt).toBeLessThanOrEqual(first.availableWidthPt + 0.01);
  });
});
