import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readDocx } from '@/word/docx-reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** The baseline of each line on the first page, by the line's first word. */
function baselines(docx: Uint8Array): Map<string, number> {
  const flow = readDocx(docx).doc;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out = new Map<string, number>();
  for (const c of laid.pages[0]!.commands) {
    if (c.type !== 'line') continue;
    const line = c.line;
    const word = line.tokens.find((t) => t.kind === 'text' && !t.isSpace);
    if (word?.kind === 'text') out.set(word.text, c.baselineY);
  }
  return out;
}

/** The gaps between consecutive named lines. */
function gaps(at: Map<string, number>, names: ReadonlyArray<string>): Array<number> {
  return names.slice(1).map((n, i) => +(at.get(n)! - at.get(names[i]!)!).toFixed(2));
}

describe('w:contextualSpacing (§17.3.1.9)', () => {
  // Word for Mac: items that ask for it stand line to line, in the body and in
  // a table cell alike; the heading above keeps its 8pt after, and the last
  // item gives up its own above a paragraph of the same style.
  const para = (text: string, ctx: boolean, style?: string): string =>
    '<w:p><w:pPr>' +
    (style ? `<w:pStyle w:val="${style}"/>` : '') +
    '<w:spacing w:before="0" w:after="160"/>' +
    (ctx ? '<w:contextualSpacing/>' : '') +
    `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const run = (tag: string): string =>
    para(`${tag}H`, false) +
    para(`${tag}A`, true) +
    para(`${tag}B`, true) +
    para(`${tag}C`, true) +
    para(`${tag}P`, false) +
    para(`${tag}Q`, true, 'Other') +
    para(`${tag}R`, false);
  const names = (tag: string): Array<string> =>
    ['H', 'A', 'B', 'C', 'P', 'Q', 'R'].map((n) => tag + n);
  const cell = (xml: string): string =>
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${xml}</w:tc></w:tr></w:tbl>`;

  it('drops the space between items in the body', () => {
    const at = baselines(buildDocxFromBody(run('b')));
    const [hA, aB, bC, cP, pQ, qR] = gaps(at, names('b'));
    expect(aB).toBeCloseTo(bC!, 6);
    expect(hA! - aB!).toBeCloseTo(8, 6);
    expect(cP).toBeCloseTo(aB!, 6);
    expect(pQ! - aB!).toBeCloseTo(8, 6);
    expect(qR! - aB!).toBeCloseTo(8, 6);
  });

  it('drops it in a table cell just as in the body', () => {
    const at = baselines(buildDocxFromBody(run('b') + cell(run('t')) + '<w:p/>'));
    expect(gaps(at, names('t'))).toEqual(gaps(at, names('b')));
  });
});
