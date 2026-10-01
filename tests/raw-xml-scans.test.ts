// The readers that take one attribute out of a raw part — embedded fonts, the
// default font, the document's language — search its text with expressions,
// beside the parser rather than through it. Searched for as they were, they
// were quadratic in text the file decides: many starts of a tag in one stretch
// with no `>` (a comment holds that as well as broken markup does), or
// elements that never close.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { tagMatches } from '@/core/opc/tag-scan';
import { readDocx } from '@/word/docx-reader';
import { detectDocxFamilyKeys, detectDocxFontFamily } from '@/word/docx-to-pdf';
import { parseFontTable } from '@/word/font-table';

const encode = (s: string): Uint8Array => new TextEncoder().encode(s);
const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** How long `f` takes, in milliseconds. */
function timed(f: () => void): number {
  const start = performance.now();
  f();
  return performance.now() - start;
}

describe('tagMatches', () => {
  it('finds what a global search with the same expression finds', () => {
    const source = String.raw`<w:rFonts[^>]*?\bw:(ascii|asciiTheme)="([^"]+)"`;
    const xml =
      '<w:rFonts w:hAnsi="x"/><w:rFonts w:ascii="Arial"/>' +
      '<w:rFonts <w:rFonts w:asciiTheme="minorHAnsi"><w:rFontsX w:ascii="B">';
    const found = [...tagMatches(xml, '<w:rFonts', false, new RegExp(source, 'y'))];
    const expected = [...xml.matchAll(new RegExp(source, 'g'))];
    expect(found.map((m) => [m.index, m[1], m[2]])).toEqual(
      expected.map((m) => [m.index, m[1], m[2]]),
    );
  });
});

describe('the font table (§17.8.3)', () => {
  it("reads each font's embedded faces", () => {
    const xml =
      '<w:fonts><w:font w:name="Body"><w:embedRegular r:id="rId1" w:fontKey="{A}"/>' +
      '<w:embedBold r:id="rId2" w:fontKey="{B}"/></w:font><w:font w:name="Plain"/></w:fonts>';
    expect(parseFontTable(encode(xml))).toEqual([
      {
        name: 'Body',
        embeds: {
          regular: { rId: 'rId1', fontKey: '{A}' },
          bold: { rId: 'rId2', fontKey: '{B}' },
        },
      },
    ]);
  });

  it('reads a table built to stall it in time linear in it', () => {
    const tables = [
      // Well-formed, and no font closes: each was read on to the end.
      `<w:fonts>${'<w:font w:name="x"/>'.repeat(60_000)}</w:fonts>`,
      // Opened over and over, never ended.
      '<w:font '.repeat(60_000),
      // One tag of many ids and no key, read again after each id.
      `<w:font w:name="x"><w:embedRegular ${'r:id="a" '.repeat(40_000)}/></w:font>`,
    ].map(encode);
    for (const table of tables) {
      expect(timed(() => expect(parseFontTable(table)).toEqual([]))).toBeLessThan(1000);
    }
  });
});

describe("the document's font, read from its raw parts", () => {
  it('takes the default font, then the commonest', () => {
    const defaults =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia"/></w:rPr></w:rPrDefault></w:docDefaults>';
    expect(detectDocxFontFamily(buildDocxFromBody(p('x'), { stylesXml: defaults }))).toBe(
      'Georgia',
    );
    const run = '<w:r><w:rPr><w:rFonts w:ascii="Verdana"/></w:rPr><w:t>x</w:t></w:r>';
    expect(detectDocxFontFamily(buildDocxFromBody(`<w:p>${run}${run}</w:p>`))).toBe('Verdana');
  });

  it('reads parts built to stall it in time linear in them', () => {
    // Well-formed defaults that name no font, each read on to the end.
    const defaults = buildDocxFromBody(p('x'), {
      stylesXml: '<w:docDefaults></w:docDefaults>'.repeat(60_000),
    });
    expect(timed(() => expect(detectDocxFontFamily(defaults)).toBeUndefined())).toBeLessThan(1000);
    // `<w:rFonts` in a comment, with no `>` before the comment ends.
    const comment = buildDocxFromBody(`${p('x')}<!-- ${'<w:rFonts '.repeat(60_000)} -->`);
    expect(timed(() => expect([...detectDocxFamilyKeys(comment)]).toEqual(['arimo']))).toBeLessThan(
      1000,
    );
  });
});

describe("the document's language", () => {
  it('is found past a comment built to stall the search, in time linear in it', () => {
    const lang = '<w:r><w:rPr><w:lang w:val="de-DE"/></w:rPr><w:t>x</w:t></w:r>';
    const docx = buildDocxFromBody(`<!-- ${'<w:lang '.repeat(60_000)} --><w:p>${lang}</w:p>`);
    expect(timed(() => expect(readDocx(docx).doc.language).toBe('de-DE'))).toBeLessThan(1000);
  });
});
