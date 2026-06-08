import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { convertDocxToPdfSync } from '@/converter';
import { parseTtf } from '@/font';
import { createHyphenator, getHyphenator } from '@/hyphenation';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('Liang hyphenation algorithm', () => {
  it('respects leftMin and rightMin defaults', () => {
    // No patterns → no break points proposed; but even with patterns, the
    // first leftMin chars and last rightMin chars are protected.
    const h = createHyphenator([], { leftMin: 2, rightMin: 3 });
    expect(h.hyphenate('any')).toEqual([]);
    expect(h.hyphenate('a')).toEqual([]);
  });

  it('finds break points via parsed patterns', () => {
    // "hy3ph" → between y and p (position 2 in "hyph"). "1y" → between any
    // preceding char and a "y" (odd, so allowed).
    const h = createHyphenator(['hy3ph'], { leftMin: 1, rightMin: 1 });
    expect(h.hyphenate('hyphen')).toEqual([2]);
  });

  it('takes the maximum weight when multiple patterns overlap', () => {
    // 2-weight at one position is even (no break); a 3-weight elsewhere wins.
    const h = createHyphenator(['ab2c', 'b3c'], { leftMin: 1, rightMin: 1 });
    const breaks = h.hyphenate('abcabc');
    // Both "ab2c" and "b3c" propose weight at position 2 (between b and c).
    // Max = 3 → odd → break.
    expect(breaks).toContain(2);
  });

  it('honours an explicit exception list', () => {
    const h = createHyphenator(['hy3ph'], {
      leftMin: 2,
      rightMin: 2,
      exceptions: ['as-so-ci-ate'],
    });
    expect(h.hyphenate('associate')).toEqual([2, 4, 6]);
    expect(h.hyphenate('hyphen')).toEqual([2]); // pattern path unchanged
  });
});

describe('Bundled language patterns', () => {
  it('hyphenates English words using the en-US bundle', async () => {
    const h = await getHyphenator('en-us');
    // Well-known en-US splits.
    expect(h.hyphenate('hyphenation')).toContain(2);
    expect(h.hyphenate('computer')).toEqual(expect.arrayContaining([3]));
    expect(h.hyphenate('typography').length).toBeGreaterThan(0);
    // Two-letter words must not be split.
    expect(h.hyphenate('it')).toEqual([]);
  });

  it('hyphenates Russian words using the ru bundle', async () => {
    const h = await getHyphenator('ru');
    // "программа" — known to break around "про-грам-ма".
    const breaks = h.hyphenate('программа');
    expect(breaks.length).toBeGreaterThan(0);
    // "ввод" is short — no breaks expected.
    expect(h.hyphenate('ввод')).toEqual([]);
  });
});

describe('Hyphenator end-to-end through convertDocxToPdfSync', () => {
  it('emits a hyphen glyph at line breaks when a hyphenator is supplied', async () => {
    // A narrow justified column packed with long words forces Knuth-Plass to
    // accept hyphenation breaks. We check that at least one Tj output ends
    // with the "-" glyph encoded in Identity-H.
    const hyphenator = await getHyphenator('en-us');
    const para =
      'Internationalization characteristics conventionally documented configuration ' +
      'implementations representativeness multidimensional approximations.';
    const body = `
      <w:p>
        <w:pPr>
          <w:jc w:val="both"/>
          <w:ind w:left="3600" w:right="3600"/>
        </w:pPr>
        <w:r><w:t>${para}</w:t></w:r>
      </w:p>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS, hyphenator });
    const text = asLatin1(pdf);

    const parsed = parseTtf(FONTS.regular);
    const hyphenGid = parsed.glyphForCodepoint('-'.codePointAt(0)!);
    const hyphenHex = hyphenGid.toString(16).padStart(4, '0').toUpperCase();
    // The hyphen glyph must appear at the end of at least one Tj string
    // (i.e. as the trailing CID in a "<...HHHH> Tj" hex literal).
    const tjMatches = [...text.matchAll(/<([0-9A-F]+)> Tj/g)];
    const endsWithHyphen = tjMatches.some((m) => m[1]!.endsWith(hyphenHex));
    expect(endsWithHyphen).toBe(true);
  });

  it('produces a different paragraph layout with vs. without hyphenation', async () => {
    const hyphenator = await getHyphenator('en-us');
    const para =
      'Internationalization documents conventional configurations comprehensive ' +
      'representativeness multidimensional approximations classifications.';
    const body = `
      <w:p>
        <w:pPr>
          <w:jc w:val="both"/>
          <w:ind w:left="3600" w:right="3600"/>
        </w:pPr>
        <w:r><w:t>${para}</w:t></w:r>
      </w:p>`;
    const docx = buildDocxFromBody(body);
    const pdfA = convertDocxToPdfSync(docx, { fonts: FONTS });
    const pdfB = convertDocxToPdfSync(docx, { fonts: FONTS, hyphenator });

    // Count Tm operators per PDF (one per line in non-justify; more in
    // justify). With hyphenation the algorithm picks different line breaks
    // → byte-level layout must differ.
    expect(asLatin1(pdfB)).not.toBe(asLatin1(pdfA));
  });
});
