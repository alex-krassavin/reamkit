// The letters Kazakh, Tatar, Bashkir, Chuvash, Uzbek, Azerbaijani, Sakha, Altai
// and Khakas add to the Russian alphabet — ә ғ қ ң ө ұ ү һ і җ ҙ ҡ ҫ ӑ ӗ ӳ ў ҳ ҹ
// ј ҝ ҕ ҥ ӧ ӱ ӌ and their capitals — drawn, and read back, as the letters they
// are. Measured against Word's own PDF of probe documents in each of the six
// faces an Office document most often names (.git/wip-turkic/notes.md).

import { readFileSync } from 'node:fs';

import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { buildDocx, buildDocxFromBody } from './fixtures/build-docx';
import type { FetchLike } from '@/core/fonts';
import { Ream } from '@/core/converter/ream';
import { parseTtf } from '@/core/font/ttf-parser';
import { clearFontCache } from '@/core/fonts/remote-fonts';
import { PdfFile } from '@/pdf-reader/document';
import { textForGlyphName } from '@/pdf-reader/glyph-names';
import { extractPageText } from '@/pdf-reader/text';

// The face Ream downloads for Arial, read from disk so no test waits on it.
const ARIMO = { regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Arimo-Regular.ttf')) };

/** Everything the page says, in order. */
async function textOf(source: Uint8Array): Promise<string> {
  const pdf = await Ream.parse(source).convert('pdf', { fonts: ARIMO });
  const file = PdfFile.parse(pdf);
  let out = '';
  for (const page of file.pages()) for (const run of extractPageText(file, page)) out += run.text;
  return out;
}

describe('one glyph for two letters: the schwa (ISO 32000-1 §9.10.3)', () => {
  it('is one glyph in the face, which is what puts the question', () => {
    // Arimo (and Tinos, Cousine) send U+0259 and U+04D9, U+018F and U+04D8, to
    // the same glyph. Were they apart, the tests below would prove nothing.
    const face = parseTtf(ARIMO.regular);
    expect(face.glyphForCodepoint(0x04d9)).toBe(face.glyphForCodepoint(0x0259));
    expect(face.glyphForCodepoint(0x04d8)).toBe(face.glyphForCodepoint(0x018f));
  });

  it('reads a Kazakh ә back as the Cyrillic letter', async () => {
    // /ToUnicode named the glyph after the lowest code reaching it, the Latin
    // schwa: every "және" of a Kazakh document in Arial, Times New Roman or
    // Courier New read back, and searched, as "жəне".
    const text = await textOf(buildDocx(['Барлық адамдар азат және тең. Әке, әже.']));
    expect(text).toContain('және');
    expect(text).toContain('Әке, әже');
    expect(/[əƏ]/u.test(text)).toBe(false);
  });

  it('keeps the Latin ə of a text in the Latin alphabet', async () => {
    const text = await textOf(buildDocx(['Azərbaycan Respublikası. Ərazi.']));
    expect(text).toContain('Azərbaycan');
    expect(text).toContain('Ərazi');
  });

  it('names a glyph both alphabets drew after the one drawn more', async () => {
    // One /ToUnicode entry per glyph: a page that sets the same glyph for both
    // letters can be right about only one, and it is the one it uses most.
    const text = await textOf(buildDocx(['және әке әже', 'Azərbaycan']));
    expect(text).toContain('және әке әже');
  });
});

// Cut-down copies of the files the CDN serves (tests/fixtures/fonts/NOTICE.md):
// Caladea to printable ASCII, which is all it draws of what the test sets, and
// Noto Serif to the Cyrillic letters the test sets.
const CALADEA = new Uint8Array(readFileSync('tests/fixtures/fonts/Caladea-subset.ttf'));
const NOTO_SERIF = new Uint8Array(readFileSync('tests/fixtures/fonts/NotoSerif-subset.ttf'));

/** The CDN, from disk: the regular face of each family the test needs, and nothing else. */
const fontFetch: FetchLike = (url) => {
  const regular = /_400Regular\.ttf$/u.test(url);
  const bytes = !regular
    ? undefined
    : url.includes('/caladea/')
      ? CALADEA
      : url.includes('/noto-serif/')
        ? NOTO_SERIF
        : url.includes('/arimo/')
          ? ARIMO.regular
          : undefined;
  return Promise.resolve({
    ok: bytes !== undefined,
    arrayBuffer: () => Promise.resolve((bytes ?? new Uint8Array(0)).buffer.slice(0)),
  });
};

/** A paragraph of one run in a named family. */
function runIn(family: string, text: string): string {
  return `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

describe("Cambria's Cyrillic, which its twin does not draw", () => {
  // The download cache is keyed by URL and outlives a test: what this one
  // serves for Caladea must not be what the next one is given.
  beforeEach(() => clearFontCache());
  afterEach(() => clearFontCache());

  it('is a face the twin lacks, which is what puts the question', () => {
    const face = parseTtf(CALADEA);
    expect(face.glyphForCodepoint('Б'.codePointAt(0)!)).toBe(0);
    expect(face.glyphForCodepoint('C'.codePointAt(0)!)).not.toBe(0);
  });

  it('is drawn in Noto Serif, squeezed to Cambria’s width, beside Caladea’s Latin', async () => {
    // Every Cyrillic letter of a Cambria run fell to the first family loaded,
    // Arimo: a grotesque in a roman line, its letters 9 to 11 % off Cambria's
    // widths on average and up to 67 % for one letter.
    const word = 'Барлық';
    const docx = buildDocxFromBody(runIn('Cambria', `Cambria ${word} адамдар және`));
    const { bytes, losses } = await Ream.parse(docx).convertWithReport('pdf', { fontFetch });
    expect(losses).toEqual([]);
    const file = PdfFile.parse(bytes);
    const runs = extractPageText(file, file.pages()[0]!);
    const text = runs.map((r) => r.text).join('');
    expect(text).toContain(`Cambria ${word} адамдар және`);
    const faceOf = (s: string): string | undefined =>
      runs.find((r) => r.text.includes(s))?.fontName;
    expect(faceOf('Cambria')).toMatch(/^caladea/u);
    expect(faceOf(word)).toMatch(/^notoserif/u);
    // The word is as wide as Noto Serif sets it, times the regular face's
    // factor: 0.894, Cambria 6.90's Cyrillic over Noto Serif's. Kerning and the
    // PDF's rounded widths move it by a fraction of a percent; unsqueezed it
    // would be 1.
    const run = runs.find((r) => r.text === word)!;
    const noto = parseTtf(NOTO_SERIF);
    const natural = [...word].reduce(
      (sum, ch) => sum + noto.advanceWidths[noto.glyphForCodepoint(ch.codePointAt(0)!)]!,
      0,
    );
    const squeeze = (run.endX - run.x) / ((natural / noto.unitsPerEm) * 11);
    expect(squeeze).toBeGreaterThan(0.88);
    expect(squeeze).toBeLessThan(0.9);
  });

  it('is fetched only for a document that holds such a letter', async () => {
    const asked: Array<string> = [];
    const spy: FetchLike = (url) => {
      asked.push(url);
      return fontFetch(url);
    };
    await Ream.parse(buildDocxFromBody(runIn('Cambria', 'Only Latin here.'))).convert('pdf', {
      fontFetch: spy,
    });
    expect(asked.some((u) => u.includes('/noto-serif/'))).toBe(false);
    await Ream.parse(buildDocxFromBody(runIn('Cambria', 'Әке'))).convert('pdf', { fontFetch: spy });
    expect(asked.some((u) => u.includes('/noto-serif/'))).toBe(true);
  });
});

/**
 * A simple Type 1 font with NO /ToUnicode, whose codes say what they are only
 * through `/Encoding /Differences` — glyph names, as a PDF from TeX or an old
 * Type 1 Cyrillic face gives them.
 */
function namedGlyphPdf(names: ReadonlyArray<string>): Uint8Array {
  const codes = names.map((_, i) => (i + 1).toString(16).padStart(2, '0')).join('');
  const content = `BT /F1 12 Tf 72 720 Td <${codes}> Tj ET`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    `<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+KzTimes /FirstChar 1 /LastChar ${String(names.length)} /Encoding 6 0 R >>`,
    `<< /Type /Encoding /Differences [1 ${names.map((n) => `/${n}`).join(' ')}] >>`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('a Cyrillic letter a font names rather than maps (§9.6.6.1, Adobe Glyph List)', () => {
  // The Adobe Glyph List's name for each of the 52 letters, lower case then
  // capital, as its glyphlist.txt spells them.
  const NAMES: ReadonlyArray<readonly [string, string]> = [
    ['schwacyrillic', 'ә'],
    ['ghestrokecyrillic', 'ғ'],
    ['kadescendercyrillic', 'қ'],
    ['endescendercyrillic', 'ң'],
    ['obarredcyrillic', 'ө'],
    ['ustraightstrokecyrillic', 'ұ'],
    ['ustraightcyrillic', 'ү'],
    ['shhacyrillic', 'һ'],
    ['icyrillic', 'і'],
    ['zhedescendercyrillic', 'җ'],
    ['zedescendercyrillic', 'ҙ'],
    ['kabashkircyrillic', 'ҡ'],
    ['esdescendercyrillic', 'ҫ'],
    ['abrevecyrillic', 'ӑ'],
    ['iebrevecyrillic', 'ӗ'],
    ['uhungarumlautcyrillic', 'ӳ'],
    ['ushortcyrillic', 'ў'],
    ['hadescendercyrillic', 'ҳ'],
    ['cheverticalstrokecyrillic', 'ҹ'],
    ['jecyrillic', 'ј'],
    ['kaverticalstrokecyrillic', 'ҝ'],
    ['ghemiddlehookcyrillic', 'ҕ'],
    ['enghecyrillic', 'ҥ'],
    ['odieresiscyrillic', 'ӧ'],
    ['udieresiscyrillic', 'ӱ'],
    ['chekhakassiancyrillic', 'ӌ'],
  ];

  it('reads every Turkic letter, and its capital, by the name the list gives it', () => {
    for (const [name, letter] of NAMES) {
      expect(textForGlyphName(name)).toBe(letter);
      expect(textForGlyphName(name[0]!.toUpperCase() + name.slice(1))).toBe(letter.toUpperCase());
    }
  });

  it('reads the afii names PostScript Cyrillic faces carry', () => {
    expect(textForGlyphName('afii10017')).toBe('А');
    expect(textForGlyphName('afii10065')).toBe('а');
    expect(textForGlyphName('afii10846')).toBe('ә');
    expect(textForGlyphName('afii10103')).toBe('і');
    expect(textForGlyphName('afii10110')).toBe('ў');
  });

  it('reads a page whose only word of it is those names', () => {
    // Without the names the codes read as nothing a reader can use.
    const file = PdfFile.parse(
      namedGlyphPdf([
        'Kadescendercyrillic',
        'afii10065',
        'zecyrillic',
        'acyrillic',
        'kadescendercyrillic',
        'schwacyrillic',
      ]),
    );
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('');
    expect(text).toBe('Қазақә');
  });
});
