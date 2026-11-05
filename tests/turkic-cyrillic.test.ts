// The letters Kazakh, Tatar, Bashkir, Chuvash, Uzbek, Azerbaijani, Sakha, Altai
// and Khakas add to the Russian alphabet — ә ғ қ ң ө ұ ү һ і җ ҙ ҡ ҫ ӑ ӗ ӳ ў ҳ ҹ
// ј ҝ ҕ ҥ ӧ ӱ ӌ and their capitals — drawn, and read back, as the letters they
// are. Measured against Word's own PDF of probe documents in each of the six
// faces an Office document most often names (.git/wip-turkic/notes.md).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocx } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { parseTtf } from '@/core/font/ttf-parser';
import { PdfFile } from '@/pdf-reader/document';
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
