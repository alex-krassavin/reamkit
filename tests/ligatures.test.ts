// `ﬀ` is not a letter of any alphabet — it is one glyph for two f's, a decision
// of the FACE. Most faces make it through their own `liga` table and give the
// presentation form no code point at all; the ones that do also answer to the
// legacy private-use slot the same glyph sat in before Unicode had a block for
// it. Both facts cost text: bug1873345.pdf reads "different" and "flavor", and
// we set "di erent" — the character reached a face that could not draw it and
// went out silently — while the ligatures the face COULD draw went into
// `/ToUnicode` as U+F001/U+F002, so our own page could not be searched.

import { describe, expect, it } from 'vitest';

import { buildDocx } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

/** Everything the page says, in order. */
async function textOf(source: Uint8Array): Promise<string> {
  const pdf = await Ream.parse(source).convert('pdf');
  const file = PdfFile.parse(pdf);
  let out = '';
  for (const page of file.pages()) for (const run of extractPageText(file, page)) out += run.text;
  return out;
}

describe('typographic ligatures (§9.10.3)', () => {
  it('sets the letters of a ligature no face on hand can draw', async () => {
    // U+FB00 is in none of the substitute faces: before, the word came out
    // "di erent" — two letters short, and nothing said.
    expect(await textOf(buildDocx(['diﬀerent']))).toContain('different');
  });

  it('reads back a ligature the face DOES carry as its letters', async () => {
    // Arimo draws U+FB01 and U+FB02, so each is set as one glyph — and that
    // glyph answers to U+F001/U+F002 as well, which is what a `/ToUnicode`
    // built by scanning upward stated. A search for "flavor" finds nothing in
    // a page whose text says U+F002 followed by "avor".
    const text = await textOf(buildDocx(['a ﬂavor speciﬁcation']));
    expect(text).toContain('flavor');
    expect(text).toContain('specification');
    // Nor does a private-use code reach the text of the page at all.
    expect(/[\ue000-\uf8ff]/u.test(text)).toBe(false);
  });

  it('leaves the letters of a word that carries no ligature', async () => {
    expect(await textOf(buildDocx(['fluffier stuff']))).toContain('fluffier stuff');
  });

  it('sets a mathematical letter as the letter it is', async () => {
    // U+1D45D is a lower case p set in italic, and no ordinary face carries it:
    // bug1529502.pdf reads "p ← trim(p)" and every one of them came out a box.
    // The slant is lost with the code point; the letter is not.
    expect(await textOf(buildDocx(['\u{1d45d} ← trim(\u{1d45d})']))).toContain('p ← trim(p)');
    // Its digits too: U+1D7D9 is a double-struck 1.
    expect(await textOf(buildDocx(['\u{1d7d9}\u{1d7da}']))).toContain('12');
    // `ℎ` is the same story in the letterlike block — but a face on hand has
    // that one, so it is drawn as written and stays the character it is.
    expect(await textOf(buildDocx(['ℎ']))).toContain('ℎ');
  });

  it('leaves alone a letterlike character that abbreviates rather than styles', async () => {
    // `№` stands for "No" and `℅` for "c/o": those are abbreviations, not one
    // letter in a style, and a face that lacks them draws what it always drew.
    expect(await textOf(buildDocx(['№ ℅']))).not.toContain('No c/o');
  });
});
