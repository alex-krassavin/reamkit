// A character no face on hand can draw is drawn as glyph 0: an empty box on
// the page and nothing in the text layer. A line of Hindi set in Latin faces
// came out exactly so, with an empty loss report, and `strict` let it through.
// It is a loss now, named by its code points, whichever target draws it.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocx, buildDocxFromBody } from './fixtures/build-docx';
import type { FetchLike } from '@/core/fonts';
import { Ream } from '@/core/converter/ream';
import { clearFontCache } from '@/core/fonts';
import { ConversionLossError } from '@/core/ir';

const font = (name: string): Uint8Array<ArrayBuffer> =>
  new Uint8Array(readFileSync(`tests/fixtures/fonts/${name}.ttf`));
const ROBOTO = { regular: font('Roboto-Regular') };

describe('a character no face has a glyph for', () => {
  it('is reported as a loss, named by its code points', async () => {
    const docx = buildDocx(['Hindi: हिन्दी']);
    const { losses } = await Ream.parse(docx).convertWithReport('pdf', { fonts: ROBOTO });
    const loss = losses.find((l) => l.feature === 'text');
    expect(loss?.severity).toBe('dropped');
    expect(loss?.where).toBe('page 1');
    // Each character once, in the order the text has them.
    expect(loss?.detail).toContain('U+0939, U+093F, U+0928, U+094D, U+0926, U+0940;');
    expect(loss?.detail).not.toContain('U+0048'); // the H of "Hindi" has its glyph
  });

  it('counts what it does not name', async () => {
    const docx = buildDocx(['ক খ গ ঘ ঙ চ ছ জ ঝ ঞ']);
    const { losses } = await Ream.parse(docx).convertWithReport('pdf', { fonts: ROBOTO });
    expect(losses.find((l) => l.feature === 'text')?.detail).toContain('U+099C and 2 more;');
  });

  it('is on the page it first appears on', async () => {
    const body =
      '<w:p><w:r><w:t>Latin</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>தமிழ்</w:t></w:r></w:p>';
    const { losses } = await Ream.parse(buildDocxFromBody(body)).convertWithReport('pdf', {
      fonts: ROBOTO,
    });
    expect(losses.find((l) => l.feature === 'text')?.where).toBe('page 2');
  });

  it('fails a strict conversion', async () => {
    const docx = buildDocx(['ગુજરાતી']);
    await expect(Ream.parse(docx).convert('pdf', { fonts: ROBOTO, strict: true })).rejects.toThrow(
      ConversionLossError,
    );
  });

  it('is reported for an SVG too', async () => {
    const docx = buildDocx(['ਪੰਜਾਬੀ']);
    const { losses } = await Ream.parse(docx).convertWithReport('svg', { fonts: ROBOTO });
    expect(losses.find((l) => l.feature === 'text')?.detail).toContain('U+0A2A');
  });

  it('is not reported when every character has a glyph', async () => {
    const docx = buildDocx(['Hello, world — «привет»', 'Tab\there']);
    const { losses } = await Ream.parse(docx).convertWithReport('pdf', { fonts: ROBOTO });
    expect(losses.filter((l) => l.feature === 'text')).toEqual([]);
  });

  it('is not reported when another loaded face draws it', async () => {
    // The dotted circle is in Arimo and not in Roboto. The Times run is drawn
    // in "Tinos", which here is Roboto, and the character goes to the Arial
    // run's face, which is Arimo.
    const served: FetchLike = (url) => {
      const bytes = url.includes('/arimo/') ? font('Arimo-Regular') : font('Roboto-Regular');
      return Promise.resolve({
        ok: true,
        arrayBuffer: () =>
          Promise.resolve(
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
          ),
      });
    };
    const run = (family: string, text: string): string =>
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr>` +
      `<w:t>${text}</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(run('Times New Roman', 'a ◌ b') + run('Arial', 'c'));
    clearFontCache();
    const { losses } = await Ream.parse(docx).convertWithReport('pdf', { fontFetch: served });
    clearFontCache();
    expect(losses.filter((l) => l.feature === 'text')).toEqual([]);
  });
});
