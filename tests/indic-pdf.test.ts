// The scripts of India, from a document to its page: the shaper's glyphs in
// its order, marks set off the baseline, spacing by cluster, a syllable a run
// boundary cuts kept whole, the join controls in the face of their word, and a
// text layer that reads as the text was written. The faces are the Noto
// fixtures cut down to the characters set here (see fixtures/fonts/NOTICE.md).

import { readFileSync } from 'node:fs';

import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { FetchLike } from '@/core/fonts';
import { Ream } from '@/core/converter/ream';
import { createFontMeasure, parseTtf } from '@/core/font';
import { shapeIndic } from '@/core/font/indic-shaper';
import { clearFontCache } from '@/core/fonts';

const fixture = (name: string): Uint8Array<ArrayBuffer> =>
  new Uint8Array(readFileSync(`tests/fixtures/fonts/${name}.ttf`));
const DEVANAGARI = fixture('NotoSansDevanagari-shaping');
const BENGALI = fixture('NotoSansBengali-shaping');

const latin1 = new TextDecoder('latin1');
const hex4 = (gid: number): string => gid.toString(16).padStart(4, '0').toUpperCase();
const utf16 = (text: string): string =>
  'FEFF' + [...text].map((c) => hex4(c.codePointAt(0)!)).join('');

const para = (text: string, rPr = ''): string =>
  `<w:p><w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function pdfBytes(body: string, regular: Uint8Array): Promise<Uint8Array> {
  return Ream.parse(buildDocxFromBody(body)).convert('pdf', { fonts: { regular } });
}

async function pdfOf(body: string, regular: Uint8Array): Promise<string> {
  return latin1.decode(await pdfBytes(body, regular));
}

/** Every Flate stream of a PDF, inflated. */
function inflated(pdf: Uint8Array): Array<string> {
  // One character per byte, so an index in the text is an offset in the bytes.
  const text = latin1.decode(pdf);
  const out: Array<string> = [];
  const re = /\/FlateDecode[^]*?stream\r?\n/gu;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    const start = m.index + m[0].length;
    const end = text.indexOf('endstream', start);
    try {
      out.push(latin1.decode(unzlibSync(pdf.subarray(start, end))));
    } catch {
      // Not a stream this test reads.
    }
  }
  return out;
}

/** The glyphs of a shaped run, as the hex a content stream shows them in, kern splices allowed. */
function shownPattern(gids: ReadonlyArray<number>): RegExp {
  return new RegExp(gids.map(hex4).join(String.raw`(?:>? ?-?[\d.]+ ?<?)?`), 'u');
}

describe('a word of the scripts of India on a PDF page', () => {
  const parsed = parseTtf(DEVANAGARI);

  it('is drawn in the shaper’s order and read back in the text’s', async () => {
    const pdf = await pdfOf(para('हिन्दी'), DEVANAGARI);
    const run = shapeIndic(parsed, 'हिन्दी')!;
    // The short i first, then ha: the order the face expects.
    expect(pdf).toMatch(shownPattern(run.gids));
    // …bracketed with the word as it was written (ISO 32000-1 §14.9.4).
    expect(pdf).toContain(`/Span <</ActualText <${utf16('हिन्दी')}>>> BDC`);
  });

  it('names in /ToUnicode a glyph only shaping reaches', async () => {
    const pdf = await pdfBytes(para('हिन्दी'), DEVANAGARI);
    const shortI = shapeIndic(parsed, 'हिन्दी')!.gids[0]!;
    expect(parsed.glyphForCodepoint(0x093f)).not.toBe(shortI);
    const cmap = inflated(pdf).find((s) => s.includes('beginbfchar'))!;
    expect(cmap).toContain(`<${hex4(shortI)}> <093F>`);
  });

  it('sets a mark off the baseline with the text rise, and puts the rise back', async () => {
    const pdf = await pdfOf(para('কোথায়'), BENGALI);
    const run = shapeIndic(parseTtf(BENGALI), 'কোথায়')!;
    const dy = run.yOffsets.find((y) => y !== 0)!;
    const size = Number(/\/F\d+ ([\d.]+) Tf/u.exec(pdf)![1]);
    const rise = (dy * size) / parseTtf(BENGALI).unitsPerEm;
    expect(pdf).toMatch(new RegExp(`${String(Number(rise.toFixed(3)))} Ts [^]*? 0 Ts`, 'u'));
  });

  it('spaces letters by cluster, keeping a vowel sign on its consonant', async () => {
    // §17.3.2.35, two points: कि ता ब are three letters, not five glyphs.
    const pdf = await pdfOf(para('किताब', '<w:spacing w:val="40"/>'), DEVANAGARI);
    expect(pdf).not.toMatch(/ 2 Tc/u);
    const measure = createFontMeasure(parsed);
    expect(measure.spacingSlots('किताब')).toBe(3);
    expect(measure.spacingSlots('kitab')).toBe(5);
    // The spacing rides the show operator, after each cluster's last glyph.
    const show = measure.showText('किताब', 10, 2);
    expect(show.match(/-200/gu)).toHaveLength(3);
  });

  it('keeps a syllable whole that a run boundary cuts', async () => {
    // A vowel sign coloured apart from its consonant is still its vowel sign.
    const body =
      '<w:p><w:r><w:t>क</w:t></w:r>' +
      '<w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>ि</w:t></w:r></w:p>';
    const pdf = await pdfOf(body, DEVANAGARI);
    const dotted = hex4(parsed.glyphForCodepoint(0x25cc));
    expect(pdf).not.toContain(dotted);
    expect(pdf).toMatch(shownPattern(shapeIndic(parsed, 'कि')!.gids));
  });
});

describe('the face a word of the scripts of India is drawn in', () => {
  it('draws its join controls too, so the syllable is shaped whole', async () => {
    // Roboto has a ZWJ of its own, and a ka-virama-ZWJ-ssa drawn half in it
    // and half in the Devanagari face came apart into a ka with its virama.
    const roboto = fixture('Roboto-Regular');
    const serve: FetchLike = (url) => {
      const bytes = url.includes('NotoSansDevanagari') ? DEVANAGARI : roboto;
      return Promise.resolve({
        ok: true,
        arrayBuffer: () => Promise.resolve(bytes.slice().buffer),
      });
    };
    clearFontCache();
    const pdf = latin1.decode(
      await Ream.parse(buildDocxFromBody(para('क्‍ष'))).convert('pdf', { fontFetch: serve }),
    );
    clearFontCache();
    const halfKa = shapeIndic(parseTtf(DEVANAGARI), 'क्‍ष')!.gids;
    expect(pdf).toMatch(shownPattern(halfKa));
  });
});

describe('a word of the scripts of India in an SVG', () => {
  it('places a mark off the baseline', async () => {
    const svg = new TextDecoder().decode(
      await Ream.parse(buildDocxFromBody(para('কোথায়'))).convert('svg', {
        fonts: { regular: BENGALI },
      }),
    );
    const dy = shapeIndic(parseTtf(BENGALI), 'কোথায়')!.yOffsets.find((y) => y !== 0)!;
    expect(svg).toContain(`y="${String(dy / 1000)}"`);
  });
});
