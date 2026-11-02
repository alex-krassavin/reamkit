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
const KANNADA = fixture('NotoSansKannada-shaping');

const latin1 = new TextDecoder('latin1');
const hex4 = (gid: number): string => gid.toString(16).padStart(4, '0').toUpperCase();
const ucs = (text: string): string => [...text].map((c) => hex4(c.codePointAt(0)!)).join('');
const utf16 = (text: string): string => 'FEFF' + ucs(text);

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

  /** The /ToUnicode CMap of the PDF of a paragraph, set in one face. */
  async function toUnicode(text: string, regular: Uint8Array): Promise<string> {
    const pdf = await pdfBytes(para(text), regular);
    return inflated(pdf).find((s) => s.includes('beginbfchar'))!;
  }

  it('names each glyph only shaping reaches by the characters it was made from', async () => {
    // पार्श्व: the conjunct श्व and the reph drawn over it are one cluster,
    // and named each after the whole of it, a reader that does not read
    // /ActualText read पार्श्वर्श्व.
    const cmap = await toUnicode('पार्श्व पक्षों', DEVANAGARI);
    const [, , conjunct, reph] = shapeIndic(parsed, 'पार्श्व')!.gids;
    expect(cmap).toContain(`<${hex4(conjunct!)}> <${ucs('श्व')}>`);
    expect(cmap).toContain(`<${hex4(reph!)}> <${ucs('र्')}>`);
    // पक्षों: the face draws the o as the aa's glyph, which the cmap names
    // aa, and its top in the anusvara's. Ksha stands for its own letters, and
    // the anusvara for the o and itself, so that the o is not lost.
    const [, ksha, aa, anusvara] = shapeIndic(parsed, 'पक्षों')!.gids;
    expect(parsed.glyphForCodepoint(0x093e)).toBe(aa);
    expect(cmap).toContain(`<${hex4(ksha!)}> <${ucs('क्ष')}>`);
    expect(cmap).toContain(`<${hex4(anusvara!)}> <${ucs('ों')}>`);
  });

  it('names a glyph of half a vowel sign by that half', async () => {
    // ಕೊ and ಕೈ draw ka with the first half of their vowel sign as one glyph,
    // and the second halves are characters of their own. Named after the
    // first word's ಕೊ, the glyph read ಕೊ in ಕೈ too, and the o twice.
    const kannada = parseTtf(KANNADA);
    const [kaE, uu] = shapeIndic(kannada, 'ಕೊ')!.gids;
    expect(shapeIndic(kannada, 'ಕೈ')!.gids[0]).toBe(kaE);
    expect(kannada.glyphForCodepoint(0x0cc2)).toBe(uu);
    expect(await toUnicode('ಕೊ ಕೈ', KANNADA)).toContain(`<${hex4(kaE!)}> <${ucs('ಕೆ')}>`);
    // কো: the e half is drawn before ka in a form of its own.
    const e = shapeIndic(parseTtf(BENGALI), 'কো')!.gids[0]!;
    expect(await toUnicode('কো', BENGALI)).toContain(`<${hex4(e)}> <${ucs('ে')}>`);
  });

  it('names a glyph a face moved by swapping two by what it stands for', async () => {
    // Noto Sans Bengali sets a reph after the ii by swapping their glyphs:
    // the reph's place takes the ii, the ii's a form of the reph.
    const bengali = parseTtf(BENGALI);
    const [ksha, ii, reph] = shapeIndic(bengali, 'র্ক্ষী')!.gids;
    expect(bengali.glyphForCodepoint(0x09c0)).toBe(ii);
    const cmap = await toUnicode('র্ক্ষী', BENGALI);
    expect(cmap).toContain(`<${hex4(ksha!)}> <${ucs('ক্ষ')}>`);
    expect(cmap).toContain(`<${hex4(reph!)}> <${ucs('র্')}>`);
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
    // In the shaper's order, the short i first: the colour switches between
    // the two glyphs, not the syllable between two shows.
    const [shortI, ka] = shapeIndic(parsed, 'कि')!.gids.map((g) => hex4(g));
    expect(pdf).toMatch(new RegExp(`<${shortI}> Tj [\\d. ]+ rg <${ka}>`, 'u'));
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

describe('a syllable a run boundary cuts, in its runs’ colours', () => {
  // Cut to क ि ृ ष ् ण र and the dotted circle (see fixtures/fonts/NOTICE.md).
  const COLOUR = fixture('NotoSansDevanagari-colour');
  const parsed = parseTtf(COLOUR);
  const run = (text: string, colour?: string): string =>
    `<w:r>${colour ? `<w:rPr><w:color w:val="${colour}"/></w:rPr>` : ''}<w:t>${text}</w:t></w:r>`;
  // कृष्ण with each sign in a colour of its own, as font issues.docx has it:
  // the half form of ष् is one glyph made of a green ष and a blue virama.
  const krishna =
    `<w:p>${run('क')}${run('ृ', 'FF0000')}${run('ष', '00FF00')}` +
    `${run('्', '0000FF')}${run('ण')}</w:p>`;
  const svgOf = async (body: string): Promise<string> =>
    new TextDecoder().decode(
      await Ream.parse(buildDocxFromBody(body)).convert('svg', { fonts: { regular: COLOUR } }),
    );

  it('names, per glyph, the first character it stands for', () => {
    const measure = createFontMeasure(parsed);
    // The short i is drawn first and stands for the second character.
    expect(measure.glyphSources('कि')).toEqual([1, 0]);
    expect(measure.glyphSources('ki')).toBeUndefined();
  });

  it('switches the fill colour around the glyph of a vowel sign coloured apart', async () => {
    const pdf = await pdfOf(`<w:p>${run('क')}${run('ि', 'FF0000')}</w:p>`, COLOUR);
    const [shortI, ka] = shapeIndic(parsed, 'कि')!.gids;
    const word = pdf.slice(pdf.indexOf('/ActualText'), pdf.indexOf('EMC'));
    // Inside the word's one bracket: the short i red, the ka black again.
    expect(word).toMatch(
      new RegExp(`BDC\\s+1 0 0 rg <${hex4(shortI!)}> Tj 0 0 0 rg <${hex4(ka!)}> Tj\\s*$`, 'u'),
    );
    expect(pdf).not.toContain(hex4(parsed.glyphForCodepoint(0x25cc)));
  });

  it('paints each glyph of कृष्ण in the colour of its first character', async () => {
    const pdf = await pdfOf(krishna, COLOUR);
    const [ka, r, ssa, nna] = shapeIndic(parsed, 'कृष्ण')!.gids.map((g) => hex4(g));
    expect(pdf.match(/BDC/gu)).toHaveLength(1);
    expect(pdf).toMatch(
      new RegExp(
        `0 0 0 rg[^]*?<${ka}[^]*?1 0 0 rg[^]*?<${r}>[^]*?0 1 0 rg[^]*?<${ssa}>` +
          `[^]*?0 0 0 rg[^]*?<${nna}>[^]*?EMC`,
        'u',
      ),
    );
    // The virama's blue has no glyph of its own to paint.
    expect(pdf).not.toContain('0 0 1 rg');
  });

  it('fills each glyph of an SVG in the colour of its first character', async () => {
    const svg = await svgOf(krishna);
    const [ka, r, ssa, nna] = shapeIndic(parsed, 'कृष्ण')!.gids;
    const use = (gid: number, fill?: string): RegExp =>
      new RegExp(`<use href="#g\\d+-${gid}"[^>]*?${fill ? ` fill="#${fill}"` : ''}/>`, 'u');
    // कृ in a black group, its vowel sign red; ष्ण in a green one, its ण black.
    expect(svg).toMatch(/fill="#000000"><use href="#g\d+-\d+"\/>/u);
    expect(svg).toMatch(use(ka!));
    expect(svg).toMatch(use(r!, 'FF0000'));
    expect(svg).toMatch(new RegExp(`fill="#00FF00"><use href="#g\\d+-${ssa}"/>`, 'u'));
    expect(svg).toMatch(use(nna!, '000000'));
    expect(svg).not.toContain('0000FF');
  });

  it('strokes a faked bold in the colour of each glyph, the weight being its own', async () => {
    // A green bold ष् and a black ण: the face has no bold, so the page fakes it.
    const bold = '<w:rPr><w:b/><w:color w:val="00FF00"/></w:rPr>';
    const svg = await svgOf(`<w:p><w:r>${bold}<w:t>ष्</w:t></w:r>${run('ण')}</w:p>`);
    const [ssa, nna] = shapeIndic(parsed, 'ष्ण')!.gids;
    expect(svg).toMatch(
      new RegExp(
        `fill="#00FF00" stroke="#00FF00"[^>]*><use href="#g\\d+-${ssa}"/>` +
          `<use href="#g\\d+-${nna}"[^>]*? fill="#000000" stroke="#000000"/>`,
        'u',
      ),
    );
  });
});
