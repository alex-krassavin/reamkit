// E-FONT F4/F6 — the writing systems the curated Latin substitutes cannot draw.
// 2145 characters of the corpus are Han, Kana, Hangul, Arabic or a geometric
// symbol, and every one of them was a notdef box: the five families are Latin.
// A face for the script a document actually holds text in is fetched when it is
// needed, and a run reaches it per CHARACTER.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { FetchLike } from '@/core/fonts';
import { Ream } from '@/core/converter/ream';
import { clearFontCache } from '@/core/fonts';
import { scriptForCodepoint, scriptsInFlow } from '@/core/fonts/scripts';
import { buildTrueType } from '@/core/font/ttf-build';

const cp = (s: string): number => s.codePointAt(0)!;

describe('the script a character belongs to', () => {
  it('names the face each writing system needs', () => {
    expect(scriptForCodepoint(cp('ا'))).toBe('arabic');
    expect(scriptForCodepoint(cp('א'))).toBe('hebrew');
    expect(scriptForCodepoint(cp('ก'))).toBe('thai');
    expect(scriptForCodepoint(cp('か'))).toBe('jp');
    expect(scriptForCodepoint(cp('한'))).toBe('kr');
    expect(scriptForCodepoint(cp('☐'))).toBe('symbols2');
    expect(scriptForCodepoint(cp('①'))).toBe('symbols1');
    expect(scriptForCodepoint(cp('क'))).toBe('devanagari');
    expect(scriptForCodepoint(cp('ক'))).toBe('bengali');
    expect(scriptForCodepoint(cp('ਕ'))).toBe('gurmukhi');
    expect(scriptForCodepoint(cp('ક'))).toBe('gujarati');
    expect(scriptForCodepoint(cp('କ'))).toBe('oriya');
    expect(scriptForCodepoint(cp('க'))).toBe('tamil');
    expect(scriptForCodepoint(cp('క'))).toBe('telugu');
    expect(scriptForCodepoint(cp('ಕ'))).toBe('kannada');
    expect(scriptForCodepoint(cp('ക'))).toBe('malayalam');
    expect(scriptForCodepoint(0xa8f2)).toBe('devanagari'); // Devanagari extended
    expect(scriptForCodepoint(0x1cd0)).toBe('devanagari'); // Vedic extensions
    // The dandas end a sentence in several of them, and every face has them.
    expect(scriptForCodepoint(cp('।'))).toBeUndefined();
    expect(scriptForCodepoint(cp('॥'))).toBeUndefined();
    // A Latin family draws these itself.
    expect(scriptForCodepoint(cp('A'))).toBeUndefined();
    expect(scriptForCodepoint(cp('Я'))).toBeUndefined();
  });

  it('lets the DOCUMENT decide which face draws unified Han', () => {
    // The character is the same in all three languages and drawn differently in
    // each; only its neighbours say which.
    expect(scriptForCodepoint(cp('行'))).toBe('sc');
    expect(scriptForCodepoint(cp('行'), 'jp')).toBe('jp');
    const kana = Ream.parse(sheet('行くこと')).flow;
    expect(scriptsInFlow(kana).hanFace).toBe('jp');
    const hangul = Ream.parse(sheet('행 行')).flow;
    expect(scriptsInFlow(hangul).hanFace).toBe('kr');
    const plain = Ream.parse(sheet('行政')).flow;
    expect(scriptsInFlow(plain).hanFace).toBe('sc');
    expect([...scriptsInFlow(plain).scripts]).toEqual(['sc']);
    // A Latin document asks for nothing, and so downloads nothing.
    expect([...scriptsInFlow(Ream.parse(sheet('hello')).flow).scripts]).toEqual([]);
  });
});

const sheet = (text: string): Uint8Array => buildXlsx([[text]]);

describe('a face for a script the substitutes lack', () => {
  // Every download answers with the one face the repo carries: the test is
  // about WHICH faces are asked for, and a real Noto is 200 KB at best.
  const serveRoboto = (asked: Array<string>): FetchLike => {
    const bytes = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
    return (url) => {
      asked.push(url.split('/').pop() ?? url);
      return Promise.resolve({
        ok: true,
        arrayBuffer: () =>
          Promise.resolve(
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
          ),
      });
    };
  };

  it('is fetched when the document holds that script', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('hello مرحبا')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('NotoSansArabic'))).toBe(true);
    // …in the regular weight only: ten megabytes of Noto Sans SC four times
    // over is not a substitute, it is a download.
    expect(asked.filter((f) => f.startsWith('NotoSansArabic'))).toHaveLength(1);
  });

  it('asks for BOTH symbol faces when a document holds one symbol', async () => {
    // The repertory did not fit in one file and Noto's two barely overlap: the
    // ballot box is in the second face, the circled one in the first. Given
    // only the second, ZapfDingbats.pdf drew its stars and boxed its crosses.
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('☐')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('NotoSansSymbols_'))).toBe(true);
    expect(asked.some((f) => f.startsWith('NotoSansSymbols2'))).toBe(true);
  });

  it('is fetched for each script of India a document holds', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('हिन्दी தமிழ் বাংলা')).convert('pdf', { fontFetch: serveRoboto(asked) });
    const noto = asked.filter((f) => f.startsWith('Noto')).sort();
    expect(noto).toEqual([
      'NotoSansBengali_400Regular.ttf',
      'NotoSansDevanagari_400Regular.ttf',
      'NotoSansTamil_400Regular.ttf',
    ]);
  });

  it('is not fetched for Devanagari when a Bengali sentence ends in a danda', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('আমি বাংলায় লিখি।')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.filter((f) => f.startsWith('Noto'))).toEqual(['NotoSansBengali_400Regular.ttf']);
  });

  it('draws the letters, and they are in the text layer', async () => {
    // A face with the six letters of हिन्दी, standing in for Noto Sans
    // Devanagari; everything else is served Roboto, which has none of them.
    const BOX = [
      { op: 'move', x: 50, y: 0 },
      { op: 'line', x: 450, y: 0 },
      { op: 'line', x: 450, y: 700 },
      { op: 'line', x: 50, y: 700 },
      { op: 'close' },
    ] as const;
    const devanagari = buildTrueType({
      family: 'Test Devanagari',
      bold: false,
      italic: false,
      postScriptName: 'TestDevanagari',
      glyphs: [
        { codePoints: [0x20], outline: [], advance: 250 },
        ...[0x0939, 0x093f, 0x0928, 0x094d, 0x0926, 0x0940].map((c) => ({
          codePoints: [c],
          outline: BOX,
          advance: 500,
        })),
      ],
      ascent: 800,
      descent: -200,
      fsType: 0,
    });
    const roboto = serveRoboto([]);
    const served: FetchLike = (url) =>
      url.includes('NotoSansDevanagari')
        ? Promise.resolve({
            ok: true,
            arrayBuffer: () => Promise.resolve(devanagari.slice().buffer),
          })
        : roboto(url);
    clearFontCache();
    const { bytes, losses } = await Ream.parse(sheet('Hindi: हिन्दी')).convertWithReport('pdf', {
      fontFetch: served,
    });
    clearFontCache();
    expect(losses.filter((l) => l.feature === 'text')).toEqual([]);
    expect(new TextDecoder('latin1').decode(bytes)).toContain('+TestDevanagari');
    const text = new TextDecoder().decode(await Ream.parse(bytes).convert('md'));
    expect(text).toContain('हिन्दी');
  });

  it('…and not when it does not', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('hello')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('Noto'))).toBe(false);
  });
});
