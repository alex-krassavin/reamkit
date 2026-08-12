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

  it('…and not when it does not', async () => {
    clearFontCache();
    const asked: Array<string> = [];
    await Ream.parse(sheet('hello')).convert('pdf', { fontFetch: serveRoboto(asked) });
    expect(asked.some((f) => f.startsWith('Noto'))).toBe(false);
  });
});
