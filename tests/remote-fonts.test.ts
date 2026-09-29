import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { FetchLike } from '@/core/fonts/remote-fonts';
import { convertDocxToPdf } from '@/core/converter';
import { Ream } from '@/core/converter/ream';
import { clearFontCache, fetchFontSet, resolveFamilyKey } from '@/core/fonts/remote-fonts';

const here = dirname(fileURLToPath(import.meta.url));
const ROBOTO = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf')));
const ROBOTO_BOLD = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf')));
const latin1 = new TextDecoder('latin1');

// A fake fetch that serves local Roboto bytes for any URL, recording the URLs
// requested so we can assert the family/variant resolution without a network.
// `failUrls` are answered with an HTTP error; `rejectUrls` are never answered:
// the request rejects, as fetch does offline or when a DNS lookup fails.
function fakeFetch(
  record: Array<string>,
  opts: { failUrls?: RegExp; rejectUrls?: RegExp } = {},
): FetchLike {
  return async (url: string) => {
    record.push(url);
    if (opts.rejectUrls && opts.rejectUrls.test(url)) throw new TypeError('fetch failed');
    if (opts.failUrls && opts.failUrls.test(url)) {
      return { ok: false, arrayBuffer: async () => new ArrayBuffer(0) };
    }
    const bytes = url.includes('700Bold') ? ROBOTO_BOLD : ROBOTO;
    return {
      ok: true,
      arrayBuffer: async () =>
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    };
  };
}

describe('resolveFamilyKey (font substitution)', () => {
  it('maps proprietary fonts to their metric-compatible twins', () => {
    expect(resolveFamilyKey('Calibri')).toBe('carlito');
    expect(resolveFamilyKey('Cambria')).toBe('caladea');
    expect(resolveFamilyKey('Arial')).toBe('arimo');
    expect(resolveFamilyKey('Times New Roman')).toBe('tinos');
    expect(resolveFamilyKey('Courier New')).toBe('cousine');
    // The heading font of a default Word theme is Calibri's light weight.
    expect(resolveFamilyKey('Calibri Light')).toBe('carlito');
  });
  it('falls back by class for families without an exact twin', () => {
    expect(resolveFamilyKey('Georgia')).toBe('tinos'); // generic serif
    expect(resolveFamilyKey('Consolas')).toBe('cousine'); // generic mono
    expect(resolveFamilyKey('Verdana')).toBe('arimo'); // generic sans
  });
  it('defaults unknown / unnamed families to sans (Arimo)', () => {
    expect(resolveFamilyKey(undefined)).toBe('arimo');
    expect(resolveFamilyKey('Totally Unknown Font')).toBe('arimo');
  });
});

describe('fetchFontSet (injected fetch — no network)', () => {
  it('downloads the regular + bold/italic variants for a family', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const set = await fetchFontSet({ family: 'Times New Roman', fetch: fakeFetch(urls) });
    expect(set.regular).toBeInstanceOf(Uint8Array);
    // Requested the Tinos package for the serif family.
    expect(urls.every((u) => u.includes('/tinos/'))).toBe(true);
    expect(urls.some((u) => u.includes('400Regular'))).toBe(true);
    expect(urls.some((u) => u.includes('700Bold'))).toBe(true);
  });

  it('builds the nested CDN path for Calibri → Carlito', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const set = await fetchFontSet({ family: 'Calibri', fetch: fakeFetch(urls) });
    expect(set.regular).toBeInstanceOf(Uint8Array);
    // Calibri resolves to its metric twin Carlito…
    expect(urls.every((u) => u.includes('/carlito/'))).toBe(true);
    // …whose @expo-google-fonts package nests each variant in a suffix folder.
    expect(urls.some((u) => u.includes('/carlito/400Regular/Carlito_400Regular.ttf'))).toBe(true);
  });

  it('still resolves when optional variants 404 (regular only)', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const set = await fetchFontSet({
      family: 'Arial',
      fetch: fakeFetch(urls, { failUrls: /(700Bold|Italic)/ }),
    });
    expect(set.regular).toBeInstanceOf(Uint8Array);
    expect(set.bold).toBeUndefined();
    expect(set.italic).toBeUndefined();
  });

  it('throws a clear error when the regular face cannot be downloaded', async () => {
    clearFontCache();
    await expect(
      fetchFontSet({ family: 'Arial', fetch: fakeFetch([], { failUrls: /.*/ }) }),
    ).rejects.toThrow(/Failed to download font/);
  });
});

describe('convertDocxToPdf (async, auto font via injected fetch)', () => {
  it('downloads a font and produces a PDF without caller-supplied fonts', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const docx = buildDocxFromBody('<w:p><w:r><w:t>Auto font hello</w:t></w:r></w:p>');
    const pdf = await convertDocxToPdf(docx, { fontFetch: fakeFetch(urls) });
    expect(latin1.decode(pdf.subarray(0, 8))).toBe('%PDF-1.7');
    expect(pdf.byteLength).toBeGreaterThan(1000);
    // It fetched at least the regular face.
    expect(urls.some((u) => u.includes('400Regular'))).toBe(true);
  });

  it('honours an explicit fontFamily override for substitution', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const docx = buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
    await convertDocxToPdf(docx, { fontFamily: 'Courier New', fontFetch: fakeFetch(urls) });
    expect(urls.every((u) => u.includes('/cousine/'))).toBe(true);
  });
});

describe('a request that rejects (offline, a failed DNS lookup, a dropped connection)', () => {
  it('is asked again by the next call, not answered with the same error for good', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const offline = fetchFontSet({ family: 'Arial', fetch: fakeFetch(urls, { rejectUrls: /./ }) });
    await expect(offline).rejects.toThrow(/Failed to download font from .*Arimo_400Regular\.ttf/);
    // The network's own error stays readable underneath.
    await expect(offline).rejects.toHaveProperty('cause.message', 'fetch failed');
    // Back online, the next call downloads the face instead of re-throwing the blip.
    const set = await fetchFontSet({ family: 'Arial', fetch: fakeFetch(urls) });
    expect(set.regular).toBeInstanceOf(Uint8Array);
    expect(urls.filter((u) => u.endsWith('Arimo_400Regular.ttf'))).toHaveLength(2);
  });

  it('costs a set only the best-effort faces it lost, as an HTTP error does', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const set = await fetchFontSet({
      family: 'Arial',
      fetch: fakeFetch(urls, { rejectUrls: /Bold/ }),
    });
    expect(set.regular).toBeInstanceOf(Uint8Array);
    expect(set.italic).toBeInstanceOf(Uint8Array);
    expect(set.bold).toBeUndefined();
    expect(set.boldItalic).toBeUndefined();
    // …and the next call asks for them again.
    const again = await fetchFontSet({ family: 'Arial', fetch: fakeFetch(urls) });
    expect(again.bold).toBeInstanceOf(Uint8Array);
    expect(again.boldItalic).toBeInstanceOf(Uint8Array);
  });

  it("costs a document a script's face as a font loss, not a throw", async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const docx = buildDocxFromBody('<w:p><w:r><w:t>hello مرحبا</w:t></w:r></w:p>');
    const { bytes, losses } = await Ream.parse(docx).convertWithReport('pdf', {
      fontFetch: fakeFetch(urls, { rejectUrls: /NotoSansArabic/ }),
    });
    expect(latin1.decode(bytes.subarray(0, 8))).toBe('%PDF-1.7');
    expect(urls.some((u) => u.includes('NotoSansArabic'))).toBe(true);
    expect(losses).toContainEqual(
      expect.objectContaining({ feature: 'font', detail: expect.stringContaining('arabic') }),
    );
  });

  it('is shared by every call waiting on it, as any download in flight is', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const offline = fakeFetch(urls, { rejectUrls: /./ });
    const lost = await Promise.allSettled([
      fetchFontSet({ family: 'Arial', fetch: offline }),
      fetchFontSet({ family: 'Arial', fetch: offline }),
    ]);
    expect(lost.map((r) => r.status)).toEqual(['rejected', 'rejected']);
    expect(urls).toHaveLength(1);
    // Asked again, the four faces are each downloaded once for both calls.
    const online = fakeFetch(urls);
    await Promise.all([
      fetchFontSet({ family: 'Arial', fetch: online }),
      fetchFontSet({ family: 'Arial', fetch: online }),
    ]);
    expect(urls).toHaveLength(5);
  });
});
