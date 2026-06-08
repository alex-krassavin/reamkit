import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { FetchLike } from '@/fonts/remote-fonts';
import { convertDocxToPdf } from '@/converter';
import { clearFontCache, fetchFontSet, resolveFamilyKey } from '@/fonts/remote-fonts';

const here = dirname(fileURLToPath(import.meta.url));
const ROBOTO = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf')));
const ROBOTO_BOLD = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf')));
const latin1 = new TextDecoder('latin1');

// A fake fetch that serves local Roboto bytes for any URL, recording the URLs
// requested so we can assert the family/variant resolution without a network.
function fakeFetch(record: Array<string>, opts: { failUrls?: RegExp } = {}): FetchLike {
  return async (url: string) => {
    record.push(url);
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
  it('maps serif families to Tinos', () => {
    expect(resolveFamilyKey('Times New Roman')).toBe('tinos');
    expect(resolveFamilyKey('Georgia')).toBe('tinos');
    expect(resolveFamilyKey('Cambria')).toBe('tinos');
  });
  it('maps monospace families to Cousine', () => {
    expect(resolveFamilyKey('Courier New')).toBe('cousine');
    expect(resolveFamilyKey('Consolas')).toBe('cousine');
  });
  it('defaults everything else (incl. Calibri/Arial) to Roboto', () => {
    expect(resolveFamilyKey('Calibri')).toBe('roboto');
    expect(resolveFamilyKey('Arial')).toBe('roboto');
    expect(resolveFamilyKey(undefined)).toBe('roboto');
    expect(resolveFamilyKey('Totally Unknown Font')).toBe('roboto');
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
