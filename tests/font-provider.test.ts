import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { FontAnswer } from '@/core/fonts/provider';
import { ConversionLossError } from '@/core/ir';
import { createConverter } from '@/core/converter/facade';
import {
  NO_FONT,
  callerFontProvider,
  chainProviders,
  isEmbeddingRestricted,
  localFontProvider,
  readOs2FsType,
  remoteFontProvider,
} from '@/core/fonts/provider';
import { clearFontCache } from '@/core/fonts/remote-fonts';

const ROBOTO = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
const ROBOTO_BOLD = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf'));

describe('chainProviders', () => {
  it('first byte answer wins; none falls through', async () => {
    const never = { id: 'never', resolve: () => Promise.resolve(NO_FONT) };
    const always = callerFontProvider({ regular: ROBOTO });
    const chain = chainProviders([never, always]);
    const a = await chain.resolve({ bold: false, italic: false });
    expect(a.kind).toBe('bytes');
    if (a.kind !== 'bytes') throw new Error('unreachable');
    expect(a.providerId).toBe('caller');
  });

  it('empty chain answers none', async () => {
    const a = await chainProviders([]).resolve({ bold: false, italic: false });
    expect(a.kind).toBe('none');
  });
});

describe('callerFontProvider', () => {
  it('picks the variant and falls back to regular', async () => {
    const p = callerFontProvider({ regular: ROBOTO, bold: ROBOTO_BOLD });
    const bold = await p.resolve({ bold: true, italic: false });
    const italic = await p.resolve({ bold: false, italic: true });
    if (bold.kind !== 'bytes' || italic.kind !== 'bytes') throw new Error('unreachable');
    expect(bold.bytes).toBe(ROBOTO_BOLD);
    expect(italic.bytes).toBe(ROBOTO); // no italic supplied → regular
  });
});

describe('OS/2 fsType licensing gate', () => {
  it('reads fsType from a real font (Roboto is installable: 0 or preview/editable)', () => {
    const fsType = readOs2FsType(ROBOTO);
    expect(fsType).toBeDefined();
    expect(isEmbeddingRestricted(fsType)).toBe(false);
  });

  it('flags a restricted-licensing font', () => {
    // Synthetic sfnt: one OS/2 table whose fsType (offset+8) = 0x0002.
    const table = new Uint8Array(12);
    table[8] = 0x00;
    table[9] = 0x02;
    const header = new Uint8Array(12 + 16 + table.length);
    header[4] = 0; // numTables hi
    header[5] = 1; // numTables lo
    header.set(new TextEncoder().encode('OS/2'), 12);
    const offset = 12 + 16;
    header[12 + 8] = 0;
    header[12 + 9] = 0;
    header[12 + 10] = (offset >> 8) & 0xff;
    header[12 + 11] = offset & 0xff;
    header.set(table, offset);
    expect(readOs2FsType(header)).toBe(2);
    expect(isEmbeddingRestricted(2)).toBe(true);
    expect(isEmbeddingRestricted(0)).toBe(false);
    expect(isEmbeddingRestricted(4)).toBe(false); // preview & print → allowed (subset)
    expect(isEmbeddingRestricted(8)).toBe(false); // editable → allowed
    expect(isEmbeddingRestricted(undefined)).toBe(false);
  });
});

describe('remoteFontProvider fallback cascade', () => {
  it('degrades boldItalic to bold when the CDN set is partial (A4 fix)', async () => {
    // fetchFontSet is best-effort: italic faces fail here, bold succeeds.
    const fakeFetch = (url: string) => {
      if (url.includes('Italic')) {
        return Promise.resolve({
          ok: false,
          status: 404,
          arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)),
        });
      }
      const bytes = url.includes('Bold') ? ROBOTO_BOLD : ROBOTO;
      return Promise.resolve({
        ok: true,
        status: 200,
        arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)),
      });
    };
    const p = remoteFontProvider({ fetch: fakeFetch });
    const a = await p.resolve({ bold: true, italic: true });
    if (a.kind !== 'bytes') throw new Error('unreachable');
    // Pre-fix this fell straight through to regular.
    expect(Buffer.from(a.bytes).equals(Buffer.from(ROBOTO_BOLD))).toBe(true);
  });
});

describe('remoteFontProvider, after a set the network lost', () => {
  // Serves the Roboto fixtures, except that the first `lost` requests reject,
  // as fetch does offline or when a DNS lookup fails. Records every URL asked.
  const flakyFetch = (urls: Array<string>, lost: number) => (url: string) => {
    urls.push(url);
    if (urls.length <= lost) return Promise.reject(new TypeError('fetch failed'));
    const bytes = url.includes('Bold') ? ROBOTO_BOLD : ROBOTO;
    return Promise.resolve({ ok: true, arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)) });
  };

  it('asks again on the next request, not answering it with the same error for good', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    // One provider for the life of the process, as a server builds its chain.
    const provider = remoteFontProvider({ fetch: flakyFetch(urls, 1) });
    const regular = { family: 'Arial', bold: false, italic: false };
    // The regular face is the one a set cannot do without, so the loss is thrown.
    await expect(provider.resolve(regular)).rejects.toThrow(
      /Failed to download font from .*Arimo_400Regular\.ttf/,
    );
    const a = await provider.resolve(regular);
    if (a.kind !== 'bytes') throw new Error('unreachable');
    expect(Buffer.from(a.bytes).equals(Buffer.from(ROBOTO))).toBe(true);
    expect(urls.filter((u) => u.endsWith('Arimo_400Regular.ttf'))).toHaveLength(2);
  });

  it('is one download for the requests made while it is in flight, lost or not', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const provider = remoteFontProvider({ fetch: flakyFetch(urls, 1) });
    // The four faces a conversion asks the chain for at once.
    const four = () =>
      Promise.allSettled(
        [false, true].flatMap((bold) =>
          [false, true].map((italic) => provider.resolve({ family: 'Arial', bold, italic })),
        ),
      );
    const errors = (await four()).flatMap((r) => (r.status === 'rejected' ? [r.reason] : []));
    expect(errors).toHaveLength(4);
    // One request, lost once, and one error for all four: each would have
    // thrown its own had the provider started a set per request.
    expect(urls).toHaveLength(1);
    expect(new Set(errors).size).toBe(1);
    // Asked again, the set is downloaded anew: each face once, for all four.
    const found = await four();
    expect(found.map((r) => r.status)).toEqual([
      'fulfilled',
      'fulfilled',
      'fulfilled',
      'fulfilled',
    ]);
    expect(urls).toHaveLength(5);
  });
});

describe('remoteFontProvider, after a face the network lost', () => {
  // Serves Roboto Bold for a bold face and Roboto for any other, and records
  // every URL asked. The first request for a face `lose` matches rejects, as
  // fetch does when the network blips; a face `refuse` matches is answered with
  // an HTTP error, as the CDN answers for a file it does not have.
  const faceFetch =
    (urls: Array<string>, faces: { lose?: RegExp; refuse?: RegExp }) => (url: string) => {
      const again = urls.includes(url);
      urls.push(url);
      if (faces.lose?.test(url) && !again) return Promise.reject(new TypeError('fetch failed'));
      if (faces.refuse?.test(url)) {
        return Promise.resolve({
          ok: false,
          arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)),
        });
      }
      const bytes = url.includes('Bold') ? ROBOTO_BOLD : ROBOTO;
      return Promise.resolve({
        ok: true,
        arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)),
      });
    };
  const LOSE_BOLD = /Arimo_700Bold\.ttf$/;
  // The weight an answer is set in, told by the fixture it carries.
  const weight = (answer: FontAnswer): 'bold' | 'regular' => {
    if (answer.kind !== 'bytes') throw new Error('unreachable');
    return Buffer.from(answer.bytes).equals(Buffer.from(ROBOTO_BOLD)) ? 'bold' : 'regular';
  };
  const arial = (bold: boolean, italic: boolean) => ({ family: 'Arial', bold, italic });

  it('asks for it again on the next request, not answering bold with regular for good', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    // One provider for the life of the process, as a server builds its chain.
    const provider = remoteFontProvider({ fetch: faceFetch(urls, { lose: LOSE_BOLD }) });
    // The set came without its bold face, so bold falls to regular…
    expect(weight(await provider.resolve(arial(true, false)))).toBe('regular');
    // …this once: the next request asks for the face again.
    expect(weight(await provider.resolve(arial(true, false)))).toBe('bold');
    expect(urls.filter((u) => u.endsWith('Arimo_700Bold.ttf'))).toHaveLength(2);
    // The three faces that came are not asked for again.
    expect(urls).toHaveLength(5);
  });

  it('asks for it once for the requests made at once, and keeps the set once whole', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    const provider = remoteFontProvider({ fetch: faceFetch(urls, { lose: LOSE_BOLD }) });
    // The four faces a conversion asks the chain for at once (regular, bold,
    // italic, bold italic), and the weight each is answered in.
    const four = async () =>
      (
        await Promise.all([
          provider.resolve(arial(false, false)),
          provider.resolve(arial(true, false)),
          provider.resolve(arial(false, true)),
          provider.resolve(arial(true, true)),
        ])
      ).map(weight);
    // A request per face, and bold, lost, falls to regular.
    expect(await four()).toEqual(['regular', 'regular', 'regular', 'bold']);
    expect(urls).toHaveLength(4);
    // One more request for the next four: the lost face, asked for again.
    expect(await four()).toEqual(['regular', 'bold', 'regular', 'bold']);
    expect(urls).toHaveLength(5);
    // Whole, the set is the provider's own: with the download cache beneath it
    // emptied, four more requests are answered without a download.
    clearFontCache();
    expect(await four()).toEqual(['regular', 'bold', 'regular', 'bold']);
    expect(urls).toHaveLength(5);
  });

  it('does not ask again for a face the CDN answered with an error', async () => {
    clearFontCache();
    const urls: Array<string> = [];
    // The CDN has no italic faces, and the network loses the bold one once.
    const provider = remoteFontProvider({
      fetch: faceFetch(urls, { lose: LOSE_BOLD, refuse: /Italic/ }),
    });
    // Bold italic falls to bold, and to regular while bold is lost.
    expect(weight(await provider.resolve(arial(true, true)))).toBe('regular');
    expect(weight(await provider.resolve(arial(true, true)))).toBe('bold');
    expect(weight(await provider.resolve(arial(true, true)))).toBe('bold');
    // The set is never whole, yet each face the CDN refused was asked for once:
    // the download cache remembers it as missing. Only the lost face was asked
    // for again.
    expect(urls.filter((u) => u.includes('Italic'))).toHaveLength(2);
    expect(urls).toHaveLength(5);
  });
});

describe('localFontProvider', () => {
  it('answers none where Local Font Access is unavailable (Node)', async () => {
    const a = await localFontProvider().resolve({ family: 'Arial', bold: false, italic: false });
    expect(a.kind).toBe('none');
  });
});

describe('facade × font chain', () => {
  const DOCX = buildDocxFromBody('<w:p><w:r><w:t>chain</w:t></w:r></w:p>');

  it('caller provider in the chain → no losses', async () => {
    const ream = createConverter();
    const r = await ream.convert(DOCX, {
      fontProviders: [callerFontProvider({ regular: ROBOTO, bold: ROBOTO_BOLD })],
    });
    expect(r.losses).toEqual([]);
    expect(r.bytes.length).toBeGreaterThan(0);
  });

  it('remote winner → substituted loss; strict throws', async () => {
    // Inject a fetch that serves the local Roboto fixtures (no network).
    const fakeFetch = (url: string) => {
      const bytes = url.includes('Bold') ? ROBOTO_BOLD : ROBOTO;
      return Promise.resolve({
        ok: true,
        status: 200,
        arrayBuffer: () => Promise.resolve(bytes.buffer.slice(0)),
      });
    };
    const providers = [
      localFontProvider(), // none in Node → falls through
      remoteFontProvider({ fetch: fakeFetch }),
    ];

    const ream = createConverter();
    const r = await ream.convert(DOCX, { fontProviders: providers });
    expect(r.losses).toHaveLength(1);
    expect(r.losses[0]!.severity).toBe('substituted');
    expect(r.losses[0]!.feature).toBe('fonts.substitution');

    await expect(ream.convert(DOCX, { fontProviders: providers, strict: true })).rejects.toThrow(
      ConversionLossError,
    );
  });
});
