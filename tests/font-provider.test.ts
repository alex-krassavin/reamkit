import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
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
