import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { parseTtf, shapeText } from '@/font';

const here = dirname(fileURLToPath(import.meta.url));

function loadFont(name: string) {
  const bytes = new Uint8Array(readFileSync(resolve(here, `fixtures/fonts/${name}`)));
  return parseTtf(bytes);
}

describe('OpenType layout: GSUB ligatures', () => {
  it('produces a non-empty ligature map for Roboto-Regular', () => {
    const font = loadFont('Roboto-Regular.ttf');
    expect(font.ligatures.size).toBeGreaterThan(0);
  });

  it('substitutes the "fi" sequence with a single ligature GID', () => {
    const font = loadFont('Roboto-Regular.ttf');
    const fGid = font.glyphForCodepoint('f'.codePointAt(0)!);
    const iGid = font.glyphForCodepoint('i'.codePointAt(0)!);
    const ligGid = font.ligatures.get(`${fGid},${iGid}`);
    expect(ligGid).toBeDefined();
    // "fi" → one glyph (ligature), so shape returns a single GID.
    const shaped = shapeText(
      'fi',
      font.glyphForCodepoint,
      font.advanceWidths,
      font.ligatures,
      font.kerning,
    );
    expect(shaped.gids).toHaveLength(1);
    expect(shaped.gids[0]).toBe(ligGid);
  });

  it('leaves text without ligatable pairs untouched', () => {
    const font = loadFont('Roboto-Regular.ttf');
    const shaped = shapeText(
      'xyz',
      font.glyphForCodepoint,
      font.advanceWidths,
      font.ligatures,
      font.kerning,
    );
    expect(shaped.gids).toHaveLength(3);
  });
});

describe('OpenType layout: GPOS kerning', () => {
  it('parses kerning pairs from Roboto-Regular', () => {
    const font = loadFont('Roboto-Regular.ttf');
    expect(font.kerning.size).toBeGreaterThan(0);
  });

  it('adjusts the advance of the previous glyph for a kerned pair', () => {
    const font = loadFont('Roboto-Regular.ttf');
    if (font.kerning.size === 0) return; // font has no kerning — nothing to test

    const [key, adj] = font.kerning.entries().next().value as [string, number];
    const [gid1, gid2] = key.split(',').map(Number) as [number, number];
    const baseline1 = font.advanceWidths[gid1] ?? 0;
    const baseline2 = font.advanceWidths[gid2] ?? 0;

    // Build a synthetic two-codepoint input where each char maps to its
    // intended GID. We disable ligatures to isolate the kerning path.
    const customShape = shapeText(
      '',
      (cp) => (cp === 1 ? gid1 : gid2),
      font.advanceWidths,
      new Map(),
      font.kerning,
    );
    expect(customShape.gids).toEqual([gid1, gid2]);
    expect(customShape.advances[0]).toBe(baseline1 + adj);
    expect(customShape.advances[1]).toBe(baseline2);
  });

  it('passes the original advance for a non-kerned pair', () => {
    const font = loadFont('Roboto-Regular.ttf');
    // Two glyphs with no kerning entry — choose space pair.
    const spaceGid = font.glyphForCodepoint(0x20);
    const shaped = shapeText(
      '  ',
      font.glyphForCodepoint,
      font.advanceWidths,
      new Map(),
      font.kerning,
    );
    expect(shaped.gids).toEqual([spaceGid, spaceGid]);
    const baseline = font.advanceWidths[spaceGid] ?? 0;
    expect(shaped.advances[0]).toBe(baseline);
    expect(shaped.advances[1]).toBe(baseline);
  });
});
