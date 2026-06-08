import { describe, expect, it } from 'vitest';

import { deobfuscateEmbeddedFont, parseFontTable } from '@/ooxml/wordproc';

const GUID = '{01014A78-CABC-4EF0-12AC-5CD89AEFDE01}';

describe('embedded font de-obfuscation (ECMA-376 §17.8.1)', () => {
  it('XORs only the first 32 bytes with the reversed fontKey and is self-inverse', () => {
    const plain = new Uint8Array(40).map((_, i) => (i + 7) & 0xff);
    plain[0] = 0x00;
    plain[1] = 0x01;
    plain[2] = 0x00;
    plain[3] = 0x00;
    const obf = deobfuscateEmbeddedFont(plain, GUID); // the XOR IS the obfuscation
    expect([...obf.slice(0, 32)]).not.toEqual([...plain.slice(0, 32)]); // head changed
    expect([...obf.slice(32)]).toEqual([...plain.slice(32)]); // tail untouched
    // XOR is its own inverse → applying it again restores the original.
    expect([...deobfuscateEmbeddedFont(obf, GUID)]).toEqual([...plain]);
  });

  it('restores the TrueType signature from an obfuscated head', () => {
    // The exact obfuscated head of saut_page.docx font1.odttf.
    const obf = new Uint8Array([0x01, 0xdf, 0xef, 0x9a, ...new Array<number>(28).fill(0)]);
    expect([...deobfuscateEmbeddedFont(obf, GUID).slice(0, 4)]).toEqual([0x00, 0x01, 0x00, 0x00]);
  });

  it('leaves data untouched for a non-GUID key', () => {
    const data = new Uint8Array([1, 2, 3, 4, 5]);
    expect([...deobfuscateEmbeddedFont(data, 'not-a-guid')]).toEqual([1, 2, 3, 4, 5]);
  });
});

describe('parseFontTable', () => {
  it('extracts font names and embed refs, skipping non-embedded fonts', () => {
    const xml = new TextEncoder().encode(
      `<w:fonts>` +
        `<w:font w:name="Times New Roman">` +
        `<w:embedRegular r:id="rId1" w:fontKey="{01014A78-CABC-4EF0-12AC-5CD89AEFDE01}"/>` +
        `<w:embedBold r:id="rId2" w:fontKey="{02014A78-CABC-4EF0-12AC-5CD89AEFDE02}"/>` +
        `</w:font>` +
        `<w:font w:name="Calibri"/>` + // no embeds → not returned
        `</w:fonts>`,
    );
    const fonts = parseFontTable(xml);
    expect(fonts).toHaveLength(1);
    expect(fonts[0]!.name).toBe('Times New Roman');
    expect(fonts[0]!.embeds.regular?.rId).toBe('rId1');
    expect(fonts[0]!.embeds.bold?.rId).toBe('rId2');
    expect(fonts[0]!.embeds.italic).toBeUndefined();
  });
});
