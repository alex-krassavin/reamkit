// E-FONT F5 — §19.2.1.13 `p:embeddedFontLst`: the faces a deck brings with it.
//
// Word and PowerPoint wrap an embedded font differently, and the difference is
// the whole story. A .docx part is the font with 32 obfuscated bytes; a .pptx
// `fntdata` part is an EOT container, and all eleven in the corpus set
// TTEMBED_TTCOMPRESSED — the font inside is packed with MicroType Express, a
// codec of its own. So a plain container is used and a packed one is a loss.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { readFntData } from '@/pptx/embedded-fonts';

const ROBOTO = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));

/** An EOT v1 container around `font`, with the flags a caller wants tested. */
function eot(font: Uint8Array, flags: number): Uint8Array {
  const names = 4 * (2 + 2); // four empty length-prefixed names, each padded
  const total = 82 + names + font.length;
  const out = new Uint8Array(total);
  const v = new DataView(out.buffer);
  v.setUint32(0, total, true); // EOTSize
  v.setUint32(4, font.length, true); // FontDataSize
  v.setUint32(8, 0x00010000, true); // Version
  v.setUint32(12, flags, true);
  v.setUint16(34, 0x504c, true); // MagicNumber
  out.set(font, 82 + names);
  return out;
}

describe('a .pptx embedded font part', () => {
  it('unwraps a container that stores the font plainly', () => {
    const face = readFntData(eot(ROBOTO, 0));
    expect(face.kind).toBe('font');
    expect(face.kind === 'font' ? [...face.bytes.subarray(0, 4)] : []).toEqual([0, 1, 0, 0]);
  });

  it('says so when the font is packed, instead of drawing nothing', () => {
    // TTEMBED_TTCOMPRESSED — every corpus part sets it.
    expect(readFntData(eot(ROBOTO, 0x04)).kind).toBe('compressed');
    // …and TTEMBED_XORENCRYPTDATA is no more readable.
    expect(readFntData(eot(ROBOTO, 0x10000000)).kind).toBe('compressed');
  });

  it('takes a part that is the bare font', () => {
    // Nothing in the spec forbids it, and a reader that insists on the wrapper
    // would drop one.
    expect(readFntData(ROBOTO).kind).toBe('font');
  });

  it('refuses bytes that are neither', () => {
    expect(readFntData(new Uint8Array(200)).kind).toBe('unreadable');
    expect(readFntData(new Uint8Array(4)).kind).toBe('unreadable');
  });
});
