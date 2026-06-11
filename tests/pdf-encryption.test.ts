import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { aesCbcNoPadDecrypt, decryptBytes, hardHash, permissionBits } from '@/pdf/encryption';
import { convertDocxToPdfSync } from '@/word/docx-to-pdf';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const BODY = '<w:p><w:r><w:t>top secret memo</w:t></w:r></w:p>';
const latin1 = (b: Uint8Array) => new TextDecoder('latin1').decode(b);
const encoder = new TextEncoder();

function hexEntry(pdf: string, key: string): Uint8Array {
  const m = new RegExp(`/${key} <([0-9a-fA-F]+)>`).exec(pdf);
  if (!m) throw new Error(`missing /${key}`);
  const hex = m[1]!;
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  return out;
}

describe('PDF encryption — AES-256 / R6 (ISO 32000-2 §7.6)', () => {
  it('round-trips: the user password unwraps the file key and decrypts content', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', {
      fonts: FONTS,
      encrypt: { userPassword: 'secret' },
    });
    const text = latin1(pdf);
    expect(text).toContain('/Filter /Standard');
    expect(text).toContain('/V 5');
    expect(text).toContain('/R 6');
    expect(text).toContain('/CFM /AESV3');
    expect(text).toContain('/Encrypt');
    // Plaintext must not leak anywhere.
    expect(text).not.toContain('top secret memo');

    const U = hexEntry(text, 'U');
    const UE = hexEntry(text, 'UE');
    const pwd = encoder.encode('secret');
    // Algorithm 11: validate the password against /U.
    const check = await hardHash(pwd, U.slice(32, 40), new Uint8Array(0));
    expect(Buffer.from(check).toString('hex')).toBe(Buffer.from(U.slice(0, 32)).toString('hex'));
    // Wrong password fails validation.
    const wrong = await hardHash(encoder.encode('nope'), U.slice(32, 40), new Uint8Array(0));
    expect(Buffer.from(wrong)).not.toEqual(Buffer.from(U.slice(0, 32)));

    // Unwrap the FEK from /UE; decrypting the streams must surface the page
    // content (our streams are raw — no Filter — so 'BT' appears directly).
    const inter = await hardHash(pwd, U.slice(40, 48), new Uint8Array(0));
    const fek = await aesCbcNoPadDecrypt(inter, new Uint8Array(16), UE);
    let sawContent = false;
    const streamRe = />>\s*stream\r?\n/g;
    let m: RegExpExecArray | null;
    while ((m = streamRe.exec(text)) !== null) {
      const start = m.index + m[0].length;
      let end = text.indexOf('endstream', start);
      // The serializer writes one EOL between the data and `endstream`.
      if (pdf[end - 1] === 0x0a) end--;
      if (pdf[end - 1] === 0x0d) end--;
      const cipher = pdf.slice(start, end);
      if (cipher.length < 32 || cipher.length % 16 !== 0) continue;
      const plain = latin1(await decryptBytes(fek, cipher));
      if (plain.includes('BT') && plain.includes('Tj')) {
        sawContent = true;
        break;
      }
    }
    expect(sawContent).toBe(true);
  });

  it('permission bits: all-allowed is -4; switching off print clears bits 3+12', () => {
    expect(permissionBits(undefined)).toBe(-4);
    const p = permissionBits({ printing: false });
    expect(p & (1 << 2)).toBe(0);
    expect(p & (1 << 11)).toBe(0);
    expect(p & (1 << 4)).not.toBe(0); // copying untouched
  });

  it('rejects encryption on the sync path and under PDF/A', async () => {
    expect(() =>
      convertDocxToPdfSync(buildDocxFromBody(BODY), {
        fonts: FONTS,
        encrypt: { userPassword: 'x' },
      }),
    ).toThrow(/async/);
    await expect(
      Ream.parse(buildDocxFromBody(BODY)).convert('pdf', {
        fonts: FONTS,
        pdfA: 'PDF/A-2b',
        encrypt: { userPassword: 'x' },
      }),
    ).rejects.toThrow(/PDF\/A/);
  });

  it('owner password unwraps the same file key via /OE', async () => {
    const pdf = await Ream.parse(buildDocxFromBody(BODY)).convert('pdf', {
      fonts: FONTS,
      encrypt: { userPassword: 'user', ownerPassword: 'owner' },
    });
    const text = latin1(pdf);
    const U = hexEntry(text, 'U');
    const UE = hexEntry(text, 'UE');
    const O = hexEntry(text, 'O');
    const OE = hexEntry(text, 'OE');
    const fekUser = await aesCbcNoPadDecrypt(
      await hardHash(encoder.encode('user'), U.slice(40, 48), new Uint8Array(0)),
      new Uint8Array(16),
      UE,
    );
    const fekOwner = await aesCbcNoPadDecrypt(
      await hardHash(encoder.encode('owner'), O.slice(40, 48), U),
      new Uint8Array(16),
      OE,
    );
    expect(Buffer.from(fekOwner)).toEqual(Buffer.from(fekUser));
    // Owner validation hash includes the full /U as extra input (Algorithm 12).
    const oCheck = await hardHash(encoder.encode('owner'), O.slice(32, 40), U);
    expect(Buffer.from(oCheck)).toEqual(Buffer.from(O.slice(0, 32)));
  });
});
