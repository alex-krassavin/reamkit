// E-PDF EP14 — encrypted-PDF reading with a NON-EMPTY password. EP9 covered the
// permissions-only (empty-password) case; here the reader must validate the
// exact password — as the user's or as the owner's, by the verifiers every
// revision of the Standard handler stores (§7.6.3.4, §7.6.4.3.3) — recover the
// file key and decrypt the content, while a wrong or missing password is
// REFUSED with WrongPasswordError, as an Office package's is.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { WrongPasswordError } from '@/core/crypto/password-error';
import { md5, rc4 } from '@/core/crypto/primitives';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const PASSWORD = 'Corre©t-Horse-9';

const OWNER = 'the-owner-9';

async function lockedPdf(text: string): Promise<Uint8Array> {
  const docx = buildDocxFromBody(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
  return Ream.parse(docx).convert('pdf', {
    fonts: FONTS,
    encrypt: { userPassword: PASSWORD, ownerPassword: OWNER },
  });
}

async function html(pdf: Uint8Array, password?: string): Promise<string> {
  const doc = password === undefined ? Ream.parse(pdf) : Ream.parse(pdf, { password });
  return new TextDecoder().decode(await doc.convert('html'));
}

/** Whether the reader says it could not decrypt the content. */
function undecrypted(losses: Ream['losses']): boolean {
  return losses.some((l) => l.detail.startsWith('encrypted PDF'));
}

describe('encrypted PDF read with a user password (E-PDF EP14)', () => {
  it('recovers content when the correct user password is supplied', async () => {
    const pdf = await lockedPdf('UnlockedWithPassword');

    const parsed = Ream.parse(pdf, { password: PASSWORD });
    expect(undecrypted(parsed.losses)).toBe(false);

    const out = new TextDecoder().decode(await parsed.convert('html'));
    expect(out.replace(/\s+/g, '')).toContain('UnlockedWithPassword');
  });

  it('opens it with the OWNER password as well', async () => {
    // Algorithm 2.A — the owner's hashes are salted from /O and the key comes
    // out of /OE. Any reader opens a file with either password.
    const pdf = await lockedPdf('UnlockedByOwner');
    expect((await html(pdf, OWNER)).replace(/\s+/g, '')).toContain('UnlockedByOwner');
  });

  it('refuses a wrong password, as an Office package does', async () => {
    // The file stores a verifier, so it KNOWS the password is wrong. It used to
    // come back as a document with no text and a line in the losses, which a
    // caller reading the text and not the losses takes for an empty file.
    const pdf = await lockedPdf('UnlockedWithPassword');
    expect(() => Ream.parse(pdf, { password: 'not-the-password' })).toThrow(WrongPasswordError);
  });

  it('refuses a locked document given no password, and says what to pass', async () => {
    const pdf = await lockedPdf('UnlockedWithPassword');
    expect(() => Ream.parse(pdf)).toThrow(WrongPasswordError);
    expect(() => Ream.parse(pdf)).toThrow(/pass it as `password`/u);
  });

  it('still opens the permissions-only (empty-password) case by default', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>PermissionsOnly</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', {
      fonts: FONTS,
      encrypt: { userPassword: '' },
    });

    const parsed = Ream.parse(pdf);
    expect(undecrypted(parsed.losses)).toBe(false);
    expect((await html(pdf)).replace(/\s+/g, '')).toContain('PermissionsOnly');
  });
});

describe('an RC4 PDF (R2, R3) checks its password too (§7.6.3.4)', () => {
  // The legacy handler derived a key from whatever password it was given and
  // decrypted with it: a wrong one came back as rubbish, with nothing said. Its
  // /U is a verifier like any other — Algorithm 4 for R2, 5 for R3 and later —
  // and its /O unlocks the user password from the owner's (Algorithm 7).

  it('opens with the user password', () => {
    for (const r of [2, 3] as const) {
      const pdf = rc4Pdf('UserOpens', { user: 'u-pass', owner: 'o-pass', r });
      expect(textOf(Ream.parse(pdf, { password: 'u-pass' }))).toContain('UserOpens');
    }
  });

  it('opens with the owner password', () => {
    for (const r of [2, 3] as const) {
      const pdf = rc4Pdf('OwnerOpens', { user: 'u-pass', owner: 'o-pass', r });
      expect(textOf(Ream.parse(pdf, { password: 'o-pass' }))).toContain('OwnerOpens');
    }
  });

  it('refuses a wrong password and a missing one', () => {
    for (const r of [2, 3] as const) {
      const pdf = rc4Pdf('Locked', { user: 'u-pass', owner: 'o-pass', r });
      expect(() => Ream.parse(pdf, { password: 'wrong' })).toThrow(WrongPasswordError);
      expect(() => Ream.parse(pdf)).toThrow(WrongPasswordError);
    }
  });

  it('opens permissions-only encryption with no password', () => {
    const pdf = rc4Pdf('PermissionsOnly', { user: '', owner: 'o-pass', r: 3 });
    expect(textOf(Ream.parse(pdf))).toContain('PermissionsOnly');
  });

  it('opens a file whose crypt filters leave the document in the clear', () => {
    // §7.6.5 — /StmF and /StrF Identity: the key is for embedded files alone
    // (`/AuthEvent /EFOpen`), and the document opens without a password.
    // auth-event-ef-open.pdf is this, and asking for one refused it.
    const pdf = rc4Pdf('InTheClear', { user: 'u-pass', owner: 'o-pass', r: 4, identity: true });
    expect(textOf(Ream.parse(pdf))).toContain('InTheClear');
  });
});

/** The text of a parsed PDF's paragraphs, run together. */
function textOf(doc: Ream): string {
  return JSON.stringify(doc.flow.body);
}

/**
 * A one-page PDF locked by the Standard handler's RC4 revisions — R2 (40-bit)
 * or R3 (128-bit) — with `user` and `owner` passwords, built by the spec's own
 * algorithms (§7.6.3.3–7.6.3.4). With `identity` it is R4 with Identity crypt
 * filters: locked, and its document in the clear.
 */
function rc4Pdf(
  text: string,
  opts: { user: string; owner: string; r: 2 | 3 | 4; identity?: boolean },
): Uint8Array {
  const PAD = Uint8Array.from([
    0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
    0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
  ]);
  const cat = (...parts: ReadonlyArray<Uint8Array>): Uint8Array => {
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of parts) {
      out.set(p, at);
      at += p.length;
    }
    return out;
  };
  const pad = (pw: string): Uint8Array => {
    const raw = new TextEncoder().encode(pw).subarray(0, 32);
    return cat(raw, PAD.subarray(0, 32 - raw.length));
  };
  const xor = (key: Uint8Array, i: number): Uint8Array => key.map((b) => b ^ i);
  const hex = (b: Uint8Array): string =>
    [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
  const { r } = opts;
  const keyLen = r === 2 ? 5 : 16;
  const id0 = Uint8Array.from({ length: 16 }, (_, i) => (i * 37 + 11) & 0xff);
  const p = -4;
  const p32 = Uint8Array.from([p & 0xff, (p >> 8) & 0xff, (p >> 16) & 0xff, (p >> 24) & 0xff]);
  // Algorithm 3 — /O from the owner password.
  let oHash = md5(pad(opts.owner));
  if (r >= 3) for (let i = 0; i < 50; i++) oHash = md5(oHash);
  const oKey = oHash.subarray(0, keyLen);
  let o = rc4(oKey, pad(opts.user));
  if (r >= 3) for (let i = 1; i <= 19; i++) o = rc4(xor(oKey, i), o);
  // Algorithm 2 — the file key from the user password.
  let key = md5(cat(pad(opts.user), o, p32, id0));
  if (r >= 3) for (let i = 0; i < 50; i++) key = md5(key.subarray(0, keyLen));
  key = key.subarray(0, keyLen);
  // Algorithm 4 (R2) / 5 (R3) — /U.
  let u: Uint8Array;
  if (r === 2) u = rc4(key, PAD);
  else {
    let x = rc4(key, md5(cat(PAD, id0)));
    for (let i = 1; i <= 19; i++) x = rc4(xor(key, i), x);
    u = cat(x, new Uint8Array(16));
  }
  // The content stream, object 4, encrypted with its object key (§7.6.3.2).
  const plain = new TextEncoder().encode(`BT /F1 12 Tf 20 700 Td (${text}) Tj ET`);
  const objKey = md5(cat(key, Uint8Array.from([4, 0, 0, 0, 0]))).subarray(
    0,
    Math.min(keyLen + 5, 16),
  );
  const content = opts.identity ? plain : rc4(objKey, plain);
  const encrypt = opts.identity
    ? `<< /Filter /Standard /V 4 /R 4 /Length 128 /CF << /StdCF << /CFM /V2 /AuthEvent /EFOpen /Length 16 >> >> /StmF /Identity /StrF /Identity /EFF /StdCF /O <${hex(o)}> /U <${hex(u)}> /P ${String(p)} >>`
    : `<< /Filter /Standard /V ${r === 2 ? 1 : 2} /R ${String(r)} /Length ${String(keyLen * 8)} /O <${hex(o)}> /U <${hex(u)}> /P ${String(p)} >>`;
  const enc = new TextEncoder();
  const objects: Array<Uint8Array> = [
    enc.encode('<< /Type /Catalog /Pages 2 0 R >>'),
    enc.encode('<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
    enc.encode(
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    ),
    cat(
      enc.encode(`<< /Length ${String(content.length)} >>\nstream\n`),
      content,
      enc.encode('\nendstream'),
    ),
    enc.encode('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'),
    enc.encode(encrypt),
  ];
  const parts: Array<Uint8Array> = [enc.encode('%PDF-1.4\n')];
  const offsets: Array<number> = [];
  let at = parts[0]!.length;
  objects.forEach((body, i) => {
    offsets.push(at);
    const chunk = cat(enc.encode(`${String(i + 1)} 0 obj\n`), body, enc.encode('\nendobj\n'));
    parts.push(chunk);
    at += chunk.length;
  });
  let tail = `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) tail += `${String(off).padStart(10, '0')} 00000 n \n`;
  tail +=
    `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R /Encrypt 6 0 R ` +
    `/ID [<${hex(id0)}> <${hex(id0)}>] >>\nstartxref\n${String(at)}\n%%EOF\n`;
  parts.push(enc.encode(tail));
  return cat(...parts);
}
