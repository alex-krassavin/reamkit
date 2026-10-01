import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildCfb } from './fixtures/build-cfb';
import {
  WrongPasswordError as PublicWrongPasswordError,
  isEncryptedPackage as publicIsEncryptedPackage,
} from '@/index';
import { Ream } from '@/core/converter/ream';
import { WrongPasswordError, decryptPackage, isEncryptedPackage } from '@/core/crypto/offcrypto';
import { sha1 } from '@/core/crypto/primitives';

const real = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(`tests/fixtures/real/${name}`));

// ECMA-376 §2.3 — the encrypted package is an OLE container, so a `.docx` that
// asks for a password does not even open as a zip.
const STANDARD = 'Encrypted_LO_Standard_abc.docx'; // EncryptionInfo 3.2
const AGILE = 'Encrypted_MSO2013_abc.docx'; // EncryptionInfo 4.4
const WORKBOOK = 'protected_passtika.xlsx';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

describe('SHA-1 (FIPS 180-4)', () => {
  const hex = (b: Uint8Array): string =>
    [...b].map((x) => x.toString(16).padStart(2, '0')).join('');

  it('matches the published vectors', () => {
    expect(hex(sha1(new Uint8Array(0)))).toBe('da39a3ee5e6b4b0d3255bfef95601890afd80709');
    expect(hex(sha1(new TextEncoder().encode('abc')))).toBe(
      'a9993e364706816aba3e25717850c26c9cd0d89d',
    );
    expect(hex(sha1(new TextEncoder().encode('a'.repeat(1000))))).toBe(
      '291e9a6c66994949b57ba5e650361e98fc36b1ba',
    );
  });
});

describe('an encrypted OOXML package (MS-OFFCRYPTO)', () => {
  it('is recognised as one, and a plain document is not', () => {
    expect(isEncryptedPackage(real(STANDARD))).toBe(true);
    expect(isEncryptedPackage(real(AGILE))).toBe(true);
    expect(isEncryptedPackage(real('Spill.xlsx'))).toBe(false);
  });

  it('opens standard encryption (3.2 — AES-ECB, 50 000 SHA-1 rounds)', () => {
    const plain = decryptPackage(real(STANDARD), 'abc');
    // What comes out is an ordinary OPC package: a zip, "PK".
    expect([...plain.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('opens agile encryption (4.4 — SHA-512, AES-CBC, 4096-byte segments)', () => {
    const plain = decryptPackage(real(AGILE), 'abc');
    expect([...plain.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('refuses a wrong password rather than returning rubbish', () => {
    // §2.3.4.9/§2.3.4.13 — the document carries a verifier for exactly this.
    expect(() => decryptPackage(real(STANDARD), 'wrong')).toThrow(WrongPasswordError);
    expect(() => decryptPackage(real(AGILE), 'wrong')).toThrow(WrongPasswordError);
  });
});

describe('the agile encryption descriptor', () => {
  /** An agile package (§2.3.4.10, version 4.4) whose descriptor is `xml`. */
  function agileWith(xml: string): Uint8Array {
    const text = new TextEncoder().encode(xml);
    const info = new Uint8Array(8 + text.length);
    info.set([4, 0, 4, 0, 0x40, 0, 0, 0]); // version 4.4, the reserved 0x40
    info.set(text, 8);
    return buildCfb([
      { name: 'EncryptionInfo', data: info },
      { name: 'EncryptedPackage', data: new Uint8Array(16) },
    ]);
  }

  // A hash it does not know is named in the error, which shows how the
  // descriptor was read.
  it('reads a name whatever its prefix and the spaces around its `=`', () => {
    const xml = '<encryption><x:keyData p:hashAlgorithm = "Whirlpool"/></encryption>';
    expect(() => decryptPackage(agileWith(xml), 'abc')).toThrow('OOXML hash (whirlpool)');
  });

  it('takes the key encryptor with a spin count — the password one', () => {
    const xml =
      '<encryption><keyData hashAlgorithm="SHA512"/>' +
      '<c:encryptedKey hashAlgorithm="SHA1"/>' +
      '<p:encryptedKey spinCount="1" hashAlgorithm="Whirlpool"/></encryption>';
    expect(() => decryptPackage(agileWith(xml), 'abc')).toThrow('OOXML hash (whirlpool)');
  });

  // The descriptor is read before any password is checked, so it is whatever
  // the file says. Each of these took the regular expressions that read it time
  // quadratic in its length: seconds here, a hundred times that at ten times the
  // length.
  it('reads past a long run of name characters no `=` follows, in one pass', () => {
    const bytes = agileWith(
      `<encryption><keyData ${'0'.repeat(100_000)} hashAlgorithm="Whirlpool"/>`,
    );
    const start = performance.now();
    expect(() => decryptPackage(bytes, 'abc')).toThrow('OOXML hash (whirlpool)');
    expect(performance.now() - start).toBeLessThan(1000);
  });

  it('gives up at the first element nothing closes', () => {
    const bytes = agileWith(`<encryption>${'<p:encryptedKey spinCount="1"'.repeat(20_000)}`);
    const start = performance.now();
    expect(() => decryptPackage(bytes, 'abc')).toThrow('No password key encryptor');
    expect(performance.now() - start).toBeLessThan(1000);
  });
});

describe('Ream.parse on an encrypted source', () => {
  it('reads the document the password opens', async () => {
    const ream = Ream.parse(real(AGILE), { password: 'abc' });
    expect(ream.format).toBe('docx');
    const pdf = await ream.convert('pdf', { fonts: FONTS });
    // The one word in that document, drawn on the page.
    expect(pdf.length).toBeGreaterThan(1000);
    const text = ream.flow.body
      .flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []))
      .map((r) => r.text)
      .join('');
    expect(text).toContain('ABC');
  });

  it('reads an encrypted WORKBOOK the same way', () => {
    const ream = Ream.parse(real(WORKBOOK), { password: 'tika' });
    expect(ream.format).toBe('xlsx');
    expect(ream.sheet).toBeDefined();
  });

  it('says what to do when no password is given', () => {
    expect(() => Ream.parse(real(STANDARD))).toThrow(/password/iu);
  });

  it('…and reports a wrong one as a wrong password', () => {
    expect(() => Ream.parse(real(STANDARD), { password: 'nope' })).toThrow(WrongPasswordError);
  });

  it('can be asked about before parsing, through the package', () => {
    // A caller that means to prompt for a password wants to know BEFORE the
    // parse throws — the same answer the reader uses to decide to decrypt.
    expect(publicIsEncryptedPackage).toBe(isEncryptedPackage);
    expect(publicIsEncryptedPackage(real(AGILE))).toBe(true);
    expect(publicIsEncryptedPackage(real('Spill.xlsx'))).toBe(false);
  });

  it('throws the error the package exports, so a caller can catch it by type', () => {
    expect(PublicWrongPasswordError).toBe(WrongPasswordError);
    try {
      Ream.parse(real(AGILE), { password: 'nope' });
      expect.unreachable('a wrong password must throw');
    } catch (err) {
      expect(err).toBeInstanceOf(PublicWrongPasswordError);
      expect((err as Error).name).toBe('WrongPasswordError');
    }
  });
});
