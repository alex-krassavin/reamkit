// E-PDF EP9 — the synchronous crypto primitives, checked against published test
// vectors (RFC 1321, FIPS 180-4, FIPS 197, the RC4 wiki vector). These underpin
// encrypted-PDF reading, so they are pinned independently of the PDF plumbing.

import { describe, expect, it } from 'vitest';

import {
  aesCbcDecrypt,
  aesCbcEncrypt,
  md5,
  rc4,
  sha256,
  sha384,
  sha512,
} from '@/pdf-reader/crypto';

const hex = (b: Uint8Array): string => [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
const bytes = (h: string): Uint8Array =>
  Uint8Array.from(h.match(/.{2}/g)!.map((x) => parseInt(x, 16)));
const ascii = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('crypto primitives (E-PDF EP9)', () => {
  it('MD5 (RFC 1321)', () => {
    expect(hex(md5(ascii('')))).toBe('d41d8cd98f00b204e9800998ecf8427e');
    expect(hex(md5(ascii('abc')))).toBe('900150983cd24fb0d6963f7d28e17f72');
    expect(hex(md5(ascii('The quick brown fox jumps over the lazy dog')))).toBe(
      '9e107d9d372bb6826bd81d3542a419d6',
    );
  });

  it('SHA-256 (FIPS 180-4)', () => {
    expect(hex(sha256(ascii('abc')))).toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    );
  });

  it('SHA-512 / SHA-384 (FIPS 180-4)', () => {
    expect(hex(sha512(ascii('abc')))).toBe(
      'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a' +
        '2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
    );
    expect(hex(sha384(ascii('abc')))).toBe(
      'cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed' +
        '8086072ba1e7cc2358baeca134c825a7',
    );
  });

  it('RC4 (key "Key", "Plaintext")', () => {
    const ct = rc4(ascii('Key'), ascii('Plaintext'));
    expect(hex(ct)).toBe('bbf316e8d940af0ad3');
    expect(new TextDecoder().decode(rc4(ascii('Key'), ct))).toBe('Plaintext');
  });

  it('AES-128 / AES-256 block decrypt (FIPS 197)', () => {
    const iv = new Uint8Array(16);
    const pt128 = aesCbcDecrypt(
      bytes('000102030405060708090a0b0c0d0e0f'),
      iv,
      bytes('69c4e0d86a7b0430d8cdb78070b4c55a'),
      false,
    );
    expect(hex(pt128)).toBe('00112233445566778899aabbccddeeff');
    const key256 = bytes('000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f');
    const pt256 = aesCbcDecrypt(key256, iv, bytes('8ea2b7ca516745bfeafc49904b496089'), false);
    expect(hex(pt256)).toBe('00112233445566778899aabbccddeeff');
    // Forward cipher (needed by R6 Algorithm 2.B).
    const ct256 = aesCbcEncrypt(key256, iv, bytes('00112233445566778899aabbccddeeff'));
    expect(hex(ct256)).toBe('8ea2b7ca516745bfeafc49904b496089');
  });
});
