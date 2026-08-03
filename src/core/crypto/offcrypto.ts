// MS-OFFCRYPTO — open a password-protected OOXML document.
//
// ECMA-376 §2.3: an encrypted .docx/.xlsx/.pptx is not a zip at all. The whole
// OPC package is encrypted into the `EncryptedPackage` stream of an OLE
// compound file, beside an `EncryptionInfo` stream describing how. Two schemes
// are in the wild and both are here:
//
//   * STANDARD (§2.3.4.5, Office 2007 and LibreOffice's own export) — a binary
//     header, AES in ECB, and a key spun out of 50 000 SHA-1 rounds.
//   * AGILE (§2.3.4.10, Office 2010 and later) — an XML descriptor naming its
//     own hash and cipher, a key encrypted under the password's key, and a
//     package cut into 4096-byte segments each with an IV of its own.
//
// What comes out is the plain OPC package: a zip the ordinary readers open.
// Decryption only — nothing here writes an encrypted document.

import { openCfb } from '@/core/ole/cfb';
import {
  aesCbcDecrypt,
  aesEcbDecrypt,
  sha1,
  sha256,
  sha384,
  sha512,
} from '@/core/crypto/primitives';

/** Thrown when the bytes are encrypted and the password does not open them. */
export class WrongPasswordError extends Error {
  constructor(message = 'Wrong password for this document') {
    super(message);
    this.name = 'WrongPasswordError';
  }
}

/** Whether the bytes are an OOXML package encrypted per ECMA-376 §2.3. */
export function isEncryptedPackage(bytes: Uint8Array): boolean {
  try {
    const cfb = openCfb(bytes);
    return cfb.hasStream('EncryptedPackage') && cfb.hasStream('EncryptionInfo');
  } catch {
    return false;
  }
}

/**
 * Decrypt an encrypted OOXML package into the plain OPC zip inside it.
 *
 * @param bytes    The compound file (`EncryptionInfo` + `EncryptedPackage`).
 * @param password The password to open it with.
 * @returns The decrypted package bytes — an ordinary `.docx`/`.xlsx` zip.
 * @throws WrongPasswordError when the password fails the document's own
 * verifier, and Error when the encryption is one this reader does not know.
 */
export function decryptPackage(bytes: Uint8Array, password: string): Uint8Array {
  const cfb = openCfb(bytes);
  const info = cfb.readStream('EncryptionInfo');
  const payload = cfb.readStream('EncryptedPackage');
  if (!info || !payload) throw new Error('Not an encrypted OOXML package');
  const v = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const major = v.getUint16(0, true);
  const minor = v.getUint16(2, true);
  // §2.3.4.10 — version 4.4 is the agile scheme, and only that one is XML.
  if (major === 4 && minor === 4) return agile(info, payload, password);
  if ((major === 3 || major === 4) && minor === 2) return standard(info, payload, password);
  throw new Error(`Unsupported OOXML encryption (version ${String(major)}.${String(minor)})`);
}

// ——— the package stream —————————————————————————————————————————————————

// §2.3.4.4 — the stream opens with the plaintext size as a 64-bit count, so the
// padding the last cipher block carries can be cut off again.
function packageSize(payload: Uint8Array): number {
  const v = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  const lo = v.getUint32(0, true);
  const hi = v.getUint32(4, true);
  return hi * 0x100000000 + lo;
}

// ——— standard encryption (§2.3.4.5–§2.3.4.9) ————————————————————————————

function standard(info: Uint8Array, payload: Uint8Array, password: string): Uint8Array {
  const v = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const headerSize = v.getUint32(8, true);
  // §2.3.2 EncryptionHeader: flags, sizeExtra, algId, algIdHash, keySize, …
  const algId = v.getUint32(12 + 8, true);
  const keyBits = v.getUint32(12 + 16, true);
  // §2.3.3 EncryptionVerifier follows the header.
  let at = 12 + headerSize;
  const saltSize = v.getUint32(at, true);
  at += 4;
  const salt = info.subarray(at, at + saltSize);
  at += saltSize;
  const encryptedVerifier = info.subarray(at, at + 16);
  at += 16;
  const verifierHashSize = v.getUint32(at, true);
  at += 4;
  const encryptedVerifierHash = info.subarray(at, at + (verifierHashSize <= 20 ? 32 : 32));
  // §2.3.4.5 — AES-128 (0x660E), AES-192 (0x660F), AES-256 (0x6610). Anything
  // else is the RC4 family, which no corpus document and no supported Office
  // version writes for OOXML.
  if (algId !== 0 && (algId < 0x660e || algId > 0x6610)) {
    throw new Error(`Unsupported OOXML cipher (algId 0x${algId.toString(16)})`);
  }
  const keyLength = Math.max(16, Math.floor((keyBits || 128) / 8));
  const key = standardKey(password, salt, keyLength);

  // §2.3.4.9 — the verifier proves the password before anything is decrypted.
  const verifier = aesEcbDecrypt(key, encryptedVerifier);
  const expected = aesEcbDecrypt(key, encryptedVerifierHash).subarray(0, 20);
  const actual = sha1(verifier);
  if (!sameBytes(actual, expected)) throw new WrongPasswordError();

  const plain = aesEcbDecrypt(key, payload.subarray(8));
  return plain.subarray(0, Math.min(packageSize(payload), plain.length));
}

// §2.3.4.7 — 50 000 rounds of SHA-1 over the salted password, a block number,
// and then the 64-byte pad the specification calls the "derived key".
function standardKey(password: string, salt: Uint8Array, keyLength: number): Uint8Array {
  let h = sha1(concat(salt, utf16le(password)));
  const counter = new Uint8Array(4);
  for (let i = 0; i < 50_000; i++) {
    new DataView(counter.buffer).setUint32(0, i, true);
    h = sha1(concat(counter, h));
  }
  h = sha1(concat(h, new Uint8Array(4))); // block 0
  const x1 = sha1(padded(h, 0x36));
  const x2 = sha1(padded(h, 0x5c));
  return concat(x1, x2).subarray(0, keyLength);
}

function padded(h: Uint8Array, byte: number): Uint8Array {
  const buf = new Uint8Array(64).fill(byte);
  for (let i = 0; i < h.length && i < 64; i++) buf[i] = buf[i]! ^ h[i]!;
  return buf;
}

// ——— agile encryption (§2.3.4.10–§2.3.4.15) ——————————————————————————————

// §2.3.4.13 — the block keys that separate the three things one password key
// is used for.
const BLOCK_VERIFIER_INPUT = Uint8Array.of(0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79);
const BLOCK_VERIFIER_VALUE = Uint8Array.of(0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e);
const BLOCK_KEY_VALUE = Uint8Array.of(0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6);

const SEGMENT = 4096;

interface AgileKeyData {
  readonly saltValue: Uint8Array;
  readonly blockSize: number;
  readonly keyBits: number;
  readonly hash: (b: Uint8Array) => Uint8Array;
  readonly cbc: boolean;
}

function agile(info: Uint8Array, payload: Uint8Array, password: string): Uint8Array {
  const xml = new TextDecoder().decode(info.subarray(8));
  const keyData = agileKeyData(attrsOf(xml, 'keyData'));
  const encAttrs = passwordKeyEncryptor(xml);
  const enc = agileKeyData(encAttrs);
  const spinCount = Number(encAttrs['spinCount'] ?? '0');
  const encryptedKeyValue = b64(encAttrs['encryptedKeyValue'] ?? '');
  const encryptedVerifierHashInput = b64(encAttrs['encryptedVerifierHashInput'] ?? '');
  const encryptedVerifierHashValue = b64(encAttrs['encryptedVerifierHashValue'] ?? '');

  // §2.3.4.11 — the spin is over the password and its salt alone, so it runs
  // ONCE; the block key only joins for the last hash. Spinning it per block
  // hashed a hundred thousand times over for each of the three.
  const spun = agileSpin(password, enc, spinCount);
  const keyFor = (block: Uint8Array): Uint8Array => agileBlockKey(spun, enc, block);

  // §2.3.4.13 — the verifier: decrypt the input, hash it, and it must match the
  // decrypted value.
  const verifier = decryptAgile(
    keyFor(BLOCK_VERIFIER_INPUT),
    enc.saltValue,
    encryptedVerifierHashInput,
    enc.cbc,
  );
  const expected = decryptAgile(
    keyFor(BLOCK_VERIFIER_VALUE),
    enc.saltValue,
    encryptedVerifierHashValue,
    enc.cbc,
  );
  const actual = enc.hash(verifier.subarray(0, enc.saltValue.length));
  if (!sameBytes(actual, expected.subarray(0, actual.length))) throw new WrongPasswordError();

  // §2.3.4.14 — and the package key itself is encrypted under the same key.
  const packageKey = decryptAgile(
    keyFor(BLOCK_KEY_VALUE),
    enc.saltValue,
    encryptedKeyValue,
    enc.cbc,
  ).subarray(0, keyData.keyBits / 8);

  // §2.3.4.15 — the package is cut into 4096-byte segments, each with an IV of
  // its own so a reader can seek without decrypting what came before.
  const cipher = payload.subarray(8);
  const out = new Uint8Array(cipher.length);
  const counter = new Uint8Array(4);
  for (let i = 0; i * SEGMENT < cipher.length; i++) {
    new DataView(counter.buffer).setUint32(0, i, true);
    const iv = keyData
      .hash(concat(keyData.saltValue, counter))
      .subarray(0, keyData.blockSize || 16);
    const block = cipher.subarray(i * SEGMENT, Math.min((i + 1) * SEGMENT, cipher.length));
    out.set(decryptAgile(packageKey, iv, block, keyData.cbc), i * SEGMENT);
  }
  return out.subarray(0, Math.min(packageSize(payload), out.length));
}

// §2.3.4.11 — the salted password hashed `spinCount` times, the round number in
// front of each. This is the whole cost of opening an agile document.
function agileSpin(password: string, enc: AgileKeyData, spinCount: number): Uint8Array {
  let h = enc.hash(concat(enc.saltValue, utf16le(password)));
  const counter = new Uint8Array(4);
  for (let i = 0; i < spinCount; i++) {
    new DataView(counter.buffer).setUint32(0, i, true);
    h = enc.hash(concat(counter, h));
  }
  return h;
}

// …and the key for one purpose: the spun hash with that purpose's block key.
function agileBlockKey(spun: Uint8Array, enc: AgileKeyData, block: Uint8Array): Uint8Array {
  const keyLength = enc.keyBits / 8;
  const derived = enc.hash(concat(spun, block));
  if (derived.length >= keyLength) return derived.subarray(0, keyLength);
  // A hash shorter than the key is padded with 0x36, as the specification says.
  const out = new Uint8Array(keyLength).fill(0x36);
  out.set(derived);
  return out;
}

function decryptAgile(key: Uint8Array, iv: Uint8Array, data: Uint8Array, cbc: boolean): Uint8Array {
  if (!cbc) return aesEcbDecrypt(key, data);
  // The IV is padded to the cipher's block size, never truncated by the caller.
  const block = iv.length >= 16 ? iv.subarray(0, 16) : concat(iv, new Uint8Array(16 - iv.length));
  return aesCbcDecrypt(key, block, data, false);
}

const HASHES: ReadonlyMap<string, (b: Uint8Array) => Uint8Array> = new Map([
  ['sha1', sha1],
  ['sha-1', sha1],
  ['sha256', sha256],
  ['sha-256', sha256],
  ['sha384', sha384],
  ['sha-384', sha384],
  ['sha512', sha512],
  ['sha-512', sha512],
]);

function agileKeyData(attrs: Record<string, string>): AgileKeyData {
  const name = (attrs['hashAlgorithm'] ?? 'sha512').toLowerCase();
  const hash = HASHES.get(name);
  if (!hash) throw new Error(`Unsupported OOXML hash (${name})`);
  const cipher = (attrs['cipherAlgorithm'] ?? 'AES').toUpperCase();
  if (!cipher.startsWith('AES')) throw new Error(`Unsupported OOXML cipher (${cipher})`);
  return {
    saltValue: b64(attrs['saltValue'] ?? ''),
    blockSize: Number(attrs['blockSize'] ?? '16'),
    keyBits: Number(attrs['keyBits'] ?? '256'),
    hash,
    cbc: (attrs['cipherChaining'] ?? 'ChainingModeCBC') !== 'ChainingModeECB',
  };
}

// The descriptor may carry several key encryptors — a certificate one beside
// the password one (Office 2013 writes both). Only the password opens a file
// from a password.
function passwordKeyEncryptor(xml: string): Record<string, string> {
  const re = /<[A-Za-z0-9]*:?encryptedKey\b([^>]*)\/?>/gu;
  let m = re.exec(xml);
  while (m) {
    const attrs = parseAttrs(m[1] ?? '');
    if (attrs['spinCount'] !== undefined) return attrs;
    m = re.exec(xml);
  }
  throw new Error('No password key encryptor in the encryption descriptor');
}

function attrsOf(xml: string, tag: string): Record<string, string> {
  const m = new RegExp(`<[A-Za-z0-9]*:?${tag}\\b([^>]*)/?>`, 'u').exec(xml);
  return m ? parseAttrs(m[1] ?? '') : {};
}

function parseAttrs(raw: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /([A-Za-z0-9_:]+)\s*=\s*"([^"]*)"/gu;
  let m = re.exec(raw);
  while (m) {
    out[m[1]!.replace(/^.*:/u, '')] = m[2]!;
    m = re.exec(raw);
  }
  return out;
}

// ——— small helpers ——————————————————————————————————————————————————————

function utf16le(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  const v = new DataView(out.buffer);
  for (let i = 0; i < s.length; i++) v.setUint16(i * 2, s.charCodeAt(i), true);
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i]! ^ b[i]!;
  return diff === 0;
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

function b64(s: string): Uint8Array {
  const clean = s.replace(/[^A-Za-z0-9+/]/gu, '');
  const out = new Uint8Array(Math.floor((clean.length * 3) / 4));
  let acc = 0;
  let bits = 0;
  let at = 0;
  for (const ch of clean) {
    const v = B64.indexOf(ch);
    if (v < 0) continue;
    acc = (acc << 6) | v;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out[at++] = (acc >> bits) & 0xff;
    }
  }
  return out.subarray(0, at);
}
