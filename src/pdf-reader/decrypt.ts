// E-PDF EP9 — the Standard security handler (ISO 32000 §7.6.3 / ISO 32000-2
// §7.6.4). Builds a Decryptor from the /Encrypt dictionary assuming the EMPTY
// user password (the common "permissions-only" encryption); a document that
// needs a real user password fails key validation and yields no decryptor
// (reported as a loss). Supports RC4 (V1/V2), AES-128 (V4 /AESV2) and AES-256
// (V5/R6 /AESV3), decrypting each string and stream with its per-object key.

import { aesCbcDecrypt, aesCbcEncrypt, md5, rc4, sha256, sha384, sha512 } from './crypto';
import type { PdfDict, PdfValue } from '@/pdf/objects';

import { PdfHexString, PdfName, PdfStream } from '@/pdf/objects';

export interface Decryptor {
  decrypt: (value: PdfValue, objNum: number, gen: number) => PdfValue;
}

type Method = 'rc4' | 'aesv2' | 'aesv3';

// §7.6.3.3 — the 32-byte password padding string.
const PAD = Uint8Array.from([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
]);
const AES_SALT = Uint8Array.from([0x73, 0x41, 0x6c, 0x54]); // "sAlT" (§7.6.2)
const ZERO16 = new Uint8Array(16);
const EMPTY = new Uint8Array(0);

export function buildDecryptor(
  encrypt: PdfDict,
  idArray: PdfValue | undefined,
): Decryptor | undefined {
  const filter = encrypt.get('Filter');
  if (!(filter instanceof PdfName) || filter.value !== 'Standard') return undefined;
  const v = numOf(encrypt.get('V'));
  const r = numOf(encrypt.get('R'));

  let fileKey: Uint8Array | undefined;
  let method: Method;
  if (r >= 5 || v >= 5) {
    fileKey = deriveKeyR6(strBytes(encrypt.get('U')), strBytes(encrypt.get('UE')));
    method = 'aesv3';
  } else {
    const keyLen = (numOf(encrypt.get('Length')) || 40) / 8;
    const id0 = Array.isArray(idArray) ? strBytes(idArray[0]) : EMPTY;
    fileKey = deriveKeyLegacy(
      strBytes(encrypt.get('O')),
      numOf(encrypt.get('P')) | 0,
      id0,
      r,
      Math.max(5, Math.min(16, keyLen)),
      encrypt.get('EncryptMetadata') !== false,
    );
    const m = cipherMethod(encrypt, v);
    if (m === undefined) return undefined;
    method = m;
  }
  if (!fileKey) return undefined;
  const key = fileKey;
  return {
    decrypt: (value, objNum, gen) => decryptValue(value, objNum, gen, key, method),
  };
}

// §7.6.3.3 Algorithm 2 — key from the empty user password (R2-R4).
function deriveKeyLegacy(
  o: Uint8Array,
  p: number,
  id0: Uint8Array,
  r: number,
  keyLen: number,
  encryptMetadata: boolean,
): Uint8Array {
  const tail = r >= 4 && !encryptMetadata ? Uint8Array.from([0xff, 0xff, 0xff, 0xff]) : EMPTY;
  let hash = md5(concat(PAD, o.subarray(0, 32), p32le(p), id0, tail));
  if (r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, keyLen));
  return hash.subarray(0, keyLen);
}

// §7.6.4.3.3 Algorithm 2.A — validate the empty user password and recover the
// 32-byte file key from /UE (R6, AES-256).
function deriveKeyR6(u: Uint8Array, ue: Uint8Array): Uint8Array | undefined {
  if (u.length < 48 || ue.length < 32) return undefined;
  const validationSalt = u.subarray(32, 40);
  const keySalt = u.subarray(40, 48);
  if (!equal(hash2B(EMPTY, validationSalt, EMPTY), u.subarray(0, 32))) return undefined; // wrong pw
  const intermediate = hash2B(EMPTY, keySalt, EMPTY);
  return aesCbcDecrypt(intermediate, ZERO16, ue.subarray(0, 32), false);
}

// §7.6.4.3.4 Algorithm 2.B — the R6 hardened hash.
function hash2B(pw: Uint8Array, salt: Uint8Array, udata: Uint8Array): Uint8Array {
  let k = sha256(concat(pw, salt, udata));
  for (let round = 1; round <= 256; round++) {
    const block = concat(pw, k, udata);
    const k1 = new Uint8Array(block.length * 64);
    for (let i = 0; i < 64; i++) k1.set(block, i * block.length);
    const e = aesCbcEncrypt(k.subarray(0, 16), k.subarray(16, 32), k1);
    let sum = 0;
    for (let i = 0; i < 16; i++) sum += e[i]!;
    const mod = sum % 3;
    k = mod === 0 ? sha256(e) : mod === 1 ? sha384(e) : sha512(e);
    if (round >= 64 && e[e.length - 1]! <= round - 32) break;
  }
  return k.subarray(0, 32);
}

// V4 stream/string crypt-filter method (/CF → /CFM); V<4 is always RC4.
function cipherMethod(encrypt: PdfDict, v: number): Method | undefined {
  if (v < 4) return 'rc4';
  const stmf = encrypt.get('StmF');
  const cfName = stmf instanceof PdfName ? stmf.value : 'StdCF';
  if (cfName === 'Identity') return undefined; // streams not encrypted
  const cf = encrypt.get('CF');
  const cfDict = cf instanceof Map ? cf.get(cfName) : undefined;
  const cfm = cfDict instanceof Map ? cfDict.get('CFM') : undefined;
  if (cfm instanceof PdfName) {
    if (cfm.value === 'AESV2') return 'aesv2';
    if (cfm.value === 'AESV3') return 'aesv3';
  }
  return 'rc4';
}

// --- per-object decryption --------------------------------------------------

function objectKey(fileKey: Uint8Array, objNum: number, gen: number, aes: boolean): Uint8Array {
  const extra = aes ? AES_SALT : EMPTY;
  const seed = new Uint8Array(fileKey.length + 5 + extra.length);
  seed.set(fileKey, 0);
  seed[fileKey.length] = objNum & 0xff;
  seed[fileKey.length + 1] = (objNum >> 8) & 0xff;
  seed[fileKey.length + 2] = (objNum >> 16) & 0xff;
  seed[fileKey.length + 3] = gen & 0xff;
  seed[fileKey.length + 4] = (gen >> 8) & 0xff;
  seed.set(extra, fileKey.length + 5);
  return md5(seed).subarray(0, Math.min(fileKey.length + 5, 16));
}

function decryptBytes(
  data: Uint8Array,
  objNum: number,
  gen: number,
  fileKey: Uint8Array,
  method: Method,
): Uint8Array {
  if (method === 'aesv3') {
    if (data.length < 16) return data;
    return aesCbcDecrypt(fileKey, data.subarray(0, 16), data.subarray(16), true);
  }
  const key = objectKey(fileKey, objNum, gen, method === 'aesv2');
  if (method === 'aesv2') {
    if (data.length < 16) return data;
    return aesCbcDecrypt(key, data.subarray(0, 16), data.subarray(16), true);
  }
  return rc4(key, data);
}

function decryptValue(
  value: PdfValue,
  objNum: number,
  gen: number,
  fileKey: Uint8Array,
  method: Method,
): PdfValue {
  if (typeof value === 'string') {
    return latin1(decryptBytes(strToBytes(value), objNum, gen, fileKey, method));
  }
  if (value instanceof PdfHexString) {
    return new PdfHexString(decryptBytes(value.bytes, objNum, gen, fileKey, method));
  }
  if (value instanceof PdfStream) {
    const dict = decryptValue(value.dict, objNum, gen, fileKey, method);
    const data = decryptBytes(value.data, objNum, gen, fileKey, method);
    return new PdfStream(dict instanceof Map ? dict : value.dict, data);
  }
  if (Array.isArray(value)) {
    return value.map((v) => decryptValue(v, objNum, gen, fileKey, method));
  }
  if (value instanceof Map) {
    const out: PdfDict = new Map();
    for (const [k, v] of value) out.set(k, decryptValue(v, objNum, gen, fileKey, method));
    return out;
  }
  return value;
}

// --- small helpers ----------------------------------------------------------

function numOf(v: PdfValue | undefined): number {
  return typeof v === 'number' ? v : 0;
}

function strBytes(v: PdfValue | undefined): Uint8Array {
  if (v instanceof PdfHexString) return v.bytes;
  if (typeof v === 'string') return strToBytes(v);
  return EMPTY;
}

function strToBytes(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

function latin1(b: Uint8Array): string {
  let s = '';
  for (const x of b) s += String.fromCharCode(x);
  return s;
}

function p32le(p: number): Uint8Array {
  const u = p >>> 0;
  return Uint8Array.from([u & 0xff, (u >> 8) & 0xff, (u >> 16) & 0xff, (u >> 24) & 0xff]);
}

function concat(...parts: Array<Uint8Array>): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

function equal(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}
