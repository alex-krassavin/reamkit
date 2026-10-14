// E-PDF EP9 — the Standard security handler (ISO 32000 §7.6.3 / ISO 32000-2
// §7.6.4). Builds a Decryptor from the /Encrypt dictionary using the supplied
// password (EP14), as the user's or as the owner's; the empty string — the
// default — opens the common "permissions-only" encryption. Every revision
// stores a verifier (/U, /O), so a password that opens the file as neither is
// REFUSED, never turned into a key that decrypts rubbish. Supports RC4 (V1/V2),
// AES-128 (V4 /AESV2) and AES-256 (V5/R6 /AESV3), decrypting each string and
// stream with its per-object key.

import type { PdfDict, PdfValue } from '@/pdf/objects';
import {
  aesCbcDecrypt,
  aesCbcEncrypt,
  md5,
  rc4,
  sha256,
  sha384,
  sha512,
} from '@/core/crypto/primitives';

import { PdfHexString, PdfName, PdfStream } from '@/pdf/objects';

/** Decrypts the strings and streams of an encrypted PDF, per indirect object. */
export interface Decryptor {
  /**
   * Decrypt one value with its owning object's number and generation (which seed
   * the per-object key). Recurses into arrays and dictionaries.
   */
  decrypt: (value: PdfValue, objNum: number, gen: number) => PdfValue;
}

type Method = 'rc4' | 'aesv2' | 'aesv3' | 'none';

/** §7.6.5 — how a file's streams and its strings are each encrypted, if at all. */
interface Methods {
  readonly stream: Method;
  readonly string: Method;
}

// §7.6.3.3 — the 32-byte password padding string.
const PAD = Uint8Array.from([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
]);
const AES_SALT = Uint8Array.from([0x73, 0x41, 0x6c, 0x54]); // "sAlT" (§7.6.2)
const ZERO16 = new Uint8Array(16);
const EMPTY = new Uint8Array(0);

/**
 * What {@link buildDecryptor} makes of a file: the decryptor, or why there is
 * none — the password opens it as neither its user nor its owner, or the file is
 * locked by a handler this does not implement.
 */
export type DecryptorResult =
  | { readonly decryptor: Decryptor }
  | { readonly refused: 'password' | 'handler' };

/**
 * Build a {@link Decryptor} from the `/Encrypt` dictionary and the file `/ID`
 * (ISO 32000 §7.6.3 / ISO 32000-2 §7.6.4). Authenticates the password as the
 * user's and then as the owner's, derives the file key for the handler's
 * revision — R6/V5 via `/UE` or `/OE` (AES-256), else the legacy MD5 key — and
 * picks the cipher (RC4, AES-128 `/AESV2`, AES-256 `/AESV3`).
 *
 * @param encrypt  The `/Encrypt` dictionary.
 * @param idArray  The trailer `/ID` array (its first element seeds the legacy key).
 * @param password The password (EP14), the user's or the owner's; the empty
 *                 string opens permissions-only encryption.
 * @returns The decryptor, or why there is none: a password the file's own
 *          verifiers refuse, or a non-Standard filter or unencrypting crypt filter.
 */
export function buildDecryptor(
  encrypt: PdfDict,
  idArray: PdfValue | undefined,
  password = '',
): DecryptorResult {
  const filter = encrypt.get('Filter');
  if (!(filter instanceof PdfName) || filter.value !== 'Standard') return { refused: 'handler' };
  const v = numOf(encrypt.get('V'));
  const r = numOf(encrypt.get('R'));
  const o = strBytes(encrypt.get('O'));
  const u = strBytes(encrypt.get('U'));
  const methods = cryptMethods(encrypt, v, r);
  // §7.6.5 — a file whose streams and strings both pass through the Identity
  // filter keeps its key for its embedded files alone (`/AuthEvent /EFOpen`):
  // the document itself is in the clear and opens without a password.
  // auth-event-ef-open.pdf is exactly that, and asking for one refused a file
  // every reader opens.
  if (methods.stream === 'none' && methods.string === 'none') {
    return { decryptor: { decrypt: (value) => value } };
  }

  let fileKey: Uint8Array | undefined;
  if (r >= 5 || v >= 5) {
    fileKey =
      userKeyR6(u, strBytes(encrypt.get('UE')), password) ??
      ownerKeyR6(o, u, strBytes(encrypt.get('OE')), password);
  } else {
    const keyLen = Math.max(5, Math.min(16, (numOf(encrypt.get('Length')) || 40) / 8));
    const id0 = Array.isArray(idArray) ? strBytes(idArray[0]) : EMPTY;
    // Algorithm 2's key from a padded user password, if /U says it is the file's.
    const keyFor = (padded: Uint8Array): Uint8Array | undefined => {
      const key = deriveKeyLegacy(
        o,
        numOf(encrypt.get('P')) | 0,
        id0,
        r,
        keyLen,
        encrypt.get('EncryptMetadata') !== false,
        padded,
      );
      return userKeyMatches(key, r, id0, u) ? key : undefined;
    };
    fileKey =
      keyFor(padPassword(password)) ?? keyFor(userPasswordFromOwner(password, o, r, keyLen));
  }
  if (!fileKey) return { refused: 'password' };
  const key = fileKey;
  return {
    decryptor: {
      decrypt: (value, objNum, gen) => decryptValue(value, objNum, gen, key, methods),
    },
  };
}

// §7.6.3.3 Algorithm 2 — key from the user password (R2-R4), given already
// padded to 32 bytes with PAD per step (a); the empty password is PAD itself.
function deriveKeyLegacy(
  o: Uint8Array,
  p: number,
  id0: Uint8Array,
  r: number,
  keyLen: number,
  encryptMetadata: boolean,
  padded: Uint8Array,
): Uint8Array {
  const tail = r >= 4 && !encryptMetadata ? Uint8Array.from([0xff, 0xff, 0xff, 0xff]) : EMPTY;
  let hash = md5(concat(padded, o.subarray(0, 32), p32le(p), id0, tail));
  if (r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, keyLen));
  return hash.subarray(0, keyLen);
}

// §7.6.3.4 Algorithm 6 — whether a key Algorithm 2 derived is the file's: /U
// computed again from it must be /U. R2 encrypts PAD under the key and stores
// all 32 bytes (Algorithm 4); R3 and later hash PAD with the file's first ID,
// pass the result through RC4 under the key and nineteen variants of it, and
// store those 16 bytes first (Algorithm 5).
//
// A /U too short to hold a verifier verifies nothing, and the key is taken as
// it was before there was a check: refusing it would close files every reader
// before this one opened.
function userKeyMatches(key: Uint8Array, r: number, id0: Uint8Array, u: Uint8Array): boolean {
  if (r <= 2) return u.length < 32 || equal(rc4(key, PAD), u.subarray(0, 32));
  if (u.length < 16) return true;
  let x = rc4(key, md5(concat(PAD, id0)));
  for (let i = 1; i <= 19; i++) x = rc4(xorKey(key, i), x);
  return equal(x, u.subarray(0, 16));
}

// §7.6.3.4 Algorithm 7 — the user password an OWNER password unlocks. The owner
// password's RC4 key (Algorithm 3, steps a–d) decrypts /O back to the padded
// user password — once for R2, through the key's twenty variants in reverse for
// R3 and later — which then opens the file as the user's does.
function userPasswordFromOwner(
  owner: string,
  o: Uint8Array,
  r: number,
  keyLen: number,
): Uint8Array {
  let hash = md5(padPassword(owner));
  if (r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash);
  const key = hash.subarray(0, r >= 3 ? keyLen : 5);
  let x = o.subarray(0, 32);
  if (r <= 2) return rc4(key, x);
  for (let i = 19; i >= 0; i--) x = rc4(xorKey(key, i), x);
  return x;
}

// The RC4 key with every byte XORed with `i` (Algorithms 5 and 7).
function xorKey(key: Uint8Array, i: number): Uint8Array {
  return key.map((b) => b ^ i);
}

// §7.6.4.3.3 Algorithm 2.A — validate the USER password and recover the 32-byte
// file key from /UE (R6, AES-256). The password is UTF-8, truncated to 127
// bytes; the empty string is the permissions-only case.
function userKeyR6(u: Uint8Array, ue: Uint8Array, password: string): Uint8Array | undefined {
  if (u.length < 48 || ue.length < 32) return undefined;
  const pw = passwordBytes(password);
  const validationSalt = u.subarray(32, 40);
  const keySalt = u.subarray(40, 48);
  if (!equal(hash2B(pw, validationSalt, EMPTY), u.subarray(0, 32))) return undefined; // wrong pw
  const intermediate = hash2B(pw, keySalt, EMPTY);
  return aesCbcDecrypt(intermediate, ZERO16, ue.subarray(0, 32), false);
}

// §7.6.4.3.3 Algorithm 2.A — the same for the OWNER password: its hashes are
// salted from /O and taken over the 48 bytes of /U, and the key comes out of
// /OE.
function ownerKeyR6(
  o: Uint8Array,
  u: Uint8Array,
  oe: Uint8Array,
  password: string,
): Uint8Array | undefined {
  if (o.length < 48 || u.length < 48 || oe.length < 32) return undefined;
  const pw = passwordBytes(password);
  const udata = u.subarray(0, 48);
  if (!equal(hash2B(pw, o.subarray(32, 40), udata), o.subarray(0, 32))) return undefined;
  const intermediate = hash2B(pw, o.subarray(40, 48), udata);
  return aesCbcDecrypt(intermediate, ZERO16, oe.subarray(0, 32), false);
}

// §7.6.3.3 step (a) — the password padded/truncated to exactly 32 bytes with the
// PAD string. PDFDocEncoding is approximated by latin1 (the common ASCII case).
function padPassword(password: string): Uint8Array {
  const raw = strToBytes(password);
  const out = new Uint8Array(32);
  const n = Math.min(raw.length, 32);
  out.set(raw.subarray(0, n), 0);
  out.set(PAD.subarray(0, 32 - n), n);
  return out;
}

// §7.6.4.3.3 — the R6 password is its UTF-8 bytes, truncated to 127 bytes.
// SASLprep (§7.6.4.3.4) is omitted; ASCII passwords are unaffected.
function passwordBytes(password: string): Uint8Array {
  return new TextEncoder().encode(password).subarray(0, 127);
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

// §7.6.5 — the crypt filter each of streams (/StmF) and strings (/StrF) goes
// through, and the method (/CFM) it names; V<4 has no filters and is RC4
// throughout. The two are separate — bug1782186.pdf encrypts its streams with
// AES and leaves its strings Identity — and one method read for both decrypts
// strings that were never encrypted.
function cryptMethods(encrypt: PdfDict, v: number, r: number): Methods {
  if (v < 4) return { stream: 'rc4', string: 'rc4' };
  const fallback: Method = r >= 5 || v >= 5 ? 'aesv3' : 'rc4';
  const methodOf = (key: 'StmF' | 'StrF'): Method => {
    const named = encrypt.get(key);
    const cfName = named instanceof PdfName ? named.value : 'StdCF';
    if (cfName === 'Identity') return 'none';
    const cf = encrypt.get('CF');
    const cfDict = cf instanceof Map ? cf.get(cfName) : undefined;
    const cfm = cfDict instanceof Map ? cfDict.get('CFM') : undefined;
    if (cfm instanceof PdfName) {
      if (cfm.value === 'AESV2') return 'aesv2';
      if (cfm.value === 'AESV3') return 'aesv3';
      if (cfm.value === 'V2') return 'rc4';
      if (cfm.value === 'None') return 'none';
    }
    return fallback;
  };
  return { stream: methodOf('StmF'), string: methodOf('StrF') };
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
  if (method === 'none') return data;
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
  methods: Methods,
): PdfValue {
  if (typeof value === 'string') {
    return latin1(decryptBytes(strToBytes(value), objNum, gen, fileKey, methods.string));
  }
  if (value instanceof PdfHexString) {
    return new PdfHexString(decryptBytes(value.bytes, objNum, gen, fileKey, methods.string));
  }
  if (value instanceof PdfStream) {
    const dict = decryptValue(value.dict, objNum, gen, fileKey, methods);
    const data = decryptBytes(value.data, objNum, gen, fileKey, methods.stream);
    return new PdfStream(dict instanceof Map ? dict : value.dict, data);
  }
  if (Array.isArray(value)) {
    return value.map((v) => decryptValue(v, objNum, gen, fileKey, methods));
  }
  if (value instanceof Map) {
    const out: PdfDict = new Map();
    for (const [k, v] of value) out.set(k, decryptValue(v, objNum, gen, fileKey, methods));
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
