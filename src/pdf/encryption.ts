// PDF encryption — ISO 32000-2 §7.6, Standard security handler, V5/R6 only
// (AES-256). The legacy handlers (RC4, AES-128) need MD5/RC4, which WebCrypto
// does not provide and this project does not hand-roll; R6 is pure SHA-2 +
// AES-CBC, all WebCrypto. Encryption is inherently non-deterministic (random
// file key, salts and per-object IVs) and asynchronous (WebCrypto) — it is
// only offered on the async conversion path.

import type { PdfValue } from '@/pdf/objects';
import { PdfHexString, PdfStream, dict, name } from '@/pdf/objects';

/** The user-access permission flags (`/P`, §7.6.4.4.7); each defaults to allowed. */
export interface PdfPermissions {
  /** Bit 3 — print the document (default true). */
  readonly printing?: boolean;
  /** Bit 4 — modify contents (default true). */
  readonly modifying?: boolean;
  /** Bit 5 — copy text and graphics (default true). */
  readonly copying?: boolean;
  /** Bit 6 — add/modify annotations (default true). */
  readonly annotating?: boolean;
  /** Bit 9 — fill in form fields (default true). */
  readonly fillingForms?: boolean;
  /** Bit 10 — extraction for accessibility (default true; forced by PDF/UA). */
  readonly contentAccessibility?: boolean;
  /** Bit 11 — assemble (insert/rotate/delete pages) (default true). */
  readonly documentAssembly?: boolean;
}

/** Passwords and permissions for {@link preparePdfEncryption}. */
export interface PdfEncryptOptions {
  /** Required to OPEN the document. Empty string = opens without a prompt. */
  readonly userPassword?: string;
  /** Overrides the permission restrictions. Defaults to the user password. */
  readonly ownerPassword?: string;
  readonly permissions?: PdfPermissions;
}

/** The output of {@link preparePdfEncryption}: the file key plus the `/Encrypt` dictionary. */
export interface PreparedEncryption {
  /** 32-byte file encryption key (FEK) — encrypts every string/stream. */
  readonly fileKey: Uint8Array;
  readonly encryptDict: PdfValue;
}

const encoder = new TextEncoder();
const subtle = globalThis.crypto.subtle;

/**
 * Compute the `/P` permission integer (§7.6.4.4.7) from the high-level flags:
 * bits 1–2 are reserved 0 and undefined bits are 1, so all-permissions is
 * `0xFFFFFFFC` (-4 as int32); each disabled flag clears its bit.
 *
 * @param p The permission flags, or `undefined` for all-allowed.
 * @returns The signed 32-bit `/P` value.
 */
export function permissionBits(p: PdfPermissions | undefined): number {
  let bits = -4; // all 1s, bits 1-2 zero
  const off = (bit: number) => (bits &= ~(1 << (bit - 1)));
  if (p?.printing === false) {
    off(3);
    off(12);
  }
  if (p?.modifying === false) off(4);
  if (p?.copying === false) off(5);
  if (p?.annotating === false) off(6);
  if (p?.fillingForms === false) off(9);
  if (p?.contentAccessibility === false) off(10);
  if (p?.documentAssembly === false) off(11);
  return bits | 0;
}

// UTF-8, truncated to 127 bytes (§7.6.4.3.2; SASLprep is out of scope — use
// ASCII-safe passwords for cross-viewer interop).
function passwordBytes(password: string): Uint8Array {
  return encoder.encode(password).slice(0, 127);
}

function concat(...parts: Array<Uint8Array>): Uint8Array {
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

async function sha(algo: 'SHA-256' | 'SHA-384' | 'SHA-512', data: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await subtle.digest(algo, data as BufferSource));
}

// AES-CBC without padding via WebCrypto (which always PKCS7-pads): for a
// 16-multiple input the output is input+16 — the trailing padding block is
// dropped. Decryption appends a crafted padding block so unpadding succeeds.
async function aesCbcNoPad(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const k = await subtle.importKey('raw', key as BufferSource, 'AES-CBC', false, ['encrypt']);
  const out = new Uint8Array(
    await subtle.encrypt({ name: 'AES-CBC', iv: iv as BufferSource }, k, data as BufferSource),
  );
  return out.slice(0, data.length);
}

/**
 * AES-CBC decryption without padding, the inverse of the internal no-pad
 * encrypt. Appends a crafted padding block so WebCrypto's PKCS7 unpad succeeds
 * and returns exactly `data.length` plain bytes.
 *
 * @param key  The 16-byte AES key.
 * @param iv   The initialization vector.
 * @param data The ciphertext (a 16-byte multiple).
 * @returns The decrypted bytes.
 */
export async function aesCbcNoPadDecrypt(
  key: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
): Promise<Uint8Array> {
  // Append E(lastCipherBlock, full-padding-block) so WebCrypto's unpad sees a
  // valid 16-byte PKCS7 tail and returns exactly `data.length` plain bytes.
  const lastBlock = data.length >= 16 ? data.slice(data.length - 16) : iv;
  const pad = new Uint8Array(16).fill(16);
  const padCipher = await aesCbcNoPad(key, lastBlock, pad);
  const k = await subtle.importKey('raw', key as BufferSource, 'AES-CBC', false, ['decrypt']);
  const full = concat(data, padCipher);
  const out = new Uint8Array(
    await subtle.decrypt({ name: 'AES-CBC', iv: iv as BufferSource }, k, full as BufferSource),
  );
  return out.slice(0, data.length);
}

/**
 * The iterated hash hardening of Algorithm 2.B (§7.6.4.3.4): SHA-256/384/512
 * rounds interleaved with AES-CBC, run to the spec's stopping condition.
 *
 * @param password The (truncated) password bytes.
 * @param salt     The validation or key salt.
 * @param userData Additional hash input (the 48-byte `/U` for owner derivation,
 *   empty otherwise).
 * @returns The derived 32-byte key.
 */
export async function hardHash(
  password: Uint8Array,
  salt: Uint8Array,
  userData: Uint8Array,
): Promise<Uint8Array> {
  let k = await sha('SHA-256', concat(password, salt, userData));
  for (let round = 0; ; round++) {
    const block = concat(password, k, userData);
    const k1 = new Uint8Array(block.length * 64);
    for (let i = 0; i < 64; i++) k1.set(block, i * block.length);
    const e = await aesCbcNoPad(k.slice(0, 16), k.slice(16, 32), k1);
    let mod = 0;
    for (let i = 0; i < 16; i++) mod = (mod + e[i]!) % 3;
    const algo = mod === 0 ? 'SHA-256' : mod === 1 ? 'SHA-384' : 'SHA-512';
    k = (await sha(algo, e)).slice(0, 64);
    if (round >= 63 && e[e.length - 1]! <= round - 31) break;
  }
  return k.slice(0, 32);
}

/**
 * Build the V5/R6 (AES-256) Standard-security-handler `/Encrypt` dictionary:
 * derive `/U`, `/UE`, `/O`, `/OE` and `/Perms` (§7.6.4.4.8–.10, Algorithms
 * 8–10) for a freshly-generated random file encryption key.
 *
 * @param options The passwords and permissions.
 * @returns The file key and the `/Encrypt` dictionary ({@link PreparedEncryption}).
 */
export async function preparePdfEncryption(
  options: PdfEncryptOptions,
): Promise<PreparedEncryption> {
  const user = passwordBytes(options.userPassword ?? '');
  const owner = passwordBytes(options.ownerPassword ?? options.userPassword ?? '');
  const fileKey = new Uint8Array(32);
  globalThis.crypto.getRandomValues(fileKey);

  const salts = new Uint8Array(16 + 16 + 4);
  globalThis.crypto.getRandomValues(salts);
  const uvs = salts.slice(0, 8);
  const uks = salts.slice(8, 16);
  const ovs = salts.slice(16, 24);
  const oks = salts.slice(24, 32);

  // /U: hash(user, validation salt) ‖ salts; /UE: FEK wrapped by the key-salt
  // derived intermediate key.
  const uHash = await hardHash(user, uvs, new Uint8Array(0));
  const U = concat(uHash, uvs, uks);
  const uInter = await hardHash(user, uks, new Uint8Array(0));
  const UE = await aesCbcNoPad(uInter, new Uint8Array(16), fileKey);

  // /O and /OE include the full 48-byte /U as additional hash input.
  const oHash = await hardHash(owner, ovs, U);
  const O = concat(oHash, ovs, oks);
  const oInter = await hardHash(owner, oks, U);
  const OE = await aesCbcNoPad(oInter, new Uint8Array(16), fileKey);

  // /Perms — §7.6.4.4.10 Algorithm 10: P ‖ 0xFF×4 ‖ 'T' ‖ 'adb' ‖ 4 noise
  // bytes, AES-ECB'd with the FEK (one block: CBC with a zero IV).
  const p = permissionBits(options.permissions);
  const permsPlain = new Uint8Array(16);
  new DataView(permsPlain.buffer).setInt32(0, p, true);
  permsPlain.set([0xff, 0xff, 0xff, 0xff], 4);
  permsPlain.set([0x54, 0x61, 0x64, 0x62], 8); // 'T' (metadata encrypted) + 'adb'
  permsPlain.set(salts.slice(32, 36), 12);
  const perms = await aesCbcNoPad(fileKey, new Uint8Array(16), permsPlain);

  const encryptDict = dict({
    Filter: name('Standard'),
    V: 5,
    R: 6,
    Length: 256,
    CF: dict({ StdCF: dict({ CFM: name('AESV3'), AuthEvent: name('DocOpen'), Length: 32 }) }),
    StmF: name('StdCF'),
    StrF: name('StdCF'),
    U: new PdfHexString(U),
    UE: new PdfHexString(UE),
    O: new PdfHexString(O),
    OE: new PdfHexString(OE),
    P: p,
    Perms: new PdfHexString(perms),
    EncryptMetadata: true,
  });

  return { fileKey, encryptDict };
}

// AESV3 content encryption (§7.6.3.3): random IV ‖ AES-256-CBC (PKCS7).
async function encryptBytes(fileKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const iv = new Uint8Array(16);
  globalThis.crypto.getRandomValues(iv);
  const k = await subtle.importKey('raw', fileKey as BufferSource, 'AES-CBC', false, ['encrypt']);
  const cipher = new Uint8Array(
    await subtle.encrypt({ name: 'AES-CBC', iv: iv as BufferSource }, k, data as BufferSource),
  );
  return concat(iv, cipher);
}

/**
 * Decrypt AESV3 content (§7.6.3.3): split the leading 16-byte IV and
 * AES-256-CBC-decrypt the remainder.
 *
 * @param fileKey The 32-byte file encryption key.
 * @param data    The `IV ‖ ciphertext` bytes.
 * @returns The decrypted plaintext.
 */
export async function decryptBytes(fileKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const iv = data.slice(0, 16);
  const k = await subtle.importKey('raw', fileKey as BufferSource, 'AES-CBC', false, ['decrypt']);
  return new Uint8Array(
    await subtle.decrypt({ name: 'AES-CBC', iv: iv as BufferSource }, k, data.slice(16)),
  );
}

/**
 * Recursively encrypt every string and stream in an object graph (AESV3). The
 * `/Encrypt` dictionary itself is added AFTER this pass and never encrypted, and
 * trailer `/ID` strings live outside the body. Strings come back as hex strings,
 * since the cipher bytes are binary; numbers, booleans, names, refs, null and
 * raw tokens pass through unchanged.
 *
 * @param value   The value to walk.
 * @param fileKey The 32-byte file encryption key.
 * @returns The value with its strings/streams replaced by their ciphertext.
 */
export async function encryptObjectGraph(value: PdfValue, fileKey: Uint8Array): Promise<PdfValue> {
  if (typeof value === 'string') {
    return new PdfHexString(await encryptBytes(fileKey, encoder.encode(value)));
  }
  if (value instanceof PdfHexString) {
    return new PdfHexString(await encryptBytes(fileKey, value.bytes));
  }
  if (value instanceof PdfStream) {
    const data = await encryptBytes(fileKey, value.data);
    const d = new Map(value.dict);
    d.set('Length', data.length);
    return new PdfStream(d, data);
  }
  if (Array.isArray(value)) {
    const out: Array<PdfValue> = [];
    for (const v of value) out.push(await encryptObjectGraph(v, fileKey));
    return out;
  }
  if (value instanceof Map) {
    const out = new Map<string, PdfValue>();
    for (const [k, v] of value) out.set(k, await encryptObjectGraph(v, fileKey));
    return out;
  }
  // numbers, booleans, names, refs, null, raw tokens — pass through
  return value;
}
