// E-PDF EP9 — synchronous cryptographic primitives for reading encrypted PDFs.
// The reader contract is synchronous (handoff v1 §4), which rules out WebCrypto
// (async) — so the Standard security handler's primitives are hand-rolled here:
// MD5 + RC4 (legacy V1/V2/V4), and SHA-256/384/512 + AES-CBC (V5/R6, AES-256).
// These are used ONLY to decrypt content the document already contains; nothing
// here generates keys or signs. Inputs are small (keys, passwords, salts), so
// clarity is preferred over speed.

// --- RC4 (legacy stream cipher, §7.6.2) -------------------------------------

/**
 * RC4 stream cipher (§7.6.2). Symmetric, so the same call both encrypts and
 * decrypts; used by the legacy V1/V2/V4 Standard security handlers.
 */
export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) s[i] = i;
  let j = 0;
  for (let i = 0; i < 256; i++) {
    j = (j + s[i]! + key[i % key.length]!) & 0xff;
    [s[i], s[j]] = [s[j]!, s[i]!];
  }
  const out = new Uint8Array(data.length);
  let a = 0;
  let b = 0;
  for (let k = 0; k < data.length; k++) {
    a = (a + 1) & 0xff;
    b = (b + s[a]!) & 0xff;
    [s[a], s[b]] = [s[b]!, s[a]!];
    out[k] = data[k]! ^ s[(s[a]! + s[b]!) & 0xff]!;
  }
  return out;
}

// --- MD5 (RFC 1321) ---------------------------------------------------------

const MD5_S = [
  7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14,
  20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6,
  10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21,
];
const MD5_K = Array.from({ length: 64 }, (_, i) =>
  Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296),
);

/** MD5 digest (RFC 1321) → 16 bytes; the key-derivation hash for V1/V2/V4. */
export function md5(input: Uint8Array): Uint8Array {
  const len = input.length;
  const withOne = ((len + 8) >> 6) + 1;
  const msg = new Uint8Array(withOne * 64);
  msg.set(input);
  msg[len] = 0x80;
  const bitLen = len * 8;
  for (let i = 0; i < 4; i++) msg[msg.length - 8 + i] = (bitLen >>> (8 * i)) & 0xff;

  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;
  const rol = (x: number, c: number): number => (x << c) | (x >>> (32 - c));

  for (let off = 0; off < msg.length; off += 64) {
    const m = new Array<number>(16);
    for (let i = 0; i < 16; i++) {
      m[i] =
        msg[off + i * 4]! |
        (msg[off + i * 4 + 1]! << 8) |
        (msg[off + i * 4 + 2]! << 16) |
        (msg[off + i * 4 + 3]! << 24);
    }
    let a = a0;
    let b = b0;
    let c = c0;
    let d = d0;
    for (let i = 0; i < 64; i++) {
      let f: number;
      let g: number;
      if (i < 16) {
        f = (b & c) | (~b & d);
        g = i;
      } else if (i < 32) {
        f = (d & b) | (~d & c);
        g = (5 * i + 1) % 16;
      } else if (i < 48) {
        f = b ^ c ^ d;
        g = (3 * i + 5) % 16;
      } else {
        f = c ^ (b | ~d);
        g = (7 * i) % 16;
      }
      f = (f + a + MD5_K[i]! + m[g]!) | 0;
      a = d;
      d = c;
      c = b;
      b = (b + rol(f, MD5_S[i]!)) | 0;
    }
    a0 = (a0 + a) | 0;
    b0 = (b0 + b) | 0;
    c0 = (c0 + c) | 0;
    d0 = (d0 + d) | 0;
  }
  const out = new Uint8Array(16);
  [a0, b0, c0, d0].forEach((v, i) => {
    out[i * 4] = v & 0xff;
    out[i * 4 + 1] = (v >>> 8) & 0xff;
    out[i * 4 + 2] = (v >>> 16) & 0xff;
    out[i * 4 + 3] = (v >>> 24) & 0xff;
  });
  return out;
}

// --- SHA-256 (FIPS 180-4) ---------------------------------------------------

const SHA256_K = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

export function sha256(input: Uint8Array): Uint8Array {
  const h = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ];
  const len = input.length;
  const blocks = ((len + 8) >> 6) + 1;
  const msg = new Uint8Array(blocks * 64);
  msg.set(input);
  msg[len] = 0x80;
  const bitLen = len * 8;
  const dv = new DataView(msg.buffer);
  dv.setUint32(msg.length - 4, bitLen >>> 0);
  dv.setUint32(msg.length - 8, Math.floor(bitLen / 4294967296));

  const rotr = (x: number, n: number): number => (x >>> n) | (x << (32 - n));
  const w = new Array<number>(64);
  for (let off = 0; off < msg.length; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(off + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15]!, 7) ^ rotr(w[i - 15]!, 18) ^ (w[i - 15]! >>> 3);
      const s1 = rotr(w[i - 2]!, 17) ^ rotr(w[i - 2]!, 19) ^ (w[i - 2]! >>> 10);
      w[i] = (w[i - 16]! + s0 + w[i - 7]! + s1) | 0;
    }
    let [a, b, c, d, e, f, g, hh] = h;
    for (let i = 0; i < 64; i++) {
      const S1 = rotr(e!, 6) ^ rotr(e!, 11) ^ rotr(e!, 25);
      const ch = (e! & f!) ^ (~e! & g!);
      const t1 = (hh! + S1 + ch + SHA256_K[i]! + w[i]!) | 0;
      const S0 = rotr(a!, 2) ^ rotr(a!, 13) ^ rotr(a!, 22);
      const maj = (a! & b!) ^ (a! & c!) ^ (b! & c!);
      const t2 = (S0 + maj) | 0;
      hh = g;
      g = f;
      f = e;
      e = (d! + t1) | 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) | 0;
    }
    h[0] = (h[0]! + a!) | 0;
    h[1] = (h[1]! + b!) | 0;
    h[2] = (h[2]! + c!) | 0;
    h[3] = (h[3]! + d!) | 0;
    h[4] = (h[4]! + e!) | 0;
    h[5] = (h[5]! + f!) | 0;
    h[6] = (h[6]! + g!) | 0;
    h[7] = (h[7]! + hh!) | 0;
  }
  const out = new Uint8Array(32);
  new DataView(out.buffer).setUint32(0, h[0]! >>> 0);
  for (let i = 0; i < 8; i++) new DataView(out.buffer).setUint32(i * 4, h[i]! >>> 0);
  return out;
}

// --- SHA-512 / SHA-384 (FIPS 180-4), BigInt for the 64-bit words ------------

const MASK64 = (1n << 64n) - 1n;
const SHA512_K: Array<bigint> = [
  0x428a2f98d728ae22n,
  0x7137449123ef65cdn,
  0xb5c0fbcfec4d3b2fn,
  0xe9b5dba58189dbbcn,
  0x3956c25bf348b538n,
  0x59f111f1b605d019n,
  0x923f82a4af194f9bn,
  0xab1c5ed5da6d8118n,
  0xd807aa98a3030242n,
  0x12835b0145706fben,
  0x243185be4ee4b28cn,
  0x550c7dc3d5ffb4e2n,
  0x72be5d74f27b896fn,
  0x80deb1fe3b1696b1n,
  0x9bdc06a725c71235n,
  0xc19bf174cf692694n,
  0xe49b69c19ef14ad2n,
  0xefbe4786384f25e3n,
  0x0fc19dc68b8cd5b5n,
  0x240ca1cc77ac9c65n,
  0x2de92c6f592b0275n,
  0x4a7484aa6ea6e483n,
  0x5cb0a9dcbd41fbd4n,
  0x76f988da831153b5n,
  0x983e5152ee66dfabn,
  0xa831c66d2db43210n,
  0xb00327c898fb213fn,
  0xbf597fc7beef0ee4n,
  0xc6e00bf33da88fc2n,
  0xd5a79147930aa725n,
  0x06ca6351e003826fn,
  0x142929670a0e6e70n,
  0x27b70a8546d22ffcn,
  0x2e1b21385c26c926n,
  0x4d2c6dfc5ac42aedn,
  0x53380d139d95b3dfn,
  0x650a73548baf63den,
  0x766a0abb3c77b2a8n,
  0x81c2c92e47edaee6n,
  0x92722c851482353bn,
  0xa2bfe8a14cf10364n,
  0xa81a664bbc423001n,
  0xc24b8b70d0f89791n,
  0xc76c51a30654be30n,
  0xd192e819d6ef5218n,
  0xd69906245565a910n,
  0xf40e35855771202an,
  0x106aa07032bbd1b8n,
  0x19a4c116b8d2d0c8n,
  0x1e376c085141ab53n,
  0x2748774cdf8eeb99n,
  0x34b0bcb5e19b48a8n,
  0x391c0cb3c5c95a63n,
  0x4ed8aa4ae3418acbn,
  0x5b9cca4f7763e373n,
  0x682e6ff3d6b2b8a3n,
  0x748f82ee5defb2fcn,
  0x78a5636f43172f60n,
  0x84c87814a1f0ab72n,
  0x8cc702081a6439ecn,
  0x90befffa23631e28n,
  0xa4506cebde82bde9n,
  0xbef9a3f7b2c67915n,
  0xc67178f2e372532bn,
  0xca273eceea26619cn,
  0xd186b8c721c0c207n,
  0xeada7dd6cde0eb1en,
  0xf57d4f7fee6ed178n,
  0x06f067aa72176fban,
  0x0a637dc5a2c898a6n,
  0x113f9804bef90daen,
  0x1b710b35131c471bn,
  0x28db77f523047d84n,
  0x32caab7b40c72493n,
  0x3c9ebe0a15c9bebcn,
  0x431d67c49c100d4cn,
  0x4cc5d4becb3e42b6n,
  0x597f299cfc657e2an,
  0x5fcb6fab3ad6faecn,
  0x6c44198c4a475817n,
];

function sha512Core(iv: Array<bigint>, input: Uint8Array, outLen: number): Uint8Array {
  const rotr = (x: bigint, n: bigint): bigint => ((x >> n) | (x << (64n - n))) & MASK64;
  const h = [...iv];
  const len = input.length;
  const blocks = ((len + 16) >> 7) + 1;
  const msg = new Uint8Array(blocks * 128);
  msg.set(input);
  msg[len] = 0x80;
  let bitLen = BigInt(len) * 8n;
  for (let i = 0; i < 16; i++) {
    msg[msg.length - 1 - i] = Number(bitLen & 0xffn);
    bitLen >>= 8n;
  }
  const w = new Array<bigint>(80);
  for (let off = 0; off < msg.length; off += 128) {
    for (let i = 0; i < 16; i++) {
      let v = 0n;
      for (let b = 0; b < 8; b++) v = (v << 8n) | BigInt(msg[off + i * 8 + b]!);
      w[i] = v;
    }
    for (let i = 16; i < 80; i++) {
      const s0 = rotr(w[i - 15]!, 1n) ^ rotr(w[i - 15]!, 8n) ^ (w[i - 15]! >> 7n);
      const s1 = rotr(w[i - 2]!, 19n) ^ rotr(w[i - 2]!, 61n) ^ (w[i - 2]! >> 6n);
      w[i] = (w[i - 16]! + s0 + w[i - 7]! + s1) & MASK64;
    }
    let [a, b, c, d, e, f, g, hh] = h;
    for (let i = 0; i < 80; i++) {
      const S1 = rotr(e!, 14n) ^ rotr(e!, 18n) ^ rotr(e!, 41n);
      const ch = (e! & f!) ^ (~e! & MASK64 & g!);
      const t1 = (hh! + S1 + ch + SHA512_K[i]! + w[i]!) & MASK64;
      const S0 = rotr(a!, 28n) ^ rotr(a!, 34n) ^ rotr(a!, 39n);
      const maj = (a! & b!) ^ (a! & c!) ^ (b! & c!);
      const t2 = (S0 + maj) & MASK64;
      hh = g;
      g = f;
      f = e;
      e = (d! + t1) & MASK64;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) & MASK64;
    }
    h[0] = (h[0]! + a!) & MASK64;
    h[1] = (h[1]! + b!) & MASK64;
    h[2] = (h[2]! + c!) & MASK64;
    h[3] = (h[3]! + d!) & MASK64;
    h[4] = (h[4]! + e!) & MASK64;
    h[5] = (h[5]! + f!) & MASK64;
    h[6] = (h[6]! + g!) & MASK64;
    h[7] = (h[7]! + hh!) & MASK64;
  }
  const out = new Uint8Array(h.length * 8);
  h.forEach((v, i) => {
    for (let b = 7; b >= 0; b--) {
      out[i * 8 + b] = Number(v & 0xffn);
      v >>= 8n;
    }
  });
  return out.slice(0, outLen);
}

export function sha512(input: Uint8Array): Uint8Array {
  return sha512Core(
    [
      0x6a09e667f3bcc908n,
      0xbb67ae8584caa73bn,
      0x3c6ef372fe94f82bn,
      0xa54ff53a5f1d36f1n,
      0x510e527fade682d1n,
      0x9b05688c2b3e6c1fn,
      0x1f83d9abfb41bd6bn,
      0x5be0cd19137e2179n,
    ],
    input,
    64,
  );
}

export function sha384(input: Uint8Array): Uint8Array {
  return sha512Core(
    [
      0xcbbb9d5dc1059ed8n,
      0x629a292a367cd507n,
      0x9159015a3070dd17n,
      0x152fecd8f70e5939n,
      0x67332667ffc00b31n,
      0x8eb44a8768581511n,
      0xdb0c2e0d64f98fa7n,
      0x47b5481dbefa4fa4n,
    ],
    input,
    48,
  );
}

// --- AES (FIPS 197), decrypt + CBC ------------------------------------------

const AES = buildAesTables();

function buildAesTables(): { sbox: Uint8Array; inv: Uint8Array } {
  const exp = new Uint8Array(256);
  const log = new Uint8Array(256);
  let x = 1;
  for (let i = 0; i < 255; i++) {
    exp[i] = x;
    log[x] = i;
    x ^= (x << 1) ^ (x & 0x80 ? 0x11b : 0); // ×3 in GF(2^8)
    x &= 0xff;
  }
  const sbox = new Uint8Array(256);
  const inv = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    const invByte = i === 0 ? 0 : exp[(255 - log[i]!) % 255]!;
    let s = invByte;
    let acc = invByte;
    for (let r = 0; r < 4; r++) {
      acc = ((acc << 1) | (acc >> 7)) & 0xff;
      s ^= acc;
    }
    s ^= 0x63;
    sbox[i] = s;
    inv[s] = i;
  }
  return { sbox, inv };
}

function gmul(a: number, b: number): number {
  let p = 0;
  for (let i = 0; i < 8; i++) {
    if (b & 1) p ^= a;
    const hi = a & 0x80;
    a = (a << 1) & 0xff;
    if (hi) a ^= 0x1b;
    b >>= 1;
  }
  return p;
}

function expandKey(key: Uint8Array): Uint8Array {
  const nk = key.length / 4;
  const nr = nk + 6;
  const total = 4 * (nr + 1);
  const w = new Uint8Array(total * 4);
  w.set(key);
  let rcon = 1;
  for (let i = nk; i < total; i++) {
    const t = [w[(i - 1) * 4]!, w[(i - 1) * 4 + 1]!, w[(i - 1) * 4 + 2]!, w[(i - 1) * 4 + 3]!];
    if (i % nk === 0) {
      const tmp = t[0]!;
      t[0] = AES.sbox[t[1]!]! ^ rcon;
      t[1] = AES.sbox[t[2]!]!;
      t[2] = AES.sbox[t[3]!]!;
      t[3] = AES.sbox[tmp]!;
      rcon = gmul(rcon, 2);
    } else if (nk > 6 && i % nk === 4) {
      for (let j = 0; j < 4; j++) t[j] = AES.sbox[t[j]!]!;
    }
    for (let j = 0; j < 4; j++) w[i * 4 + j] = w[(i - nk) * 4 + j]! ^ t[j]!;
  }
  return w;
}

function decryptBlock(
  input: Uint8Array,
  w: Uint8Array,
  nr: number,
  out: Uint8Array,
  outOff: number,
): void {
  const s = input.slice(0, 16);
  const addRoundKey = (round: number): void => {
    for (let i = 0; i < 16; i++) s[i] = s[i]! ^ w[round * 16 + i]!;
  };
  addRoundKey(nr);
  for (let round = nr - 1; round >= 1; round--) {
    invShiftRows(s);
    invSubBytes(s);
    addRoundKey(round);
    invMixColumns(s);
  }
  invShiftRows(s);
  invSubBytes(s);
  addRoundKey(0);
  out.set(s, outOff);
}

function invSubBytes(s: Uint8Array): void {
  for (let i = 0; i < 16; i++) s[i] = AES.inv[s[i]!]!;
}

function invShiftRows(s: Uint8Array): void {
  // State is column-major; row r is bytes r, r+4, r+8, r+12. Rotate right by r.
  const t = s.slice();
  for (let r = 1; r < 4; r++) {
    for (let c = 0; c < 4; c++) s[c * 4 + r] = t[((c - r + 4) % 4) * 4 + r]!;
  }
}

function invMixColumns(s: Uint8Array): void {
  for (let c = 0; c < 4; c++) {
    const o = c * 4;
    const a0 = s[o]!;
    const a1 = s[o + 1]!;
    const a2 = s[o + 2]!;
    const a3 = s[o + 3]!;
    s[o] = gmul(a0, 14) ^ gmul(a1, 11) ^ gmul(a2, 13) ^ gmul(a3, 9);
    s[o + 1] = gmul(a0, 9) ^ gmul(a1, 14) ^ gmul(a2, 11) ^ gmul(a3, 13);
    s[o + 2] = gmul(a0, 13) ^ gmul(a1, 9) ^ gmul(a2, 14) ^ gmul(a3, 11);
    s[o + 3] = gmul(a0, 11) ^ gmul(a1, 13) ^ gmul(a2, 9) ^ gmul(a3, 14);
  }
}

// AES-CBC decrypt. `stripPad` removes PKCS#7 padding (PDF content); leave off
// for fixed-size key material.
export function aesCbcDecrypt(
  key: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
  stripPad: boolean,
): Uint8Array {
  const w = expandKey(key);
  const nr = key.length / 4 + 6;
  const blocks = Math.floor(data.length / 16);
  const out = new Uint8Array(blocks * 16);
  let prev = iv;
  for (let i = 0; i < blocks; i++) {
    const cipher = data.subarray(i * 16, i * 16 + 16);
    decryptBlock(cipher, w, nr, out, i * 16);
    for (let j = 0; j < 16; j++) out[i * 16 + j] = out[i * 16 + j]! ^ prev[j]!;
    prev = cipher;
  }
  if (!stripPad || out.length === 0) return out;
  const pad = out[out.length - 1]!;
  return pad >= 1 && pad <= 16 && pad <= out.length ? out.subarray(0, out.length - pad) : out;
}

// The forward cipher is needed only by the R6 key-derivation hash (Algorithm
// 2.B uses AES-128-CBC encryption, no padding).
export function aesCbcEncrypt(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array {
  const w = expandKey(key);
  const nr = key.length / 4 + 6;
  const blocks = Math.floor(data.length / 16);
  const out = new Uint8Array(blocks * 16);
  let prev = iv;
  for (let i = 0; i < blocks; i++) {
    const block = data.slice(i * 16, i * 16 + 16);
    for (let j = 0; j < 16; j++) block[j] = block[j]! ^ prev[j]!;
    encryptBlock(block, w, nr, out, i * 16);
    prev = out.subarray(i * 16, i * 16 + 16);
  }
  return out;
}

function encryptBlock(
  input: Uint8Array,
  w: Uint8Array,
  nr: number,
  out: Uint8Array,
  off: number,
): void {
  const s = input.slice(0, 16);
  const addRoundKey = (round: number): void => {
    for (let i = 0; i < 16; i++) s[i] = s[i]! ^ w[round * 16 + i]!;
  };
  addRoundKey(0);
  for (let round = 1; round < nr; round++) {
    subBytes(s);
    shiftRows(s);
    mixColumns(s);
    addRoundKey(round);
  }
  subBytes(s);
  shiftRows(s);
  addRoundKey(nr);
  out.set(s, off);
}

function subBytes(s: Uint8Array): void {
  for (let i = 0; i < 16; i++) s[i] = AES.sbox[s[i]!]!;
}

function shiftRows(s: Uint8Array): void {
  const t = s.slice();
  for (let r = 1; r < 4; r++) {
    for (let c = 0; c < 4; c++) s[c * 4 + r] = t[((c + r) % 4) * 4 + r]!;
  }
}

function mixColumns(s: Uint8Array): void {
  for (let c = 0; c < 4; c++) {
    const o = c * 4;
    const a0 = s[o]!;
    const a1 = s[o + 1]!;
    const a2 = s[o + 2]!;
    const a3 = s[o + 3]!;
    s[o] = gmul(a0, 2) ^ gmul(a1, 3) ^ a2 ^ a3;
    s[o + 1] = a0 ^ gmul(a1, 2) ^ gmul(a2, 3) ^ a3;
    s[o + 2] = a0 ^ a1 ^ gmul(a2, 2) ^ gmul(a3, 3);
    s[o + 3] = gmul(a0, 3) ^ a1 ^ a2 ^ gmul(a3, 2);
  }
}
