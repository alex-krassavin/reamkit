// ASN.1 DER (ITU-T X.690) — a minimal encoder + reader, just enough to build a
// CMS/PKCS#7 SignedData blob and to pull the issuer + serial out of an X.509
// certificate. No third-party code: this is the only way to produce the
// detached signature a PDF signature dictionary needs.

const textEncoder = new TextEncoder();

export function concat(arrays: ReadonlyArray<Uint8Array>): Uint8Array {
  let total = 0;
  for (const a of arrays) total += a.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) {
    out.set(a, off);
    off += a.length;
  }
  return out;
}

// §8.1.3 — definite-form length octets (short form < 128, else long form).
function encodeLength(len: number): Uint8Array {
  if (len < 0x80) return new Uint8Array([len]);
  const bytes: Array<number> = [];
  let n = len;
  while (n > 0) {
    bytes.unshift(n & 0xff);
    n = Math.floor(n / 256);
  }
  return new Uint8Array([0x80 | bytes.length, ...bytes]);
}

// Tag-length-value: wrap `content` in an identifier octet + length.
export function tlv(tag: number, content: Uint8Array): Uint8Array {
  const len = encodeLength(content.length);
  const out = new Uint8Array(1 + len.length + content.length);
  out[0] = tag;
  out.set(len, 1);
  out.set(content, 1 + len.length);
  return out;
}

export const seq = (...items: Array<Uint8Array>): Uint8Array => tlv(0x30, concat(items));
export const set = (...items: Array<Uint8Array>): Uint8Array => tlv(0x31, concat(items));
export const octetString = (bytes: Uint8Array): Uint8Array => tlv(0x04, bytes);
export const nullValue = (): Uint8Array => tlv(0x05, new Uint8Array(0));
// [n] context-tag, constructed (EXPLICIT or IMPLICIT-over-constructed).
export const explicit = (tagNum: number, content: Uint8Array): Uint8Array =>
  tlv(0xa0 | tagNum, content);

// §8.19 — OBJECT IDENTIFIER. First two arcs fold into 40*a+b; each subsequent
// arc is base-128 big-endian with the high bit set on all but the last octet.
function base128(v: number): Array<number> {
  const out = [v & 0x7f];
  let n = Math.floor(v / 128);
  while (n > 0) {
    out.unshift((n & 0x7f) | 0x80);
    n = Math.floor(n / 128);
  }
  return out;
}

export function oid(dotted: string): Uint8Array {
  const parts = dotted.split('.').map((p) => Number(p));
  const body: Array<number> = [...base128(40 * parts[0]! + parts[1]!)];
  for (let i = 2; i < parts.length; i++) body.push(...base128(parts[i]!));
  return tlv(0x06, new Uint8Array(body));
}

// §8.3 — INTEGER (two's complement, minimal). Accepts a small non-negative
// number or raw magnitude bytes; prepends 0x00 when the top bit is set so the
// value stays positive.
export function integer(value: number | Uint8Array): Uint8Array {
  let bytes: Array<number>;
  if (typeof value === 'number') {
    bytes = [];
    let n = value;
    if (n === 0) bytes = [0];
    while (n > 0) {
      bytes.unshift(n & 0xff);
      n = Math.floor(n / 256);
    }
  } else {
    bytes = [...value];
    while (bytes.length > 1 && bytes[0] === 0 && (bytes[1]! & 0x80) === 0) bytes.shift();
  }
  if (bytes.length === 0) bytes = [0];
  if (bytes[0]! & 0x80) bytes.unshift(0);
  return tlv(0x02, new Uint8Array(bytes));
}

const pad2 = (n: number): string => n.toString().padStart(2, '0');

// CMS Time (RFC 5652 §11.3): UTCTime (YYMMDDHHMMSSZ) for years 1950–2049,
// GeneralizedTime (YYYYMMDDHHMMSSZ) otherwise.
export function cmsTime(d: Date): Uint8Array {
  const y = d.getUTCFullYear();
  const mmddhhmmss =
    pad2(d.getUTCMonth() + 1) +
    pad2(d.getUTCDate()) +
    pad2(d.getUTCHours()) +
    pad2(d.getUTCMinutes()) +
    pad2(d.getUTCSeconds());
  if (y >= 1950 && y < 2050) {
    return tlv(0x17, textEncoder.encode(pad2(y % 100) + mmddhhmmss + 'Z'));
  }
  return tlv(0x18, textEncoder.encode(String(y) + mmddhhmmss + 'Z'));
}

// DER SET OF (§11.6): components ordered ascending by their full encoding. Used
// for the signedAttrs whose byte order must be canonical for the signature.
function compareDer(a: Uint8Array, b: Uint8Array): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    if (a[i]! !== b[i]!) return a[i]! - b[i]!;
  }
  return a.length - b.length;
}

// The sorted, concatenated content of a SET OF — without the outer tag, so the
// same canonical body can be wrapped once as SET (0x31, for signing) and once
// as [0] IMPLICIT (0xa0, inside the SignerInfo).
export function setOfBody(items: ReadonlyArray<Uint8Array>): Uint8Array {
  return concat([...items].sort(compareDer));
}

// ---- Minimal reader (only what cert issuer/serial extraction needs) ----

export interface Tlv {
  readonly tag: number;
  readonly contentStart: number;
  readonly contentEnd: number;
  readonly end: number;
}

export function children(buf: Uint8Array, parent: Tlv): Array<Tlv> {
  const out: Array<Tlv> = [];
  let p = parent.contentStart;
  while (p < parent.contentEnd) {
    const t = readTlv(buf, p);
    out.push(t);
    p = t.end;
  }
  return out;
}

export function readTlv(buf: Uint8Array, offset: number): Tlv {
  const tag = buf[offset]!;
  let i = offset + 1;
  let len = buf[i]!;
  i++;
  if (len & 0x80) {
    const n = len & 0x7f;
    len = 0;
    for (let k = 0; k < n; k++) {
      len = len * 256 + buf[i]!;
      i++;
    }
  }
  return { tag, contentStart: i, contentEnd: i + len, end: i + len };
}

// Certificate ::= SEQ { tbsCertificate SEQ { [0] version?, serialNumber INTEGER,
//   signature AlgId, issuer Name, ... }, ... }. Returns the issuer Name and
// serialNumber as raw DER TLV slices (verbatim, for IssuerAndSerialNumber).
export function certIssuerAndSerial(cert: Uint8Array): {
  readonly issuer: Uint8Array;
  readonly serial: Uint8Array;
} {
  const certSeq = readTlv(cert, 0);
  const tbs = readTlv(cert, certSeq.contentStart);
  let p = tbs.contentStart;
  let t = readTlv(cert, p);
  if (t.tag === 0xa0) {
    p = t.end; // skip the optional [0] version
    t = readTlv(cert, p);
  }
  const serial = cert.slice(p, t.end); // serialNumber INTEGER TLV
  p = t.end;
  t = readTlv(cert, p); // signature AlgorithmIdentifier
  p = t.end;
  t = readTlv(cert, p); // issuer Name
  const issuer = cert.slice(p, t.end);
  return { issuer, serial };
}
