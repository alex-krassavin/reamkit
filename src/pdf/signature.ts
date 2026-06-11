// ISO 32000-1 §12.8 — PDF digital signatures. A signature covers two byte
// ranges of the file (everything except the /Contents hex hole); the hole holds
// a detached PKCS#7/CMS blob over the SHA-256 of those ranges. We build the
// signature dictionary with placeholders during the normal write, then splice
// the real ByteRange + signature in afterwards so no other byte shifts.

import type { PdfDict, PdfRef, PdfValue } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import { buildPkcs7Detached } from '@/core/crypto';
import * as der from '@/core/crypto/asn1';
import { PdfHexString, PdfRawToken, dict, name, ref } from '@/pdf/objects';

// A fixed-width /ByteRange placeholder: four 10-digit slots → overwritten in
// place with the real offsets (right-padded with spaces, same total length).
const BYTE_RANGE_PLACEHOLDER = '[0 0000000000 0000000000 0000000000]';
const DEFAULT_RESERVE_BYTES = 8192; // /Contents hole (16 384 hex chars)

export interface SignaturePlaceholder {
  readonly reason?: string;
  readonly location?: string;
  // Signer name recorded in the dictionary (/Name).
  readonly name?: string;
  readonly contactInfo?: string;
  // Claimed signing time (→ /M and the CMS signingTime attribute).
  readonly signingTime?: Date;
  // Signature field name (default "Signature1").
  readonly fieldName?: string;
  // Bytes reserved for the /Contents signature hole (default 8192).
  readonly reserveBytes?: number;
}

export interface SignerCredentials {
  // Signer's X.509 certificate (DER).
  readonly certificate: Uint8Array;
  // WebCrypto private key. An RSASSA-PKCS1-v1_5 key (default) or — when
  // `algorithm` is 'ecdsa' — an ECDSA key (P-256/P-384/P-521).
  readonly privateKey: CryptoKey;
  // Signature algorithm of the key (default 'rsa').
  readonly algorithm?: 'rsa' | 'ecdsa';
  // Optional chain certificates (DER) to embed alongside the signer cert.
  readonly extraCertificates?: ReadonlyArray<Uint8Array>;
  readonly signingTime?: Date;
}

// Converter convenience: everything needed to emit + sign in one call.
export interface SignatureOptions extends SignaturePlaceholder, SignerCredentials {}

function pad2(n: number): string {
  return n.toString().padStart(2, '0');
}

// PDF date string (§7.9.4): D:YYYYMMDDHHmmSS+00'00' (always emitted as UTC).
function pdfDate(d: Date): string {
  return (
    `D:${d.getUTCFullYear()}${pad2(d.getUTCMonth() + 1)}${pad2(d.getUTCDate())}` +
    `${pad2(d.getUTCHours())}${pad2(d.getUTCMinutes())}${pad2(d.getUTCSeconds())}+00'00'`
  );
}

// Add an (invisible) signature field + its signature dictionary to a document
// being built. The dictionary carries placeholder /ByteRange and /Contents that
// signPdf fills later. Returns the field ref (attach to the page's /Annots) and
// the /AcroForm dict (attach to the catalog).
export function addSignaturePlaceholder(
  doc: PdfDocument,
  ph: SignaturePlaceholder,
  pageRef: PdfRef,
): { readonly fieldRef: PdfRef; readonly acroForm: PdfDict } {
  const reserve = ph.reserveBytes ?? DEFAULT_RESERVE_BYTES;
  const sigEntries: Record<string, PdfValue> = {
    Type: name('Sig'),
    Filter: name('Adobe.PPKLite'),
    SubFilter: name('adbe.pkcs7.detached'),
    // Insertion order matters: ByteRange precedes Contents so signPdf's forward
    // scan finds the hole after the range slot.
    ByteRange: new PdfRawToken(BYTE_RANGE_PLACEHOLDER),
    Contents: new PdfHexString(new Uint8Array(reserve)),
  };
  if (ph.signingTime) sigEntries['M'] = pdfDate(ph.signingTime);
  if (ph.reason) sigEntries['Reason'] = ph.reason;
  if (ph.location) sigEntries['Location'] = ph.location;
  if (ph.name) sigEntries['Name'] = ph.name;
  if (ph.contactInfo) sigEntries['ContactInfo'] = ph.contactInfo;
  const sigRef = doc.add(dict(sigEntries));

  // A signature field merged with its widget annotation (§12.7.4.5). Invisible:
  // zero Rect, Print + Locked flags.
  const fieldRef = doc.add(
    dict({
      FT: name('Sig'),
      Type: name('Annot'),
      Subtype: name('Widget'),
      T: ph.fieldName ?? 'Signature1',
      V: ref(sigRef.id),
      P: ref(pageRef.id),
      Rect: [0, 0, 0, 0],
      F: 132,
    }),
  );

  // SigFlags 3 = SignaturesExist | AppendOnly (§12.7.2).
  const acroForm = dict({ Fields: [ref(fieldRef.id)], SigFlags: 3 });
  return { fieldRef, acroForm };
}

function indexOfAscii(buf: Uint8Array, needle: string, from = 0): number {
  const n = needle.length;
  outer: for (let i = from; i <= buf.length - n; i++) {
    for (let j = 0; j < n; j++) {
      if (buf[i + j] !== needle.charCodeAt(j)) continue outer;
    }
    return i;
  }
  return -1;
}

function writeAscii(buf: Uint8Array, at: number, s: string): void {
  for (let i = 0; i < s.length; i++) buf[at + i] = s.charCodeAt(i);
}

function getSubtle(): SubtleCrypto {
  const c = (globalThis as { crypto?: Crypto }).crypto;
  if (!c?.subtle) {
    throw new Error('Signing requires WebCrypto (globalThis.crypto.subtle)');
  }
  return c.subtle;
}

// WebCrypto wants an ArrayBuffer-backed view; TS 5.7 widened Uint8Array's buffer
// to ArrayBufferLike (which BufferSource rejects), so re-assert / copy.
function asData(u: Uint8Array): Uint8Array<ArrayBuffer> {
  return u.buffer instanceof ArrayBuffer ? (u as Uint8Array<ArrayBuffer>) : new Uint8Array(u);
}

// Convert a WebCrypto raw ECDSA signature (r‖s, each curve-order bytes) into the
// DER Ecdsa-Sig-Value the CMS expects: SEQUENCE { r INTEGER, s INTEGER }.
// der.integer applies the minimal-positive-integer encoding (leading-zero strip
// + 0x00 sign guard) to each half.
function ecdsaRawToDer(raw: Uint8Array): Uint8Array {
  const half = raw.length >> 1;
  return der.seq(der.integer(raw.slice(0, half)), der.integer(raw.slice(half)));
}

const HEX = '0123456789ABCDEF';
function toHex(bytes: Uint8Array): string {
  let s = '';
  for (const b of bytes) s += HEX[b >> 4]! + HEX[b & 0x0f]!;
  return s;
}

// Sign a PDF that already carries a signature placeholder (from
// addSignaturePlaceholder / options.signaturePlaceholder). Computes the real
// ByteRange, hashes the covered bytes, builds the detached CMS, and writes it
// into the reserved /Contents hole — leaving every other byte untouched.
export async function signPdf(pdf: Uint8Array, cred: SignerCredentials): Promise<Uint8Array> {
  const out = pdf.slice();
  // Match the FULL fixed-width placeholder, not a bare '/ByteRange' — an
  // embedded file (e.g. a PDF attached via /AF) or a font stream could
  // legitimately contain that name earlier in the file (oop-design §8, B4).
  const brKey = indexOfAscii(out, `/ByteRange ${BYTE_RANGE_PLACEHOLDER}`);
  if (brKey < 0) throw new Error('signPdf: no /ByteRange placeholder found');
  const brOpen = indexOfAscii(out, '[', brKey);
  const brClose = indexOfAscii(out, ']', brOpen);
  const ctKey = indexOfAscii(out, '/Contents', brClose);
  const ctOpen = indexOfAscii(out, '<', ctKey);
  const ctClose = indexOfAscii(out, '>', ctOpen);
  if (brOpen < 0 || brClose < 0 || ctOpen < 0 || ctClose < 0) {
    throw new Error('signPdf: malformed signature placeholder');
  }

  // ByteRange covers [0 .. the "<"] and ["<...>" excluded] [the ">" .. EOF].
  const a = 0;
  const b = ctOpen + 1;
  const c = ctClose;
  const d = out.length - ctClose;

  const brStr = `[${a} ${b} ${c} ${d}]`;
  const brWidth = brClose - brOpen + 1;
  if (brStr.length > brWidth) throw new Error('signPdf: /ByteRange placeholder too small');
  // Pad with spaces before the closing bracket, keeping total width identical.
  writeAscii(out, brOpen, brStr.slice(0, -1) + ' '.repeat(brWidth - brStr.length) + ']');

  // Hash the two covered segments (with the now-final ByteRange bytes in seg 1).
  const segment = new Uint8Array(b + d);
  segment.set(out.subarray(a, a + b), 0);
  segment.set(out.subarray(c, c + d), b);
  const subtle = getSubtle();
  const digest = new Uint8Array(await subtle.digest('SHA-256', asData(segment)));

  const signingTime = cred.signingTime ?? new Date();
  const isEcdsa = cred.algorithm === 'ecdsa';
  const cms = await buildPkcs7Detached({
    certificate: cred.certificate,
    ...(cred.extraCertificates ? { extraCertificates: cred.extraCertificates } : {}),
    messageDigest: digest,
    signingTime,
    signatureAlgorithm: isEcdsa ? 'ecdsa' : 'rsa',
    sign: async (attrs) => {
      if (isEcdsa) {
        // WebCrypto ECDSA returns a raw r‖s (IEEE P1363); CMS needs a DER
        // Ecdsa-Sig-Value (SEQUENCE { r INTEGER, s INTEGER }).
        const raw = new Uint8Array(
          await subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, cred.privateKey, asData(attrs)),
        );
        return ecdsaRawToDer(raw);
      }
      return new Uint8Array(await subtle.sign('RSASSA-PKCS1-v1_5', cred.privateKey, asData(attrs)));
    },
  });

  const holeLen = ctClose - ctOpen - 1; // hex chars available
  const hex = toHex(cms);
  if (hex.length > holeLen) {
    throw new Error(
      `signPdf: signature (${hex.length} hex) exceeds reserved /Contents (${holeLen}); raise reserveBytes`,
    );
  }
  writeAscii(out, ctOpen + 1, hex + '0'.repeat(holeLen - hex.length));
  return out;
}
