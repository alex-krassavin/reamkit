import { webcrypto } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { buildPkcs7Detached } from '@/crypto';
import * as asn1 from '@/crypto/asn1';

const here = dirname(fileURLToPath(import.meta.url));
const KEY = new Uint8Array(readFileSync(resolve(here, 'fixtures/sign/signer-key.pkcs8')));
const CERT = new Uint8Array(readFileSync(resolve(here, 'fixtures/sign/signer-cert.der')));
const subtle = webcrypto.subtle;

// Pull the SubjectPublicKeyInfo (SPKI) out of the cert so we can verify the
// signature — walk the TBSCertificate children tracking each START offset.
function spkiSlice(cert: Uint8Array): Uint8Array {
  const certSeq = asn1.readTlv(cert, 0);
  const tbs = asn1.readTlv(cert, certSeq.contentStart);
  let p = tbs.contentStart;
  const starts: Array<number> = [];
  while (p < tbs.contentEnd) {
    starts.push(p);
    p = asn1.readTlv(cert, p).end;
  }
  const base = asn1.readTlv(cert, starts[0]!).tag === 0xa0 ? 1 : 0;
  const spkiStart = starts[base + 5]!; // SPKI is the 6th field after the optional version
  return cert.slice(spkiStart, asn1.readTlv(cert, spkiStart).end);
}

function indexOfBytes(hay: Uint8Array, needle: Uint8Array): number {
  outer: for (let i = 0; i <= hay.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}

describe('CMS / PKCS#7 detached signature (RFC 5652)', () => {
  it('builds a SignedData whose signature verifies against the certificate', async () => {
    const data = new TextEncoder().encode('the signed PDF byte range');
    const digest = new Uint8Array(await subtle.digest('SHA-256', data));

    const privKey = await subtle.importKey(
      'pkcs8',
      KEY,
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false,
      ['sign'],
    );

    const cms = await buildPkcs7Detached({
      certificate: CERT,
      messageDigest: digest,
      signingTime: new Date('2026-01-02T03:04:05Z'),
      sign: async (attrs) => new Uint8Array(await subtle.sign('RSASSA-PKCS1-v1_5', privKey, attrs)),
    });

    // ContentInfo SEQ { OID id-signedData, [0] SignedData }
    const ci = asn1.readTlv(cms, 0);
    expect(ci.tag).toBe(0x30);
    const ciKids = asn1.children(cms, ci);
    expect(ciKids[0]!.tag).toBe(0x06); // OID
    const signedData = asn1.readTlv(cms, ciKids[1]!.contentStart); // inside [0]

    // SignedData fields: version, digestAlgs SET, encap, [0]certs, signerInfos SET
    const sdKids = asn1.children(cms, signedData);
    const signerInfos = sdKids.find((k) => k.tag === 0x31 && k !== sdKids[1])!;
    const signerInfo = asn1.readTlv(cms, signerInfos.contentStart);
    const siKids = asn1.children(cms, signerInfo);
    // version, sid, digestAlg, [0]signedAttrs, sigAlg, signature
    const signedAttrs = siKids.find((k) => k.tag === 0xa0)!;
    const signatureOctet = siKids[siKids.length - 1]!;
    expect(signatureOctet.tag).toBe(0x04);
    const signature = cms.slice(signatureOctet.contentStart, signatureOctet.contentEnd);

    // Re-tag signedAttrs [0] → SET (0x31): that is the signed message.
    const attrsBody = cms.slice(signedAttrs.contentStart, signedAttrs.contentEnd);
    const signedMessage = asn1.tlv(0x31, attrsBody);

    const pubKey = await subtle.importKey(
      'spki',
      spkiSlice(CERT),
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false,
      ['verify'],
    );
    const ok = await subtle.verify('RSASSA-PKCS1-v1_5', pubKey, signature, signedMessage);
    expect(ok).toBe(true);

    // The messageDigest attribute must carry our digest.
    const attrs = asn1.children(cms, signedAttrs);
    let found: Uint8Array | undefined;
    for (const a of attrs) {
      const ak = asn1.children(cms, a);
      const oid = cms.slice(ak[0]!.contentStart, ak[0]!.contentEnd);
      // messageDigest OID 1.2.840.113549.1.9.4 → bytes 2a 86 48 86 f7 0d 01 09 04
      if (oid[oid.length - 1] === 0x04 && oid[oid.length - 2] === 0x09) {
        const setVal = ak[1]!;
        const octet = asn1.readTlv(cms, setVal.contentStart);
        found = cms.slice(octet.contentStart, octet.contentEnd);
      }
    }
    expect(found).toBeDefined();
    expect([...found!]).toEqual([...digest]);
  });

  it('builds an ECDSA SignedData whose signature verifies (P-256)', async () => {
    const data = new TextEncoder().encode('the signed PDF byte range');
    const digest = new Uint8Array(await subtle.digest('SHA-256', data));
    const pair = await subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, [
      'sign',
      'verify',
    ]);

    const cms = await buildPkcs7Detached({
      certificate: CERT, // for issuer/serial only (key-agnostic)
      messageDigest: digest,
      signingTime: new Date('2026-01-02T03:04:05Z'),
      signatureAlgorithm: 'ecdsa',
      sign: async (attrs) => {
        const raw = new Uint8Array(
          await subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, pair.privateKey, attrs),
        );
        const half = raw.length / 2;
        // raw r‖s → DER Ecdsa-Sig-Value SEQUENCE { r INTEGER, s INTEGER }.
        return asn1.seq(asn1.integer(raw.slice(0, half)), asn1.integer(raw.slice(half)));
      },
    });

    // The SignerInfo signatureAlgorithm is ecdsa-with-SHA256 (1.2.840.10045.4.3.2).
    expect(indexOfBytes(cms, asn1.oid('1.2.840.10045.4.3.2'))).toBeGreaterThanOrEqual(0);

    // Pull out the embedded signature + the signed message, then verify via WebCrypto.
    const ci = asn1.readTlv(cms, 0);
    const signedData = asn1.readTlv(cms, asn1.children(cms, ci)[1]!.contentStart);
    const sdKids = asn1.children(cms, signedData);
    const signerInfos = sdKids.find((k) => k.tag === 0x31 && k !== sdKids[1])!;
    const signerInfo = asn1.readTlv(cms, signerInfos.contentStart);
    const siKids = asn1.children(cms, signerInfo);
    const signedAttrs = siKids.find((k) => k.tag === 0xa0)!;
    const sigOctet = siKids[siKids.length - 1]!;
    const derSig = cms.slice(sigOctet.contentStart, sigOctet.contentEnd);
    const signedMessage = asn1.tlv(
      0x31,
      cms.slice(signedAttrs.contentStart, signedAttrs.contentEnd),
    );

    // DER Ecdsa-Sig-Value → raw r‖s (32+32 for P-256).
    const sigSeq = asn1.readTlv(derSig, 0);
    const rs = asn1.children(derSig, sigSeq).map((t) => {
      let b = derSig.slice(t.contentStart, t.contentEnd);
      while (b.length > 32 && b[0] === 0) b = b.slice(1);
      const o = new Uint8Array(32);
      o.set(b, 32 - b.length);
      return o;
    });
    const raw = new Uint8Array(64);
    raw.set(rs[0]!, 0);
    raw.set(rs[1]!, 32);

    const ok = await subtle.verify(
      { name: 'ECDSA', hash: 'SHA-256' },
      pair.publicKey,
      raw,
      signedMessage,
    );
    expect(ok).toBe(true);
  });
});
