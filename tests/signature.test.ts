import { webcrypto } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement } from '@/document-model';
import { convertDocxToPdf } from '@/converter';
import { FontRegistry } from '@/font';
import { EMPTY_STYLE_SHEET } from '@/ooxml/wordproc';
import { renderStyledPdf, signPdf } from '@/pdf';
import * as asn1 from '@/crypto/asn1';

const here = dirname(fileURLToPath(import.meta.url));
const REGULAR = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf')));
const KEY = new Uint8Array(readFileSync(resolve(here, 'fixtures/sign/signer-key.pkcs8')));
const CERT = new Uint8Array(readFileSync(resolve(here, 'fixtures/sign/signer-cert.der')));
const subtle = webcrypto.subtle;
const latin1 = new TextDecoder('latin1');

const registry = FontRegistry.fromBytes({ regular: REGULAR });
const BODY: Array<BodyElement> = [
  {
    kind: 'paragraph',
    paragraph: { properties: {}, runs: [{ text: 'Sign me.', properties: {} }] },
  },
];

function placeholderPdf(reserveBytes?: number): Uint8Array {
  return renderStyledPdf(BODY, {
    registry,
    styles: EMPTY_STYLE_SHEET,
    signaturePlaceholder: {
      reason: 'I approve',
      name: 'Test Signer',
      signingTime: new Date('2026-01-02T03:04:05Z'),
      ...(reserveBytes !== undefined ? { reserveBytes } : {}),
    },
  });
}

async function importKey(): Promise<CryptoKey> {
  return subtle.importKey('pkcs8', KEY, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, [
    'sign',
  ]);
}

function hexToBytes(hex: string): Uint8Array {
  const out = new Uint8Array(hex.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.substring(i * 2, i * 2 + 2), 16);
  return out;
}

describe('PDF digital signature (ISO 32000 §12.8)', () => {
  it('emits an unsigned placeholder with AcroForm + zeroed ByteRange/Contents', () => {
    const text = latin1.decode(placeholderPdf());
    expect(text).toContain('/Type /Sig');
    expect(text).toContain('/SubFilter /adbe.pkcs7.detached');
    expect(text).toContain('/ByteRange [0 0000000000 0000000000 0000000000]');
    expect(text).toContain('/SigFlags 3');
    expect(text).toMatch(/\/FT \/Sig/);
    expect(text).toMatch(/\/Annots \[\d+ 0 R\]/);
    // The /Contents hole is all zeros before signing.
    expect(text).toMatch(/\/Contents <0{1000,}>/);
  });

  it('signs the placeholder: real ByteRange + a CMS that covers exactly those bytes', async () => {
    const privateKey = await importKey();
    const signed = await signPdf(placeholderPdf(), {
      certificate: CERT,
      privateKey,
      signingTime: new Date('2026-01-02T03:04:05Z'),
    });
    const text = latin1.decode(signed);

    // ByteRange now carries real offsets (placeholder zeros are gone).
    const m = /\/ByteRange \[(\d+) (\d+) (\d+) (\d+)/.exec(text);
    expect(m).not.toBeNull();
    const [a, b, c, d] = [Number(m![1]), Number(m![2]), Number(m![3]), Number(m![4])];
    expect(a).toBe(0);

    // The gap between the two segments is exactly the /Contents hex hole. Find
    // the signature's /Contents (the one after /ByteRange), not the page's.
    const brClose = text.indexOf(']', text.indexOf('/ByteRange'));
    const open = text.indexOf('<', text.indexOf('/Contents', brClose));
    const close = text.indexOf('>', open);
    expect(b).toBe(open + 1); // segment 1 ends just after the '<'
    expect(c).toBe(close); // segment 2 starts at the '>'
    expect(d).toBe(signed.length - close);

    // The signed byte range must NOT include any of the hex hole.
    expect(a + b).toBeLessThanOrEqual(open + 1);
    expect(c).toBeGreaterThanOrEqual(close);

    // Recompute SHA-256 over the two covered segments and confirm it equals the
    // messageDigest carried in the embedded CMS → the signature covers the file.
    const covered = new Uint8Array(b + d);
    covered.set(signed.subarray(a, a + b), 0);
    covered.set(signed.subarray(c, c + d), b);
    const digest = new Uint8Array(await subtle.digest('SHA-256', covered));

    const cmsHex = text.slice(open + 1, close).replace(/0+$/, '');
    const cmsBytes = hexToBytes(cmsHex.length % 2 ? cmsHex.slice(0, -1) : cmsHex);
    const cmsLen = asn1.readTlv(cmsBytes, 0).end; // trim trailing zero padding
    const cms = cmsBytes.slice(0, cmsLen);
    expect(messageDigestOf(cms)).toEqual([...digest]);
  });

  it('is deterministic for a fixed signing time (RSA PKCS#1 v1.5)', async () => {
    const privateKey = await importKey();
    const t = new Date('2026-01-02T03:04:05Z');
    const one = await signPdf(placeholderPdf(), { certificate: CERT, privateKey, signingTime: t });
    const two = await signPdf(placeholderPdf(), { certificate: CERT, privateKey, signingTime: t });
    expect(latin1.decode(two)).toBe(latin1.decode(one));
  });

  it('rejects when the signature does not fit the reserved /Contents', async () => {
    const privateKey = await importKey();
    await expect(signPdf(placeholderPdf(64), { certificate: CERT, privateKey })).rejects.toThrow(
      /exceeds reserved/,
    );
  });

  it('signs end-to-end through convertDocxToPdf (fonts supplied → no network)', async () => {
    const privateKey = await importKey();
    const docx = buildDocxFromBody('<w:p><w:r><w:t>Hello signed</w:t></w:r></w:p>');
    const signed = await convertDocxToPdf(docx, {
      fonts: { regular: REGULAR },
      signature: {
        certificate: CERT,
        privateKey,
        reason: 'Approved',
        signingTime: new Date('2026-01-02T03:04:05Z'),
      },
    });
    const text = latin1.decode(signed);
    expect(text).toContain('/SubFilter /adbe.pkcs7.detached');
    // ByteRange is filled with real offsets (placeholder zeros are gone).
    expect(text).toMatch(/\/ByteRange \[0 \d+ \d+ \d+/);
    expect(text).not.toContain('[0 0000000000 0000000000 0000000000]');
  });
});

// Pull the messageDigest attribute octets out of a CMS SignedData.
function messageDigestOf(cms: Uint8Array): Array<number> {
  const ci = asn1.readTlv(cms, 0);
  const sd = asn1.readTlv(cms, asn1.children(cms, ci)[1]!.contentStart);
  const sdKids = asn1.children(cms, sd);
  const signerInfos = sdKids.filter((k) => k.tag === 0x31).at(-1)!;
  const si = asn1.readTlv(cms, signerInfos.contentStart);
  const signedAttrs = asn1.children(cms, si).find((k) => k.tag === 0xa0)!;
  for (const a of asn1.children(cms, signedAttrs)) {
    const ak = asn1.children(cms, a);
    const oid = cms.slice(ak[0]!.contentStart, ak[0]!.contentEnd);
    if (oid.at(-1) === 0x04 && oid.at(-2) === 0x09) {
      const octet = asn1.readTlv(cms, ak[1]!.contentStart);
      return [...cms.slice(octet.contentStart, octet.contentEnd)];
    }
  }
  throw new Error('no messageDigest attribute');
}
