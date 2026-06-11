// RFC 5652 — Cryptographic Message Syntax (PKCS#7). Builds the detached
// SignedData blob that goes into a PDF signature dictionary's /Contents with
// SubFilter /adbe.pkcs7.detached: the message digest is the SHA-256 of the
// signed PDF byte range, and the data itself is NOT carried in the CMS.

import * as der from '@/core/crypto/asn1';

const OID = {
  sha256: '2.16.840.1.101.3.4.2.1',
  rsaEncryption: '1.2.840.113549.1.1.1',
  ecdsaWithSHA256: '1.2.840.10045.4.3.2',
  idData: '1.2.840.113549.1.7.1',
  idSignedData: '1.2.840.113549.1.7.2',
  contentType: '1.2.840.113549.1.9.3',
  messageDigest: '1.2.840.113549.1.9.4',
  signingTime: '1.2.840.113549.1.9.5',
} as const;

// AlgorithmIdentifier for SHA-256 carries ABSENT parameters (RFC 5754 §2);
// rsaEncryption carries explicit NULL parameters (RFC 3447); ecdsa-with-SHA256
// carries ABSENT parameters (RFC 5758 §3.2).
const sha256AlgId = (): Uint8Array => der.seq(der.oid(OID.sha256));
const rsaAlgId = (): Uint8Array => der.seq(der.oid(OID.rsaEncryption), der.nullValue());
const ecdsaAlgId = (): Uint8Array => der.seq(der.oid(OID.ecdsaWithSHA256));

const attribute = (oidStr: string, value: Uint8Array): Uint8Array =>
  der.seq(der.oid(oidStr), der.set(value));

export interface CmsParams {
  // Signer's X.509 certificate (DER). Included in the CMS so a verifier can
  // build the trust path; its issuer + serial identify the signer.
  readonly certificate: Uint8Array;
  // Optional additional chain certificates (DER), e.g. intermediates.
  readonly extraCertificates?: ReadonlyArray<Uint8Array>;
  // SHA-256 digest of the signed PDF byte range (the detached content).
  readonly messageDigest: Uint8Array;
  readonly signingTime: Date;
  // Public-key algorithm of the signer's key (default 'rsa'). Selects the
  // SignerInfo signatureAlgorithm (rsaEncryption vs ecdsa-with-SHA256); the
  // digest stays SHA-256 either way.
  readonly signatureAlgorithm?: 'rsa' | 'ecdsa';
  // Signs the DER of the signedAttrs over SHA-256. For 'rsa' this is
  // RSASSA-PKCS1-v1_5; for 'ecdsa' it must return a DER Ecdsa-Sig-Value
  // (SEQUENCE { r INTEGER, s INTEGER }), not the raw r‖s.
  readonly sign: (signedAttrsDer: Uint8Array) => Promise<Uint8Array>;
}

// Build a complete CMS ContentInfo (id-signedData) wrapping a single SignerInfo
// with authenticated attributes (contentType, messageDigest, signingTime).
export async function buildPkcs7Detached(params: CmsParams): Promise<Uint8Array> {
  const { issuer, serial } = der.certIssuerAndSerial(params.certificate);

  // signedAttrs — its canonical (DER SET OF) body is signed; the same body is
  // re-tagged [0] IMPLICIT inside the SignerInfo.
  const attrs = [
    attribute(OID.contentType, der.oid(OID.idData)),
    attribute(OID.messageDigest, der.octetString(params.messageDigest)),
    attribute(OID.signingTime, der.cmsTime(params.signingTime)),
  ];
  const attrsBody = der.setOfBody(attrs);
  const signedAttrsForSigning = der.tlv(0x31, attrsBody); // SET OF, the signed bytes
  const signedAttrsTagged = der.tlv(0xa0, attrsBody); // [0] IMPLICIT in SignerInfo
  const signature = await params.sign(signedAttrsForSigning);

  const sigAlgId = params.signatureAlgorithm === 'ecdsa' ? ecdsaAlgId() : rsaAlgId();
  const signerInfo = der.seq(
    der.integer(1), // version (issuerAndSerialNumber ⇒ 1)
    der.seq(issuer, serial), // sid: IssuerAndSerialNumber
    sha256AlgId(),
    signedAttrsTagged,
    sigAlgId,
    der.octetString(signature),
  );

  const certBlobs = [params.certificate, ...(params.extraCertificates ?? [])];
  const certificates = der.tlv(0xa0, der.concat(certBlobs)); // [0] IMPLICIT certificates

  const signedData = der.seq(
    der.integer(1), // version
    der.set(sha256AlgId()), // digestAlgorithms
    der.seq(der.oid(OID.idData)), // encapContentInfo (detached: no eContent)
    certificates,
    der.set(signerInfo), // signerInfos
  );

  return der.seq(der.oid(OID.idSignedData), der.explicit(0, signedData));
}
