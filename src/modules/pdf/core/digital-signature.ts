/**
 * PDF digital signature — verification and creation.
 *
 * Implements:
 * - ASN.1 DER decode/encode (shared codec)
 * - PKCS#7 / CMS SignedData parse and build
 * - X.509 certificate public key extraction
 * - PDF /ByteRange extraction and hash computation
 * - Signature verification (RSA PKCS#1 v1.5 + SHA-256)
 * - Signature creation (with ByteRange placeholder/backfill)
 *
 * Uses platform-native RSA via `@utils/crypto` (node:crypto on Node,
 * Web Crypto API in browsers).
 *
 * @see RFC 5652 — CMS (Cryptographic Message Syntax)
 * @see ITU-T X.690 — ASN.1 DER encoding rules
 * @see ISO 32000-2:2020 §12.8 — Digital Signatures in PDF
 */

import {
  TAG_INTEGER,
  TAG_NULL,
  TAG_OCTET_STRING,
  TAG_OID,
  TAG_SEQUENCE,
  TAG_SET,
  asn1Decode,
  asn1Parse
} from "@pdf/core/asn1";
import type { Asn1Node } from "@pdf/core/asn1";
import { PdfSignaturePlaceholderMissingError } from "@pdf/errors";
import { sha256, md5, hash, rsaVerify, rsaSign } from "@utils/crypto";
import { encodeHex } from "@utils/hex";

export { asn1Parse } from "@pdf/core/asn1";
export type { Asn1Node } from "@pdf/core/asn1";

/**
 * Parse all ASN.1 DER elements from a buffer (for SEQUENCE content with multiple children).
 */
function asn1ParseAll(data: Uint8Array): Asn1Node[] {
  const nodes: Asn1Node[] = [];
  let offset = 0;
  while (offset < data.length) {
    const result = asn1Decode(data, offset);
    nodes.push(result.node);
    offset = result.end;
  }
  return nodes;
}

// =============================================================================
// ASN.1 DER — Encode
// =============================================================================

/**
 * Encode an ASN.1 length in DER format.
 */
function asn1EncodeLength(length: number): Uint8Array {
  if (length < 0x80) {
    return new Uint8Array([length]);
  }
  const bytes: number[] = [];
  let l = length;
  while (l > 0) {
    bytes.unshift(l & 0xff);
    l >>= 8;
  }
  return new Uint8Array([0x80 | bytes.length, ...bytes]);
}

/**
 * Encode an ASN.1 TLV (tag-length-value).
 */
function asn1Encode(tag: number, value: Uint8Array): Uint8Array {
  const length = asn1EncodeLength(value.length);
  const result = new Uint8Array(1 + length.length + value.length);
  result[0] = tag;
  result.set(length, 1);
  result.set(value, 1 + length.length);
  return result;
}

/**
 * Encode a SEQUENCE.
 */
function asn1Sequence(...children: Uint8Array[]): Uint8Array {
  let totalLen = 0;
  for (const c of children) {
    totalLen += c.length;
  }
  const body = new Uint8Array(totalLen);
  let offset = 0;
  for (const c of children) {
    body.set(c, offset);
    offset += c.length;
  }
  return asn1Encode(TAG_SEQUENCE, body);
}

/**
 * Encode a SET.
 */
function asn1Set(...children: Uint8Array[]): Uint8Array {
  let totalLen = 0;
  for (const c of children) {
    totalLen += c.length;
  }
  const body = new Uint8Array(totalLen);
  let offset = 0;
  for (const c of children) {
    body.set(c, offset);
    offset += c.length;
  }
  return asn1Encode(TAG_SET, body);
}

/**
 * Encode an OID.
 */
function asn1Oid(oid: string): Uint8Array {
  const parts = oid.split(".").map(Number);
  const bytes: number[] = [40 * parts[0] + parts[1]];
  for (let i = 2; i < parts.length; i++) {
    let v = parts[i];
    if (v < 128) {
      bytes.push(v);
    } else {
      const enc: number[] = [];
      enc.push(v & 0x7f);
      v >>= 7;
      while (v > 0) {
        enc.push(0x80 | (v & 0x7f));
        v >>= 7;
      }
      enc.reverse();
      bytes.push(...enc);
    }
  }
  return asn1Encode(TAG_OID, new Uint8Array(bytes));
}

/**
 * Encode an INTEGER (unsigned, from bytes).
 */
function asn1Integer(value: Uint8Array): Uint8Array {
  // Prepend 0x00 if high bit is set (positive integer)
  if (value.length > 0 && value[0] & 0x80) {
    const padded = new Uint8Array(value.length + 1);
    padded.set(value, 1);
    return asn1Encode(TAG_INTEGER, padded);
  }
  return asn1Encode(TAG_INTEGER, value);
}

/**
 * Encode an OCTET STRING.
 */
function asn1OctetString(value: Uint8Array): Uint8Array {
  return asn1Encode(TAG_OCTET_STRING, value);
}

/**
 * Encode a context-tagged explicit wrapper [N] EXPLICIT.
 */
function asn1ContextExplicit(tagNum: number, value: Uint8Array): Uint8Array {
  return asn1Encode(0xa0 | tagNum, value);
}

// =============================================================================
// OID Constants
// =============================================================================

const OID_PKCS7_SIGNED_DATA = "1.2.840.113549.1.7.2";
const OID_PKCS7_DATA = "1.2.840.113549.1.7.1";
const OID_SHA256 = "2.16.840.1.101.3.4.2.1";
const OID_SHA256_WITH_RSA = "1.2.840.113549.1.1.11";
const OID_CONTENT_TYPE = "1.2.840.113549.1.9.3";
const OID_MESSAGE_DIGEST = "1.2.840.113549.1.9.4";
const OID_SIGNING_TIME = "1.2.840.113549.1.9.5";

// =============================================================================
// OID Helpers
// =============================================================================

/**
 * Decode an OID from DER bytes to dotted string.
 */
function decodeOid(bytes: Uint8Array): string {
  if (bytes.length === 0) {
    return "";
  }
  const parts: number[] = [Math.floor(bytes[0] / 40), bytes[0] % 40];
  let value = 0;
  for (let i = 1; i < bytes.length; i++) {
    value = (value << 7) | (bytes[i] & 0x7f);
    if ((bytes[i] & 0x80) === 0) {
      parts.push(value);
      value = 0;
    }
  }
  return parts.join(".");
}

// =============================================================================
// X.509 Certificate — Public Key Extraction
// =============================================================================

/**
 * Extract the SubjectPublicKeyInfo (SPKI) DER bytes from an X.509 certificate.
 * This is what platform RSA verify APIs expect.
 */
function extractSpkiFromCert(certDer: Uint8Array): Uint8Array {
  const cert = asn1Parse(certDer);
  // Certificate → TBSCertificate → SubjectPublicKeyInfo
  // TBSCertificate is the first child of Certificate (SEQUENCE)
  const tbs = cert.children[0];
  if (!tbs) {
    throw new Error("Invalid X.509 certificate: missing TBSCertificate");
  }

  // SubjectPublicKeyInfo is at index 6 of TBSCertificate (after version, serial,
  // signature alg, issuer, validity, subject). If there's an explicit [0] version
  // tag, it shifts indices by 1.
  let spkiIndex = 5; // without explicit version
  if (tbs.children.length > 0 && (tbs.children[0].tag & 0xe0) === 0xa0) {
    spkiIndex = 6; // with explicit version [0]
  }

  const spki = tbs.children[spkiIndex];
  if (!spki || spki.tag !== TAG_SEQUENCE) {
    throw new Error("Invalid X.509 certificate: missing SubjectPublicKeyInfo");
  }

  // Re-encode the SPKI node as DER
  return asn1Encode(spki.tag, spki.bytes);
}

/** The TBSCertificate fields this module reads, by name. */
interface TbsFields {
  serial: Asn1Node;
  issuer: Asn1Node;
  validity: Asn1Node;
  subject: Asn1Node;
  extensions: Asn1Node | undefined;
}

/** Split a certificate's TBSCertificate into the fields this module reads. */
function tbsFields(certDer: Uint8Array): TbsFields | undefined {
  const tbs = asn1Parse(certDer).children[0];
  if (!tbs) {
    return undefined;
  }
  // [0] version?, serial, signature, issuer, validity, subject, spki, …, [3] extensions?
  const at = tbs.children.length > 0 && (tbs.children[0].tag & 0xe0) === 0xa0 ? 1 : 0;
  const [serial, , issuer, validity, subject] = tbs.children.slice(at);
  if (!serial || !issuer || !validity || !subject) {
    return undefined;
  }
  const extensions = tbs.children.find(child => child.tag === 0xa3);
  return { serial, issuer, validity, subject, extensions };
}

/**
 * Decode an X.520 `DirectoryString` (or one of the legacy string types a certificate may
 * use instead) to text.
 */
function decodeDirectoryString(node: Asn1Node): string {
  const bytes = node.bytes;
  switch (node.tag) {
    case 0x1e: {
      // BMPString — UTF-16BE
      let out = "";
      for (let i = 0; i + 1 < bytes.length; i += 2) {
        out += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
      }
      return out;
    }
    case 0x1c: {
      // UniversalString — UTF-32BE
      let out = "";
      for (let i = 0; i + 3 < bytes.length; i += 4) {
        const cp =
          ((bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3]) >>> 0;
        out += cp <= 0x10ffff ? String.fromCodePoint(cp) : "\ufffd";
      }
      return out;
    }
    case 0x14: // TeletexString (T61) — in practice Latin-1
      return new TextDecoder("latin1").decode(bytes);
    default:
      // UTF8String, PrintableString, IA5String, VisibleString
      return new TextDecoder().decode(bytes);
  }
}

/**
 * The subject common name (`CN`, OID 2.5.4.3) of an X.509 certificate, or `undefined`
 * when the certificate has none or cannot be read.
 */
export function certificateCommonName(certDer: Uint8Array): string | undefined {
  try {
    const subject = tbsFields(certDer)?.subject;
    if (!subject || subject.tag !== TAG_SEQUENCE) {
      return undefined;
    }
    for (const rdn of subject.children) {
      for (const attribute of rdn.children) {
        const [oid, value] = attribute.children;
        if (oid && value && oid.tag === TAG_OID && decodeOid(oid.bytes) === "2.5.4.3") {
          return decodeDirectoryString(value);
        }
      }
    }
  } catch {
    // An unreadable certificate has no name to report.
  }
  return undefined;
}

/** Parse an ASN.1 `UTCTime` or `GeneralizedTime` (UTC forms only). */
function decodeAsn1Time(node: Asn1Node): Date | undefined {
  const text = new TextDecoder("latin1").decode(node.bytes);
  const match =
    node.tag === 0x17
      ? /^(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?Z$/.exec(text)
      : node.tag === 0x18
        ? /^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?(?:\.\d+)?Z$/.exec(text)
        : null;
  if (!match) {
    return undefined;
  }
  let year = Number(match[1]);
  if (node.tag === 0x17) {
    // RFC 5280 §4.1.2.5.1: 50–99 → 19xx, 00–49 → 20xx
    year += year >= 50 ? 1900 : 2000;
  }
  const date = new Date(
    Date.UTC(
      year,
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4]),
      Number(match[5]),
      Number(match[6] ?? 0)
    )
  );
  return Number.isNaN(date.getTime()) ? undefined : date;
}

/**
 * The validity period (`notBefore`, `notAfter`) of an X.509 certificate, or `undefined`
 * when it cannot be read.
 */
export function certificateValidity(
  certDer: Uint8Array
): { notBefore: Date; notAfter: Date } | undefined {
  try {
    const validity = tbsFields(certDer)?.validity;
    const [from, to] = validity?.children ?? [];
    const notBefore = from && decodeAsn1Time(from);
    const notAfter = to && decodeAsn1Time(to);
    return notBefore && notAfter ? { notBefore, notAfter } : undefined;
  } catch {
    return undefined;
  }
}

/** The `subjectKeyIdentifier` extension (OID 2.5.29.14) of a certificate. */
function subjectKeyIdentifier(fields: TbsFields): Uint8Array | undefined {
  for (const extension of fields.extensions?.children[0]?.children ?? []) {
    const [oid] = extension.children;
    const value = extension.children[extension.children.length - 1];
    if (oid?.tag === TAG_OID && decodeOid(oid.bytes) === "2.5.29.14" && value) {
      const inner = asn1Parse(value.bytes);
      return inner.tag === TAG_OCTET_STRING ? inner.bytes : undefined;
    }
  }
  return undefined;
}

function sameNode(a: Asn1Node, b: Asn1Node): boolean {
  return a.tag === b.tag && bytesEqual(a.bytes, b.bytes);
}

/**
 * The certificate a SignerInfo's `sid` names: by issuer and serial number, or by
 * subject key identifier (RFC 5652 §5.3). `undefined` when no embedded certificate
 * matches — the signature then names a key the file does not carry.
 */
function findSignerCertificate(certificates: Uint8Array[], sid: Asn1Node): Uint8Array | undefined {
  for (const certificate of certificates) {
    let fields: TbsFields | undefined;
    try {
      fields = tbsFields(certificate);
    } catch {
      continue;
    }
    if (!fields) {
      continue;
    }
    if (sid.tag === TAG_SEQUENCE) {
      const [issuer, serial] = sid.children;
      if (issuer && serial && sameNode(issuer, fields.issuer) && sameNode(serial, fields.serial)) {
        return certificate;
      }
    } else if (sid.tag === 0x80) {
      let ski: Uint8Array | undefined;
      try {
        ski = subjectKeyIdentifier(fields);
      } catch {
        ski = undefined;
      }
      if (ski && bytesEqual(ski, sid.bytes)) {
        return certificate;
      }
    }
  }
  return undefined;
}

// =============================================================================
// PKCS#7 / CMS SignedData — Parse
// =============================================================================

/** Parsed CMS SignedData info for verification. */
export interface CmsSignedData {
  /**
   * The signer's certificate (DER): the embedded certificate the first SignerInfo names
   * by issuer and serial number or by subject key identifier — not merely the first
   * certificate in the set, which may belong to an issuer.
   */
  certificate: Uint8Array;
  /** The signature value. */
  signature: Uint8Array;
  /** The digest algorithm OID. */
  digestAlgorithmOid: string;
  /** The signed attributes (DER-encoded SET for hash computation). */
  signedAttrsRaw: Uint8Array;
  /** The message digest from signed attributes. */
  messageDigest: Uint8Array;
}

/**
 * Parse a PKCS#7 / CMS SignedData structure from DER bytes.
 * Extracts the first signer's info for verification.
 *
 * @throws when no embedded certificate matches the signer's identifier.
 */
export function parseCmsSignedData(derBytes: Uint8Array): CmsSignedData {
  const root = asn1Parse(derBytes);

  // ContentInfo: SEQUENCE { contentType OID, content [0] EXPLICIT }
  if (root.tag !== TAG_SEQUENCE || root.children.length < 2) {
    throw new Error("Invalid PKCS#7: not a ContentInfo SEQUENCE");
  }

  const contentTypeNode = root.children[0];
  const oid = decodeOid(contentTypeNode.bytes);
  if (oid !== OID_PKCS7_SIGNED_DATA) {
    throw new Error(`Invalid PKCS#7: expected SignedData OID, got ${oid}`);
  }

  // content [0] EXPLICIT → SignedData SEQUENCE
  const contentWrapper = root.children[1];
  const signedData = contentWrapper.children[0];
  if (!signedData || signedData.tag !== TAG_SEQUENCE) {
    throw new Error("Invalid PKCS#7: missing SignedData SEQUENCE");
  }

  // SignedData: version, digestAlgorithms, encapContentInfo, [0] certificates, [1] crls, signerInfos
  const children = signedData.children;

  // Find certificates [0] IMPLICIT
  const certificates: Uint8Array[] = [];
  for (const child of children) {
    if ((child.tag & 0xf0) === 0xa0 && (child.tag & 0x0f) === 0) {
      for (const certNode of child.children) {
        if (certNode.tag === TAG_SEQUENCE) {
          certificates.push(asn1Encode(certNode.tag, certNode.bytes));
        }
      }
      break;
    }
  }
  if (certificates.length === 0) {
    throw new Error("PKCS#7: no certificate found");
  }

  // Find signerInfos SET (last SET in SignedData)
  let signerInfosSet: Asn1Node | null = null;
  for (let i = children.length - 1; i >= 0; i--) {
    if (children[i].tag === TAG_SET) {
      signerInfosSet = children[i];
      break;
    }
  }
  if (!signerInfosSet || signerInfosSet.children.length === 0) {
    throw new Error("PKCS#7: no signerInfos found");
  }

  const signerInfo = signerInfosSet.children[0];
  // SignerInfo: version, sid, digestAlgorithm, [0] signedAttrs, signatureAlgorithm, signature
  const siChildren = signerInfo.children;

  const sid = siChildren[1];
  const certificate = sid ? findSignerCertificate(certificates, sid) : undefined;
  if (!certificate) {
    throw new Error("PKCS#7: no embedded certificate matches the signer identifier");
  }

  // digestAlgorithm
  const digestAlgSeq = siChildren[2];
  const digestAlgOid = digestAlgSeq
    ? decodeOid(digestAlgSeq.children[0]?.bytes ?? new Uint8Array())
    : "";

  // signedAttrs [0] IMPLICIT
  let signedAttrsRaw: Uint8Array = new Uint8Array();
  let messageDigest: Uint8Array = new Uint8Array();
  for (const child of siChildren) {
    if ((child.tag & 0xf0) === 0xa0 && (child.tag & 0x0f) === 0) {
      // Re-encode as SET OF for hash computation (per CMS spec §5.4)
      signedAttrsRaw = asn1Encode(TAG_SET, child.bytes);

      // Extract messageDigest attribute
      const attrs = asn1ParseAll(child.bytes);
      for (const attr of attrs) {
        if (attr.tag !== TAG_SEQUENCE || attr.children.length < 2) {
          continue;
        }
        const attrOid = decodeOid(attr.children[0].bytes);
        if (attrOid === OID_MESSAGE_DIGEST) {
          const attrValueSet = attr.children[1];
          if (attrValueSet.children.length > 0) {
            messageDigest = new Uint8Array(attrValueSet.children[0].bytes);
          }
        }
      }
      break;
    }
  }

  // signature — last OCTET_STRING in signerInfo
  let signature: Uint8Array = new Uint8Array();
  for (let i = siChildren.length - 1; i >= 0; i--) {
    if (siChildren[i].tag === TAG_OCTET_STRING) {
      signature = new Uint8Array(siChildren[i].bytes);
      break;
    }
  }

  return {
    certificate,
    signature,
    digestAlgorithmOid: digestAlgOid,
    signedAttrsRaw,
    messageDigest
  };
}

// =============================================================================
// PKCS#7 / CMS SignedData — Build
// =============================================================================

/** Options for building a CMS SignedData for PDF signing. */
export interface SignOptions {
  /** DER-encoded X.509 certificate. */
  certificate: Uint8Array;
  /** DER-encoded PKCS#8 private key. */
  privateKey: Uint8Array;
  /** The data to sign (the PDF byte ranges). */
  data: Uint8Array;
}

/**
 * Build a CMS SignedData (PKCS#7) structure for a PDF signature.
 *
 * Uses SHA-256 for digest and RSA PKCS#1 v1.5 for signing.
 * The signature is created over signed attributes that include
 * the content-type, message-digest, and signing-time.
 */
export async function buildCmsSignedData(options: SignOptions): Promise<Uint8Array> {
  const { certificate, privateKey, data } = options;

  // Compute message digest
  const digest = sha256(data);

  // Build signed attributes
  const now = new Date();
  const signingTimeStr = formatUtcTime(now);

  const contentTypeAttr = asn1Sequence(asn1Oid(OID_CONTENT_TYPE), asn1Set(asn1Oid(OID_PKCS7_DATA)));
  const messageDigestAttr = asn1Sequence(
    asn1Oid(OID_MESSAGE_DIGEST),
    asn1Set(asn1OctetString(digest))
  );
  const signingTimeAttr = asn1Sequence(
    asn1Oid(OID_SIGNING_TIME),
    asn1Set(asn1Encode(0x17, new TextEncoder().encode(signingTimeStr))) // UTCTime
  );

  // Signed attrs as SET for DER encoding
  const signedAttrsContent = concatDer(contentTypeAttr, signingTimeAttr, messageDigestAttr);
  const signedAttrsForHash = asn1Encode(TAG_SET, signedAttrsContent);

  // Sign the signed attributes
  const signatureBytes = await rsaSign(privateKey, signedAttrsForHash);

  // Build signed attrs as [0] IMPLICIT for embedding in SignerInfo
  const signedAttrsImplicit = asn1Encode(0xa0, signedAttrsContent);

  // Extract issuer and serial from certificate for SignerIdentifier
  const cert = asn1Parse(certificate);
  const tbs = cert.children[0];
  let issuer: Uint8Array;
  let serial: Uint8Array;

  if (tbs.children[0] && (tbs.children[0].tag & 0xe0) === 0xa0) {
    // Has explicit version
    serial = asn1Encode(tbs.children[1].tag, tbs.children[1].bytes);
    issuer = asn1Encode(tbs.children[3].tag, tbs.children[3].bytes);
  } else {
    serial = asn1Encode(tbs.children[0].tag, tbs.children[0].bytes);
    issuer = asn1Encode(tbs.children[2].tag, tbs.children[2].bytes);
  }

  // SignerInfo
  const signerInfo = asn1Sequence(
    asn1Integer(new Uint8Array([1])), // version 1
    asn1Sequence(issuer, serial), // issuerAndSerialNumber
    asn1Sequence(asn1Oid(OID_SHA256), asn1Encode(TAG_NULL, new Uint8Array())), // digestAlgorithm
    signedAttrsImplicit, // signedAttrs [0] IMPLICIT
    asn1Sequence(asn1Oid(OID_SHA256_WITH_RSA), asn1Encode(TAG_NULL, new Uint8Array())), // signatureAlgorithm
    asn1OctetString(signatureBytes) // signature
  );

  // SignedData
  const signedData = asn1Sequence(
    asn1Integer(new Uint8Array([1])), // version 1
    asn1Set(asn1Sequence(asn1Oid(OID_SHA256), asn1Encode(TAG_NULL, new Uint8Array()))), // digestAlgorithms
    asn1Sequence(asn1Oid(OID_PKCS7_DATA)), // encapContentInfo (detached — no eContent)
    asn1ContextExplicit(0, asn1Encode(cert.tag, cert.bytes)), // certificates [0]
    asn1Set(signerInfo) // signerInfos
  );

  // ContentInfo wrapper
  return asn1Sequence(asn1Oid(OID_PKCS7_SIGNED_DATA), asn1ContextExplicit(0, signedData));
}

// =============================================================================
// PDF Signature Verification
// =============================================================================

/** Result of verifying a PDF signature. */
/**
 * What `Pdf.Signature.verify` reports for one signature — the same field names as
 * `PdfSignatureReport`.
 */
export interface PdfSignatureCheck {
  /** Whether the signature is cryptographically intact over the given byte ranges. */
  intact: boolean;
  /** Whether the signed byte ranges cover the entire file (no unsigned gaps). */
  coversWholeFile: boolean;
  /** Digest algorithm OID the signature uses; empty when it could not be read. */
  digestAlgorithm: string;
  /** Why the signature is not {@link intact}, when it is not. */
  failure?: string;
}

/** What the deprecated `Pdf.verifySignature` returns; see {@link PdfSignatureCheck}. */
export interface SignatureVerificationResult {
  /** Whether the signature is cryptographically valid. */
  valid: boolean;
  /** Whether the signed byte ranges cover the entire file (no unsigned gaps). */
  coversWholeFile: boolean;
  /** Digest algorithm used. */
  digestAlgorithm: string;
  /** Reason for failure, if any. */
  reason?: string;
}

/**
 * Verify a digital signature in a PDF document.
 *
 * @param pdfData - The complete PDF file bytes
 * @param signatureHex - The hex-encoded PKCS#7 signature from the /Contents field
 * @param byteRange - The /ByteRange array [offset1, length1, offset2, length2]
 */
export async function verifyPdfSignature(
  pdfData: Uint8Array,
  signatureHex: string,
  byteRange: [number, number, number, number]
): Promise<SignatureVerificationResult> {
  let sigBytes: Uint8Array;
  try {
    sigBytes = hexToBytes(signatureHex);
  } catch (err) {
    return {
      valid: false,
      coversWholeFile: false,
      digestAlgorithm: "",
      reason: `Signature verification error: ${err instanceof Error ? err.message : String(err)}`
    };
  }
  return verifyPdfSignatureBytes(pdfData, sigBytes, byteRange);
}

/** {@link verifyPdfSignature} for a signature already decoded from hex. */
export async function verifyPdfSignatureBytes(
  pdfData: Uint8Array,
  sigBytes: Uint8Array,
  byteRange: [number, number, number, number]
): Promise<SignatureVerificationResult> {
  try {
    const cms = parseCmsSignedData(sigBytes);

    // Extract the signed byte ranges from the PDF
    const [off1, len1, off2, len2] = byteRange;
    const range1 = pdfData.subarray(off1, off1 + len1);
    const range2 = pdfData.subarray(off2, off2 + len2);
    const signedData = new Uint8Array(len1 + len2);
    signedData.set(range1);
    signedData.set(range2, len1);

    // Verify message digest using the algorithm from the signature
    const computedDigest = hashByOid(cms.digestAlgorithmOid, signedData);
    if (!bytesEqual(computedDigest, cms.messageDigest)) {
      return {
        valid: false,
        coversWholeFile: checkCoversWholeFile(byteRange, pdfData.length),
        digestAlgorithm: cms.digestAlgorithmOid,
        reason: "Message digest mismatch — PDF content was modified after signing"
      };
    }

    // Verify RSA signature over signed attributes
    const spki = extractSpkiFromCert(cms.certificate);
    const valid = await rsaVerify(spki, cms.signature, cms.signedAttrsRaw);

    return {
      valid,
      coversWholeFile: checkCoversWholeFile(byteRange, pdfData.length),
      digestAlgorithm: cms.digestAlgorithmOid,
      reason: valid ? undefined : "RSA signature verification failed"
    };
  } catch (err) {
    return {
      valid: false,
      coversWholeFile: false,
      digestAlgorithm: "",
      reason: `Signature verification error: ${err instanceof Error ? err.message : String(err)}`
    };
  }
}

// =============================================================================
// PDF Signature Creation — ByteRange Placeholder
// =============================================================================

/**
 * Estimated maximum size (in bytes) for the PKCS#7 signature hex string.
 * A 2048-bit RSA signature with certificate is typically ~3000 bytes DER,
 * which is ~6000 hex chars. We use 8192 to be safe.
 */
const SIGNATURE_PLACEHOLDER_SIZE = 8192;

/**
 * Create a PDF signature dictionary string with a placeholder /Contents.
 * Returns the dict string and the placeholder that will be replaced.
 *
 * @param signerName - Optional signer name for /Name field
 * @param reason - Optional reason for /Reason field
 */
export function buildSignatureDictPlaceholder(options?: {
  name?: string;
  reason?: string;
  location?: string;
  contactInfo?: string;
}): { dictString: string; placeholder: string } {
  const placeholder = "0".repeat(SIGNATURE_PLACEHOLDER_SIZE * 2); // hex chars

  let dict = "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached";
  dict += ` /Contents <${placeholder}>`;
  dict += " /ByteRange [0 0000000000 0000000000 0000000000]"; // placeholder, will be patched

  if (options?.name) {
    dict += ` /Name (${escPdfString(options.name)})`;
  }
  if (options?.reason) {
    dict += ` /Reason (${escPdfString(options.reason)})`;
  }
  if (options?.location) {
    dict += ` /Location (${escPdfString(options.location)})`;
  }
  if (options?.contactInfo) {
    dict += ` /ContactInfo (${escPdfString(options.contactInfo)})`;
  }

  // Add /M (signing time)
  const now = new Date();
  const m = formatPdfDate(now);
  dict += ` /M (${m})`;
  dict += " >>";

  return { dictString: dict, placeholder };
}

/** Signing material for `Pdf.Signature.fillPlaceholder`. */
export interface SignPdfOptions {
  /** DER-encoded X.509 certificate. */
  certificate: Uint8Array;
  /** DER-encoded PKCS#8 private key. */
  privateKey: Uint8Array;
}

/**
 * Patch a PDF with a real signature after the /ByteRange placeholder has been written.
 *
 * @param pdfBytes - The PDF bytes with placeholder /Contents and /ByteRange
 * @param certificate - DER-encoded X.509 certificate
 * @param privateKey - DER-encoded PKCS#8 private key
 * @returns The signed PDF bytes
 */
export async function signPdf(
  pdfBytes: Uint8Array,
  certificate: Uint8Array,
  privateKey: Uint8Array
): Promise<Uint8Array> {
  const result = new Uint8Array(pdfBytes);

  // Find /ByteRange first — this uniquely identifies the signature dictionary
  const byteRangePattern = findPattern(result, "/ByteRange [");
  if (byteRangePattern === -1) {
    throw new PdfSignaturePlaceholderMissingError(
      "The PDF has no signature placeholder: no /ByteRange [...] was found."
    );
  }

  // Search for /Contents <hex> near /ByteRange (within the same object, search backwards)
  // The signature dict typically has /Contents before /ByteRange, but search both directions
  const searchStart = Math.max(0, byteRangePattern - 20000); // signature hex can be ~16K
  const searchEnd = Math.min(result.length, byteRangePattern + 200);
  let contentsPattern = -1;
  const contentsBytes = new TextEncoder().encode("/Contents <");
  for (let i = searchStart; i < searchEnd; i++) {
    let matched = true;
    for (let j = 0; j < contentsBytes.length; j++) {
      if (result[i + j] !== contentsBytes[j]) {
        matched = false;
        break;
      }
    }
    if (matched) {
      contentsPattern = i;
      break;
    }
  }
  if (contentsPattern === -1) {
    throw new PdfSignaturePlaceholderMissingError(
      "The PDF has no signature placeholder: no /Contents <...> was found near its /ByteRange."
    );
  }

  const hexStart = contentsPattern + "/Contents <".length;
  // Find the closing >
  let hexEnd = hexStart;
  while (hexEnd < result.length && result[hexEnd] !== 0x3e /* > */) {
    hexEnd++;
  }

  const brStart = byteRangePattern + "/ByteRange [".length;
  let brEnd = brStart;
  while (brEnd < result.length && result[brEnd] !== 0x5d /* ] */) {
    brEnd++;
  }

  // Compute actual byte range: before <hex> and after <hex>
  const sigDictContentsStart = hexStart - 1; // position of <
  const sigDictContentsEnd = hexEnd + 1; // position after >

  const byteRange: [number, number, number, number] = [
    0,
    sigDictContentsStart,
    sigDictContentsEnd,
    result.length - sigDictContentsEnd
  ];

  // Patch the ByteRange value
  const brValue = `${byteRange[0]} ${byteRange[1]} ${byteRange[2]} ${byteRange[3]}`;
  const brPadded = brValue.padEnd(brEnd - brStart, " ");
  for (let i = 0; i < brPadded.length; i++) {
    result[brStart + i] = brPadded.charCodeAt(i);
  }

  // Compute the signed data from byte ranges
  const range1 = result.subarray(byteRange[0], byteRange[0] + byteRange[1]);
  const range2 = result.subarray(byteRange[2], byteRange[2] + byteRange[3]);
  const signedData = new Uint8Array(byteRange[1] + byteRange[3]);
  signedData.set(range1);
  signedData.set(range2, byteRange[1]);

  // Build CMS SignedData
  const cms = await buildCmsSignedData({ certificate, privateKey, data: signedData });

  // Hex-encode the signature
  const hexSig = encodeHex(cms).padEnd(hexEnd - hexStart, "0");
  for (let i = 0; i < hexSig.length && i < hexEnd - hexStart; i++) {
    result[hexStart + i] = hexSig.charCodeAt(i);
  }

  return result;
}

// =============================================================================
// Helpers
// =============================================================================

function hexToBytes(hex: string): Uint8Array {
  const clean = hex.replace(/\s/g, "");
  const bytes = new Uint8Array(clean.length / 2);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = parseInt(clean.substring(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a[i] ^ b[i];
  }
  return diff === 0;
}

function checkCoversWholeFile(
  byteRange: [number, number, number, number],
  fileSize: number
): boolean {
  // The two ranges should cover everything except the /Contents hex value
  return byteRange[0] === 0 && byteRange[2] + byteRange[3] === fileSize;
}

/** Map digest algorithm OID to a hash function. Falls back to sha256 for unknown OIDs. */
function hashByOid(oid: string, data: Uint8Array): Uint8Array {
  switch (oid) {
    case "1.3.14.3.2.26": // SHA-1
      return hash("SHA-1", data);
    case OID_SHA256: // SHA-256
      return sha256(data);
    case "2.16.840.1.101.3.4.2.2": // SHA-384
      return hash("SHA-384", data);
    case "2.16.840.1.101.3.4.2.3": // SHA-512
      return hash("SHA-512", data);
    case "1.2.840.113549.2.5": // MD5
      return md5(data);
    default:
      // Fallback to SHA-256 for unrecognized OIDs
      return sha256(data);
  }
}

function findPattern(data: Uint8Array, pattern: string): number {
  const patBytes = new TextEncoder().encode(pattern);
  for (let i = 0; i <= data.length - patBytes.length; i++) {
    let matched = true;
    for (let j = 0; j < patBytes.length; j++) {
      if (data[i + j] !== patBytes[j]) {
        matched = false;
        break;
      }
    }
    if (matched) {
      return i;
    }
  }
  return -1;
}

function concatDer(...parts: Uint8Array[]): Uint8Array {
  let totalLen = 0;
  for (const p of parts) {
    totalLen += p.length;
  }
  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (const p of parts) {
    result.set(p, offset);
    offset += p.length;
  }
  return result;
}

function escPdfString(s: string): string {
  return s.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function formatUtcTime(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  const yr = String(d.getUTCFullYear()).slice(-2);
  return `${yr}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

function formatPdfDate(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `D:${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}
