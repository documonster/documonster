/**
 * ASN.1 DER decoding, and the structural check of signing material.
 *
 * Its own module so the builder can check a certificate and key when `sign()` is called
 * without statically pulling the CMS and RSA code of `digital-signature.ts`, which it loads
 * only when a document is actually signed.
 *
 * @see ITU-T X.690 — ASN.1 DER encoding rules
 */

import { PdfArgumentError, PdfStructureError } from "@pdf/errors";

// =============================================================================
// ASN.1 DER — Types
// =============================================================================

/** ASN.1 tag classes. */
export const ASN1_CONSTRUCTED = 0x20;

/** Common ASN.1 tags. */
export const TAG_INTEGER = 0x02;
export const TAG_OCTET_STRING = 0x04;
export const TAG_NULL = 0x05;
export const TAG_OID = 0x06;
export const TAG_SEQUENCE = 0x30;
export const TAG_SET = 0x31;

/** Parsed ASN.1 node. */
export interface Asn1Node {
  tag: number;
  /** Raw bytes of the value (for primitive types). */
  bytes: Uint8Array;
  /** Child nodes (for constructed types). */
  children: Asn1Node[];
}

// =============================================================================
// ASN.1 DER — Decode
// =============================================================================

/**
 * Decode a single ASN.1 DER element from `data` starting at `offset`.
 * Returns the parsed node and the offset after the element.
 */
export function asn1Decode(data: Uint8Array, offset: number): { node: Asn1Node; end: number } {
  if (offset + 2 > data.length) {
    throw new PdfStructureError("ASN.1: unexpected end of data");
  }

  const tag = data[offset++];
  let length = data[offset++];

  // Long-form length. DER has no indefinite form (`0x80`), and four length bytes already
  // exceed any buffer this can hold; reading either as a length used to produce a node
  // whose bytes were silently truncated.
  if (length & 0x80) {
    const numBytes = length & 0x7f;
    if (numBytes === 0 || numBytes > 4) {
      throw new PdfStructureError(
        numBytes === 0
          ? "ASN.1: indefinite length is not DER"
          : `ASN.1: a ${numBytes}-byte length is longer than any buffer`
      );
    }
    if (offset + numBytes > data.length) {
      throw new PdfStructureError("ASN.1: unexpected end of data in a length");
    }
    length = 0;
    for (let i = 0; i < numBytes; i++) {
      length = length * 256 + data[offset++];
    }
  }

  const valueStart = offset;
  const valueEnd = offset + length;
  if (valueEnd > data.length) {
    throw new PdfStructureError(
      `ASN.1: an element of ${length} bytes runs past the end of the data`
    );
  }
  const bytes = data.subarray(valueStart, valueEnd);

  const children: Asn1Node[] = [];
  const isConstructed = (tag & ASN1_CONSTRUCTED) !== 0;

  if (isConstructed) {
    let childOffset = valueStart;
    while (childOffset < valueEnd) {
      const result = asn1Decode(data.subarray(0, valueEnd), childOffset);
      children.push(result.node);
      childOffset = result.end;
    }
  }

  return { node: { tag, bytes, children }, end: valueEnd };
}

/**
 * Parse ASN.1 DER data from the root.
 */
export function asn1Parse(data: Uint8Array): Asn1Node {
  return asn1Decode(data, 0).node;
}

/** DER of the rsaEncryption OID, 1.2.840.113549.1.1.1. */
const RSA_ENCRYPTION_OID = [0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x01, 0x01];

/**
 * Throw {@link PdfArgumentError} unless `certificate` is a DER X.509 certificate and
 * `privateKey` a DER PKCS#8 RSA private key — what signing (RSA PKCS#1 v1.5 + SHA-256) needs.
 *
 * Checked when the material is given: other bytes used to pass `sign()` and fail inside
 * `toBuffer()` with a `TypeError` from the certificate walk, or a raw `node:crypto` error from
 * the key import. Only the structure is read; whether the key matches the certificate is
 * what signing itself finds out.
 */
export function checkSigningMaterial(
  where: string,
  certificate: Uint8Array,
  privateKey: Uint8Array
): void {
  const certificateProblem = (): string | undefined => {
    let cert: Asn1Node;
    try {
      cert = asn1Parse(certificate);
    } catch (err) {
      return err instanceof Error ? err.message : String(err);
    }
    const tbs = cert.children[0];
    if (cert.tag !== TAG_SEQUENCE || cert.children.length < 3 || tbs?.tag !== TAG_SEQUENCE) {
      return "it is not a Certificate SEQUENCE { tbsCertificate, signatureAlgorithm, signature }";
    }
    // tbsCertificate: [0] version (optional), serialNumber, signature, issuer, …
    const first = (tbs.children[0]?.tag ?? 0) & 0xe0;
    const serialAt = first === 0xa0 ? 1 : 0;
    if (
      tbs.children[serialAt]?.tag !== TAG_INTEGER ||
      tbs.children[serialAt + 2]?.tag !== TAG_SEQUENCE
    ) {
      return "its tbsCertificate has no serial number and issuer";
    }
    return undefined;
  };
  const keyProblem = (): string | undefined => {
    let key: Asn1Node;
    try {
      key = asn1Parse(privateKey);
    } catch (err) {
      return err instanceof Error ? err.message : String(err);
    }
    const [version, algorithm, body] = key.children;
    if (
      key.tag !== TAG_SEQUENCE ||
      version?.tag !== TAG_INTEGER ||
      algorithm?.tag !== TAG_SEQUENCE ||
      body?.tag !== TAG_OCTET_STRING
    ) {
      return "it is not a PKCS#8 PrivateKeyInfo SEQUENCE { version, algorithm, privateKey }";
    }
    const oid = algorithm.children[0];
    if (
      oid?.tag !== TAG_OID ||
      oid.bytes.length !== RSA_ENCRYPTION_OID.length ||
      oid.bytes.some((byte, i) => byte !== RSA_ENCRYPTION_OID[i])
    ) {
      return "it is not an RSA key, and signing uses RSA PKCS#1 v1.5";
    }
    return undefined;
  };
  const certificateError = certificateProblem();
  if (certificateError !== undefined) {
    throw new PdfArgumentError(
      "options.certificate",
      `${where}: options.certificate is not a DER X.509 certificate — ${certificateError}`
    );
  }
  const keyError = keyProblem();
  if (keyError !== undefined) {
    throw new PdfArgumentError(
      "options.privateKey",
      `${where}: options.privateKey is not a DER PKCS#8 RSA private key — ${keyError}`
    );
  }
}
