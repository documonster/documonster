/**
 * DOCX Digital Signatures (OPC Package Digital Signature)
 *
 * DOCX documents may contain digital signatures stored in:
 *   _xmlsignatures/origin.sigs
 *   _xmlsignatures/sig1.xml, sig2.xml, ...
 *   _xmlsignatures/_rels/origin.sigs.rels
 *
 * The signatures use W3C XML Digital Signature (XMLDSig) format with
 * Office-specific OfficeObject extensions per MS-OFFCRYPTO.
 *
 * This module provides:
 * - Detection and extraction of signature metadata
 * - Preservation of signatures through round-trip (via opaqueParts)
 *
 * Note: This module does NOT validate signature integrity (would require
 * full XMLDSig + Canonical XML implementation). Signatures are preserved
 * verbatim but become invalid if the package contents change.
 *
 * References:
 *   - MS-OFFCRYPTO §3.2.6 Digital Signatures
 *   - XMLDSig: https://www.w3.org/TR/xmldsig-core/
 */

import { xmlDecode } from "@xml/encode";

import { utf8Decoder } from "../core/internal-utils";
import { getFileName } from "../core/opc-paths";

/** Parsed digital signature metadata. */
export interface DigitalSignatureInfo {
  /** Signer's display name (from OfficeObject > SignatureInfoV1 > SignatureText). */
  readonly signer?: string;
  /** Sign date (ISO 8601). */
  readonly signDate?: string;
  /** Signature commitment type URI. */
  readonly commitmentType?: string;
  /** Signature comments/reason. */
  readonly signatureComments?: string;
  /** Signature purpose. */
  readonly purpose?: string;
  /** Signature provider URL. */
  readonly providerUrl?: string;
  /** Certificate subject CN. */
  readonly certificateSubject?: string;
  /** Certificate issuer CN. */
  readonly certificateIssuer?: string;
  /** Certificate serial number. */
  readonly certificateSerialNumber?: string;
  /** Hash of the signature (base64). */
  readonly signatureValue?: string;
  /**
   * Cryptographic verification status.
   *
   * `"not-verified"` is the only value this module ever produces — full
   * verification requires a complete XMLDSig + Canonical XML implementation
   * which is intentionally out of scope. The field is exposed so callers
   * are not tempted to interpret a missing value as "valid".
   */
  readonly cryptographicStatus: "not-verified";
  /** Raw XML for preservation. */
  readonly rawXml: string;
  /** Signature file name (e.g. "sig1.xml"). */
  readonly fileName: string;
}

/**
 * Check if a document has digital signatures.
 *
 * @param opaquePaths - Set of paths in the package (typically from opaqueParts).
 * @returns True if signatures are present.
 */
export function hasDigitalSignatures(opaquePaths: readonly string[]): boolean {
  return opaquePaths.some(p => p.startsWith("_xmlsignatures/sig"));
}

/**
 * Extract digital signature metadata from sig XML content.
 *
 * @param xmlStr - The signature XML content.
 * @param fileName - The file name (e.g. "sig1.xml").
 * @returns Parsed signature info.
 */
export function parseSignatureXml(xmlStr: string, fileName: string): DigitalSignatureInfo {
  const info: {
    -readonly [P in keyof DigitalSignatureInfo]: DigitalSignatureInfo[P];
  } = {
    rawXml: xmlStr,
    fileName,
    cryptographicStatus: "not-verified"
  };

  // Extract Office-specific metadata from <SignatureInfoV1>
  const sigTextMatch = /<SignatureText[^>]*>([^<]*)<\/SignatureText>/.exec(xmlStr);
  if (sigTextMatch) {
    info.signer = xmlDecode(sigTextMatch[1]);
  }

  const sigCommentsMatch = /<SignatureComments[^>]*>([^<]*)<\/SignatureComments>/.exec(xmlStr);
  if (sigCommentsMatch) {
    info.signatureComments = xmlDecode(sigCommentsMatch[1]);
  }

  const purposeMatch = /<SignaturePurpose[^>]*>([^<]*)<\/SignaturePurpose>/.exec(xmlStr);
  if (purposeMatch) {
    info.purpose = xmlDecode(purposeMatch[1]);
  }

  const dateMatch = /<SignatureDate[^>]*>([^<]*)<\/SignatureDate>/.exec(xmlStr);
  if (dateMatch) {
    info.signDate = xmlDecode(dateMatch[1]);
  }

  const providerMatch = /<SignatureProviderUrl[^>]*>([^<]*)<\/SignatureProviderUrl>/.exec(xmlStr);
  if (providerMatch) {
    info.providerUrl = xmlDecode(providerMatch[1]);
  }

  // Commitment type
  const commitMatch =
    /<CommitmentType[^>]*>\s*<CommitmentTypeIndication[^>]*>\s*<CommitmentTypeId>([^<]*)<\/CommitmentTypeId>/.exec(
      xmlStr
    );
  if (commitMatch) {
    info.commitmentType = xmlDecode(commitMatch[1]);
  }

  // Extract signature value (base64)
  const sigValMatch = /<SignatureValue[^>]*>([^]*?)<\/SignatureValue>/.exec(xmlStr);
  if (sigValMatch) {
    info.signatureValue = sigValMatch[1].trim();
  }

  // Certificate details from <X509Data>
  const certSubjectMatch = /<X509SubjectName[^>]*>([^<]*)<\/X509SubjectName>/.exec(xmlStr);
  if (certSubjectMatch) {
    info.certificateSubject = xmlDecode(certSubjectMatch[1]);
  }

  const certIssuerMatch = /<X509IssuerName[^>]*>([^<]*)<\/X509IssuerName>/.exec(xmlStr);
  if (certIssuerMatch) {
    info.certificateIssuer = xmlDecode(certIssuerMatch[1]);
  }

  const certSerialMatch = /<X509SerialNumber[^>]*>([^<]*)<\/X509SerialNumber>/.exec(xmlStr);
  if (certSerialMatch) {
    info.certificateSerialNumber = xmlDecode(certSerialMatch[1]);
  }

  return info;
}

/**
 * Extract all digital signatures from opaque parts of a document.
 *
 * @param opaqueParts - Opaque parts (from DocxDocument.opaqueParts).
 * @returns Array of parsed signature info.
 */
export function extractSignatures(
  opaqueParts: readonly { path: string; data: Uint8Array }[] | undefined
): DigitalSignatureInfo[] {
  if (!opaqueParts) {
    return [];
  }
  const signatures: DigitalSignatureInfo[] = [];
  for (const part of opaqueParts) {
    if (part.path.startsWith("_xmlsignatures/sig") && part.path.endsWith(".xml")) {
      const xmlStr = utf8Decoder.decode(part.data);
      const fileName = getFileName(part.path);
      signatures.push(parseSignatureXml(xmlStr, fileName));
    }
  }
  return signatures;
}

/**
 * Check that a parsed signature has the structural elements XMLDSig
 * requires (`Signature`, `SignedInfo`, `SignatureValue`, `KeyInfo`).
 *
 * This is **not** a cryptographic check — see `cryptographicStatus`. It is
 * also tolerant of namespace prefixes (`<ds:Signature>` etc.) which the
 * previous implementation missed.
 *
 * @returns True if the signature XML carries the required elements.
 */
export function isWellFormedSignature(info: DigitalSignatureInfo): boolean {
  // Allow optional namespace prefix and either an attribute-bearing or
  // self-closing form — XMLDSig signatures in real DOCX files commonly
  // use `<ds:Signature ...>` rather than the default-namespace form.
  const hasElement = (local: string): boolean =>
    new RegExp(`<(?:[\\w-]+:)?${local}(?:\\s|>|/>)`).test(info.rawXml);
  return (
    hasElement("Signature") &&
    hasElement("SignedInfo") &&
    hasElement("SignatureValue") &&
    hasElement("KeyInfo")
  );
}
