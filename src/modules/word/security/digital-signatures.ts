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

import { utf8Decoder } from "@word/core/internal-utils";
import { getFileName } from "@word/core/opc-paths";
import { xmlDecode } from "@xml/encode";

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

  // Each `<TagName ...>...</TagName>` lookup used to be a regex of the
  // form `/<Tag[^>]*>([^<]*)<\/Tag>/.exec(xmlStr)`. Although `[^>]*` and
  // `[^<]*` are linear in isolation, running ten such regexes against
  // attacker-controlled signature XML triggers CodeQL's
  // `js/polynomial-redos` rule. `extractTextElement` performs the same
  // job in a single linear scan and cannot exhibit super-linear runtime.
  const signer = extractTextElement(xmlStr, "SignatureText");
  if (signer !== undefined) {
    info.signer = xmlDecode(signer);
  }

  const sigComments = extractTextElement(xmlStr, "SignatureComments");
  if (sigComments !== undefined) {
    info.signatureComments = xmlDecode(sigComments);
  }

  const purpose = extractTextElement(xmlStr, "SignaturePurpose");
  if (purpose !== undefined) {
    info.purpose = xmlDecode(purpose);
  }

  const signDate = extractTextElement(xmlStr, "SignatureDate");
  if (signDate !== undefined) {
    info.signDate = xmlDecode(signDate);
  }

  const providerUrl = extractTextElement(xmlStr, "SignatureProviderUrl");
  if (providerUrl !== undefined) {
    info.providerUrl = xmlDecode(providerUrl);
  }

  // Commitment type — nested element. Read the full `<CommitmentType>`
  // body (which contains nested elements, hence `allowAngleBrackets`)
  // then look for `<CommitmentTypeId>` inside.
  const commitmentBlock = extractTextElement(xmlStr, "CommitmentType", {
    allowAngleBrackets: true
  });
  if (commitmentBlock !== undefined) {
    const commitmentId = extractTextElement(commitmentBlock, "CommitmentTypeId");
    if (commitmentId !== undefined) {
      info.commitmentType = xmlDecode(commitmentId);
    }
  }

  // Signature value (base64 — may legitimately span newlines, so don't strip).
  const signatureValue = extractTextElement(xmlStr, "SignatureValue", { allowAngleBrackets: true });
  if (signatureValue !== undefined) {
    info.signatureValue = signatureValue.trim();
  }

  // Certificate details from <X509Data>
  const certSubject = extractTextElement(xmlStr, "X509SubjectName");
  if (certSubject !== undefined) {
    info.certificateSubject = xmlDecode(certSubject);
  }

  const certIssuer = extractTextElement(xmlStr, "X509IssuerName");
  if (certIssuer !== undefined) {
    info.certificateIssuer = xmlDecode(certIssuer);
  }

  const certSerial = extractTextElement(xmlStr, "X509SerialNumber");
  if (certSerial !== undefined) {
    info.certificateSerialNumber = xmlDecode(certSerial);
  }

  return info;
}

/**
 * Find the first occurrence of `<tagName ...>...</tagName>` in `xml` and
 * return the inner text (verbatim — the caller is responsible for entity
 * decoding via `xmlDecode`).
 *
 * Implemented as a linear index scan rather than a regex match. The previous
 * regex-based implementation tripped CodeQL's polynomial-regex detector
 * because the input is attacker-controlled signature XML.
 *
 * @param xml - The XML text to search.
 * @param tagName - Local element name (no namespace prefix). The match
 *   ignores any namespace prefix actually present in the document.
 * @param options.allowAngleBrackets - When true, the inner text is read up
 *   to the literal `</tagName>` close tag rather than the next `<`. This
 *   is appropriate for elements like `SignatureValue` where the body is
 *   base64 and cannot legitimately contain `<` anyway, but lets the function
 *   tolerate accidental whitespace/newlines that some signers insert.
 */
function extractTextElement(
  xml: string,
  tagName: string,
  options: { allowAngleBrackets?: boolean } = {}
): string | undefined {
  const n = xml.length;
  let from = 0;
  while (from < n) {
    const lt = xml.indexOf("<", from);
    if (lt < 0) {
      return undefined;
    }
    // Skip an optional namespace prefix: <ns:Tag ...>.
    let nameStart = lt + 1;
    // Look ahead for either the bare tag name or `prefix:tagName`. We do
    // a forward scan rather than an unbounded regex match.
    const colon = xml.indexOf(":", nameStart);
    const ws = findTagNameEnd(xml, nameStart);
    if (colon > 0 && colon < ws) {
      nameStart = colon + 1;
    }
    if (
      xml.slice(nameStart, nameStart + tagName.length) !== tagName ||
      !isTagNameBoundary(xml.charCodeAt(nameStart + tagName.length))
    ) {
      from = lt + 1;
      continue;
    }
    // Found `<…tagName`. Find the closing `>` of the open tag.
    const openEnd = xml.indexOf(">", nameStart + tagName.length);
    if (openEnd < 0) {
      return undefined;
    }
    // Self-closing? Then the element has no text content.
    if (xml.charCodeAt(openEnd - 1) === 0x2f /* '/' */) {
      return "";
    }
    const bodyStart = openEnd + 1;
    if (options.allowAngleBrackets) {
      // Search for the matching close tag (allowing namespace prefix).
      const closeIdx = findCloseTag(xml, bodyStart, tagName);
      if (closeIdx < 0) {
        return undefined;
      }
      return xml.slice(bodyStart, closeIdx);
    }
    // Default behaviour: text content has no `<`. Stop at the next `<`.
    const lt2 = xml.indexOf("<", bodyStart);
    if (lt2 < 0) {
      return undefined;
    }
    return xml.slice(bodyStart, lt2);
  }
  return undefined;
}

function findTagNameEnd(xml: string, start: number): number {
  const n = xml.length;
  let i = start;
  while (i < n) {
    const c = xml.charCodeAt(i);
    if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === 0x2f || c === 0x3e) {
      return i;
    }
    i++;
  }
  return n;
}

function isTagNameBoundary(c: number): boolean {
  return (
    c === 0x20 || // space
    c === 0x09 || // tab
    c === 0x0a || // LF
    c === 0x0d || // CR
    c === 0x2f || // '/'
    c === 0x3e // '>'
  );
}

function findCloseTag(xml: string, from: number, tagName: string): number {
  const n = xml.length;
  let i = from;
  while (i < n) {
    const lt = xml.indexOf("</", i);
    if (lt < 0) {
      return -1;
    }
    let p = lt + 2;
    // Optional namespace prefix
    const colon = xml.indexOf(":", p);
    const gt = xml.indexOf(">", p);
    if (gt < 0) {
      return -1;
    }
    if (colon > 0 && colon < gt) {
      p = colon + 1;
    }
    if (
      xml.slice(p, p + tagName.length) === tagName &&
      // Allow trailing whitespace before '>' but require a boundary char.
      isTagNameBoundary(xml.charCodeAt(p + tagName.length))
    ) {
      return lt;
    }
    i = lt + 2;
  }
  return -1;
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
