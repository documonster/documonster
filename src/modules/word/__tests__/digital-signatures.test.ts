/**
 * DOCX Module - Digital Signatures Tests
 */

import { describe, it, expect } from "vitest";

import {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "../security/digital-signatures";

const SAMPLE_SIGNATURE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#" Id="idPackageSignature">
  <SignedInfo>
    <CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
    <SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
    <Reference URI="">
      <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      <DigestValue>abc123</DigestValue>
    </Reference>
  </SignedInfo>
  <SignatureValue>SIG_VALUE_BASE64</SignatureValue>
  <KeyInfo>
    <X509Data>
      <X509SubjectName>CN=John Doe, O=Example Corp</X509SubjectName>
      <X509IssuerName>CN=Example CA, O=Example</X509IssuerName>
      <X509SerialNumber>12345</X509SerialNumber>
    </X509Data>
  </KeyInfo>
  <Object>
    <SignatureProperties>
      <SignatureProperty>
        <SignatureInfoV1 xmlns="http://schemas.microsoft.com/office/2006/digsig">
          <SignatureText>John Doe</SignatureText>
          <SignatureComments>Approved</SignatureComments>
          <SignaturePurpose>Approval</SignaturePurpose>
          <SignatureProviderUrl>https://example.com/provider</SignatureProviderUrl>
          <SignatureDate>2024-01-15T10:30:00Z</SignatureDate>
        </SignatureInfoV1>
        <CommitmentType>
          <CommitmentTypeIndication>
            <CommitmentTypeId>http://uri.etsi.org/01903/v1.2.2#ProofOfApproval</CommitmentTypeId>
          </CommitmentTypeIndication>
        </CommitmentType>
      </SignatureProperty>
    </SignatureProperties>
  </Object>
</Signature>`;

describe("hasDigitalSignatures", () => {
  it("returns true when signature parts exist", () => {
    expect(
      hasDigitalSignatures(["word/document.xml", "_xmlsignatures/sig1.xml", "[Content_Types].xml"])
    ).toBe(true);
  });

  it("returns false when no signature parts", () => {
    expect(hasDigitalSignatures(["word/document.xml", "[Content_Types].xml"])).toBe(false);
  });

  it("returns false for empty paths array", () => {
    expect(hasDigitalSignatures([])).toBe(false);
  });

  it("matches multiple signature files", () => {
    expect(hasDigitalSignatures(["_xmlsignatures/sig1.xml", "_xmlsignatures/sig2.xml"])).toBe(true);
  });
});

describe("parseSignatureXml", () => {
  it("extracts signer name from SignatureText", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.signer).toBe("John Doe");
  });

  it("extracts signature comments", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.signatureComments).toBe("Approved");
  });

  it("extracts signature purpose", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.purpose).toBe("Approval");
  });

  it("extracts signature date", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.signDate).toBe("2024-01-15T10:30:00Z");
  });

  it("extracts provider URL", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.providerUrl).toBe("https://example.com/provider");
  });

  it("extracts commitment type", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.commitmentType).toBe("http://uri.etsi.org/01903/v1.2.2#ProofOfApproval");
  });

  it("extracts certificate subject", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.certificateSubject).toBe("CN=John Doe, O=Example Corp");
  });

  it("extracts certificate issuer", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.certificateIssuer).toBe("CN=Example CA, O=Example");
  });

  it("extracts certificate serial number", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.certificateSerialNumber).toBe("12345");
  });

  it("extracts signature value", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.signatureValue).toBe("SIG_VALUE_BASE64");
  });

  it("preserves rawXml", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.rawXml).toBe(SAMPLE_SIGNATURE_XML);
  });

  it("preserves fileName", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig2.xml");
    expect(info.fileName).toBe("sig2.xml");
  });

  it("decodes XML entities in text fields", () => {
    const xml = `<Signature xmlns="http://www.w3.org/2000/09/xmldsig#">
      <Object><SignatureProperties><SignatureProperty>
        <SignatureInfoV1 xmlns="http://schemas.microsoft.com/office/2006/digsig">
          <SignatureText>O&#39;Connor &amp; Sons</SignatureText>
        </SignatureInfoV1>
      </SignatureProperty></SignatureProperties></Object>
    </Signature>`;
    const info = parseSignatureXml(xml, "sig.xml");
    expect(info.signer).toBe("O'Connor & Sons");
  });
});

describe("extractSignatures", () => {
  it("returns empty array when opaqueParts is undefined", () => {
    expect(extractSignatures(undefined)).toEqual([]);
  });

  it("returns empty array when no signature parts present", () => {
    const parts = [{ path: "word/document.xml", data: new TextEncoder().encode("<doc/>") }];
    expect(extractSignatures(parts)).toEqual([]);
  });

  it("extracts signatures from _xmlsignatures/ paths", () => {
    const parts = [
      { path: "word/document.xml", data: new TextEncoder().encode("<doc/>") },
      { path: "_xmlsignatures/sig1.xml", data: new TextEncoder().encode(SAMPLE_SIGNATURE_XML) }
    ];
    const signatures = extractSignatures(parts);
    expect(signatures.length).toBe(1);
    expect(signatures[0].signer).toBe("John Doe");
    expect(signatures[0].fileName).toBe("sig1.xml");
  });

  it("extracts multiple signatures", () => {
    const parts = [
      { path: "_xmlsignatures/sig1.xml", data: new TextEncoder().encode(SAMPLE_SIGNATURE_XML) },
      { path: "_xmlsignatures/sig2.xml", data: new TextEncoder().encode(SAMPLE_SIGNATURE_XML) }
    ];
    expect(extractSignatures(parts).length).toBe(2);
  });

  it("ignores non-sig files in _xmlsignatures/ folder", () => {
    const parts = [
      { path: "_xmlsignatures/origin.sigs", data: new Uint8Array() },
      { path: "_xmlsignatures/sig1.xml", data: new TextEncoder().encode(SAMPLE_SIGNATURE_XML) }
    ];
    expect(extractSignatures(parts).length).toBe(1);
  });
});

describe("isWellFormedSignature", () => {
  it("returns true for valid signature XML", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(isWellFormedSignature(info)).toBe(true);
  });

  it("returns false when missing required elements", () => {
    const info = parseSignatureXml("<root>no signature here</root>", "sig.xml");
    expect(isWellFormedSignature(info)).toBe(false);
  });

  it("recognises namespace-prefixed Signature elements", () => {
    const xml =
      '<ds:Signature xmlns:ds="http://www.w3.org/2000/09/xmldsig#">' +
      "<ds:SignedInfo></ds:SignedInfo>" +
      "<ds:SignatureValue>abc==</ds:SignatureValue>" +
      "<ds:KeyInfo></ds:KeyInfo>" +
      "</ds:Signature>";
    const info = parseSignatureXml(xml, "sig.xml");
    expect(isWellFormedSignature(info)).toBe(true);
  });

  it("reports cryptographicStatus 'not-verified' on every parsed signature", () => {
    const info = parseSignatureXml(SAMPLE_SIGNATURE_XML, "sig1.xml");
    expect(info.cryptographicStatus).toBe("not-verified");
  });
});
