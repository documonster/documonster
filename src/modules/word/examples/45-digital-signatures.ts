/**
 * Word Example 45 — Digital signatures (read-only inspection)
 *
 * The library exposes inspection helpers for XMLDSig signatures embedded
 * in `_xmlsignatures/sigN.xml` parts. Cryptographic verification is
 * intentionally out of scope (it requires a full XMLDSig + Canonical XML
 * implementation), so every parsed signature carries
 * `cryptographicStatus: "not-verified"`.
 *
 * Covers:
 *   - hasDigitalSignatures(opaquePaths)
 *   - extractSignatures(opaqueParts)
 *   - parseSignatureXml(xmlStr, fileName)
 *   - isWellFormedSignature(info)
 *
 * Output: tmp/word-examples/45-signatures.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "../crypto";
import { Document } from "../index";
import type { OpaquePart } from "../index";
import { packageDocx } from "../writer/docx-packager";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const lines: string[] = [];
function log(s: string): void {
  console.log(s);
  lines.push(s);
}

// ---------------------------------------------------------------------------
// 1. A made-up signature XML snippet (real ones are produced by Word's
//    digital-signing UI). It is structurally well-formed XMLDSig with the
//    Office-specific SignatureInfoV1 metadata wrapper inside <Object>.
// ---------------------------------------------------------------------------
const signatureXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Signature xmlns="http://www.w3.org/2000/09/xmldsig#" Id="idPackageSignature">
  <SignedInfo>
    <CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>
    <SignatureMethod Algorithm="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"/>
    <Reference URI="#idOfficeObject">
      <DigestMethod Algorithm="http://www.w3.org/2001/04/xmlenc#sha256"/>
      <DigestValue>BASE64DIGEST==</DigestValue>
    </Reference>
  </SignedInfo>
  <SignatureValue>BASE64SIGNATURE==</SignatureValue>
  <KeyInfo>
    <X509Data>
      <X509Certificate>BASE64CERT==</X509Certificate>
    </X509Data>
  </KeyInfo>
  <Object Id="idOfficeObject">
    <SignatureProperties>
      <SignatureProperty Id="idOfficeV1Details" Target="#idPackageSignature">
        <SignatureInfoV1 xmlns="http://schemas.microsoft.com/office/2006/digsig">
          <SignatureText>Jane Q. Public</SignatureText>
          <SignatureComments>Approved for release</SignatureComments>
          <WindowsVersion>10.0</WindowsVersion>
          <OfficeVersion>16.0</OfficeVersion>
          <SignatureProviderUrl>https://example.com/idprovider</SignatureProviderUrl>
          <SignaturePurpose>Approval</SignaturePurpose>
        </SignatureInfoV1>
      </SignatureProperty>
    </SignatureProperties>
  </Object>
</Signature>`;

// ---------------------------------------------------------------------------
// 2. parseSignatureXml on the snippet
// ---------------------------------------------------------------------------
const parsed = parseSignatureXml(signatureXml, "sig1.xml");
log(`  parseSignatureXml:`);
log(`    fileName:               ${parsed.fileName}`);
log(`    signer:                 ${parsed.signer ?? "(none)"}`);
log(`    signatureComments:      ${parsed.signatureComments ?? "(none)"}`);
log(`    purpose:                ${parsed.purpose ?? "(none)"}`);
log(`    providerUrl:            ${parsed.providerUrl ?? "(none)"}`);
log(`    cryptographicStatus:    ${parsed.cryptographicStatus}`);
log(`    isWellFormedSignature:  ${isWellFormedSignature(parsed)}`);

// ---------------------------------------------------------------------------
// 3. Build a docx whose opaque parts contain the signature, then verify the
//    discovery helpers.
// ---------------------------------------------------------------------------
const baseDoc = (() => {
  const dd = Document.create();
  Document.useDefaultStyles(dd);
  Document.addParagraph(dd, "Document with a (synthetic) digital signature.");
  return Document.build(dd);
})();
const sigPart: OpaquePart = {
  path: "_xmlsignatures/sig1.xml",
  data: new TextEncoder().encode(signatureXml),
  contentType: "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
};
const docWithSig = { ...baseDoc, opaqueParts: [sigPart] };

const opaquePaths = (docWithSig.opaqueParts ?? []).map(p => p.path);
log(`\n  hasDigitalSignatures:    ${hasDigitalSignatures(opaquePaths)}`);

const sigs = extractSignatures(docWithSig.opaqueParts);
log(`  extractSignatures count: ${sigs.length}`);
for (const s of sigs) {
  log(`    - ${s.fileName}: signer=${s.signer ?? "(?)"}, well-formed=${isWellFormedSignature(s)}`);
}

// The default security policy strips _xmlsignatures/* on every save (since
// re-serialising the document invalidates any embedded signatures). For this
// demo we pass dropSignaturesOnModify:false so the signature stays in the
// resulting DOCX — note the signature is no longer cryptographically valid,
// only structurally present.
const bytes = await packageDocx(docWithSig, {
  securityPolicy: { dropSignaturesOnModify: false }
});
fs.writeFileSync(path.join(outDir, "45-signed.docx"), bytes);
log(`\n  → 45-signed.docx (${bytes.length} bytes)`);

// ---------------------------------------------------------------------------
// 4. Edge case: a malformed signature (missing SignatureValue) is detected
// ---------------------------------------------------------------------------
const broken = parseSignatureXml(
  `<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo/></Signature>`,
  "sig-broken.xml"
);
log(`  isWellFormedSignature(broken): ${isWellFormedSignature(broken)}`);

// ---------------------------------------------------------------------------
// 5. Edge case: extractSignatures on a doc without opaqueParts → []
// ---------------------------------------------------------------------------
log(`  extractSignatures(undefined): ${extractSignatures(undefined).length}`);

fs.writeFileSync(path.join(outDir, "45-signatures.txt"), lines.join("\n"));
console.log(`  → 45-signatures.txt`);
