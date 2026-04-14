/**
 * Example: PDF Digital Signatures
 *
 * Covers:
 *   1. Build a signature dictionary placeholder
 *   2. ASN.1 DER parsing
 *   3. Signature verification — graceful handling of invalid data
 *   4. Full sign + verify roundtrip with a generated RSA key pair
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-signatures.ts
 */

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import {
  PdfDocumentBuilder,
  buildSignatureDictPlaceholder,
  verifyPdfSignature,
  signPdf,
  asn1Parse
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-signature-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Signature Dictionary Placeholder
// =============================================================================

{
  const { dictString, placeholder } = buildSignatureDictPlaceholder({
    name: "John Doe",
    reason: "Document approval",
    location: "New York, NY",
    contactInfo: "john@example.com"
  });

  console.log("1. Signature dict placeholder:");
  console.log(`   Contains /Type /Sig: ${dictString.includes("/Type /Sig")}`);
  console.log(`   Contains /Filter: ${dictString.includes("/Adobe.PPKLite")}`);
  console.log(`   Contains /SubFilter: ${dictString.includes("/adbe.pkcs7.detached")}`);
  console.log(`   Contains /ByteRange: ${dictString.includes("/ByteRange")}`);
  console.log(`   Contains /Name: ${dictString.includes("/Name")}`);
  console.log(`   Contains /Reason: ${dictString.includes("/Reason")}`);
  console.log(`   Placeholder length: ${placeholder.length} hex chars`);
}

// =============================================================================
// 2. ASN.1 Parser — Parse DER Structures
// =============================================================================

{
  // SEQUENCE { INTEGER 42, NULL }
  const derData = new Uint8Array([0x30, 0x05, 0x02, 0x01, 0x2a, 0x05, 0x00]);
  const node = asn1Parse(derData);

  console.log("\n2. ASN.1 parser:");
  console.log(`   Root tag: 0x${node.tag.toString(16)} (SEQUENCE)`);
  console.log(`   Children: ${node.children.length}`);
  console.log(`   Child[0] tag: 0x${node.children[0].tag.toString(16)} (INTEGER)`);
  console.log(`   Child[0] value: ${node.children[0].bytes[0]}`);
  console.log(`   Child[1] tag: 0x${node.children[1].tag.toString(16)} (NULL)`);
}

// =============================================================================
// 3. Signature Verification — Graceful Handling of Invalid Data
// =============================================================================

{
  const doc = new PdfDocumentBuilder();
  const page = doc.addPage();
  page.drawText("Test document for signature verification", { x: 72, y: 770, fontSize: 14 });
  const pdfBytes = await doc.build();
  fs.writeFileSync(path.join(outDir, "unsigned.pdf"), pdfBytes);

  // Attempt to verify with fake signature data — should fail gracefully
  const result = await verifyPdfSignature(pdfBytes, "00112233", [
    0,
    100,
    200,
    pdfBytes.length - 200
  ]);

  console.log("\n3. Signature verification (invalid data — expected to fail gracefully):");
  console.log(`   Valid: ${result.valid}`);
  console.log(`   Covers whole file: ${result.coversWholeFile}`);
  console.log(`   Reason: ${result.reason}`);
}

// =============================================================================
// 4. Full Sign + Verify Roundtrip
// =============================================================================

{
  console.log("\n4. Full sign + verify roundtrip:");

  // Step 1: Generate a temporary RSA key pair + self-signed certificate
  const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", {
    modulusLength: 2048,
    publicKeyEncoding: { type: "spki", format: "der" },
    privateKeyEncoding: { type: "pkcs8", format: "der" }
  });
  const pubKeyDer = new Uint8Array(publicKey);
  const privKeyDer = new Uint8Array(privateKey);

  // Create a self-signed X.509 certificate using node:crypto
  const certDer = createSelfSignedCert(pubKeyDer, privKeyDer);
  console.log(`   Generated RSA-2048 key pair + self-signed cert (${certDer.length} bytes DER)`);

  // Step 2: Build a PDF with the signature placeholder embedded
  const { dictString } = buildSignatureDictPlaceholder({
    name: "Test Signer",
    reason: "Roundtrip test"
  });

  // Build a minimal PDF that includes the signature dict as a Sig annotation
  const pdfWithPlaceholder = buildPdfWithSignaturePlaceholder(dictString);
  fs.writeFileSync(path.join(outDir, "with-placeholder.pdf"), pdfWithPlaceholder);
  console.log(`   Built PDF with signature placeholder (${pdfWithPlaceholder.length} bytes)`);

  // Step 3: Sign the PDF
  const signed = await signPdf(pdfWithPlaceholder, certDer, privKeyDer);
  fs.writeFileSync(path.join(outDir, "signed.pdf"), signed);
  console.log(`   Signed PDF (${signed.length} bytes)`);

  // Step 4: Extract signature info and verify
  const { signatureHex, byteRange } = extractSignatureInfo(signed);
  console.log(`   Extracted signature: ${signatureHex.length} hex chars`);
  console.log(`   ByteRange: [${byteRange.join(", ")}]`);

  const result = await verifyPdfSignature(signed, signatureHex, byteRange);
  console.log(`   Verification result:`);
  console.log(`     Valid: ${result.valid}`);
  console.log(`     Covers whole file: ${result.coversWholeFile}`);
  console.log(`     Digest algorithm: ${result.digestAlgorithm}`);
  if (result.reason) {
    console.log(`     Reason: ${result.reason}`);
  }

  // Step 5: Tamper with the PDF and verify again — should fail
  const tampered = new Uint8Array(signed);
  // Modify a byte in the first range (before the signature)
  tampered[50] = tampered[50] ^ 0xff;
  const tamperedResult = await verifyPdfSignature(tampered, signatureHex, byteRange);
  console.log(`\n   Tampered PDF verification:`);
  console.log(`     Valid: ${tamperedResult.valid} (expected: false)`);
  console.log(`     Reason: ${tamperedResult.reason}`);
}

console.log(`\nAll examples written to: ${outDir}`);

// =============================================================================
// Helper: Build a minimal PDF with a signature placeholder
// =============================================================================

function buildPdfWithSignaturePlaceholder(sigDictString: string): Uint8Array {
  // Build a minimal valid PDF with visible content + a /Sig annotation
  // Build page content stream
  const streamContent = [
    "BT /F1 20 Tf 72 750 Td (Digitally Signed Document) Tj ET",
    "BT /F1 11 Tf 72 725 Td (This PDF contains a PKCS#7 digital signature.) Tj ET",
    "BT /F1 11 Tf 72 710 Td (The signature widget is in the box below.) Tj ET",
    "0.9 0.95 1 rg 72 680 228 40 re f",
    "0.2 0.4 0.8 RG 72 680 228 40 re S",
    "BT /F1 10 Tf 80 696 Td (Signed by: Test Signer) Tj ET",
    "BT /F1 8 Tf 80 684 Td (Reason: Roundtrip test) Tj ET"
  ].join("\n");

  const lines = [
    "%PDF-2.0",
    "",
    "1 0 obj",
    "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] /SigFlags 3 >> >>",
    "endobj",
    "",
    "2 0 obj",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "endobj",
    "",
    "3 0 obj",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Annots [4 0 R] /Contents 6 0 R /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> >>",
    "endobj",
    "",
    "4 0 obj",
    `<< /Type /Annot /Subtype /Widget /FT /Sig /Rect [72 680 300 720] /T (Signature1) /V 5 0 R >>`,
    "endobj",
    "",
    "5 0 obj",
    sigDictString,
    "endobj",
    "",
    "6 0 obj",
    `<< /Length ${streamContent.length} >>`,
    "stream",
    streamContent,
    "endstream",
    "endobj",
    ""
  ];

  const body = lines.join("\n");

  // Build xref
  const offsets: number[] = [];
  for (let objNum = 1; objNum <= 6; objNum++) {
    const marker = `${objNum} 0 obj`;
    offsets.push(body.indexOf(marker));
  }

  const xrefOffset = body.length;
  let xref = "xref\n";
  xref += `0 7\n`;
  xref += "0000000000 65535 f \n";
  for (const off of offsets) {
    xref += `${String(off).padStart(10, "0")} 00000 n \n`;
  }
  xref += "trailer\n";
  xref += "<< /Size 7 /Root 1 0 R >>\n";
  xref += "startxref\n";
  xref += `${xrefOffset}\n`;
  xref += "%%EOF\n";

  return new TextEncoder().encode(body + xref);
}

// =============================================================================
// Helper: Create a minimal self-signed X.509 certificate (DER)
// =============================================================================

function createSelfSignedCert(publicKeyDer: Uint8Array, privateKeyDer: Uint8Array): Uint8Array {
  // Build a minimal X.509 v3 TBSCertificate, then sign it with RSA-SHA256

  // OIDs
  const sha256WithRsa = new Uint8Array([
    0x30, 0x0d, 0x06, 0x09, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x01, 0x0b, 0x05, 0x00
  ]);

  // Version [0] EXPLICIT INTEGER 2 (v3)
  const version = new Uint8Array([0xa0, 0x03, 0x02, 0x01, 0x02]);

  // Serial number: INTEGER 1
  const serial = new Uint8Array([0x02, 0x01, 0x01]);

  // Issuer/Subject: SEQUENCE { SET { SEQUENCE { OID(CN), UTF8String "Test" } } }
  const cn = new Uint8Array([
    0x30, 0x13, 0x31, 0x11, 0x30, 0x0f, 0x06, 0x03, 0x55, 0x04, 0x03, 0x0c, 0x08, 0x54, 0x65, 0x73,
    0x74, 0x43, 0x65, 0x72, 0x74
  ]);

  // Validity: not before/after (UTCTime)
  const validity = new Uint8Array([
    0x30, 0x1e, 0x17, 0x0d, 0x32, 0x35, 0x30, 0x31, 0x30, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30,
    0x5a, 0x17, 0x0d, 0x33, 0x35, 0x30, 0x31, 0x30, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x5a
  ]);

  // TBSCertificate
  const tbsContent = concatAll(version, serial, sha256WithRsa, cn, validity, cn, publicKeyDer);
  const tbs = derWrap(0x30, tbsContent);

  // Sign TBS
  const key = crypto.createPrivateKey({
    key: Buffer.from(privateKeyDer),
    format: "der",
    type: "pkcs8"
  });
  const signer = crypto.createSign("SHA256");
  signer.update(tbs);
  const sig = signer.sign(key);

  // BIT STRING wrapping of signature
  const sigBits = new Uint8Array(sig.length + 1);
  sigBits[0] = 0; // no unused bits
  sigBits.set(sig, 1);
  const sigBitString = derWrap(0x03, sigBits);

  // Certificate = SEQUENCE { TBS, AlgorithmIdentifier, Signature }
  return derWrap(0x30, concatAll(tbs, sha256WithRsa, sigBitString));
}

function derWrap(tag: number, value: Uint8Array): Uint8Array {
  const len = derLength(value.length);
  const result = new Uint8Array(1 + len.length + value.length);
  result[0] = tag;
  result.set(len, 1);
  result.set(value, 1 + len.length);
  return result;
}

function derLength(length: number): Uint8Array {
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

function concatAll(...arrays: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const a of arrays) {
    total += a.length;
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    result.set(a, offset);
    offset += a.length;
  }
  return result;
}

// =============================================================================
// Helper: Extract signature hex and byte range from a signed PDF
// =============================================================================

function extractSignatureInfo(pdf: Uint8Array): {
  signatureHex: string;
  byteRange: [number, number, number, number];
} {
  const text = new TextDecoder().decode(pdf);

  // Extract /Contents <hex>
  const contentsMatch = text.match(/\/Contents\s*<([0-9a-fA-F]+)>/);
  if (!contentsMatch) {
    throw new Error("No /Contents found");
  }
  const signatureHex = contentsMatch[1];

  // Extract /ByteRange [n n n n]
  const brMatch = text.match(/\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/);
  if (!brMatch) {
    throw new Error("No /ByteRange found");
  }
  const byteRange: [number, number, number, number] = [
    parseInt(brMatch[1]),
    parseInt(brMatch[2]),
    parseInt(brMatch[3]),
    parseInt(brMatch[4])
  ];

  return { signatureHex, byteRange };
}
