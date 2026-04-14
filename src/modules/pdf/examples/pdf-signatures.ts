/**
 * Example: PDF Digital Signatures
 *
 * Covers:
 *   1. Build a signature dictionary placeholder
 *   2. Verify a PDF signature (demonstration with synthetic data)
 *   3. Full sign + verify roundtrip (requires RSA key pair)
 *
 * Note: Digital signature creation requires a DER-encoded X.509 certificate
 * and PKCS#8 private key. This example demonstrates the API shape and
 * placeholder creation. For a full roundtrip, provide real keys.
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-signatures.ts
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  PdfDocumentBuilder,
  buildSignatureDictPlaceholder,
  verifyPdfSignature,
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
  // Create a test PDF
  const doc = new PdfDocumentBuilder();
  const page = doc.addPage();
  page.drawText("Test document for signature verification", { x: 72, y: 770, fontSize: 14 });
  const pdfBytes = await doc.build();
  fs.writeFileSync(path.join(outDir, "unsigned.pdf"), pdfBytes);

  // Attempt to verify with fake signature data
  const result = await verifyPdfSignature(
    pdfBytes,
    "00112233", // fake PKCS#7 hex
    [0, 100, 200, pdfBytes.length - 200] // fake byte range
  );

  console.log("\n3. Signature verification (invalid data — expected to fail gracefully):");
  console.log(`   Valid: ${result.valid}`);
  console.log(`   Covers whole file: ${result.coversWholeFile}`);
  console.log(`   Reason: ${result.reason}`);
}

// =============================================================================
// 4. Full Signing Workflow (API Demonstration)
// =============================================================================

console.log("\n4. Full signing workflow (API shape):");
console.log("   // Step 1: Create PDF with signature placeholder");
console.log("   const doc = new PdfDocumentBuilder();");
console.log("   doc.addPage().drawText('Signed Document', { x: 72, y: 770 });");
console.log("   const pdfBytes = await doc.build();");
console.log("");
console.log("   // Step 2: Sign with certificate and private key");
console.log("   import { signPdf } from '@cj-tech-master/excelts/pdf';");
console.log("   const signed = await signPdf(pdfWithPlaceholder, certificate, privateKey);");
console.log("");
console.log("   // Step 3: Verify the signature");
console.log("   import { verifyPdfSignature } from '@cj-tech-master/excelts/pdf';");
console.log("   const result = await verifyPdfSignature(signed, signatureHex, byteRange);");
console.log("   console.log(result.valid, result.coversWholeFile);");

console.log(`\nAll examples written to: ${outDir}`);
