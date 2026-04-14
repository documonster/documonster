/**
 * Example: PDF Digital Signatures
 *
 * Covers:
 *   1. Sign a new PDF with PdfDocumentBuilder.sign()
 *   2. Sign an existing PDF with PdfEditor.sign()
 *   3. Verify a signature with verifyPdfSignature()
 *   4. Tamper detection — modified PDF fails verification
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-signatures.ts
 */

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import { PdfDocumentBuilder, PdfEditor, verifyPdfSignature, readPdf } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-signature-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// Generate a temporary RSA key pair + self-signed certificate for this demo
const { certDer, privKeyDer } = generateTestCredentials();

// =============================================================================
// 1. Sign a New PDF (PdfDocumentBuilder)
// =============================================================================

{
  console.log("1. Sign a new PDF with PdfDocumentBuilder:");

  const doc = new PdfDocumentBuilder();
  doc.setMetadata({ title: "Signed Agreement", author: "excelts" });

  const page = doc.addPage();
  page.drawText("Service Agreement", { x: 72, y: 750, fontSize: 24, bold: true });
  page.drawText("This document is digitally signed to ensure authenticity.", {
    x: 72,
    y: 720,
    fontSize: 12
  });
  page.drawText("Any modification after signing will invalidate the signature.", {
    x: 72,
    y: 700,
    fontSize: 11,
    color: { r: 0.4, g: 0.4, b: 0.4 }
  });

  // One method call — certificate, key, and metadata
  doc.sign({
    certificate: certDer,
    privateKey: privKeyDer,
    name: "Alice Johnson",
    reason: "Contract approval",
    location: "San Francisco, CA"
  });

  const signedPdf = await doc.build();
  fs.writeFileSync(path.join(outDir, "01-builder-signed.pdf"), signedPdf);
  console.log(`   Output: 01-builder-signed.pdf (${signedPdf.length} bytes)`);

  // Verify it
  const verification = await verifyFromPdf(signedPdf);
  console.log(`   Valid: ${verification.valid}`);
  console.log(`   Covers whole file: ${verification.coversWholeFile}`);
}

// =============================================================================
// 2. Sign an Existing PDF (PdfEditor)
// =============================================================================

{
  console.log("\n2. Sign an existing PDF with PdfEditor:");

  // First, create an unsigned PDF
  const doc = new PdfDocumentBuilder();
  const page = doc.addPage();
  page.drawText("Invoice #12345", { x: 72, y: 750, fontSize: 20, bold: true });
  page.drawText("Amount: $1,500.00", { x: 72, y: 720, fontSize: 14 });
  page.drawText("Due: 2026-05-01", { x: 72, y: 700, fontSize: 12 });
  const unsignedPdf = await doc.build();
  fs.writeFileSync(path.join(outDir, "02-unsigned.pdf"), unsignedPdf);
  console.log(`   Unsigned: 02-unsigned.pdf (${unsignedPdf.length} bytes)`);

  // Load and sign it
  const editor = PdfEditor.load(unsignedPdf);
  const signedPdf = await editor.sign({
    certificate: certDer,
    privateKey: privKeyDer,
    name: "Bob Smith",
    reason: "Invoice approval"
  });
  fs.writeFileSync(path.join(outDir, "02-editor-signed.pdf"), signedPdf);
  console.log(`   Signed:   02-editor-signed.pdf (${signedPdf.length} bytes)`);

  // Verify
  const verification = await verifyFromPdf(signedPdf);
  console.log(`   Valid: ${verification.valid}`);
  console.log(`   Covers whole file: ${verification.coversWholeFile}`);
}

// =============================================================================
// 3. Tamper Detection
// =============================================================================

{
  console.log("\n3. Tamper detection:");

  // Create and sign
  const doc = new PdfDocumentBuilder();
  doc.addPage().drawText("Tamper test", { x: 72, y: 750, fontSize: 16 });
  doc.sign({ certificate: certDer, privateKey: privKeyDer, name: "Security Test" });
  const signedPdf = await doc.build();

  // Verify original — should pass
  const original = await verifyFromPdf(signedPdf);
  console.log(`   Original: valid=${original.valid}`);

  // Tamper with a byte and verify — should fail
  const tampered = new Uint8Array(signedPdf);
  tampered[50] = tampered[50] ^ 0xff;
  const tamperedResult = await verifyFromPdf(tampered);
  console.log(`   Tampered: valid=${tamperedResult.valid}`);
  console.log(`   Reason:   ${tamperedResult.reason}`);
}

// =============================================================================
// 4. Read Signed PDF — Confirm Signature Presence
// =============================================================================

{
  console.log("\n4. Read signed PDF to confirm signature field:");

  const doc = new PdfDocumentBuilder();
  doc.addPage().drawText("Document with signature field", { x: 72, y: 750, fontSize: 14 });
  doc.sign({ certificate: certDer, privateKey: privKeyDer, name: "Reader Test" });
  const signedPdf = await doc.build();

  const result = await readPdf(signedPdf);
  const sigField = result.formFields.find(f => f.type === "signature");
  console.log(`   Form fields: ${result.formFields.length}`);
  console.log(`   Signature field found: ${sigField !== undefined}`);
  if (sigField) {
    console.log(`   Field name: "${sigField.name}"`);
  }
}

console.log(`\nAll examples written to: ${outDir}`);

// =============================================================================
// Helpers
// =============================================================================

/** Extract signature hex + byte range from a signed PDF and verify. */
async function verifyFromPdf(pdf: Uint8Array) {
  const text = new TextDecoder().decode(pdf);

  const contentsMatch = text.match(/\/Contents\s*<([0-9a-fA-F]+)>/);
  if (!contentsMatch) {
    return { valid: false, coversWholeFile: false, digestAlgorithm: "", reason: "No /Contents" };
  }

  const brMatch = text.match(/\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/);
  if (!brMatch) {
    return { valid: false, coversWholeFile: false, digestAlgorithm: "", reason: "No /ByteRange" };
  }

  return verifyPdfSignature(pdf, contentsMatch[1], [
    parseInt(brMatch[1]),
    parseInt(brMatch[2]),
    parseInt(brMatch[3]),
    parseInt(brMatch[4])
  ]);
}

/** Generate a test RSA key pair + self-signed X.509 certificate. */
function generateTestCredentials(): { certDer: Uint8Array; privKeyDer: Uint8Array } {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", {
    modulusLength: 2048,
    publicKeyEncoding: { type: "spki", format: "der" },
    privateKeyEncoding: { type: "pkcs8", format: "der" }
  });

  const pubKeyDer = new Uint8Array(publicKey);
  const privKeyDer = new Uint8Array(privateKey);

  // Build minimal self-signed X.509 certificate
  const sha256WithRsa = new Uint8Array([
    0x30, 0x0d, 0x06, 0x09, 0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x01, 0x0b, 0x05, 0x00
  ]);
  const version = new Uint8Array([0xa0, 0x03, 0x02, 0x01, 0x02]);
  const serial = new Uint8Array([0x02, 0x01, 0x01]);
  const cn = new Uint8Array([
    0x30, 0x13, 0x31, 0x11, 0x30, 0x0f, 0x06, 0x03, 0x55, 0x04, 0x03, 0x0c, 0x08, 0x54, 0x65, 0x73,
    0x74, 0x43, 0x65, 0x72, 0x74
  ]);
  const validity = new Uint8Array([
    0x30, 0x1e, 0x17, 0x0d, 0x32, 0x35, 0x30, 0x31, 0x30, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30,
    0x5a, 0x17, 0x0d, 0x33, 0x35, 0x30, 0x31, 0x30, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x5a
  ]);

  const tbsContent = concat(version, serial, sha256WithRsa, cn, validity, cn, pubKeyDer);
  const tbs = derWrap(0x30, tbsContent);

  const key = crypto.createPrivateKey({
    key: Buffer.from(privKeyDer),
    format: "der",
    type: "pkcs8"
  });
  const signer = crypto.createSign("SHA256");
  signer.update(tbs);
  const sig = signer.sign(key);

  const sigBits = new Uint8Array(sig.length + 1);
  sigBits[0] = 0;
  sigBits.set(sig, 1);

  const certDer = derWrap(0x30, concat(tbs, sha256WithRsa, derWrap(0x03, sigBits)));
  return { certDer, privKeyDer };
}

function derWrap(tag: number, value: Uint8Array): Uint8Array {
  const len =
    value.length < 0x80
      ? new Uint8Array([value.length])
      : (() => {
          const bytes: number[] = [];
          let l = value.length;
          while (l > 0) {
            bytes.unshift(l & 0xff);
            l >>= 8;
          }
          return new Uint8Array([0x80 | bytes.length, ...bytes]);
        })();
  const result = new Uint8Array(1 + len.length + value.length);
  result[0] = tag;
  result.set(len, 1);
  result.set(value, 1 + len.length);
  return result;
}

function concat(...arrays: Uint8Array[]): Uint8Array {
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
