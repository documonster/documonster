/**
 * Example: PDF Digital Signatures
 *
 * Covers:
 *   1. Sign a new PDF with PdfDocumentBuilder.sign()
 *   2. Sign an existing PDF with PdfEditor.sign()
 *   3. Find and verify every signature with Pdf.verifySignatures()
 *   4. Tamper detection — modified PDF fails verification
 *   5. An unsigned update after signing — still intact, no longer the whole file
 *   6. Low-level: verify one signature with Pdf.Signature.verify()
 *
 * Run: pnpm example --filter pdf-signatures
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Pdf } from "../index";
import { generateSelfSignedCertificate } from "./utils/self-signed-certificate";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// A signing identity is an *input* to this example, not something the library
// produces. This generates a throwaway self-signed one so the example runs
// unattended; in production the certificate and key come from a CA, a keystore
// or an HSM. See `./utils/self-signed-certificate.ts` for the openssl equivalent.
const { certificate, privateKey } = await generateSelfSignedCertificate("Example Signer");

// =============================================================================
// 1. Sign a New PDF (PdfDocumentBuilder)
// =============================================================================

{
  console.log("1. Sign a new PDF with PdfDocumentBuilder:");

  const doc = new Pdf.Builder();
  doc.setMetadata({ title: "Signed Agreement", author: "documonster" });

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
    certificate,
    privateKey,
    name: "Alice Johnson",
    reason: "Contract approval",
    location: "San Francisco, CA"
  });

  const signedPdf = await doc.toBuffer();
  fs.writeFileSync(path.join(outDir, "01-builder-signed.pdf"), signedPdf);
  console.log(`   Output: 01-builder-signed.pdf (${signedPdf.length} bytes)`);

  // Verify it: every signature field in the file, each with its own report
  const [verification] = await Pdf.verifySignatures(signedPdf);
  // `intact` is integrity only: the bytes are unchanged and the signature matches the
  // certificate it carries. Whether to trust that certificate — and so the name in it —
  // is for your own trust store to decide; `verification.certificate` is the DER to check.
  console.log(`   Field: ${verification.fieldName}, claimed signer: ${verification.claimedSigner}`);
  console.log(`   Name: ${verification.name}, reason: ${verification.reason}`);
  console.log(`   Intact: ${verification.intact}`);
  console.log(`   Covers whole file: ${verification.coversWholeFile}`);
}

// =============================================================================
// 2. Sign an Existing PDF (PdfEditor)
// =============================================================================

{
  console.log("\n2. Sign an existing PDF with PdfEditor:");

  // First, create an unsigned PDF
  const doc = new Pdf.Builder();
  const page = doc.addPage();
  page.drawText("Invoice #12345", { x: 72, y: 750, fontSize: 20, bold: true });
  page.drawText("Amount: $1,500.00", { x: 72, y: 720, fontSize: 14 });
  page.drawText("Due: 2026-05-01", { x: 72, y: 700, fontSize: 12 });
  const unsignedPdf = await doc.toBuffer();
  fs.writeFileSync(path.join(outDir, "02-unsigned.pdf"), unsignedPdf);
  console.log(`   Unsigned: 02-unsigned.pdf (${unsignedPdf.length} bytes)`);

  // Load and sign — one method call
  const editor = Pdf.Editor.read(unsignedPdf);
  const signedPdf = await editor.sign({
    certificate,
    privateKey,
    name: "Bob Smith",
    reason: "Invoice approval"
  });
  fs.writeFileSync(path.join(outDir, "02-editor-signed.pdf"), signedPdf);
  console.log(`   Signed:   02-editor-signed.pdf (${signedPdf.length} bytes)`);

  const [verification] = await Pdf.verifySignatures(signedPdf);
  console.log(`   Intact: ${verification.intact}`);
  console.log(`   Covers whole file: ${verification.coversWholeFile}`);
}

// =============================================================================
// 3. Tamper Detection
// =============================================================================

{
  console.log("\n3. Tamper detection:");

  const doc = new Pdf.Builder();
  doc.addPage().drawText("Tamper test", { x: 72, y: 750, fontSize: 16 });
  doc.sign({ certificate, privateKey, name: "Security Test" });
  const signedPdf = await doc.toBuffer();

  // Verify original — should pass
  const [original] = await Pdf.verifySignatures(signedPdf);
  console.log(`   Original: intact=${original.intact}`);

  // Tamper with a byte and verify — should fail
  const tampered = new Uint8Array(signedPdf);
  tampered[50] = tampered[50] ^ 0xff;
  const [tamperedResult] = await Pdf.verifySignatures(tampered);
  console.log(`   Tampered: intact=${tamperedResult.intact}`);
  console.log(`   Failure:  ${tamperedResult.failure}`);
}

// =============================================================================
// 4. Read Signed PDF — Confirm Signature Presence
// =============================================================================

{
  console.log("\n4. Read signed PDF to confirm signature field:");

  const doc = new Pdf.Builder();
  doc.addPage().drawText("Document with signature field", { x: 72, y: 750, fontSize: 14 });
  doc.sign({ certificate, privateKey, name: "Reader Test" });
  const signedPdf = await doc.toBuffer();

  const result = await Pdf.extract(signedPdf);
  const sigField = result.formFields.find(f => f.type === "signature");
  console.log(`   Form fields: ${result.formFields.length}`);
  console.log(`   Signature field found: ${sigField !== undefined}`);
  if (sigField) {
    console.log(`   Field name: "${sigField.name}"`);
  }
}

// =============================================================================
// 5. An Unsigned Update After Signing
// =============================================================================

{
  console.log("\n5. Incremental update after signing:");

  const doc = new Pdf.Builder();
  doc.addPage().drawText("Signed, then annotated", { x: 72, y: 750, fontSize: 14 });
  doc.sign({ certificate, privateKey, name: "Revision Test" });
  const signedPdf = await doc.toBuffer();

  // An incremental update leaves the signed bytes alone, so the signature still holds —
  // but it covers the revision it signed, not the annotation appended after it.
  const editor = Pdf.Editor.read(signedPdf);
  editor.getPage(0).drawText("Received", { x: 72, y: 700, fontSize: 12 });
  const updated = await editor.toBuffer({ incremental: true });
  fs.writeFileSync(path.join(outDir, "05-signed-then-updated.pdf"), updated);

  const [report] = await Pdf.verifySignatures(updated);
  console.log(`   Intact: ${report.intact}`);
  console.log(`   Covers whole file: ${report.coversWholeFile}`);
}

// =============================================================================
// 6. Low-level: Verify One Signature
// =============================================================================

{
  console.log("\n6. Pdf.Signature.verify with a known /ByteRange:");

  const doc = new Pdf.Builder();
  doc.addPage().drawText("Low-level verification", { x: 72, y: 750, fontSize: 14 });
  doc.sign({ certificate, privateKey, name: "Low Level" });
  const signedPdf = await doc.toBuffer();

  // `verifySignatures` locates the dictionary for you; `Signature.verify` takes the
  // pieces directly — here, the hex string sitting in the gap between the two ranges.
  const [{ byteRange }] = await Pdf.verifySignatures(signedPdf);
  const gap = signedPdf.subarray(byteRange[0] + byteRange[1] + 1, byteRange[2] - 1);
  const result = await Pdf.Signature.verify(signedPdf, new TextDecoder().decode(gap), byteRange);
  console.log(`   Intact: ${result.intact}`);
}

console.log(`\nAll examples written to: ${outDir}`);
