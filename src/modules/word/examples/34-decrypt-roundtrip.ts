/**
 * Word Example 34 — Decryption round-trip
 *
 * Covers:
 *   - encryptDocx → file → readDocx({password}) — the most common path,
 *     readDocx handles decryption transparently when a password is supplied.
 *   - isEncryptedDocx — detect whether bytes are an encrypted CFB.
 *   - decryptDocx (low-level) — decrypt to raw ZIP bytes.
 *   - Wrong password → DocxDecryptionError.
 *   - Round-trip after edit (decrypt → mutate → encrypt back).
 *
 * Output: tmp/word-examples/34-decrypt/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { isEncryptedDocx, decryptDocx } from "../crypto";
import {
  Document,
  toBuffer,
  readDocx,
  encryptDocx,
  isDocxError,
  DocxDecryptionError,
  replaceText
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/34-decrypt"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build a normal document, encrypt it
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addHeading(d, "Sensitive document", 1);
Document.addParagraph(d, "Replace SECRET with the actual value when authorised.");
const plainBytes = await toBuffer(Document.build(d));

const password = "letmein-2026";
const encryptedBytes = await encryptDocx(plainBytes, password, {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 50_000 // smaller spin count for example speed
});
fs.writeFileSync(path.join(outDir, "01-plain.docx"), plainBytes);
fs.writeFileSync(path.join(outDir, "02-encrypted.docx"), encryptedBytes);
console.log(`  plain: ${plainBytes.length} bytes, encrypted: ${encryptedBytes.length} bytes`);

// ---------------------------------------------------------------------------
// 2. isEncryptedDocx detection
// ---------------------------------------------------------------------------
console.log(`  isEncryptedDocx(plain):     ${isEncryptedDocx(plainBytes)}`);
console.log(`  isEncryptedDocx(encrypted): ${isEncryptedDocx(encryptedBytes)}`);

// ---------------------------------------------------------------------------
// 3. readDocx with password — transparent decryption
// ---------------------------------------------------------------------------
const docModel = await readDocx(encryptedBytes, { password });
console.log(`  readDocx({password}) → body length: ${docModel.body.length}`);

// ---------------------------------------------------------------------------
// 4. Wrong password → DocxDecryptionError
// ---------------------------------------------------------------------------
try {
  await readDocx(encryptedBytes, { password: "definitely-not-the-password" });
  console.log("  ERROR: expected wrong-password to throw");
} catch (err) {
  if (err instanceof DocxDecryptionError) {
    console.log(`  wrong password → DocxDecryptionError: "${err.message}"`);
  } else if (isDocxError(err)) {
    console.log(`  wrong password → DocxError (${err.constructor.name}): "${err.message}"`);
  } else {
    throw err;
  }
}

// ---------------------------------------------------------------------------
// 5. decryptDocx (low-level) — get the raw ZIP bytes back, then read normally
// ---------------------------------------------------------------------------
const zipBytes = await decryptDocx(encryptedBytes, password);
console.log(`  decryptDocx() → ${zipBytes.length} bytes of plain ZIP`);
const fromZip = await readDocx(zipBytes);
console.log(`  re-read decrypted bytes → body length: ${fromZip.body.length}`);

// ---------------------------------------------------------------------------
// 6. Edit-in-place round trip: decrypt → modify → re-encrypt
// ---------------------------------------------------------------------------
const editable = await readDocx(encryptedBytes, { password });
const replacements = replaceText(editable, "SECRET", "approved value 42");
console.log(`  replaceText replaced ${replacements} occurrence(s)`);
const editedBytes = await toBuffer(editable);
const reEncrypted = await encryptDocx(editedBytes, password, {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 50_000
});
fs.writeFileSync(path.join(outDir, "03-edited-encrypted.docx"), reEncrypted);
console.log(`  → 03-edited-encrypted.docx (${reEncrypted.length} bytes)`);

// Verify the edit took effect after another decrypt cycle
const verify = await readDocx(reEncrypted, { password });
const flat = JSON.stringify(verify.body);
console.log(`  edit visible after re-decrypt: ${flat.includes("approved value 42")}`);
