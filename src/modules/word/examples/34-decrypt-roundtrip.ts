/**
 * Word Example 34 — Decryption round-trip
 *
 * Covers:
 *   - WordCrypto.encrypt → file → Io.read({password}) — the most common path,
 *     Io.read handles decryption transparently when a password is supplied.
 *   - WordCrypto.isEncrypted — detect whether bytes are an encrypted CFB.
 *   - WordCrypto.decrypt (low-level) — decrypt to raw ZIP bytes.
 *   - Wrong password → DocxDecryptionError.
 *   - Round-trip after edit (decrypt → mutate → encrypt back).
 *
 * Output: tmp/word-examples/34-decrypt/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { WordCrypto } from "../crypto";
import { Document, DocxDecryptionError, isDocxError, Io } from "../index";

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
const plainBytes = await Io.toBuffer(d);

const password = "letmein-2026";
const encryptedBytes = await WordCrypto.encrypt(plainBytes, password, {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 50_000 // smaller spin count for example speed
});
fs.writeFileSync(path.join(outDir, "01-plain.docx"), plainBytes);
fs.writeFileSync(path.join(outDir, "02-encrypted.docx"), encryptedBytes);
console.log(`  plain: ${plainBytes.length} bytes, encrypted: ${encryptedBytes.length} bytes`);

// ---------------------------------------------------------------------------
// 2. WordCrypto.isEncrypted detection
// ---------------------------------------------------------------------------
console.log(`  WordCrypto.isEncrypted(plain):     ${WordCrypto.isEncrypted(plainBytes)}`);
console.log(`  WordCrypto.isEncrypted(encrypted): ${WordCrypto.isEncrypted(encryptedBytes)}`);

// ---------------------------------------------------------------------------
// 3. Io.read with password — transparent decryption
// ---------------------------------------------------------------------------
const docModel = await Io.read(encryptedBytes, { password });
console.log(`  Io.read({password}) → body length: ${docModel.body.length}`);

// ---------------------------------------------------------------------------
// 4. Wrong password → DocxDecryptionError
// ---------------------------------------------------------------------------
try {
  await Io.read(encryptedBytes, { password: "definitely-not-the-password" });
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
// 5. WordCrypto.decrypt (low-level) — get the raw ZIP bytes back, then read normally
// ---------------------------------------------------------------------------
const zipBytes = await WordCrypto.decrypt(encryptedBytes, password);
console.log(`  WordCrypto.decrypt() → ${zipBytes.length} bytes of plain ZIP`);
const fromZip = await Io.read(zipBytes);
console.log(`  re-read decrypted bytes → body length: ${fromZip.body.length}`);

// ---------------------------------------------------------------------------
// 6. Edit-in-place round trip: decrypt → modify → re-encrypt
// ---------------------------------------------------------------------------
const editable = await Io.read(encryptedBytes, { password });
const replacements = Document.replaceText(editable, "SECRET", "approved value 42");
console.log(`  replaceText replaced ${replacements} occurrence(s)`);
const editedBytes = await Io.toBuffer(editable);
const reEncrypted = await WordCrypto.encrypt(editedBytes, password, {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 50_000
});
fs.writeFileSync(path.join(outDir, "03-edited-encrypted.docx"), reEncrypted);
console.log(`  → 03-edited-encrypted.docx (${reEncrypted.length} bytes)`);

// Verify the edit took effect after another decrypt cycle
const verify = await Io.read(reEncrypted, { password });
const flat = JSON.stringify(verify.body);
console.log(`  edit visible after re-decrypt: ${flat.includes("approved value 42")}`);
