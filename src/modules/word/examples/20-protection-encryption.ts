/**
 * Word Example 20 — Document protection & encryption
 *
 * Covers:
 *   - protectDocument with various edit restrictions (readOnly / comments /
 *     trackedChanges / forms)
 *   - With and without a password (different hash algorithms / spin counts)
 *   - isDocumentProtected, getProtectionState, verifyProtectionPassword
 *   - unprotectDocument
 *   - encryptDocx — wrap a finished DOCX in MS-Agile encryption (the file
 *     can only be opened by clients that know the password)
 *
 * Output: tmp/word-examples/20-protection/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Io, Security } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/20-protection"
);
fs.mkdirSync(outDir, { recursive: true });

// Build a small base document
function makeDoc(): ReturnType<typeof Document.build> {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Confidential", 1);
  Document.addParagraph(d, "This document demonstrates document protection.");
  return Document.build(d);
}

// ---------------------------------------------------------------------------
// 1. Read-only edit restriction with password (SHA-256, 100k spin)
//
//    NOTE on how this shows up in Word:
//    `protectDocument({ edit: "readOnly" })` writes OOXML
//    <w:documentProtection> — an *editing restriction*, NOT "mark as
//    final" and NOT whole-file encryption. Word does NOT pop up a password
//    dialog when opening the file. Instead the document opens read-only;
//    to make edits you go to the Review tab → "Restrict Editing" →
//    "Stop Protection", at which point Word prompts for the password
//    (`swordfish`). (Some Word builds also surface a "Restrict Editing"
//    task pane automatically.)
// ---------------------------------------------------------------------------
const readOnlyProtected = await Security.protect(makeDoc(), {
  edit: "readOnly",
  password: "swordfish",
  hashAlgorithm: "SHA-256",
  spinCount: 100_000
});
console.log("  isDocumentProtected:", Security.isProtected(readOnlyProtected));
console.log("  protection state:", Security.getState(readOnlyProtected));
console.log(
  "  verifyProtectionPassword('swordfish'):",
  await Security.verifyPassword(readOnlyProtected, "swordfish")
);
console.log(
  "  verifyProtectionPassword('wrong'):",
  await Security.verifyPassword(readOnlyProtected, "wrong")
);
fs.writeFileSync(path.join(outDir, "01-readonly.docx"), await Io.toBuffer(readOnlyProtected));

// ---------------------------------------------------------------------------
// 2. Comments-only restriction (no password — anyone can disable in Word)
// ---------------------------------------------------------------------------
const commentsOnly = await Security.protect(makeDoc(), { edit: "comments" });
fs.writeFileSync(path.join(outDir, "02-comments-only.docx"), await Io.toBuffer(commentsOnly));

// ---------------------------------------------------------------------------
// 3. Tracked-changes-only with SHA-512
// ---------------------------------------------------------------------------
const tracked = await Security.protect(makeDoc(), {
  edit: "trackedChanges",
  password: "p@ssw0rd",
  hashAlgorithm: "SHA-512",
  spinCount: 50_000,
  formatting: true
});
fs.writeFileSync(path.join(outDir, "03-tracked-changes.docx"), await Io.toBuffer(tracked));

// ---------------------------------------------------------------------------
// 4. Forms only — typical for "fillable form" DOCX
// ---------------------------------------------------------------------------
const formsOnly = await Security.protect(makeDoc(), { edit: "forms", password: "fill" });
fs.writeFileSync(path.join(outDir, "04-forms-only.docx"), await Io.toBuffer(formsOnly));

// ---------------------------------------------------------------------------
// 5. Unprotect
// ---------------------------------------------------------------------------
const noLongerProtected = Security.unprotect(readOnlyProtected);
console.log("  after unprotect, isProtected =", Security.isProtected(noLongerProtected));
fs.writeFileSync(path.join(outDir, "05-unprotected.docx"), await Io.toBuffer(noLongerProtected));

// ---------------------------------------------------------------------------
// 6. Agile encryption — the whole DOCX is encrypted, not just settings.
//    The output is an OLE2 (CFB) container that Word opens with the password.
// ---------------------------------------------------------------------------
const plainBuf = await Io.toBuffer(makeDoc());
fs.writeFileSync(path.join(outDir, "06-plain.docx"), plainBuf);
const encrypted = await Security.encrypt(plainBuf, "topsecret", {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 100_000
});
fs.writeFileSync(path.join(outDir, "06-encrypted.docx"), encrypted);
console.log(
  `  → 06-plain.docx (${plainBuf.length} bytes), 06-encrypted.docx (${encrypted.length} bytes)`
);
