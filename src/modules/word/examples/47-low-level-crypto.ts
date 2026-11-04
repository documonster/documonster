/**
 * Word Example 47 — Low-level crypto building blocks
 *
 * Complements `34-decrypt-roundtrip.ts` by exercising the lower-level helpers
 * exposed on `WordCrypto` (`documonster/word/crypto`):
 *   - WordCrypto.verifyPassword — check a password against an encrypted file
 *     in one call, without decrypting it.
 *   - WordCrypto.Cfb.read / WordCrypto.Cfb.write — parse and synthesise OLE2 Compound Files. The
 *     encrypted DOCX wire format wraps a ZIP package inside CFB streams; the
 *     same primitives also underpin .doc, .ppt, .xls (legacy formats) and OLE
 *     embedded objects.
 *   - WordCrypto.Agile.parseEncryptionInfo — turn an EncryptionInfo XML string into an
 *     AgileEncryptionInfo struct.
 *   - WordCrypto.Agile.verifyPassword — the same check against parsed info.
 *   - WordCrypto.Agile.BLOCK_KEYS — the constant block-key table consumed by
 *     WordCrypto.Agile.deriveKey when implementing custom workflows.
 *   - WordCrypto.Agile.decryptPackage — decrypt a raw EncryptedPackage stream.
 *
 * Output: tmp/word-examples/47-low-crypto.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { WordCrypto } from "../crypto";
import type { CfbEntry } from "../crypto";
import { Document, Io } from "../index";

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
// 1. Agile.BLOCK_KEYS — exposed for WordCrypto.Agile.deriveKey() callers
// ---------------------------------------------------------------------------
log(`  Agile.BLOCK_KEYS keys: ${Object.keys(WordCrypto.Agile.BLOCK_KEYS).join(", ")}`);
for (const [k, v] of Object.entries(WordCrypto.Agile.BLOCK_KEYS)) {
  log(
    `    ${k.padEnd(28)} = ${Array.from(v as Uint8Array)
      .map(b => b.toString(16).padStart(2, "0"))
      .join(" ")}`
  );
}

// ---------------------------------------------------------------------------
// 2. writeCfb / readCfb — round-trip a synthetic CFB file containing two
//    named streams. This is exactly the wire format that encrypted DOCX
//    files use (with streams "EncryptionInfo" + "EncryptedPackage").
// ---------------------------------------------------------------------------
const inputEntries: CfbEntry[] = [
  { name: "Hello", data: new TextEncoder().encode("hello world") },
  { name: "Bin", data: Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x05]) }
];
const cfbBytes = WordCrypto.Cfb.write(inputEntries);
fs.writeFileSync(path.join(outDir, "47-test.cfb"), cfbBytes);
log(`\n  writeCfb → ${cfbBytes.length} bytes`);

const decoded = WordCrypto.Cfb.read(cfbBytes);
log(`  readCfb  → ${decoded.length} streams:`);
for (const e of decoded) {
  log(`    ${e.name.padEnd(20)} ${e.data.length} bytes`);
}

// ---------------------------------------------------------------------------
// 3. WordCrypto.Agile.parseEncryptionInfo — extract Agile encryption parameters from an
//    EncryptionInfo XML string. Real encrypted docx files store this as a
//    stream inside the CFB envelope; we synthesise one with WordCrypto.encrypt then
//    locate its EncryptionInfo stream and parse it.
// ---------------------------------------------------------------------------
const plain = await Io.toBuffer(
  (() => {
    const dd = Document.create();
    Document.useDefaultStyles(dd);
    Document.addParagraph(dd, "Crypto plumbing demo");
    return dd;
  })()
);
const password = "knock-knock";
const encrypted = await WordCrypto.encrypt(plain, password, {
  keyBits: 256,
  hashAlgorithm: "SHA512",
  spinCount: 50_000
});
const encryptedEntries = WordCrypto.Cfb.read(encrypted);
const infoEntry = encryptedEntries.find(e => e.name === "EncryptionInfo");
if (infoEntry) {
  // Skip the 8-byte version header at the start of the EncryptionInfo stream
  // (4 bytes major/minor version, 4 bytes flags) so the trailing bytes are
  // the XML payload.
  const xmlBytes = infoEntry.data.slice(8);
  const xmlText = new TextDecoder().decode(xmlBytes);
  const info = WordCrypto.Agile.parseEncryptionInfo(xmlText);
  log(`\n  parseEncryptionInfo:`);
  log(
    `    keyBits=${info.keyBits}, hashAlgorithm=${info.hashAlgorithm}, spinCount=${info.spinCount}`
  );
  log(`    cipherAlgorithm=${info.cipherAlgorithm}, cipherChaining=${info.cipherChaining}`);
  log(`    blockSize=${info.blockSize}, hashSize=${info.hashSize}`);
  log(
    `    keySalt=${info.keySalt.length}B, encryptedVerifierHashInput=${info.encryptedVerifierHashInput.length}B`
  );

  // 4. verifyEncryptionPassword — check the password against the parsed info, without
  //    actually decrypting the (potentially large) data stream.
  const valid = await WordCrypto.Agile.verifyPassword(info, password);
  log(`  verifyEncryptionPassword("${password}"): ${valid}`);

  const wrong = await WordCrypto.Agile.verifyPassword(info, "wrong-pw");
  log(`  verifyEncryptionPassword("wrong-pw"):    ${wrong}`);

  // 5. deriveKey — the per-block key derivation that powers
  //    verifyEncryptionPassword and decryptPackage. Useful when implementing custom
  //    encryption workflows (e.g. re-keying without re-encrypting the
  //    package data, or interoperating with non-OOXML CFB containers).
  //    Each entry in AGILE_BLOCK_KEYS gives a different key for a different
  //    purpose: encryptedKey, dataIntegrity, verifierHashInput, verifierHashValue.
  const dataIntegrityKey = await WordCrypto.Agile.deriveKey(
    password,
    info,
    WordCrypto.Agile.BLOCK_KEYS.dataIntegrityKey
  );
  log(`  deriveKey(dataIntegrityKey) → ${dataIntegrityKey.length} bytes (keyBits=${info.keyBits})`);

  // 6. decryptPackage — the lowest-level decryption primitive, taking the
  //    raw EncryptedPackage stream + the parsed info + the password and
  //    returning the decrypted ZIP bytes. `WordCrypto.decrypt()` is a thin wrapper
  //    around this; using it directly is useful when the encrypted bytes
  //    arrive separately from the EncryptionInfo (e.g. from a streaming
  //    source) or when integrating with a non-OOXML container.
  const encryptedPackage = encryptedEntries.find(e => e.name === "EncryptedPackage");
  if (encryptedPackage) {
    const decrypted = await WordCrypto.Agile.decryptPackage(encryptedPackage.data, info, password);
    log(
      `  decryptPackage → ${decrypted.length} bytes (matches plain=${decrypted.length === plain.length})`
    );
  }
} else {
  log(`  WARN: EncryptionInfo stream not found`);
}

// ---------------------------------------------------------------------------
// 7. WordCrypto.verifyPassword — the one-call form of steps 2–4: hand it the
//    encrypted file and a password, get a boolean. Useful to validate a
//    password prompt before committing to a full decrypt.
// ---------------------------------------------------------------------------
log(
  `\n  WordCrypto.verifyPassword(file, "${password}"): ${await WordCrypto.verifyPassword(encrypted, password)}`
);
log(
  `  WordCrypto.verifyPassword(file, "wrong-pw"):    ${await WordCrypto.verifyPassword(encrypted, "wrong-pw")}`
);

fs.writeFileSync(path.join(outDir, "47-low-crypto.txt"), lines.join("\n"));
console.log(`\n  → 47-low-crypto.txt`);
