/**
 * `WordCrypto` namespace surface — DOCX encryption and digital signatures, at
 * `documonster/word/crypto`.
 *
 * `import { WordCrypto } from "documonster/word/crypto"` →
 *   `WordCrypto.isEncrypted(bytes)`, `WordCrypto.decrypt(bytes, pw)`,
 *   `WordCrypto.encrypt(zipBytes, pw)`, `WordCrypto.listSignatures(doc)`,
 *   `WordCrypto.parseSignature(xml, name)`, `WordCrypto.verifyPassword(bytes, pw)`.
 *
 * The low-level primitives behind them live in two sub-namespaces:
 * `WordCrypto.Agile` (encryption-info parsing, key derivation, package
 * decryption) and `WordCrypto.Cfb` (the compound-file container). Font
 * obfuscation is a font concern and lives in `Font` in `documonster/word`.
 */
import { isRecordArgument } from "@utils/object";
import {
  checkHandle,
  checkLimit,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import {
  extractSignatures,
  hasDigitalSignatures,
  isWellFormedSignature as isWellFormedSignatureImpl,
  parseSignatureXml
} from "@word/security/digital-signatures";
import type { DigitalSignatureInfo } from "@word/security/digital-signatures";
import { decryptDocx, verifyDocxPassword } from "@word/security/encryption";
import type { DecryptOptions } from "@word/security/encryption";
import { DEFAULT_SECURITY_POLICY } from "@word/security/policy";
import type { DocxDocument } from "@word/types";

export { isEncryptedDocx as isEncrypted, encryptDocx as encrypt } from "@word/security/encryption";

/**
 * Whether `password` opens the encrypted DOCX `buffer`, checked against its
 * password verifier without decrypting the package.
 *
 * @throws {DocxParseError} when `buffer` is not a readable encrypted package.
 * @throws {DocxDecryptionError} `reason: "malformed"` when its verifier is damaged — not answered as a wrong password.
 */
export function verifyPassword(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  password: string
): Promise<boolean> {
  checkString(password, "password", "WordCrypto.verifyPassword");
  return verifyDocxPassword(buffer, password);
}
export * as Agile from "@word/surface/crypto-agile";
export * as Cfb from "@word/surface/crypto-cfb";

/**
 * Parse the metadata of one signature part (`_xmlsignatures/sig1.xml`): its XML
 * text and the part's file name (`"sig1.xml"`).
 */
export function parseSignature(xml: string, fileName: string): DigitalSignatureInfo {
  const where = "WordCrypto.parseSignature";
  checkString(xml, "xml", where);
  checkString(fileName, "fileName", where);
  return parseSignatureXml(xml, fileName);
}

/**
 * Whether a parsed signature has the elements XMLDSig requires (`Signature`,
 * `SignedInfo`, `SignatureValue`, `KeyInfo`). Not a cryptographic check — see
 * `cryptographicStatus`.
 */
export function isWellFormedSignature(info: DigitalSignatureInfo): boolean {
  const where = "WordCrypto.isWellFormedSignature";
  checkRecord(info, "info", where);
  checkString(info.rawXml, "info.rawXml", where);
  return isWellFormedSignatureImpl(info);
}

export type { DecryptOptions } from "@word/security/encryption";

/**
 * Decrypt a password-protected (Agile-encrypted) DOCX.
 *
 * @param buffer - The encrypted DOCX file (CFB format).
 * @param password - The password to decrypt with.
 * @param options - Decryption limits.
 * @returns The decrypted DOCX ZIP bytes, ready for `Document.read`.
 * @throws {DocxDecryptionError} If the file is not encrypted, the password is
 *   wrong, or decryption fails.
 */
export async function decrypt(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  password: string,
  options?: DecryptOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "WordCrypto.decrypt";
  checkString(password, "password", where);
  checkRecord(options, "options", where, { optional: true });
  checkLimit(options?.maxDecryptedSize, "options.maxDecryptedSize", where, { optional: true });
  // The cap `Document.read` applies, so a package this returns is one `Document.read` opens.
  return decryptDocx(
    buffer,
    password,
    options?.maxDecryptedSize ?? DEFAULT_SECURITY_POLICY.maxPackageSize
  );
}

/** Whether a document carries any digital signature parts. */
export function hasSignatures(doc: DocxDocument): boolean;
/**
 * Whether a package with these part names carries any digital signature
 * parts — so a file can be checked without reading it into a document:
 * `WordCrypto.hasSignatures(await Package.listParts(bytes))`.
 */
export function hasSignatures(partNames: readonly string[]): boolean;
export function hasSignatures(source: DocxDocument | readonly string[]): boolean {
  if (isRecordArgument(source)) {
    checkHandle(source, "document", "source", "WordCrypto.hasSignatures");
    return hasDigitalSignatures(((source as DocxDocument).opaqueParts ?? []).map(p => p.path));
  }
  if (!Array.isArray(source) || source.some(name => typeof name !== "string")) {
    invalidArgument(
      "WordCrypto.hasSignatures",
      "source",
      "a document or an array of part names",
      source
    );
  }
  return hasDigitalSignatures((source as readonly string[]).map(name => name.replace(/^\/+/, "")));
}

/** List the digital signatures a document carries, parsed from its signature parts. */
export function listSignatures(doc: DocxDocument): DigitalSignatureInfo[] {
  checkHandle(doc, "document", "doc", "WordCrypto.listSignatures");
  return extractSignatures(doc.opaqueParts);
}
