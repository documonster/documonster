import { checkLimit, checkString } from "@word/core/argument-checks";
/**
 * DOCX Module - Encryption & Digital Signatures (Subpath Entry)
 *
 * Public API at `documonster/word/crypto`. Referenced by
 * `package.json#exports["./word/crypto"]`. Import this subpath separately to
 * keep encryption / CFB / signature code out of bundles that only need core
 * document building.
 *
 * @example
 * ```ts
 * import { WordCrypto } from "documonster/word/crypto";
 *
 * if (WordCrypto.isEncrypted(bytes)) {
 *   const zip = await WordCrypto.decrypt(bytes, "secret");
 * }
 * ```
 */
import {
  deobfuscateFont as deobfuscateFontImpl,
  obfuscateFont as obfuscateFontImpl,
  generateFontKey as generateFontKeyImpl
} from "@word/font/font-obfuscation";
import { readCfb as readCfbImpl, writeCfb as writeCfbImpl } from "@word/security/cfb-reader";
import {
  extractSignatures as extractSignaturesImpl,
  hasDigitalSignatures as hasDigitalSignaturesImpl,
  parseSignatureXml as parseSignatureXmlImpl,
  isWellFormedSignature as isWellFormedSignatureImpl
} from "@word/security/digital-signatures";
import {
  AGILE_BLOCK_KEYS as AGILE_BLOCK_KEYS_IMPL,
  decryptDocx as decryptDocxImpl,
  decryptPackage as decryptPackageImpl,
  deriveEncryptionKey as deriveEncryptionKeyImpl,
  encryptDocx as encryptDocxImpl,
  isEncryptedDocx as isEncryptedDocxImpl,
  parseEncryptionInfoXml as parseEncryptionInfoXmlImpl,
  verifyPassword as verifyPasswordImpl
} from "@word/security/encryption";
import type { AgileEncryptionInfo } from "@word/security/encryption";

export * as WordCrypto from "@word/surface/crypto";
export type { AgileEncryptionInfo, EncryptOptions } from "@word/security/encryption";
export type { CfbEntry } from "@word/security/cfb-reader";
export type { DigitalSignatureInfo } from "@word/security/digital-signatures";

// The errors this entry's functions throw (also exported from the main entry):
// `Io.read` → `DocxEncryptedError`, `WordCrypto.decrypt` → `DocxDecryptionError`,
// a wrong argument → `DocxArgumentError`.
export {
  DocxArgumentError,
  DocxDecryptionError,
  DocxEncryptedError,
  isDocxError
} from "@word/errors";

/** @deprecated Use {@link WordCrypto.isEncrypted}. Removed in 1.0. */
export const isEncryptedDocx = isEncryptedDocxImpl;
/** @deprecated Use {@link WordCrypto.decrypt}. Removed in 1.0. */
export const decryptDocx = decryptDocxImpl;
/** @deprecated Use {@link WordCrypto.encrypt}. Removed in 1.0. */
export const encryptDocx = encryptDocxImpl;
/** @deprecated Use {@link WordCrypto.listSignatures}. Removed in 1.0. */
export const extractSignatures = extractSignaturesImpl;
/**
 * @deprecated Use {@link WordCrypto.hasSignatures}, which takes the document or
 * the package's part names (`Package.listParts(bytes)`). Removed in 1.0.
 */
export const hasDigitalSignatures = hasDigitalSignaturesImpl;
/** @deprecated Use {@link WordCrypto.isWellFormedSignature}. Removed in 1.0. */
export const isWellFormedSignature = isWellFormedSignatureImpl;
/** @deprecated Use {@link WordCrypto.parseSignature}. Removed in 1.0. */
export const parseSignatureXml = parseSignatureXmlImpl;
/**
 * @deprecated Use {@link WordCrypto.Agile.verifyPassword}, which takes
 * `(info, password)` instead of `(password, info)`. Removed in 1.0.
 */
export const verifyPassword = verifyPasswordImpl;
/**
 * @deprecated Use {@link WordCrypto.Agile.decryptPackage}, which takes the cap in
 * its options (`{ maxDecryptedSize }`). Removed in 1.0.
 */
export const decryptPackage = (
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  password: string,
  maxDecryptedSize?: number
): Promise<Uint8Array> => {
  const where = "WordCrypto.Agile.decryptPackage";
  checkString(password, "password", where);
  checkLimit(maxDecryptedSize, "maxDecryptedSize", where, { optional: true });
  return decryptPackageImpl(encryptedPackage, info, password, maxDecryptedSize);
};
/** @deprecated Use {@link WordCrypto.Agile.parseEncryptionInfo}. Removed in 1.0. */
export const parseEncryptionInfoXml = parseEncryptionInfoXmlImpl;
/** @deprecated Use {@link WordCrypto.Agile.deriveKey}. Removed in 1.0. */
export const deriveEncryptionKey = deriveEncryptionKeyImpl;
/** @deprecated Use {@link WordCrypto.Agile.BLOCK_KEYS}. Removed in 1.0. */
export const AGILE_BLOCK_KEYS = AGILE_BLOCK_KEYS_IMPL;
/** @deprecated Use {@link WordCrypto.Cfb.read}. Removed in 1.0. */
export const readCfb = readCfbImpl;
/** @deprecated Use {@link WordCrypto.Cfb.write}. Removed in 1.0. */
export const writeCfb = writeCfbImpl;
/** @deprecated Use `Font.obfuscate` from `documonster/word`. Removed in 1.0. */
export const obfuscateFont = obfuscateFontImpl;
/** @deprecated Use `Font.deobfuscate` from `documonster/word`. Removed in 1.0. */
export const deobfuscateFont = deobfuscateFontImpl;
/** @deprecated Use `Font.generateObfuscationKey` from `documonster/word`. Removed in 1.0. */
export const generateFontKey = generateFontKeyImpl;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
