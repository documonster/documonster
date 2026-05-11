/**
 * DOCX Module - Encryption & Digital Signatures (Subpath Entry)
 *
 * Re-exports the crypto API at `excelts/word/crypto`. This file is referenced
 * by `package.json#exports["./word/crypto"]`; it forwards to the implementation
 * under `./security/crypto`.
 *
 * @example
 * ```ts
 * import { isEncryptedDocx, decryptDocx, encryptDocx } from "excelts/word/crypto";
 * ```
 */

export {
  isEncryptedDocx,
  verifyPassword,
  decryptPackage,
  decryptDocx,
  encryptDocx,
  parseEncryptionInfoXml,
  deriveEncryptionKey,
  AGILE_BLOCK_KEYS,
  DocxEncryptedError,
  readCfb,
  writeCfb,
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature,
  deobfuscateFont,
  obfuscateFont,
  generateFontKey
} from "./security/crypto";
export type {
  AgileEncryptionInfo,
  EncryptOptions,
  CfbEntry,
  DigitalSignatureInfo
} from "./security/crypto";
