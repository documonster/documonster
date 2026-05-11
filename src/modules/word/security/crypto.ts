/**
 * DOCX Module - Encryption & Digital Signatures (Subpath Export)
 *
 * Import separately to avoid pulling crypto code into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { isEncryptedDocx, decryptDocx, encryptDocx } from "excelts/word/crypto";
 * import { extractSignatures } from "excelts/word/crypto";
 * ```
 */

// Encryption utilities
export {
  isEncryptedDocx,
  verifyPassword,
  decryptPackage,
  decryptDocx,
  encryptDocx,
  parseEncryptionInfoXml,
  deriveEncryptionKey,
  AGILE_BLOCK_KEYS
} from "./encryption";
export type { AgileEncryptionInfo, EncryptOptions } from "./encryption";

// Encryption error type (also exported from main entry)
export { DocxEncryptedError } from "../errors";

// CFB (Compound File Binary) reader/writer
export { readCfb, writeCfb } from "./cfb-reader";
export type { CfbEntry } from "./cfb-reader";

// Digital signature utilities
export {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "./digital-signatures";
export type { DigitalSignatureInfo } from "./digital-signatures";

// Font obfuscation utilities
export { deobfuscateFont, obfuscateFont, generateFontKey } from "../font/font-obfuscation";
