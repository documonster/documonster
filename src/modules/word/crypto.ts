/**
 * DOCX Module - Encryption & Digital Signatures (Subpath Export)
 *
 * Import separately to avoid pulling crypto code into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { isEncryptedDocx, decryptPackage } from "excelts/word/crypto";
 * import { extractSignatures } from "excelts/word/crypto";
 * ```
 */

// Encryption utilities
export {
  isEncryptedDocx,
  verifyPassword,
  decryptPackage,
  parseEncryptionInfoXml,
  deriveEncryptionKey,
  AGILE_BLOCK_KEYS
} from "./encryption";
export type { AgileEncryptionInfo } from "./encryption";

// Digital signature utilities
export {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "./digital-signatures";
export type { DigitalSignatureInfo } from "./digital-signatures";

// Font obfuscation utilities
export { deobfuscateFont, obfuscateFont, generateFontKey } from "./font-obfuscation";
