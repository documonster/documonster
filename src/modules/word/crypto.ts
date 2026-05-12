/**
 * DOCX Module - Encryption & Digital Signatures (Subpath Entry)
 *
 * Public API at `excelts/word/crypto`. Referenced by
 * `package.json#exports["./word/crypto"]`. Import this subpath separately to
 * keep encryption / CFB / signature / font obfuscation code out of bundles
 * that only need core document building.
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
} from "./security/encryption";
export type { AgileEncryptionInfo, EncryptOptions } from "./security/encryption";

// Encryption error type (also exported from main entry)
export { DocxEncryptedError } from "./errors";

// CFB (Compound File Binary) reader/writer
export { readCfb, writeCfb } from "./security/cfb-reader";
export type { CfbEntry } from "./security/cfb-reader";

// Digital signature utilities
export {
  hasDigitalSignatures,
  parseSignatureXml,
  extractSignatures,
  isWellFormedSignature
} from "./security/digital-signatures";
export type { DigitalSignatureInfo } from "./security/digital-signatures";

// Font obfuscation utilities
export { deobfuscateFont, obfuscateFont, generateFontKey } from "./font/font-obfuscation";
