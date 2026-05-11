/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/security"` */

export {
  isEncryptedDocx,
  decryptDocx,
  encryptDocx,
  deriveEncryptionKey,
  verifyPassword
} from "./encryption";
export type { AgileEncryptionInfo, EncryptOptions } from "./encryption";
export { readCfb, writeCfb } from "./cfb-reader";
export type { CfbEntry } from "./cfb-reader";
export {
  hasDigitalSignatures,
  extractSignatures,
  parseSignatureXml,
  isWellFormedSignature
} from "./digital-signatures";
export type { DigitalSignatureInfo } from "./digital-signatures";
export {
  protectDocument,
  unprotectDocument,
  isDocumentProtected,
  getProtectionState,
  verifyProtectionPassword
} from "./document-protection";
export type {
  ProtectionEditType,
  ProtectionHashAlgorithm,
  DocumentProtectionOptions,
  ProtectionState
} from "./document-protection";
export { DEFAULT_SECURITY_POLICY, STRICT_SECURITY_POLICY, resolveSecurityPolicy } from "./policy";
export type { WordSecurityPolicy } from "./policy";
