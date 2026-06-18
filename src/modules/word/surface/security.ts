/**
 * `Security` namespace surface — encryption, document protection, security
 * policies.
 *
 * `import { Security } from "documonster/word"` →
 *   `Security.encrypt(...)`, `Security.protect(...)`, `Security.isProtected(doc)`,
 *   `Security.DEFAULT_SECURITY_POLICY`, … — tree-shaken via
 *   `export * as Security`.
 */
export { encryptDocx as encrypt } from "@word/security/encryption";
export {
  protectDocument as protect,
  unprotectDocument as unprotect,
  isDocumentProtected as isProtected,
  getProtectionState as getState,
  verifyProtectionPassword as verifyPassword
} from "@word/security/document-protection";
export {
  DEFAULT_SECURITY_POLICY,
  STRICT_SECURITY_POLICY,
  resolveSecurityPolicy
} from "@word/security/policy";
