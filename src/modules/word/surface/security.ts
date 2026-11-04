/**
 * `Security` namespace surface — document protection and security policies.
 * Encryption lives in `WordCrypto` (`documonster/word/crypto`).
 *
 * `import { Security } from "documonster/word"` →
 *   `Security.protect(...)`, `Security.isProtected(doc)`,
 *   `Security.verifyPassword(doc, pw)` (the protection password; the file-encryption
 *   password is `WordCrypto.verifyPassword`),
 *   `Security.DEFAULT_SECURITY_POLICY`, … — tree-shaken via
 *   `export * as Security`.
 */
import type { encryptDocx } from "@word/security/encryption";
import { resolveSecurityPolicy as resolveSecurityPolicyImpl } from "@word/security/policy";

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
  resolveSecurityPolicy as resolvePolicy
} from "@word/security/policy";

/**
 * @deprecated Use `WordCrypto.encrypt` from `documonster/word/crypto`, beside
 * `decrypt` and `verifyPassword`. Removed in 1.0.
 */
export const encrypt: typeof encryptDocx = async (zipBytes, password, options) =>
  // Loaded on call, so `Security` does not bundle the encryption code for a deprecated forward.
  (await import("@word/security/encryption")).encryptDocx(zipBytes, password, options);
/** @deprecated Use {@link resolvePolicy}. Removed in 1.0. */
export const resolveSecurityPolicy = resolveSecurityPolicyImpl;
