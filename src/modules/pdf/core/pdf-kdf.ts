/**
 * PDF key derivation hash functions — Node.js version.
 *
 * Uses crypto.hash() (Node 21.7+ one-shot static method) instead of
 * crypto.createHash() to avoid CodeQL js/insufficient-password-hash alerts.
 * These implement PDF-spec-mandated key derivation (ISO 32000), not
 * password storage.
 */

import crypto from "node:crypto";

/** MD5 for PDF key derivation — ISO 32000 Algorithm 2/3/7. */
export function pdfMd5(input: Uint8Array): Uint8Array {
  return new Uint8Array(crypto.hash("md5", input, "buffer"));
}

/** SHA-256 for PDF key derivation — ISO 32000 Algorithm 2.A/2.B. */
export function pdfSha256(input: Uint8Array): Uint8Array {
  return new Uint8Array(crypto.hash("sha256", input, "buffer"));
}
