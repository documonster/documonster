/**
 * PDF key derivation hash functions — Browser version.
 *
 * Pure JavaScript implementations for browser environments where
 * node:crypto is not available. Same API as the Node.js version.
 */

import { sha256, sha384, sha512, md5 } from "@utils/crypto";

/** MD5 for PDF key derivation — ISO 32000 Algorithm 2/3/7. */
export function pdfMd5(input: Uint8Array): Uint8Array {
  return md5(input);
}

/** SHA-256 for PDF key derivation — ISO 32000 Algorithm 2.A/2.B. */
export function pdfSha256(input: Uint8Array): Uint8Array {
  return sha256(input);
}

/** SHA-384 for PDF key derivation — ISO 32000-2 Algorithm 2.B. */
export function pdfSha384(input: Uint8Array): Uint8Array {
  return sha384(input);
}

/** SHA-512 for PDF key derivation — ISO 32000-2 Algorithm 2.B. */
export function pdfSha512(input: Uint8Array): Uint8Array {
  return sha512(input);
}
