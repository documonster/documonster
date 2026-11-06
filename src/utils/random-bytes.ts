/**
 * Cryptographically random bytes, from Web Crypto — which Node (as `globalThis.crypto`) and every supported browser
 * provide, so one platform-free module serves both.
 *
 * It is its own module rather than a member of `@utils/crypto` because that module also holds SHA-256, MD5 and the
 * AES block cipher, and a bundler splits chunks by module: a ZIP writer that needs only an encryption header's salt
 * carried all of them into the first chunk.
 */

/** `getRandomValues` refuses more than this many bytes in one call (Web Crypto's quota). */
const MAX_PER_CALL = 65_536;

/** `length` random bytes. Filled in quota-sized pieces, so a length above Web Crypto's per-call quota is not refused. */
export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  for (let offset = 0; offset < length; offset += MAX_PER_CALL) {
    globalThis.crypto.getRandomValues(bytes.subarray(offset, offset + MAX_PER_CALL));
  }
  return bytes;
}
