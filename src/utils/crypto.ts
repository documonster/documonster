/**
 * Cryptographic primitives — Node.js version.
 *
 * Uses `node:crypto` for maximum performance where possible.
 * Pure JS fallbacks for operations where `node:crypto` would be inconvenient
 * (e.g., AES-CBC with specific PDF padding semantics).
 *
 * The browser counterpart (`crypto.browser.ts`) provides the same API using
 * pure JS implementations for synchronous ops and Web Crypto for async ops.
 *
 * Shared by: PDF (encryption/decryption), Archive (via re-export if needed),
 * and digital signature infrastructure.
 *
 * @see FIPS 197 — AES
 * @see FIPS 180-4 — SHA-256
 * @see RFC 1321 — MD5
 * @see RFC 2104 — HMAC
 */

import crypto from "node:crypto";

// =============================================================================
// SHA-256
// =============================================================================

/**
 * SHA-256 hash function (FIPS 180-4).
 * @returns 32-byte digest
 */
export function sha256(input: Uint8Array): Uint8Array {
  return new Uint8Array(crypto.hash("sha256", input, "buffer"));
}

// =============================================================================
// Generic hash
// =============================================================================

/**
 * HMAC-SHA256 (RFC 2104).
 * @returns 32-byte MAC
 */
export function hmacSha256(key: Uint8Array, message: Uint8Array): Uint8Array {
  return new Uint8Array(crypto.createHmac("sha256", key).update(message).digest());
}

// =============================================================================
// MD5
// =============================================================================

/**
 * MD5 hash function (RFC 1321).
 *
 * **Security note:** MD5 is cryptographically broken for password hashing.
 * This function exists solely for PDF specification compliance — ISO 32000
 * mandates MD5 in its key derivation algorithms for RC4 and AES-128
 * encryption (Algorithm 2, Algorithm 3 in PDF 1.7 Reference). It cannot
 * be replaced with a stronger hash without breaking compatibility with
 * every existing encrypted PDF file.
 *
 * @returns 16-byte digest
 */
export function md5(input: Uint8Array): Uint8Array {
  return new Uint8Array(crypto.hash("md5", input, "buffer"));
}

// =============================================================================
// AES-CBC
// =============================================================================

/**
 * AES-CBC encryption with PKCS#7 padding.
 * Supports AES-128 (16-byte key) and AES-256 (32-byte key).
 */
export function aesCbcEncrypt(plaintext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const algo = key.length === 16 ? "aes-128-cbc" : "aes-256-cbc";
  const cipher = crypto.createCipheriv(algo, key, iv);
  return new Uint8Array(Buffer.concat([cipher.update(plaintext), cipher.final()]));
}

/**
 * AES-CBC decryption with PKCS#7 padding removal.
 * Supports AES-128 (16-byte key) and AES-256 (32-byte key).
 */
export function aesCbcDecrypt(ciphertext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const algo = key.length === 16 ? "aes-128-cbc" : "aes-256-cbc";
  const decipher = crypto.createDecipheriv(algo, key, iv);
  try {
    return new Uint8Array(Buffer.concat([decipher.update(ciphertext), decipher.final()]));
  } catch {
    // Invalid padding — return raw decrypted data without padding removal
    // (matches pure JS behavior for NIST test vectors and PDF key derivation)
    const decipher2 = crypto.createDecipheriv(algo, key, iv);
    decipher2.setAutoPadding(false);
    return new Uint8Array(Buffer.concat([decipher2.update(ciphertext), decipher2.final()]));
  }
}

/**
 * AES-CBC decryption WITHOUT PKCS#7 padding removal.
 * Used for key derivation where the output length is known.
 */
export function aesCbcDecryptRaw(
  ciphertext: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Uint8Array {
  const algo = key.length === 16 ? "aes-128-cbc" : "aes-256-cbc";
  const decipher = crypto.createDecipheriv(algo, key, iv);
  decipher.setAutoPadding(false);
  return new Uint8Array(Buffer.concat([decipher.update(ciphertext), decipher.final()]));
}

/**
 * AES-CBC encryption WITHOUT PKCS#7 padding.
 * Used when the plaintext is already block-aligned.
 *
 * @throws if plaintext length is not a multiple of 16.
 */
export function aesCbcEncryptRaw(
  plaintext: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Uint8Array {
  if (plaintext.length % 16 !== 0) {
    throw new Error("aesCbcEncryptRaw: plaintext length must be a multiple of 16");
  }
  const algo = key.length === 16 ? "aes-128-cbc" : "aes-256-cbc";
  const cipher = crypto.createCipheriv(algo, key, iv);
  cipher.setAutoPadding(false);
  return new Uint8Array(Buffer.concat([cipher.update(plaintext), cipher.final()]));
}

/**
 * AES-ECB encryption of a single 16-byte block (no padding, no IV).
 */
export function aesEcbEncrypt(block: Uint8Array, key: Uint8Array): Uint8Array {
  const algo = key.length === 16 ? "aes-128-ecb" : "aes-256-ecb";
  const cipher = crypto.createCipheriv(algo, key, null);
  cipher.setAutoPadding(false);
  return new Uint8Array(Buffer.concat([cipher.update(block), cipher.final()]));
}

// =============================================================================
// RC4 (legacy)
// =============================================================================

/**
 * RC4 stream cipher.
 *
 * Required by ISO 32000 for reading PDFs encrypted with the RC4 algorithm
 * (PDF 1.4 standard handler V=1/V=2). Modern Node `crypto` no longer
 * exposes RC4, so this is a pure-JS implementation.
 */
export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  // Node's crypto doesn't expose RC4 in modern versions, use pure JS
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    s[i] = i;
  }
  let j = 0;
  for (let i = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 0xff;
    [s[i], s[j]] = [s[j], s[i]];
  }

  const result = new Uint8Array(data.length);
  let ii = 0;
  let jj = 0;
  for (let k = 0; k < data.length; k++) {
    ii = (ii + 1) & 0xff;
    jj = (jj + s[ii]) & 0xff;
    [s[ii], s[jj]] = [s[jj], s[ii]];
    result[k] = data[k] ^ s[(s[ii] + s[jj]) & 0xff];
  }
  return result;
}

// =============================================================================
// Random bytes
// =============================================================================

/**
 * Generate cryptographically secure random bytes.
 */
export function randomBytes(length: number): Uint8Array {
  return new Uint8Array(crypto.randomBytes(length));
}

// =============================================================================
// Generic hash
// =============================================================================

/**
 * Compute a hash digest using any algorithm supported by the platform.
 *
 * @param algorithm - Hash algorithm name (e.g., "SHA-256", "SHA-512", "SHA-1", "MD5").
 *   Normalized internally: hyphens removed, lowercased.
 * @param data - Data to hash
 * @returns The digest bytes
 */
export function hash(algorithm: string, data: Uint8Array): Uint8Array {
  const algo = algorithm.toLowerCase().replace(/-/g, "");
  return new Uint8Array(crypto.hash(algo, data, "buffer"));
}

/**
 * Async version of `hash()` — same behavior, but returns a Promise for API
 * parity with the browser version.
 */
export async function hashAsync(algorithm: string, data: Uint8Array): Promise<Uint8Array> {
  return hash(algorithm, data);
}

// =============================================================================
// RSA signature operations (async — for digital signatures)
// =============================================================================

/**
 * Verify an RSA PKCS#1 v1.5 signature.
 *
 * @param publicKeyDer - DER-encoded SubjectPublicKeyInfo
 * @param signature - The signature bytes
 * @param data - The signed data (will be hashed with SHA-256)
 */
export async function rsaVerify(
  publicKeyDer: Uint8Array,
  signature: Uint8Array,
  data: Uint8Array
): Promise<boolean> {
  const key = crypto.createPublicKey({
    key: Buffer.from(publicKeyDer),
    format: "der",
    type: "spki"
  });
  const verifier = crypto.createVerify("SHA256");
  verifier.update(data);
  return verifier.verify(key, signature);
}

/**
 * Create an RSA PKCS#1 v1.5 signature.
 *
 * @param privateKeyDer - DER-encoded PKCS#8 private key
 * @param data - The data to sign (will be hashed with SHA-256)
 */
export async function rsaSign(privateKeyDer: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const key = crypto.createPrivateKey({
    key: Buffer.from(privateKeyDer),
    format: "der",
    type: "pkcs8"
  });
  const signer = crypto.createSign("SHA256");
  signer.update(data);
  return new Uint8Array(signer.sign(key));
}
