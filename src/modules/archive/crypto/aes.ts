/**
 * AES encryption for ZIP files (WinZip AE-1 / AE-2).
 *
 * Implements the WinZip AES encryption format as documented in:
 * - https://www.winzip.com/en/support/aes-encryption/
 *
 * This implementation supports both AE-1 and AE-2 encryption modes:
 * - AE-1: Uses both CRC-32 and HMAC-SHA1 authentication
 * - AE-2: Uses only HMAC-SHA1 authentication (CRC-32 is set to 0)
 *
 * Key derivation uses PBKDF2 with HMAC-SHA1 as the PRF.
 * Encryption uses AES in CTR mode.
 *
 * Works in both Node.js and browsers using the Web Crypto API.
 */

import { checkBytes } from "@archive/core/entry-checks";
import { ArchiveError } from "@archive/core/errors";
import { toArrayBuffer } from "@archive/core/text";
import type { AesKeyStrength } from "@archive/crypto/aes-format";
import {
  AES_AUTH_CODE_LENGTH,
  AES_KEY_LENGTH,
  AES_PASSWORD_VERIFY_LENGTH,
  AES_SALT_LENGTH,
  checkKeyStrength
} from "@archive/crypto/aes-format";
import { checkZipPassword } from "@archive/crypto/zip-crypto";
import { stringToUint8Array as encodeUtf8 } from "@utils/binary";
import { randomBytes } from "@utils/random-bytes";

export {
  AES_AUTH_CODE_LENGTH,
  AES_EXTRA_FIELD_ID,
  AES_KEY_LENGTH,
  AES_PASSWORD_VERIFY_LENGTH,
  AES_SALT_LENGTH,
  AES_STRENGTH_BYTE,
  AES_STRENGTH_FROM_BYTE,
  AES_VENDOR_ID,
  AES_VERSION_AE1,
  AES_VERSION_AE2,
  COMPRESSION_METHOD_AES,
  buildAesExtraField,
  checkKeyStrength,
  parseAesExtraField
} from "@archive/crypto/aes-format";

/**
 * Derived key material from PBKDF2.
 */
export interface AesDerivedKeys {
  /** AES encryption key */
  encryptionKey: Uint8Array;
  /** HMAC-SHA1 key */
  hmacKey: Uint8Array;
  /** Password verification value (2 bytes) */
  passwordVerify: Uint8Array;
}

/**
 * Parsed AES encrypted data components.
 */
export interface AesEncryptedComponents {
  /** Salt (8/12/16 bytes depending on key strength) */
  salt: Uint8Array;
  /** Password verification value (2 bytes) */
  storedVerify: Uint8Array;
  /** Encrypted ciphertext */
  ciphertext: Uint8Array;
  /** HMAC-SHA1 authentication code (10 bytes) */
  storedHmac: Uint8Array;
}

/**
 * Get the Web Crypto API (works in both Node.js and browsers).
 */
function getWebCrypto(): SubtleCrypto {
  if (typeof globalThis.crypto?.subtle !== "undefined") {
    return globalThis.crypto.subtle;
  }
  throw new ArchiveError("Web Crypto API not available");
}

/**
 * Derive AES keys from password using PBKDF2.
 *
 * The derived key material is split into:
 * - Encryption key (16/24/32 bytes)
 * - HMAC key (16/24/32 bytes, same length as the encryption key)
 * - Password verification (2 bytes)
 */
export async function aesDerive(
  password: string | Uint8Array,
  salt: Uint8Array,
  keyStrength: AesKeyStrength
): Promise<AesDerivedKeys> {
  const crypto = getWebCrypto();
  const passwordBytes = typeof password === "string" ? encodeUtf8(password) : password;
  const keyLen = AES_KEY_LENGTH[keyStrength];

  // WinZip AES uses an HMAC-SHA1 key of the SAME length as the AES key
  // (16/24/32 bytes for AES-128/192/256), NOT a fixed 32 bytes. Hard-coding 32
  // only happens to be correct for AES-256; for AES-128/192 it shifts the
  // password-verification bytes, so interop readers (WinZip, 7-Zip, pyzipper)
  // reject the entry with "bad password".
  // Total derived bytes: encryption key + HMAC key (keyLen) + verification (2 bytes)
  const derivedLen = keyLen + keyLen + 2;

  // Import password as PBKDF2 key
  const baseKey = await crypto.importKey("raw", toArrayBuffer(passwordBytes), "PBKDF2", false, [
    "deriveBits"
  ]);

  // Derive key material using PBKDF2 with HMAC-SHA1
  const derivedBits = await crypto.deriveBits(
    {
      name: "PBKDF2",
      salt: toArrayBuffer(salt),
      iterations: 1000,
      hash: "SHA-1"
    },
    baseKey,
    derivedLen * 8
  );

  const derived = new Uint8Array(derivedBits);

  return {
    encryptionKey: derived.subarray(0, keyLen),
    hmacKey: derived.subarray(keyLen, keyLen + keyLen),
    passwordVerify: derived.subarray(keyLen + keyLen, keyLen + keyLen + 2)
  };
}

/**
 * Compute HMAC-SHA1 authentication code.
 *
 * @param key - HMAC key (16/24/32 bytes, matching the AES key strength)
 * @param data - Data to authenticate
 * @returns 10-byte authentication code (truncated from 20 bytes)
 */
export async function aesComputeHmac(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const crypto = getWebCrypto();

  const hmacKey = await crypto.importKey(
    "raw",
    toArrayBuffer(key),
    { name: "HMAC", hash: "SHA-1" },
    false,
    ["sign"]
  );

  const signature = await crypto.sign("HMAC", hmacKey, toArrayBuffer(data));
  // Truncate to 10 bytes per WinZip spec
  return new Uint8Array(signature, 0, AES_AUTH_CODE_LENGTH);
}

/**
 * AES-CTR counter block.
 * WinZip uses a little-endian counter starting at 1, held in the low bytes.
 */
function createCtrCounter(counter: number): Uint8Array {
  const block = new Uint8Array(16);
  const view = new DataView(block.buffer);
  view.setUint32(0, counter, true); // Little-endian counter
  return block;
}

/**
 * AES-CTR encryption/decryption (WinZip AES variant).
 *
 * WinZip uses a custom CTR mode that differs from the WebCrypto built-in
 * AES-CTR counter semantics:
 * - 16-byte blocks
 * - A LITTLE-ENDIAN counter starting at 1, incremented per block
 *
 * The WebCrypto `AES-CTR` algorithm increments the counter as a BIG-ENDIAN
 * integer, so it can only be relied upon for the first block. To stay
 * interoperable with WinZip / 7-Zip / pyzipper we generate the keystream one
 * block at a time: for each block we AES-encrypt that block's little-endian
 * counter (by running WebCrypto AES-CTR over 16 zero bytes, which yields
 * `AES(counterBlock) XOR 0 = AES(counterBlock)`), then XOR it into the data.
 *
 * Note: AES-CTR is symmetric — encrypt and decrypt are the same operation.
 */
export async function aesCtr(
  key: Uint8Array,
  data: Uint8Array,
  _encrypt: boolean = true
): Promise<Uint8Array> {
  const crypto = getWebCrypto();

  const aesKey = await crypto.importKey("raw", toArrayBuffer(key), { name: "AES-CTR" }, false, [
    "encrypt"
  ]);

  const out = new Uint8Array(data.length);
  const zeroBlock = new Uint8Array(16);
  const blockCount = Math.ceil(data.length / 16);

  for (let i = 0; i < blockCount; i++) {
    // WinZip counter starts at 1 for the first block, little-endian.
    const counterBlock = createCtrCounter(i + 1);
    const keystreamBuf = await crypto.encrypt(
      { name: "AES-CTR", counter: toArrayBuffer(counterBlock), length: 128 },
      aesKey,
      toArrayBuffer(zeroBlock)
    );
    const keystream = new Uint8Array(keystreamBuf);

    const start = i * 16;
    const end = Math.min(start + 16, data.length);
    for (let j = start; j < end; j++) {
      out[j] = data[j]! ^ keystream[j - start]!;
    }
  }

  return out;
}

/**
 * Extract components from AES-encrypted data.
 * Shared helper to avoid code duplication in aesDecrypt, aesCheckPasswordOnly, aesCheckSignature.
 */
function extractAesComponents(
  encryptedData: Uint8Array,
  keyStrength: AesKeyStrength
): AesEncryptedComponents | null {
  const saltLen = AES_SALT_LENGTH[keyStrength];
  const minLen = saltLen + AES_PASSWORD_VERIFY_LENGTH + AES_AUTH_CODE_LENGTH;

  if (encryptedData.length < minLen) {
    return null;
  }

  return {
    salt: encryptedData.subarray(0, saltLen),
    storedVerify: encryptedData.subarray(saltLen, saltLen + AES_PASSWORD_VERIFY_LENGTH),
    ciphertext: encryptedData.subarray(
      saltLen + AES_PASSWORD_VERIFY_LENGTH,
      encryptedData.length - AES_AUTH_CODE_LENGTH
    ),
    storedHmac: encryptedData.subarray(encryptedData.length - AES_AUTH_CODE_LENGTH)
  };
}

/**
 * Verify password verification bytes.
 */
function verifyPasswordBytes(derived: Uint8Array, stored: Uint8Array): boolean {
  return derived[0] === stored[0] && derived[1] === stored[1];
}

/**
 * Constant-time comparison of two Uint8Arrays.
 * Prevents timing attacks by always comparing all bytes.
 */
function constantTimeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a[i]! ^ b[i]!;
  }
  return diff === 0;
}

/**
 * Decrypt AES-encrypted ZIP entry data.
 *
 * Structure of encrypted data:
 * - Salt: saltLen bytes (8/12/16 for AES-128/192/256)
 * - Password verification: 2 bytes
 * - Encrypted data: variable
 * - Authentication code: 10 bytes
 *
 * @param encryptedData - Full encrypted data
 * @param password - Password string or bytes
 * @param keyStrength - AES key strength (128, 192, or 256)
 * @returns Decrypted data or null if authentication fails
 */
export async function aesDecrypt(
  encryptedData: Uint8Array,
  password: string | Uint8Array,
  keyStrength: AesKeyStrength
): Promise<Uint8Array> {
  checkBytes("Archive.Crypto.decryptAes", "encryptedData", encryptedData);
  checkZipPassword("Archive.Crypto.decryptAes", password);
  checkKeyStrength("Archive.Crypto.decryptAes", keyStrength);
  const components = extractAesComponents(encryptedData, keyStrength);
  if (!components) {
    throw new ArchiveError("Encrypted data too short");
  }

  const { salt, storedVerify, ciphertext, storedHmac } = components;

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Verify password
  if (!verifyPasswordBytes(keys.passwordVerify, storedVerify)) {
    throw new ArchiveError("Password verification failed");
  }

  // Verify HMAC (constant-time comparison to prevent timing attacks)
  const computedHmac = await aesComputeHmac(keys.hmacKey, ciphertext);
  if (!constantTimeEqual(computedHmac, storedHmac)) {
    throw new ArchiveError("HMAC verification failed");
  }

  // Decrypt data
  return aesCtr(keys.encryptionKey, ciphertext, false);
}

/**
 * Encrypt data using AES for ZIP.
 *
 * @param data - Plain data to encrypt
 * @param password - Password string or bytes
 * @param keyStrength - AES key strength (128, 192, or 256)
 * @returns Encrypted data with salt, verification, ciphertext, and HMAC
 */
export async function aesEncrypt(
  data: Uint8Array,
  password: string | Uint8Array,
  keyStrength: AesKeyStrength
): Promise<Uint8Array> {
  checkBytes("Archive.Crypto.encryptAes", "data", data);
  checkZipPassword("Archive.Crypto.encryptAes", password);
  checkKeyStrength("Archive.Crypto.encryptAes", keyStrength);
  const saltLen = AES_SALT_LENGTH[keyStrength];

  // Generate random salt
  const salt = randomBytes(saltLen);

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Encrypt data
  const ciphertext = await aesCtr(keys.encryptionKey, data, true);

  // Compute HMAC
  const hmac = await aesComputeHmac(keys.hmacKey, ciphertext);

  // Build output: salt + verify + ciphertext + hmac
  const output = new Uint8Array(
    saltLen + AES_PASSWORD_VERIFY_LENGTH + ciphertext.length + AES_AUTH_CODE_LENGTH
  );

  let offset = 0;
  output.set(salt, offset);
  offset += saltLen;

  output.set(keys.passwordVerify, offset);
  offset += AES_PASSWORD_VERIFY_LENGTH;

  output.set(ciphertext, offset);
  offset += ciphertext.length;

  output.set(hmac, offset);

  return output;
}

/**
 * Check if a password is valid for AES-encrypted data without full decryption.
 * This is useful for quick password validation without decompressing data.
 *
 * @param encryptedData - Full encrypted data
 * @param password - Password string or bytes
 * @param keyStrength - AES key strength (128, 192, or 256)
 * @returns true if password verification passes, false otherwise
 */
export async function aesCheckPasswordOnly(
  encryptedData: Uint8Array,
  password: string | Uint8Array,
  keyStrength: AesKeyStrength
): Promise<boolean> {
  const saltLen = AES_SALT_LENGTH[keyStrength];
  const minLen = saltLen + AES_PASSWORD_VERIFY_LENGTH;

  if (encryptedData.length < minLen) {
    return false;
  }

  // Extract only salt and verification (don't need full components)
  const salt = encryptedData.subarray(0, saltLen);
  const storedVerify = encryptedData.subarray(saltLen, saltLen + AES_PASSWORD_VERIFY_LENGTH);

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Check password verification bytes
  return verifyPasswordBytes(keys.passwordVerify, storedVerify);
}

/**
 * Alias for aesCheckPasswordOnly for API consistency with zipCryptoVerifyPassword.
 */
export { aesCheckPasswordOnly as aesVerifyPassword };

/**
 * Verify the HMAC signature of AES-encrypted data without decryption.
 * This is useful for integrity verification without decompressing data.
 *
 * @param encryptedData - Full encrypted data
 * @param password - Password string or bytes
 * @param keyStrength - AES key strength (128, 192, or 256)
 * @returns true if HMAC verification passes, false otherwise
 * @throws Error if password verification fails
 */
export async function aesCheckSignature(
  encryptedData: Uint8Array,
  password: string | Uint8Array,
  keyStrength: AesKeyStrength
): Promise<boolean> {
  const components = extractAesComponents(encryptedData, keyStrength);
  if (!components) {
    throw new ArchiveError("Encrypted data too short");
  }

  const { salt, storedVerify, ciphertext, storedHmac } = components;

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Verify password first
  if (!verifyPasswordBytes(keys.passwordVerify, storedVerify)) {
    throw new ArchiveError("Password verification failed");
  }

  // Compute and compare HMAC (constant-time)
  const computedHmac = await aesComputeHmac(keys.hmacKey, ciphertext);
  return constantTimeEqual(computedHmac, storedHmac);
}
