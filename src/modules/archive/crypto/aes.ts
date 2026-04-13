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

import { stringToUint8Array as encodeUtf8 } from "@utils/binary";
import { randomBytes } from "@utils/crypto";
import { toArrayBuffer } from "@archive/shared/text";

/**
 * AES key strength options.
 */
export type AesKeyStrength = 128 | 192 | 256;

/**
 * AES vendor ID for WinZip format.
 */
export const AES_VENDOR_ID = 0x4541; // "AE" in little-endian

/**
 * AES vendor version for AE-2 (CRC-32 not used).
 */
export const AES_VERSION_AE2 = 0x0002;

/**
 * AES vendor version for AE-1 (CRC-32 used).
 */
export const AES_VERSION_AE1 = 0x0001;

/**
 * Compression method value indicating AES encryption.
 */
export const COMPRESSION_METHOD_AES = 99;

/**
 * AES extra field header ID.
 */
export const AES_EXTRA_FIELD_ID = 0x9901;

/**
 * Salt lengths for different key strengths.
 */
export const AES_SALT_LENGTH: Record<AesKeyStrength, number> = {
  128: 8,
  192: 12,
  256: 16
};

/**
 * Key lengths for different strengths.
 */
export const AES_KEY_LENGTH: Record<AesKeyStrength, number> = {
  128: 16,
  192: 24,
  256: 32
};

/**
 * Key strength byte values used in AES extra field.
 */
export const AES_STRENGTH_BYTE: Record<AesKeyStrength, number> = {
  128: 1,
  192: 2,
  256: 3
};

/**
 * Reverse mapping from strength byte to key strength.
 */
export const AES_STRENGTH_FROM_BYTE: Record<number, AesKeyStrength> = {
  1: 128,
  2: 192,
  3: 256
};

/**
 * HMAC-SHA1 authentication code length (truncated to 10 bytes).
 */
export const AES_AUTH_CODE_LENGTH = 10;

/**
 * Password verification value length.
 */
export const AES_PASSWORD_VERIFY_LENGTH = 2;

/**
 * AES encryption info parsed from extra field.
 */
export interface AesExtraFieldInfo {
  /** AE format version (1 or 2) */
  version: number;
  /** Vendor ID (should be 0x4541 "AE") */
  vendorId: number;
  /** Key strength (128, 192, or 256) */
  keyStrength: AesKeyStrength;
  /** Original compression method */
  compressionMethod: number;
}

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
  throw new Error("Web Crypto API not available");
}

/**
 * Derive AES keys from password using PBKDF2.
 *
 * The derived key material is split into:
 * - Encryption key (16/24/32 bytes)
 * - HMAC key (32 bytes)
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

  // Total derived bytes: encryption key + HMAC key (32 bytes) + verification (2 bytes)
  const derivedLen = keyLen + 32 + 2;

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
    hmacKey: derived.subarray(keyLen, keyLen + 32),
    passwordVerify: derived.subarray(keyLen + 32, keyLen + 34)
  };
}

/**
 * Compute HMAC-SHA1 authentication code.
 *
 * @param key - HMAC key (32 bytes)
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
 * WinZip uses little-endian counter starting at 1.
 */
function createCtrCounter(counter: number): Uint8Array {
  const block = new Uint8Array(16);
  const view = new DataView(block.buffer);
  view.setUint32(0, counter, true); // Little-endian counter
  return block;
}

/**
 * AES-CTR encryption/decryption.
 *
 * WinZip uses a custom CTR mode with:
 * - 16-byte blocks
 * - Little-endian counter starting at 1
 * - Counter in the low bytes of the IV
 *
 * Note: AES-CTR is symmetric - encrypt and decrypt are the same operation.
 */
export async function aesCtr(
  key: Uint8Array,
  data: Uint8Array,
  _encrypt: boolean = true
): Promise<Uint8Array> {
  const crypto = getWebCrypto();

  // Import AES key
  const aesKey = await crypto.importKey("raw", toArrayBuffer(key), { name: "AES-CTR" }, false, [
    "encrypt",
    "decrypt"
  ]);

  // WinZip AES uses counter starting at 1, not 0
  const counter = createCtrCounter(1);

  // AES-CTR is symmetric, so we can always use encrypt
  const result = await crypto.encrypt(
    { name: "AES-CTR", counter: toArrayBuffer(counter), length: 128 },
    aesKey,
    toArrayBuffer(data)
  );
  return new Uint8Array(result);
}

/**
 * Parse AES extra field data.
 *
 * Structure (7 bytes):
 * - Bytes 0-1: Version (0x0001 or 0x0002)
 * - Bytes 2-3: Vendor ID ("AE" = 0x4541)
 * - Byte 4: Key strength (1=128, 2=192, 3=256)
 * - Bytes 5-6: Actual compression method
 */
export function parseAesExtraField(data: Uint8Array): AesExtraFieldInfo | null {
  if (data.length < 7) {
    return null;
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const version = view.getUint16(0, true);
  const vendorId = view.getUint16(2, true);
  const strengthByte = data[4]!;
  const compressionMethod = view.getUint16(5, true);

  if (vendorId !== AES_VENDOR_ID) {
    return null;
  }

  const keyStrength = AES_STRENGTH_FROM_BYTE[strengthByte];
  if (!keyStrength) {
    return null;
  }

  return {
    version,
    vendorId,
    keyStrength,
    compressionMethod
  };
}

/**
 * Build AES extra field data.
 */
export function buildAesExtraField(
  version: 1 | 2,
  keyStrength: AesKeyStrength,
  compressionMethod: number
): Uint8Array {
  const data = new Uint8Array(11); // 4 bytes header + 7 bytes data
  const view = new DataView(data.buffer);

  // Extra field header
  view.setUint16(0, AES_EXTRA_FIELD_ID, true);
  view.setUint16(2, 7, true); // Data size

  // AES extra field data
  view.setUint16(4, version, true); // AE version
  view.setUint16(6, AES_VENDOR_ID, true); // "AE"
  data[8] = AES_STRENGTH_BYTE[keyStrength]; // Key strength
  view.setUint16(9, compressionMethod, true); // Original compression method

  return data;
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
  const components = extractAesComponents(encryptedData, keyStrength);
  if (!components) {
    throw new Error("Encrypted data too short");
  }

  const { salt, storedVerify, ciphertext, storedHmac } = components;

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Verify password
  if (!verifyPasswordBytes(keys.passwordVerify, storedVerify)) {
    throw new Error("Password verification failed");
  }

  // Verify HMAC (constant-time comparison to prevent timing attacks)
  const computedHmac = await aesComputeHmac(keys.hmacKey, ciphertext);
  if (!constantTimeEqual(computedHmac, storedHmac)) {
    throw new Error("HMAC verification failed");
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
 * Calculate the total encrypted size for a given plaintext size.
 */
export function aesEncryptedSize(plaintextSize: number, keyStrength: AesKeyStrength): number {
  const saltLen = AES_SALT_LENGTH[keyStrength];
  return saltLen + AES_PASSWORD_VERIFY_LENGTH + plaintextSize + AES_AUTH_CODE_LENGTH;
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
    throw new Error("Encrypted data too short");
  }

  const { salt, storedVerify, ciphertext, storedHmac } = components;

  // Derive keys
  const keys = await aesDerive(password, salt, keyStrength);

  // Verify password first
  if (!verifyPasswordBytes(keys.passwordVerify, storedVerify)) {
    throw new Error("Password verification failed");
  }

  // Compute and compare HMAC (constant-time)
  const computedHmac = await aesComputeHmac(keys.hmacKey, ciphertext);
  return constantTimeEqual(computedHmac, storedHmac);
}
