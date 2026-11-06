/**
 * The WinZip AES extra field (0x9901): its constants, and building and reading it.
 *
 * Split from `aes.ts`, which holds the cipher, because every ZIP writer and reader needs the field — to recognise an
 * encrypted entry, or to say that one is not — while only an encrypted entry needs PBKDF2, HMAC and AES-CTR. A bundler
 * splits chunks by module, so while the two shared one, a writer that encrypts nothing (every Excel and Word package)
 * loaded the cipher with its first chunk.
 */

import { checkArchiveArgument } from "@archive/core/arguments";
import { checkUint } from "@archive/core/entry-checks";
import { s } from "@utils/schema";

/**
 * AES key strength options.
 */
export type AesKeyStrength = 128 | 192 | 256;

/** Refuse a key strength the AES extra field cannot name, before any table lookup. */
export function checkKeyStrength(where: string, keyStrength: unknown, field = "keyStrength"): void {
  checkArchiveArgument(
    s.oneOf([128, 192, 256], { expected: "128, 192 or 256" }),
    keyStrength,
    where,
    field
  );
}

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
  checkArchiveArgument(
    s.oneOf([1, 2], { expected: "1 or 2" }),
    version,
    "Archive.Crypto.createAesExtraField",
    "version"
  );
  checkKeyStrength("Archive.Crypto.createAesExtraField", keyStrength);
  checkUint("Archive.Crypto.createAesExtraField", "compressionMethod", compressionMethod, 0xffff);
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
 * Calculate the total encrypted size for a given plaintext size.
 */
export function aesEncryptedSize(plaintextSize: number, keyStrength: AesKeyStrength): number {
  checkUint(
    "Archive.Crypto.getAesEncryptedSize",
    "plaintextSize",
    plaintextSize,
    Number.MAX_SAFE_INTEGER
  );
  checkKeyStrength("Archive.Crypto.getAesEncryptedSize", keyStrength);
  const saltLen = AES_SALT_LENGTH[keyStrength];
  return saltLen + AES_PASSWORD_VERIFY_LENGTH + plaintextSize + AES_AUTH_CODE_LENGTH;
}
