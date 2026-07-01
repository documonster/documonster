/**
 * Encryption-related types and utilities for ZIP files.
 */

import type { AesKeyStrength } from "@archive/crypto/aes";

/**
 * Encryption method used for a ZIP entry.
 */
export type ZipEncryptionMethod = "none" | "zipcrypto" | "aes-128" | "aes-192" | "aes-256";

/**
 * Encryption metadata for a ZIP entry.
 */
export interface ZipEncryptionInfo {
  /** Encryption method */
  method: ZipEncryptionMethod;

  /** For AES: the original compression method before encryption */
  originalCompressionMethod?: number;

  /** For AES: the AE version (1 or 2) */
  aesVersion?: 1 | 2;

  /** For AES: the key strength */
  aesKeyStrength?: AesKeyStrength;
}

/**
 * Password options for ZIP operations.
 */
export interface ZipPasswordOptions {
  /** Password as string or bytes */
  password?: string | Uint8Array;
}

/**
 * Encryption options for ZIP creation.
 */
export interface ZipEncryptionOptions extends ZipPasswordOptions {
  /** Encryption method to use (default: none) */
  encryptionMethod?: ZipEncryptionMethod;

  /** For AES: use AE-1 (with CRC) or AE-2 (without CRC). Default: AE-2 */
  aesVersion?: 1 | 2;
}

/** Mapping from encryption method to display name */
const ENCRYPTION_METHOD_NAMES: Record<ZipEncryptionMethod, string> = {
  none: "None",
  zipcrypto: "ZipCrypto (Traditional PKWARE)",
  "aes-128": "AES-128",
  "aes-192": "AES-192",
  "aes-256": "AES-256"
};

/** Mapping from encryption method to AES key strength */
const AES_KEY_STRENGTHS: Partial<Record<ZipEncryptionMethod, AesKeyStrength>> = {
  "aes-128": 128,
  "aes-192": 192,
  "aes-256": 256
};

/** Mapping from AES key strength to encryption method */
const ENCRYPTION_METHOD_FROM_STRENGTH: Record<AesKeyStrength, ZipEncryptionMethod> = {
  128: "aes-128",
  192: "aes-192",
  256: "aes-256"
};

/**
 * Get encryption method display name.
 */
export function getEncryptionMethodName(method: ZipEncryptionMethod): string {
  return ENCRYPTION_METHOD_NAMES[method];
}

/**
 * Check if encryption method is AES-based.
 */
export function isAesEncryption(method: ZipEncryptionMethod): boolean {
  return method in AES_KEY_STRENGTHS;
}

/**
 * Get AES key strength from encryption method.
 */
export function getAesKeyStrength(method: ZipEncryptionMethod): AesKeyStrength | undefined {
  return AES_KEY_STRENGTHS[method];
}

/**
 * Get encryption method from AES key strength.
 */
export function encryptionMethodFromAesKeyStrength(strength: AesKeyStrength): ZipEncryptionMethod {
  return ENCRYPTION_METHOD_FROM_STRENGTH[strength];
}
