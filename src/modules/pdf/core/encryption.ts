/**
 * PDF encryption support (Standard Security Handler, V=5, R=5).
 *
 * Implements AES-256 encryption compatible with PDF 2.0 (ISO 32000-2:2020).
 * Supports:
 * - User password (required to open the document)
 * - Owner password (grants full access)
 * - Permission flags (print, copy, modify, etc.)
 *
 * The file encryption key (FEK) is a random 256-bit key.
 * All streams and strings are encrypted using AES-256-CBC with a random
 * 16-byte IV prepended to each encrypted value.
 *
 * @see ISO 32000-2:2020, §7.6 — Encryption
 */

import { pdfSha256 } from "@pdf/core/pdf-kdf";
import { concatUint8Arrays } from "@utils/binary";
import { aesCbcEncrypt, aesCbcEncryptRaw, aesEcbEncrypt, randomBytes } from "@utils/crypto";

// =============================================================================
// Types
// =============================================================================

/**
 * PDF encryption options.
 */
export interface PdfEncryptionOptions {
  /** User password (required to open the document). Empty string = no open password. */
  userPassword?: string;

  /** Owner password (grants full permissions). Required. */
  ownerPassword: string;

  /** Permissions to grant when opened with user password. */
  permissions?: Partial<PdfPermissions>;
}

/**
 * PDF document permissions (what a user-password holder can do).
 * All default to false.
 */
export interface PdfPermissions {
  /** Allow printing */
  print: boolean;
  /** Allow modifying content */
  modify: boolean;
  /** Allow copying text/images */
  copy: boolean;
  /** Allow adding/modifying annotations */
  annotate: boolean;
  /** Allow filling form fields */
  fillForms: boolean;
  /** Allow extracting content for accessibility */
  accessibility: boolean;
  /** Allow assembling (insert/rotate/delete pages) */
  assemble: boolean;
  /** Allow high-quality printing */
  printHighQuality: boolean;
}

/**
 * Encryption state used during PDF generation (V=5, R=5, AES-256).
 */
export interface EncryptionState {
  /** 32-byte file encryption key */
  encryptionKey: Uint8Array;
  /** 48-byte O value: hash(32) + validation salt(8) + key salt(8) */
  oValue: Uint8Array;
  /** 48-byte U value: hash(32) + validation salt(8) + key salt(8) */
  uValue: Uint8Array;
  /** 32-byte encrypted owner key (OE) */
  oeValue: Uint8Array;
  /** 32-byte encrypted user key (UE) */
  ueValue: Uint8Array;
  /** 16-byte encrypted permissions (Perms) */
  permsValue: Uint8Array;
  /** Permissions integer (P value) */
  permissions: number;
  /** File identifier (16 bytes) */
  fileId: Uint8Array;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Initialize encryption state for AES-256 (V=5, R=5).
 */
export function initEncryption(options: PdfEncryptionOptions): EncryptionState {
  const userPwd = truncatePassword(options.userPassword ?? "");
  const ownerPwd = truncatePassword(options.ownerPassword);
  const perms = computePermissions(options.permissions);

  // Step 1: Generate random 32-byte file encryption key
  const encryptionKey = randomBytes(32);

  // Step 2: Generate random salts
  const uValidationSalt = randomBytes(8);
  const uKeySalt = randomBytes(8);
  const oValidationSalt = randomBytes(8);
  const oKeySalt = randomBytes(8);

  // Step 3: Compute U value
  // U hash = SHA-256(userPassword + uValidationSalt)
  const uHash = pdfSha256(concatUint8Arrays([userPwd, uValidationSalt]));
  const uValue = concatUint8Arrays([uHash, uValidationSalt, uKeySalt]);

  // Step 4: Compute UE value
  // UE = AES-256-CBC-encrypt(encryptionKey, SHA-256(userPassword + uKeySalt), zeroIV)
  // Actually: the key for encrypting UE is SHA-256(password + key_salt),
  // and we encrypt the file encryption key with it.
  const ueKey = pdfSha256(concatUint8Arrays([userPwd, uKeySalt]));
  const zeroIv = new Uint8Array(16);
  const ueValue = aesCbcEncryptRaw(encryptionKey, ueKey, zeroIv);

  // Step 5: Compute O value
  // O hash = SHA-256(ownerPassword + oValidationSalt + U(0..47))
  const oHash = pdfSha256(concatUint8Arrays([ownerPwd, oValidationSalt, uValue]));
  const oValue = concatUint8Arrays([oHash, oValidationSalt, oKeySalt]);

  // Step 6: Compute OE value
  // OE = AES-256-CBC-encrypt(encryptionKey, SHA-256(ownerPassword + oKeySalt + U(0..47)), zeroIV)
  const oeKey = pdfSha256(concatUint8Arrays([ownerPwd, oKeySalt, uValue]));
  const oeValue = aesCbcEncryptRaw(encryptionKey, oeKey, zeroIv);

  // Step 7: Compute Perms value
  // 16-byte block: P(4 LE bytes) + 0xFF(4 bytes) + 'T' or 'F' (encryptMetadata) + 'a' 'd' 'b' + 0(3 bytes)
  const permsBlock = new Uint8Array(16);
  const permsView = new DataView(permsBlock.buffer);
  permsView.setInt32(0, perms, true); // P value in little-endian
  permsBlock[4] = 0xff;
  permsBlock[5] = 0xff;
  permsBlock[6] = 0xff;
  permsBlock[7] = 0xff;
  permsBlock[8] = 0x54; // 'T' — EncryptMetadata = true
  permsBlock[9] = 0x61; // 'a'
  permsBlock[10] = 0x64; // 'd'
  permsBlock[11] = 0x62; // 'b'
  // bytes 12-15 are zero

  const permsValue = aesEcbEncrypt(permsBlock, encryptionKey);

  // File ID (random 16 bytes, used in trailer)
  const fileId = randomBytes(16);

  return {
    encryptionKey,
    oValue,
    uValue,
    oeValue,
    ueValue,
    permsValue,
    permissions: perms,
    fileId
  };
}

/**
 * Encrypt data for a PDF object using AES-256-CBC.
 *
 * For V=5/R=5, the file encryption key is used directly (no per-object key derivation).
 * A random 16-byte IV is prepended to the ciphertext.
 */
export function encryptData(
  data: Uint8Array,
  _objectNumber: number,
  _generation: number,
  encryptionKey: Uint8Array
): Uint8Array {
  const iv = randomBytes(16);
  const ciphertext = aesCbcEncrypt(data, encryptionKey, iv);

  // Prepend IV to ciphertext per PDF spec
  const result = new Uint8Array(16 + ciphertext.length);
  result.set(iv);
  result.set(ciphertext, 16);
  return result;
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Truncate password to 127 bytes (UTF-8) per PDF 2.0 spec.
 */
function truncatePassword(password: string): Uint8Array {
  const bytes = new TextEncoder().encode(password);
  return bytes.length > 127 ? bytes.subarray(0, 127) : bytes;
}

/**
 * Compute the permissions integer (P value) from permission flags.
 */
function computePermissions(perms?: Partial<PdfPermissions>): number {
  // Start with all reserved bits set to 1
  let p = 0xfffff000 | 0b11000000;

  if (perms?.print) {
    p |= 1 << 2;
  }
  if (perms?.modify) {
    p |= 1 << 3;
  }
  if (perms?.copy) {
    p |= 1 << 4;
  }
  if (perms?.annotate) {
    p |= 1 << 5;
  }
  if (perms?.fillForms) {
    p |= 1 << 8;
  }
  if (perms?.accessibility) {
    p |= 1 << 9;
  }
  if (perms?.assemble) {
    p |= 1 << 10;
  }
  if (perms?.printHighQuality) {
    p |= 1 << 11;
  }

  return p | 0;
}
