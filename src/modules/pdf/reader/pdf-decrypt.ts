/**
 * PDF decryption for reading encrypted PDFs.
 *
 * Supports:
 * - Standard Security Handler (V1/V2/V4/V5, R2/R3/R4/R5)
 * - RC4 encryption (40-bit and 128-bit)
 * - AES-128 encryption (PDF 1.6+)
 * - AES-256 encryption (PDF 2.0, V=5, R=5)
 *
 * Reuses the existing RC4 and MD5 implementations from the writer module.
 *
 * @see PDF Reference 1.7, §3.5 - Encryption
 * @see PDF 2.0 (ISO 32000-2), §7.6 - Encryption
 */

import { rc4, md5, encryptData } from "../core/encryption";
import type { PdfDictValue } from "./pdf-parser";
import { dictGetNumber, dictGetName, dictGetBytes, dictGetArray, dictGetBool } from "./pdf-parser";
import type { PdfDocument } from "./pdf-document";
import { PdfStructureError } from "../errors";

// =============================================================================
// Constants
// =============================================================================

/** PDF password padding string (32 bytes) per PDF spec §3.5.2 */
const PASSWORD_PADDING = new Uint8Array([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a
]);

/** Cached TextEncoder instance */
const textEncoder = new TextEncoder();

// =============================================================================
// Public API
// =============================================================================

/**
 * Initialize decryption for a PDF document.
 * Returns true if decryption was successfully initialized, false if
 * the password was incorrect.
 *
 * @param doc - The PDF document
 * @param password - User or owner password (empty string for no password)
 */
export function initDecryption(doc: PdfDocument, password = ""): boolean {
  const encryptDict = doc.derefDict(doc.trailer.get("Encrypt"));
  if (!encryptDict) {
    return true; // Not encrypted
  }

  const filter = dictGetName(encryptDict, "Filter");
  if (filter !== "Standard") {
    throw new PdfStructureError(`Unsupported encryption filter: ${filter}`);
  }

  const v = dictGetNumber(encryptDict, "V") ?? 0;
  const r = dictGetNumber(encryptDict, "R") ?? 0;
  const keyLength = (dictGetNumber(encryptDict, "Length") ?? 40) / 8; // bits → bytes
  const permissions = dictGetNumber(encryptDict, "P") ?? 0;
  const oValue = dictGetBytes(encryptDict, "O");
  const uValue = dictGetBytes(encryptDict, "U");

  if (!oValue || !uValue) {
    throw new PdfStructureError("Missing /O or /U values in Encrypt dictionary");
  }

  // Get file ID from trailer
  const idArray = dictGetArray(doc.trailer, "ID");
  const fileId =
    idArray && idArray.length > 0 && idArray[0] instanceof Uint8Array
      ? idArray[0]
      : new Uint8Array(0);

  // Determine EncryptMetadata flag (default true per spec)
  const encryptMetadata = readEncryptMetadata(encryptDict);

  // Handle V=5 (AES-256, PDF 2.0)
  if (v === 5) {
    return initDecryptionV5(doc, encryptDict, password, r, oValue, uValue, permissions, fileId);
  }

  // Determine if we should use AES
  const useAes = v === 4 && isAesCryptFilter(encryptDict);

  // Try user password first, then owner password
  let encryptionKey = tryUserPassword(
    password,
    oValue,
    permissions,
    fileId,
    r,
    keyLength,
    uValue,
    encryptMetadata
  );

  if (!encryptionKey) {
    // Try as owner password
    const derivedUser = deriveUserPasswordFromOwner(password, oValue, r, keyLength);
    encryptionKey = tryUserPassword(
      derivedUser,
      oValue,
      permissions,
      fileId,
      r,
      keyLength,
      uValue,
      encryptMetadata
    );
  }

  if (!encryptionKey) {
    // Try empty password
    if (password !== "") {
      encryptionKey = tryUserPassword(
        "",
        oValue,
        permissions,
        fileId,
        r,
        keyLength,
        uValue,
        encryptMetadata
      );
    }
  }

  if (!encryptionKey) {
    return false; // Password incorrect
  }

  // Set up decryption function
  const finalKey = encryptionKey;
  if (useAes) {
    doc.decryptFn = (data, objNum, gen) => decryptAes128(data, objNum, gen, finalKey);
  } else {
    doc.decryptFn = (data, objNum, gen) => encryptData(data, objNum, gen, finalKey);
  }

  return true;
}

/**
 * Check if the document is encrypted.
 */
export function isEncrypted(doc: PdfDocument): boolean {
  return doc.trailer.has("Encrypt");
}

// =============================================================================
// V5 (AES-256) Decryption
// =============================================================================

/**
 * Initialize decryption for V=5 (AES-256, PDF 2.0).
 * Supports R=5 using SHA-256 based key derivation (Algorithm 2.A).
 */
function initDecryptionV5(
  doc: PdfDocument,
  encryptDict: PdfDictValue,
  password: string,
  revision: number,
  oValue: Uint8Array,
  uValue: Uint8Array,
  _permissions: number,
  _fileId: Uint8Array
): boolean {
  if (revision === 6) {
    throw new PdfStructureError(
      "R=6 (PDF 2.0 extension) requires SHA-384/SHA-512 which is not yet supported"
    );
  }

  if (revision !== 5) {
    throw new PdfStructureError(`Unsupported revision ${revision} for V=5 encryption`);
  }

  const oeValue = dictGetBytes(encryptDict, "OE");
  const ueValue = dictGetBytes(encryptDict, "UE");

  if (!oeValue || !ueValue) {
    throw new PdfStructureError("Missing /OE or /UE values in V=5 Encrypt dictionary");
  }

  // O value layout: 32 bytes hash + 8 bytes validation salt + 8 bytes key salt
  // U value layout: 32 bytes hash + 8 bytes validation salt + 8 bytes key salt
  if (oValue.length < 48 || uValue.length < 48) {
    throw new PdfStructureError("Invalid /O or /U length for V=5 encryption");
  }

  const passwordBytes = truncatePassword(password);

  // Try user password (Algorithm 2.A step a - user)
  let encryptionKey = tryUserPasswordV5(passwordBytes, uValue, ueValue);

  if (!encryptionKey) {
    // Try owner password (Algorithm 2.A step a - owner)
    encryptionKey = tryOwnerPasswordV5(passwordBytes, oValue, oeValue, uValue);
  }

  if (!encryptionKey) {
    // Try empty password
    if (password !== "") {
      const emptyBytes = new Uint8Array(0);
      encryptionKey = tryUserPasswordV5(emptyBytes, uValue, ueValue);
      if (!encryptionKey) {
        encryptionKey = tryOwnerPasswordV5(emptyBytes, oValue, oeValue, uValue);
      }
    }
  }

  if (!encryptionKey) {
    return false;
  }

  // V=5 always uses AES-256 with the file encryption key directly (no per-object key derivation)
  const finalKey = encryptionKey;
  doc.decryptFn = (data, _objNum, _gen) => decryptAes256Direct(data, finalKey);

  return true;
}

/**
 * Truncate password to 127 bytes (UTF-8) per PDF 2.0 spec.
 */
function truncatePassword(password: string): Uint8Array {
  const bytes = textEncoder.encode(password);
  return bytes.length > 127 ? bytes.subarray(0, 127) : bytes;
}

/**
 * Try user password for V=5/R=5.
 * Validates using SHA-256(password + validation salt from U).
 * If valid, derives file encryption key using SHA-256(password + key salt from U).
 */
function tryUserPasswordV5(
  passwordBytes: Uint8Array,
  uValue: Uint8Array,
  ueValue: Uint8Array
): Uint8Array | null {
  // U = hash(32) + validation salt(8) + key salt(8)
  const uHash = uValue.subarray(0, 32);
  const uValidationSalt = uValue.subarray(32, 40);
  const uKeySalt = uValue.subarray(40, 48);

  // Validate: SHA-256(password + validation salt) == first 32 bytes of U
  const validateInput = concatArrays(passwordBytes, uValidationSalt);
  const computedHash = sha256(validateInput);

  if (!arraysEqual(computedHash, uHash)) {
    return null;
  }

  // Derive key: SHA-256(password + key salt) => use as AES-256 key to decrypt UE
  const keyInput = concatArrays(passwordBytes, uKeySalt);
  const keyHash = sha256(keyInput);

  // Decrypt UE with this key using AES-256-CBC with zero IV
  const zeroIv = new Uint8Array(16);
  return aes256CbcDecryptRaw(ueValue.subarray(0, 32), keyHash, zeroIv);
}

/**
 * Try owner password for V=5/R=5.
 * Validates using SHA-256(password + validation salt from O + U(48)).
 * If valid, derives file encryption key using SHA-256(password + key salt from O + U(48)).
 */
function tryOwnerPasswordV5(
  passwordBytes: Uint8Array,
  oValue: Uint8Array,
  oeValue: Uint8Array,
  uValue: Uint8Array
): Uint8Array | null {
  // O = hash(32) + validation salt(8) + key salt(8)
  const oHash = oValue.subarray(0, 32);
  const oValidationSalt = oValue.subarray(32, 40);
  const oKeySalt = oValue.subarray(40, 48);
  const u48 = uValue.subarray(0, 48);

  // Validate: SHA-256(password + validation salt + U(0..47)) == first 32 bytes of O
  const validateInput = concatArrays(passwordBytes, oValidationSalt, u48);
  const computedHash = sha256(validateInput);

  if (!arraysEqual(computedHash, oHash)) {
    return null;
  }

  // Derive key: SHA-256(password + key salt + U(0..47))
  const keyInput = concatArrays(passwordBytes, oKeySalt, u48);
  const keyHash = sha256(keyInput);

  // Decrypt OE with this key using AES-256-CBC with zero IV
  const zeroIv = new Uint8Array(16);
  return aes256CbcDecryptRaw(oeValue.subarray(0, 32), keyHash, zeroIv);
}

/**
 * Decrypt data using AES-256 directly (no per-object key derivation).
 * For V=5, the file encryption key is used directly. The first 16 bytes are IV.
 */
function decryptAes256Direct(data: Uint8Array, encryptionKey: Uint8Array): Uint8Array {
  if (data.length < 16) {
    return data;
  }

  const iv = data.subarray(0, 16);
  const ciphertext = data.subarray(16);

  if (ciphertext.length === 0 || ciphertext.length % 16 !== 0) {
    return data;
  }

  return aesCbcDecrypt(ciphertext, encryptionKey, iv);
}

// =============================================================================
// Password Verification (V1-V4)
// =============================================================================

/**
 * Try to authenticate with the user password.
 * Returns the encryption key if successful, null otherwise.
 */
function tryUserPassword(
  password: string,
  oValue: Uint8Array,
  permissions: number,
  fileId: Uint8Array,
  revision: number,
  keyLength: number,
  uValue: Uint8Array,
  encryptMetadata: boolean
): Uint8Array | null {
  const key = computeEncryptionKeyForReading(
    password,
    oValue,
    permissions,
    fileId,
    revision,
    keyLength,
    encryptMetadata
  );

  // Verify against U value
  if (revision === 2) {
    // R2: encrypt password padding with key, compare to U
    const encrypted = rc4(key, PASSWORD_PADDING);
    if (arraysEqual(encrypted, uValue.subarray(0, 32))) {
      return key;
    }
  } else if (revision >= 3) {
    // R3/R4: MD5(padding + fileId), encrypt, iterate 19 times
    const hashInput = new Uint8Array(32 + fileId.length);
    hashInput.set(PASSWORD_PADDING);
    hashInput.set(fileId, 32);
    const hash = md5(hashInput);
    let result = rc4(key, hash);

    for (let i = 1; i <= 19; i++) {
      const modKey = new Uint8Array(key.length);
      for (let j = 0; j < key.length; j++) {
        modKey[j] = key[j] ^ i;
      }
      result = rc4(modKey, result);
    }

    // Compare first 16 bytes
    if (arraysEqual(result.subarray(0, 16), uValue.subarray(0, 16))) {
      return key;
    }
  }

  return null;
}

/**
 * Compute the encryption key for reading (Algorithm 2, PDF spec §3.5.2).
 */
function computeEncryptionKeyForReading(
  password: string,
  oValue: Uint8Array,
  permissions: number,
  fileId: Uint8Array,
  revision: number,
  keyLength: number,
  encryptMetadata: boolean
): Uint8Array {
  const paddedPwd = padPassword(password);
  // When encryptMetadata is false and revision >= 4, append 4 bytes of 0xFF
  const extraBytes = revision >= 4 && !encryptMetadata ? 4 : 0;
  const input = new Uint8Array(32 + 32 + 4 + fileId.length + extraBytes);
  let offset = 0;

  input.set(paddedPwd, offset);
  offset += 32;
  input.set(oValue.subarray(0, 32), offset);
  offset += 32;

  // P value as 4 LE bytes
  input[offset] = permissions & 0xff;
  input[offset + 1] = (permissions >> 8) & 0xff;
  input[offset + 2] = (permissions >> 16) & 0xff;
  input[offset + 3] = (permissions >> 24) & 0xff;
  offset += 4;

  input.set(fileId, offset);
  offset += fileId.length;

  // If EncryptMetadata is false and revision >= 4, append 0xFFFFFFFF
  if (revision >= 4 && !encryptMetadata) {
    input[offset] = 0xff;
    input[offset + 1] = 0xff;
    input[offset + 2] = 0xff;
    input[offset + 3] = 0xff;
    offset += 4;
  }

  let hash = md5(input.subarray(0, offset));

  // For revision >= 3, hash 50 more times
  if (revision >= 3) {
    for (let i = 0; i < 50; i++) {
      hash = md5(hash.subarray(0, keyLength));
    }
  }

  return hash.subarray(0, keyLength);
}

/**
 * Derive the user password from the owner password.
 * Uses Algorithm 7 from PDF spec §3.5.2.
 */
function deriveUserPasswordFromOwner(
  ownerPassword: string,
  oValue: Uint8Array,
  revision: number,
  keyLength: number
): string {
  let hash = md5(padPassword(ownerPassword));

  if (revision >= 3) {
    for (let i = 0; i < 50; i++) {
      hash = md5(hash.subarray(0, keyLength));
    }
  }

  const key = hash.subarray(0, keyLength);
  let result: Uint8Array = new Uint8Array(oValue.subarray(0, 32));

  if (revision === 2) {
    result = rc4(key, result);
  } else if (revision >= 3) {
    for (let i = 19; i >= 0; i--) {
      const modKey = new Uint8Array(key.length);
      for (let j = 0; j < key.length; j++) {
        modKey[j] = key[j] ^ i;
      }
      result = rc4(modKey, result);
    }
  }

  // Convert result bytes to password string
  let pwd = "";
  for (let i = 0; i < 32; i++) {
    if (
      result[i] === PASSWORD_PADDING[0] &&
      arraysEqual(
        result.subarray(i, i + Math.min(32 - i, 32)),
        PASSWORD_PADDING.subarray(0, Math.min(32 - i, 32))
      )
    ) {
      break;
    }
    pwd += String.fromCharCode(result[i]);
  }

  return pwd;
}

// =============================================================================
// AES-128 Decryption
// =============================================================================

/**
 * Decrypt data using AES-128-CBC.
 * Per PDF spec, the first 16 bytes of the data are the IV.
 */
function decryptAes128(
  data: Uint8Array,
  objectNumber: number,
  generation: number,
  encryptionKey: Uint8Array
): Uint8Array {
  if (data.length < 16) {
    return data;
  }

  // Compute per-object key: MD5(encryptionKey + objNum(3LE) + genNum(2LE) + "sAlT")
  const keyInput = new Uint8Array(encryptionKey.length + 5 + 4);
  keyInput.set(encryptionKey);
  keyInput[encryptionKey.length] = objectNumber & 0xff;
  keyInput[encryptionKey.length + 1] = (objectNumber >> 8) & 0xff;
  keyInput[encryptionKey.length + 2] = (objectNumber >> 16) & 0xff;
  keyInput[encryptionKey.length + 3] = generation & 0xff;
  keyInput[encryptionKey.length + 4] = (generation >> 8) & 0xff;
  // AES salt
  keyInput[encryptionKey.length + 5] = 0x73; // s
  keyInput[encryptionKey.length + 6] = 0x41; // A
  keyInput[encryptionKey.length + 7] = 0x6c; // l
  keyInput[encryptionKey.length + 8] = 0x54; // T

  const objKey = md5(keyInput);
  const keyLen = Math.min(encryptionKey.length + 5, 16);
  const aesKey = objKey.subarray(0, keyLen);

  // Extract IV (first 16 bytes) and ciphertext
  const iv = data.subarray(0, 16);
  const ciphertext = data.subarray(16);

  if (ciphertext.length === 0 || ciphertext.length % 16 !== 0) {
    return data;
  }

  return aesCbcDecrypt(ciphertext, aesKey, iv);
}

// =============================================================================
// AES Implementation (Pure JavaScript)
// =============================================================================

/** AES S-Box */
const SBOX = new Uint8Array([
  0x63, 0x7c, 0x77, 0x7b, 0xf2, 0x6b, 0x6f, 0xc5, 0x30, 0x01, 0x67, 0x2b, 0xfe, 0xd7, 0xab, 0x76,
  0xca, 0x82, 0xc9, 0x7d, 0xfa, 0x59, 0x47, 0xf0, 0xad, 0xd4, 0xa2, 0xaf, 0x9c, 0xa4, 0x72, 0xc0,
  0xb7, 0xfd, 0x93, 0x26, 0x36, 0x3f, 0xf7, 0xcc, 0x34, 0xa5, 0xe5, 0xf1, 0x71, 0xd8, 0x31, 0x15,
  0x04, 0xc7, 0x23, 0xc3, 0x18, 0x96, 0x05, 0x9a, 0x07, 0x12, 0x80, 0xe2, 0xeb, 0x27, 0xb2, 0x75,
  0x09, 0x83, 0x2c, 0x1a, 0x1b, 0x6e, 0x5a, 0xa0, 0x52, 0x3b, 0xd6, 0xb3, 0x29, 0xe3, 0x2f, 0x84,
  0x53, 0xd1, 0x00, 0xed, 0x20, 0xfc, 0xb1, 0x5b, 0x6a, 0xcb, 0xbe, 0x39, 0x4a, 0x4c, 0x58, 0xcf,
  0xd0, 0xef, 0xaa, 0xfb, 0x43, 0x4d, 0x33, 0x85, 0x45, 0xf9, 0x02, 0x7f, 0x50, 0x3c, 0x9f, 0xa8,
  0x51, 0xa3, 0x40, 0x8f, 0x92, 0x9d, 0x38, 0xf5, 0xbc, 0xb6, 0xda, 0x21, 0x10, 0xff, 0xf3, 0xd2,
  0xcd, 0x0c, 0x13, 0xec, 0x5f, 0x97, 0x44, 0x17, 0xc4, 0xa7, 0x7e, 0x3d, 0x64, 0x5d, 0x19, 0x73,
  0x60, 0x81, 0x4f, 0xdc, 0x22, 0x2a, 0x90, 0x88, 0x46, 0xee, 0xb8, 0x14, 0xde, 0x5e, 0x0b, 0xdb,
  0xe0, 0x32, 0x3a, 0x0a, 0x49, 0x06, 0x24, 0x5c, 0xc2, 0xd3, 0xac, 0x62, 0x91, 0x95, 0xe4, 0x79,
  0xe7, 0xc8, 0x37, 0x6d, 0x8d, 0xd5, 0x4e, 0xa9, 0x6c, 0x56, 0xf4, 0xea, 0x65, 0x7a, 0xae, 0x08,
  0xba, 0x78, 0x25, 0x2e, 0x1c, 0xa6, 0xb4, 0xc6, 0xe8, 0xdd, 0x74, 0x1f, 0x4b, 0xbd, 0x8b, 0x8a,
  0x70, 0x3e, 0xb5, 0x66, 0x48, 0x03, 0xf6, 0x0e, 0x61, 0x35, 0x57, 0xb9, 0x86, 0xc1, 0x1d, 0x9e,
  0xe1, 0xf8, 0x98, 0x11, 0x69, 0xd9, 0x8e, 0x94, 0x9b, 0x1e, 0x87, 0xe9, 0xce, 0x55, 0x28, 0xdf,
  0x8c, 0xa1, 0x89, 0x0d, 0xbf, 0xe6, 0x42, 0x68, 0x41, 0x99, 0x2d, 0x0f, 0xb0, 0x54, 0xbb, 0x16
]);

/** Inverse S-Box for decryption */
const INV_SBOX = new Uint8Array(256);
/* @__PURE__ */ (() => {
  for (let i = 0; i < 256; i++) {
    INV_SBOX[SBOX[i]] = i;
  }
})();

/** AES round constants */
const RCON = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1b, 0x36];

/** GF(2^8) multiplication by 2 */
function gf2(a: number): number {
  return a < 128 ? a << 1 : (a << 1) ^ 0x11b;
}

/** GF(2^8) multiplication */
function gfMul(a: number, b: number): number {
  let result = 0;
  let aa = a;
  let bb = b;
  while (bb > 0) {
    if (bb & 1) {
      result ^= aa;
    }
    aa = gf2(aa);
    bb >>= 1;
  }
  return result;
}

/**
 * AES key expansion. Supports AES-128 (16-byte key) and AES-256 (32-byte key).
 */
function aesKeyExpansion(key: Uint8Array): Uint8Array[] {
  const nk = key.length / 4; // 4 for AES-128, 8 for AES-256
  const nr = nk + 6; // 10 for AES-128, 14 for AES-256
  const w: Uint8Array[] = [];

  // First nk words are the key
  for (let i = 0; i < nk; i++) {
    w.push(new Uint8Array([key[4 * i], key[4 * i + 1], key[4 * i + 2], key[4 * i + 3]]));
  }

  for (let i = nk; i < 4 * (nr + 1); i++) {
    const temp = new Uint8Array(w[i - 1]);
    if (i % nk === 0) {
      // RotWord + SubWord + Rcon
      const t0 = temp[0];
      temp[0] = SBOX[temp[1]] ^ RCON[i / nk - 1];
      temp[1] = SBOX[temp[2]];
      temp[2] = SBOX[temp[3]];
      temp[3] = SBOX[t0];
    } else if (nk > 6 && i % nk === 4) {
      // AES-256 extra SubWord
      temp[0] = SBOX[temp[0]];
      temp[1] = SBOX[temp[1]];
      temp[2] = SBOX[temp[2]];
      temp[3] = SBOX[temp[3]];
    }
    const word = new Uint8Array(4);
    for (let j = 0; j < 4; j++) {
      word[j] = w[i - nk][j] ^ temp[j];
    }
    w.push(word);
  }

  return w;
}

/**
 * Decrypt a single AES block (16 bytes).
 * State layout: column-major, state[4*col + row] = element at (row, col).
 * Input bytes are read in order: byte i → state[i] = state[4*(i>>2) + (i&3)]... no,
 * FIPS 197: input[r + 4*c] = state[r,c], so state[4*c + r] = input[4*c + r].
 * Since we read input straight into state[0..15], state[i] = input[i] where
 * i = 4*col + row with col = i >> 2, row = i & 3... wait:
 *
 * FIPS 197 §3.4: in[4*c + r] = s[r,c]. With our flat array state[4*c + r]:
 * - Row r: indices {r, r+4, r+8, r+12}
 * - Col c: indices {4c, 4c+1, 4c+2, 4c+3}
 */
function aesDecryptBlock(block: Uint8Array, roundKeys: Uint8Array[]): Uint8Array {
  const nr = roundKeys.length / 4 - 1; // 10 for AES-128, 14 for AES-256
  const state = new Uint8Array(16);
  state.set(block);

  // Initial round key addition: state[4*c + r] ^= roundKey[nr*4 + c][r]
  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[nr * 4 + c][r];
    }
  }

  // Rounds (nr-1) down to 1
  for (let round = nr - 1; round >= 1; round--) {
    // InvShiftRows: row r shifts right by r positions
    // Row 0 (indices 0,4,8,12): no shift
    // Row 1 (indices 1,5,9,13): shift right by 1
    let tmp: number;
    tmp = state[13];
    state[13] = state[9];
    state[9] = state[5];
    state[5] = state[1];
    state[1] = tmp;
    // Row 2 (indices 2,6,10,14): shift right by 2
    tmp = state[2];
    state[2] = state[10];
    state[10] = tmp;
    tmp = state[6];
    state[6] = state[14];
    state[14] = tmp;
    // Row 3 (indices 3,7,11,15): shift right by 3
    tmp = state[3];
    state[3] = state[7];
    state[7] = state[11];
    state[11] = state[15];
    state[15] = tmp;

    // InvSubBytes
    for (let i = 0; i < 16; i++) {
      state[i] = INV_SBOX[state[i]];
    }

    // AddRoundKey: state[4*c + r] ^= roundKey[round*4 + c][r]
    const keyOffset = round * 4;
    for (let c = 0; c < 4; c++) {
      for (let r = 0; r < 4; r++) {
        state[4 * c + r] ^= roundKeys[keyOffset + c][r];
      }
    }

    // InvMixColumns: operates on columns, col c = indices {4c, 4c+1, 4c+2, 4c+3}
    for (let c = 0; c < 4; c++) {
      const s0 = state[4 * c];
      const s1 = state[4 * c + 1];
      const s2 = state[4 * c + 2];
      const s3 = state[4 * c + 3];
      state[4 * c] = gfMul(s0, 14) ^ gfMul(s1, 11) ^ gfMul(s2, 13) ^ gfMul(s3, 9);
      state[4 * c + 1] = gfMul(s0, 9) ^ gfMul(s1, 14) ^ gfMul(s2, 11) ^ gfMul(s3, 13);
      state[4 * c + 2] = gfMul(s0, 13) ^ gfMul(s1, 9) ^ gfMul(s2, 14) ^ gfMul(s3, 11);
      state[4 * c + 3] = gfMul(s0, 11) ^ gfMul(s1, 13) ^ gfMul(s2, 9) ^ gfMul(s3, 14);
    }
  }

  // Final round (no InvMixColumns)
  // InvShiftRows
  let tmp: number;
  tmp = state[13];
  state[13] = state[9];
  state[9] = state[5];
  state[5] = state[1];
  state[1] = tmp;
  tmp = state[2];
  state[2] = state[10];
  state[10] = tmp;
  tmp = state[6];
  state[6] = state[14];
  state[14] = tmp;
  tmp = state[3];
  state[3] = state[7];
  state[7] = state[11];
  state[11] = state[15];
  state[15] = tmp;

  // InvSubBytes
  for (let i = 0; i < 16; i++) {
    state[i] = INV_SBOX[state[i]];
  }

  // AddRoundKey (round 0)
  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[c][r];
    }
  }

  return state;
}

/**
 * AES-CBC decryption. Supports both AES-128 and AES-256 based on key length.
 */
function aesCbcDecrypt(ciphertext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const roundKeys = aesKeyExpansion(key);
  const numBlocks = ciphertext.length / 16;
  const output = new Uint8Array(ciphertext.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = ciphertext.subarray(b * 16, (b + 1) * 16);
    const decrypted = aesDecryptBlock(block, roundKeys);

    // XOR with previous ciphertext block (or IV for first block)
    for (let i = 0; i < 16; i++) {
      output[b * 16 + i] = decrypted[i] ^ prevBlock[i];
    }
    prevBlock = block;
  }

  // Remove PKCS#7 padding
  if (output.length > 0) {
    const padLen = output[output.length - 1];
    if (padLen > 0 && padLen <= 16) {
      let validPadding = true;
      for (let i = 0; i < padLen; i++) {
        if (output[output.length - 1 - i] !== padLen) {
          validPadding = false;
          break;
        }
      }
      if (validPadding) {
        return output.subarray(0, output.length - padLen);
      }
    }
  }

  return output;
}

/**
 * AES-256-CBC decryption without PKCS#7 padding removal.
 * Used for key derivation in V=5 where the output is exactly 32 bytes.
 */
function aes256CbcDecryptRaw(ciphertext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const roundKeys = aesKeyExpansion(key);
  const numBlocks = ciphertext.length / 16;
  const output = new Uint8Array(ciphertext.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = ciphertext.subarray(b * 16, (b + 1) * 16);
    const decrypted = aesDecryptBlock(block, roundKeys);

    for (let i = 0; i < 16; i++) {
      output[b * 16 + i] = decrypted[i] ^ prevBlock[i];
    }
    prevBlock = block;
  }

  return output;
}

// =============================================================================
// SHA-256 Implementation (Pure JavaScript)
// =============================================================================

/** SHA-256 initial hash values */
const SHA256_H = new Uint32Array([
  0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
]);

/** SHA-256 round constants */
const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
]);

/**
 * SHA-256 hash function (pure JavaScript implementation).
 * @param input - The data to hash
 * @returns 32-byte hash
 */
function sha256(input: Uint8Array): Uint8Array {
  // Pre-processing: pad message
  const msgLen = input.length;
  // Message must be padded to 64-byte (512-bit) blocks
  // Padding: 1 bit, then zeros, then 64-bit big-endian length
  const paddedLen = Math.ceil((msgLen + 9) / 64) * 64;
  const padded = new Uint8Array(paddedLen);
  padded.set(input);
  padded[msgLen] = 0x80;

  // Append length in bits as 64-bit big-endian
  const bitLen = msgLen * 8;
  // For messages < 2^32 bytes, high 32 bits are 0
  const view = new DataView(padded.buffer, padded.byteOffset, padded.byteLength);
  view.setUint32(paddedLen - 8, 0, false); // high 32 bits
  view.setUint32(paddedLen - 4, bitLen, false); // low 32 bits

  // Initialize hash values
  let h0 = SHA256_H[0];
  let h1 = SHA256_H[1];
  let h2 = SHA256_H[2];
  let h3 = SHA256_H[3];
  let h4 = SHA256_H[4];
  let h5 = SHA256_H[5];
  let h6 = SHA256_H[6];
  let h7 = SHA256_H[7];

  // Process each 64-byte block
  const w = new Uint32Array(64);
  for (let offset = 0; offset < paddedLen; offset += 64) {
    // Prepare message schedule
    for (let i = 0; i < 16; i++) {
      w[i] = view.getUint32(offset + i * 4, false);
    }
    for (let i = 16; i < 64; i++) {
      const s0 = rotr32(w[i - 15], 7) ^ rotr32(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr32(w[i - 2], 17) ^ rotr32(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }

    // Initialize working variables
    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    let f = h5;
    let g = h6;
    let h = h7;

    // Compression function
    for (let i = 0; i < 64; i++) {
      const S1 = rotr32(e, 6) ^ rotr32(e, 11) ^ rotr32(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + SHA256_K[i] + w[i]) >>> 0;
      const S0 = rotr32(a, 2) ^ rotr32(a, 13) ^ rotr32(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;

      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    // Add compressed chunk to hash value
    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
    h5 = (h5 + f) >>> 0;
    h6 = (h6 + g) >>> 0;
    h7 = (h7 + h) >>> 0;
  }

  // Produce the final hash value (big-endian)
  const result = new Uint8Array(32);
  const resultView = new DataView(result.buffer);
  resultView.setUint32(0, h0, false);
  resultView.setUint32(4, h1, false);
  resultView.setUint32(8, h2, false);
  resultView.setUint32(12, h3, false);
  resultView.setUint32(16, h4, false);
  resultView.setUint32(20, h5, false);
  resultView.setUint32(24, h6, false);
  resultView.setUint32(28, h7, false);

  return result;
}

/** 32-bit right rotation */
function rotr32(x: number, n: number): number {
  return ((x >>> n) | (x << (32 - n))) >>> 0;
}

// =============================================================================
// Helpers
// =============================================================================

function padPassword(password: string): Uint8Array {
  const result = new Uint8Array(32);
  const bytes = textEncoder.encode(password);
  const len = Math.min(bytes.length, 32);
  result.set(bytes.subarray(0, len));
  result.set(PASSWORD_PADDING.subarray(0, 32 - len), len);
  return result;
}

function arraysEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

/**
 * Concatenate multiple Uint8Arrays.
 */
function concatArrays(...arrays: Uint8Array[]): Uint8Array {
  let totalLen = 0;
  for (const arr of arrays) {
    totalLen += arr.length;
  }
  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (const arr of arrays) {
    result.set(arr, offset);
    offset += arr.length;
  }
  return result;
}

/**
 * Read the EncryptMetadata flag from the encrypt dictionary.
 * Per spec, defaults to true if not present.
 * Checks both the top-level dict and CF/StdCF sub-dictionary.
 */
function readEncryptMetadata(encryptDict: PdfDictValue): boolean {
  // Check top-level EncryptMetadata first
  const topLevel = dictGetBool(encryptDict, "EncryptMetadata");
  if (topLevel !== undefined) {
    return topLevel;
  }

  // Check CF/StdCF/EncryptMetadata
  const cf = encryptDict.get("CF");
  if (cf && cf instanceof Map) {
    const stdCF = cf.get("StdCF");
    if (stdCF && stdCF instanceof Map) {
      const cfVal = stdCF.get("EncryptMetadata");
      if (typeof cfVal === "boolean") {
        return cfVal;
      }
    }
  }

  // Default per spec
  return true;
}

/**
 * Check if V4 encryption uses AES (vs RC4).
 */
function isAesCryptFilter(encryptDict: PdfDictValue): boolean {
  const cf = encryptDict.get("CF");
  if (!cf || !(cf instanceof Map)) {
    return false;
  }
  // Check StdCF filter
  const stdCF = cf.get("StdCF");
  if (!stdCF || !(stdCF instanceof Map)) {
    return false;
  }
  const cfm = stdCF.get("CFM");
  return cfm === "AESV2";
}

// =============================================================================
// Test-only Exports
// =============================================================================

/**
 * @internal — Exported only for unit testing crypto primitives against known test vectors.
 * Not part of the public API.
 */
export const _testInternals = {
  aesCbcDecrypt,
  sha256
};
