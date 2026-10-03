/**
 * PDF decryption for reading encrypted PDFs.
 *
 * Supports:
 * - Standard Security Handler (V1/V2/V4/V5, R2/R3/R4/R5)
 * - RC4 encryption (40-bit and 128-bit)
 * - AES-128 encryption (PDF 1.6+)
 * - AES-256 encryption (PDF 2.0, V=5, R=5)
 *
 * @see PDF Reference 1.7, §3.5 - Encryption
 * @see PDF 2.0 (ISO 32000-2), §7.6 - Encryption
 */

import { pdfMd5, pdfSha256, pdfSha384, pdfSha512 } from "@pdf/core/pdf-kdf";
import { PdfStructureError } from "@pdf/errors";
import type { PdfCryptKind, PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";
import {
  dictGetNumber,
  dictGetName,
  dictGetBytes,
  dictGetArray,
  dictGetBool
} from "@pdf/reader/pdf-parser";
import { concatUint8Arrays } from "@utils/binary";
import {
  rc4,
  aesCbcDecrypt,
  aesCbcDecryptRaw,
  aesCbcEncrypt,
  aesCbcEncryptRaw,
  randomBytes
} from "@utils/crypto";

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

  // V4 names a crypt filter for strings and one for streams; either may be
  // /Identity. Resolve them before the password so a malformed /CF fails loudly.
  const stringMethod = cryptFilterMethod(encryptDict, v, "StrF");
  const streamMethod = cryptFilterMethod(encryptDict, v, "StmF");

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

  installSecurityHandler(
    doc,
    createCipher(stringMethod, encryptionKey),
    createCipher(streamMethod, encryptionKey),
    // /EncryptMetadata is only defined from V4 on (ISO 32000-1 Table 20).
    v >= 4 ? encryptMetadata : true
  );
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
  if (revision !== 5 && revision !== 6) {
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
  // R=5 hashes once with SHA-256; R=6 uses Algorithm 2.B.
  const kdf: PasswordHash =
    revision === 6
      ? hashR6
      : (pwd, salt, udata) => pdfSha256(concatUint8Arrays([pwd, salt, udata]));
  let encryptionKey = tryUserPasswordV5(passwordBytes, uValue, ueValue, kdf);

  if (!encryptionKey) {
    // Try owner password (Algorithm 2.A step a - owner)
    encryptionKey = tryOwnerPasswordV5(passwordBytes, oValue, oeValue, uValue, kdf);
  }

  if (!encryptionKey) {
    // Try empty password
    if (password !== "") {
      const emptyBytes = new Uint8Array(0);
      encryptionKey = tryUserPasswordV5(emptyBytes, uValue, ueValue, kdf);
      if (!encryptionKey) {
        encryptionKey = tryOwnerPasswordV5(emptyBytes, oValue, oeValue, uValue, kdf);
      }
    }
  }

  if (!encryptionKey) {
    return false;
  }

  // V=5 uses AES-256 with the file encryption key directly (no per-object
  // key derivation), through whichever crypt filters /StrF and /StmF name.
  installSecurityHandler(
    doc,
    createCipher(cryptFilterMethod(encryptDict, 5, "StrF"), encryptionKey),
    createCipher(cryptFilterMethod(encryptDict, 5, "StmF"), encryptionKey),
    readEncryptMetadata(encryptDict)
  );
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
  ueValue: Uint8Array,
  kdf: PasswordHash
): Uint8Array | null {
  // U = hash(32) + validation salt(8) + key salt(8)
  const uHash = uValue.subarray(0, 32);
  const uValidationSalt = uValue.subarray(32, 40);
  const uKeySalt = uValue.subarray(40, 48);

  // Validate: SHA-256(password + validation salt) == first 32 bytes of U
  const computedHash = kdf(passwordBytes, uValidationSalt, EMPTY);

  if (!arraysEqual(computedHash, uHash)) {
    return null;
  }

  // Derive key: SHA-256(password + key salt) => use as AES-256 key to decrypt UE
  const keyHash = kdf(passwordBytes, uKeySalt, EMPTY);

  // Decrypt UE with this key using AES-256-CBC with zero IV
  const zeroIv = new Uint8Array(16);
  return aesCbcDecryptRaw(ueValue.subarray(0, 32), keyHash, zeroIv);
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
  uValue: Uint8Array,
  kdf: PasswordHash
): Uint8Array | null {
  // O = hash(32) + validation salt(8) + key salt(8)
  const oHash = oValue.subarray(0, 32);
  const oValidationSalt = oValue.subarray(32, 40);
  const oKeySalt = oValue.subarray(40, 48);
  const u48 = uValue.subarray(0, 48);

  // Validate: SHA-256(password + validation salt + U(0..47)) == first 32 bytes of O
  const computedHash = kdf(passwordBytes, oValidationSalt, u48);

  if (!arraysEqual(computedHash, oHash)) {
    return null;
  }

  // Derive key: SHA-256(password + key salt + U(0..47))
  const keyHash = kdf(passwordBytes, oKeySalt, u48);

  // Decrypt OE with this key using AES-256-CBC with zero IV
  const zeroIv = new Uint8Array(16);
  return aesCbcDecryptRaw(oeValue.subarray(0, 32), keyHash, zeroIv);
}

/** Password hash of Algorithm 2.A: `hash(password, salt, userData)`. */
type PasswordHash = (password: Uint8Array, salt: Uint8Array, userData: Uint8Array) => Uint8Array;

const EMPTY = new Uint8Array(0);

/**
 * Algorithm 2.B (ISO 32000-2 §7.6.4.3.4), the R=6 password hash: SHA-256 of
 * the input, then at least 64 rounds that AES-128-CBC encrypt 64 copies of
 * `password ‖ K ‖ userData` under K and rehash with SHA-256/384/512 chosen by
 * the first 16 bytes of the result mod 3.
 */
function hashR6(password: Uint8Array, salt: Uint8Array, userData: Uint8Array): Uint8Array {
  let k = pdfSha256(concatUint8Arrays([password, salt, userData]));
  for (let round = 0; ; round++) {
    const unit = concatUint8Arrays([password, k, userData]);
    const k1 = new Uint8Array(unit.length * 64);
    for (let i = 0; i < 64; i++) {
      k1.set(unit, i * unit.length);
    }
    const e = aesCbcEncryptRaw(k1, k.subarray(0, 16), k.subarray(16, 32));
    let sum = 0;
    for (let i = 0; i < 16; i++) {
      sum += e[i];
    }
    const pick = sum % 3;
    k = pick === 0 ? pdfSha256(e) : pick === 1 ? pdfSha384(e) : pdfSha512(e);
    if (round >= 63 && e[e.length - 1] <= round - 31) {
      break;
    }
  }
  return k.subarray(0, 32);
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
    const hash = pdfMd5(hashInput);
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

  let hash = pdfMd5(input.subarray(0, offset));

  // For revision >= 3, hash 50 more times
  if (revision >= 3) {
    for (let i = 0; i < 50; i++) {
      hash = pdfMd5(hash.subarray(0, keyLength));
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
  let hash = pdfMd5(padPassword(ownerPassword));

  if (revision >= 3) {
    for (let i = 0; i < 50; i++) {
      hash = pdfMd5(hash.subarray(0, keyLength));
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
 * Per-object key (Algorithm 1, ISO 32000-1 §7.6.2): MD5 of the file key, the
 * object number (3 bytes LE), the generation (2 bytes LE) and, for AES, the
 * salt "sAlT"; truncated to min(n + 5, 16) bytes.
 */
function objectKey(
  encryptionKey: Uint8Array,
  objectNumber: number,
  generation: number,
  aes: boolean
): Uint8Array {
  const n = encryptionKey.length;
  const keyInput = new Uint8Array(n + 5 + (aes ? 4 : 0));
  keyInput.set(encryptionKey);
  keyInput[n] = objectNumber & 0xff;
  keyInput[n + 1] = (objectNumber >> 8) & 0xff;
  keyInput[n + 2] = (objectNumber >> 16) & 0xff;
  keyInput[n + 3] = generation & 0xff;
  keyInput[n + 4] = (generation >> 8) & 0xff;
  if (aes) {
    keyInput.set([0x73, 0x41, 0x6c, 0x54], n + 5); // "sAlT"
  }
  return pdfMd5(keyInput).subarray(0, Math.min(n + 5, 16));
}

/** Decrypt data using RC4 with the per-object key. */
function decryptRc4PerObject(
  data: Uint8Array,
  objectNumber: number,
  generation: number,
  encryptionKey: Uint8Array
): Uint8Array {
  return rc4(objectKey(encryptionKey, objectNumber, generation, false), data);
}

/**
 * Decrypt data using AES-128-CBC with the per-object key.
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
  const iv = data.subarray(0, 16);
  const ciphertext = data.subarray(16);
  if (ciphertext.length === 0 || ciphertext.length % 16 !== 0) {
    return data;
  }
  return aesCbcDecrypt(ciphertext, objectKey(encryptionKey, objectNumber, generation, true), iv);
}

/** AES-CBC with PKCS#7 padding and a random IV prepended, as PDF stores it. */
function aesCbcEncryptWithIv(data: Uint8Array, key: Uint8Array): Uint8Array {
  const iv = randomBytes(16);
  return concatUint8Arrays([iv, aesCbcEncrypt(data, key, iv)]);
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

// =============================================================================
// Crypt filters (ISO 32000-1 §7.6.5)
// =============================================================================

/** How one class of data (strings or streams) is encrypted. */
type CryptMethod = "identity" | "rc4" | "aesv2" | "aesv3";

/** Per-object encryption in both directions; `undefined` means Identity. */
interface Cipher {
  decrypt(data: Uint8Array, objNum: number, gen: number): Uint8Array;
  encrypt(data: Uint8Array, objNum: number, gen: number): Uint8Array;
}

/**
 * Resolve the crypt filter `/StrF` or `/StmF` names to its method. Before V4
 * there are no crypt filters and everything is RC4. From V4 on an absent
 * entry means `/Identity` (Table 20), and a named filter is looked up in
 * `/CF` (Table 25), where `/CFM /None` also leaves the data as it is.
 */
function cryptFilterMethod(
  encryptDict: PdfDictValue,
  v: number,
  entry: "StrF" | "StmF"
): CryptMethod {
  if (v < 4) {
    return "rc4";
  }
  const name = dictGetName(encryptDict, entry) ?? "Identity";
  if (name === "Identity") {
    return "identity";
  }
  const cf = encryptDict.get("CF");
  const filter = cf instanceof Map ? cf.get(name) : undefined;
  if (!(filter instanceof Map)) {
    throw new PdfStructureError(`/${entry} names crypt filter /${name}, which /CF does not define`);
  }
  const cfm = filter.get("CFM") ?? "None";
  switch (cfm) {
    case "None":
      return "identity";
    case "V2":
      return "rc4";
    case "AESV2":
      return "aesv2";
    case "AESV3":
      return "aesv3";
    default:
      throw new PdfStructureError(`Unsupported crypt filter method: ${String(cfm)}`);
  }
}

function createCipher(method: CryptMethod, key: Uint8Array): Cipher | undefined {
  switch (method) {
    case "identity":
      return undefined;
    case "rc4": {
      // RC4 is its own inverse.
      const apply = (data: Uint8Array, objNum: number, gen: number): Uint8Array =>
        decryptRc4PerObject(data, objNum, gen, key);
      return { decrypt: apply, encrypt: apply };
    }
    case "aesv2":
      return {
        decrypt: (data, objNum, gen) => decryptAes128(data, objNum, gen, key),
        encrypt: (data, objNum, gen) => aesCbcEncryptWithIv(data, objectKey(key, objNum, gen, true))
      };
    case "aesv3":
      return {
        decrypt: data => decryptAes256Direct(data, key),
        encrypt: data => aesCbcEncryptWithIv(data, key)
      };
  }
}

/** Install the document's security handler: Identity filters pass data through. */
function installSecurityHandler(
  doc: PdfDocument,
  strings: Cipher | undefined,
  streams: Cipher | undefined,
  encryptMetadata: boolean
): void {
  const pick = (kind: PdfCryptKind): Cipher | undefined => (kind === "string" ? strings : streams);
  doc.decryptFn = (data, objNum, gen, kind) => pick(kind)?.decrypt(data, objNum, gen) ?? data;
  doc.encryptFn = (data, objNum, gen, kind) => pick(kind)?.encrypt(data, objNum, gen) ?? data;
  doc.encryptMetadata = encryptMetadata;
}
