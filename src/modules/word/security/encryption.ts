/**
 * DOCX Agile Encryption (MS-OFFCRYPTO)
 *
 * Implements ECMA-376 Agile Encryption (used by Office 2010+) for encrypted
 * DOCX files. Encrypted DOCX files use the CFB (Compound File Binary) format
 * to wrap the ZIP content.
 *
 * This module provides:
 * - Password-based encryption/decryption of DOCX content
 * - EncryptionInfo XML generation/parsing
 *
 * Note: A full implementation requires a CFB container reader/writer for the
 * outer compound document format. For now, we expose primitives that can be
 * combined with an external CFB library if needed.
 *
 * References:
 *   - MS-OFFCRYPTO: Office Document Cryptography Structure
 *   - ECMA-376 Part 3: Markup Compatibility and Extensibility
 */

import {
  base64ToBytes,
  bytesToBase64,
  randomBytes,
  utf8Decoder,
  utf8Encoder
} from "../core/internal-utils";
import { DocxDecryptionError } from "../errors";
import { readCfb, writeCfb } from "./cfb-reader";

/** Agile encryption parameters. */
export interface AgileEncryptionInfo {
  /** Cipher algorithm (AES). */
  readonly cipherAlgorithm: "AES";
  /** Cipher chaining mode. */
  readonly cipherChaining: "ChainingModeCBC";
  /** Key bit length (128, 192, 256). */
  readonly keyBits: 128 | 192 | 256;
  /** Hash algorithm. */
  readonly hashAlgorithm: "SHA1" | "SHA256" | "SHA384" | "SHA512";
  /** Hash output size in bytes. */
  readonly hashSize: number;
  /** Number of iterations (spin count). */
  readonly spinCount: number;
  /** Salt for key derivation (16 bytes). */
  readonly keySalt: Uint8Array;
  /** Encrypted verifier hash input. */
  readonly encryptedVerifierHashInput: Uint8Array;
  /** Encrypted verifier hash value. */
  readonly encryptedVerifierHashValue: Uint8Array;
  /** Encrypted key value. */
  readonly encryptedKeyValue: Uint8Array;
  /** Block size (16 for AES). */
  readonly blockSize: number;
  /** Data integrity salt. */
  readonly dataIntegritySalt?: Uint8Array;
  /** Data integrity HMAC value. */
  readonly dataIntegrityHmac?: Uint8Array;
}

/** Detect whether a buffer is an encrypted Office document (CFB format). */
export function isEncryptedDocx(buffer: Uint8Array): boolean {
  // CFB signature: D0CF11E0A1B11AE1
  if (buffer.length < 8) {
    return false;
  }
  return (
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0 &&
    buffer[4] === 0xa1 &&
    buffer[5] === 0xb1 &&
    buffer[6] === 0x1a &&
    buffer[7] === 0xe1
  );
}

/**
 * Derive an encryption key from a password using the agile encryption KDF.
 *
 * Per MS-OFFCRYPTO 2.3.4.11:
 *   H_0 = H(salt + password)
 *   H_i = H(iterator + H_{i-1})  for i = 0..spinCount-1
 *   H_final = H(H_{spinCount} + blockKey)
 *   Key = first keySize bytes of H_final
 */
export async function deriveEncryptionKey(
  password: string,
  info: {
    keySalt: Uint8Array;
    spinCount: number;
    hashAlgorithm: string;
    keyBits: number;
  },
  blockKey: Uint8Array
): Promise<Uint8Array> {
  const pwdBytes = stringToUtf16LE(password);
  const hash = mapHashName(info.hashAlgorithm);

  // H_0 = H(salt + password)
  let h = await sha(hash, concat(info.keySalt, pwdBytes));

  // Iterate: H_i = H(iterator (4 bytes LE) + H_{i-1})
  for (let i = 0; i < info.spinCount; i++) {
    const iter = new Uint8Array(4);
    new DataView(iter.buffer).setUint32(0, i, true);
    h = await sha(hash, concat(iter, h));
  }

  // Final: H_final = H(H + blockKey)
  const final = await sha(hash, concat(h, blockKey));

  // Truncate to keySize. The hash size MUST be at least the requested
  // key length — otherwise we'd hand WebCrypto a short key buffer that
  // either fails AES import (`OperationError`) or produces a key the
  // counterparty can't validate. Reject misconfigured EncryptionInfo
  // up-front with a clear error.
  const keyBytes = info.keyBits / 8;
  if (final.length < keyBytes) {
    throw new Error(
      `deriveEncryptionKey: hash output of ${final.length} bytes is too ` +
        `short for keyBits=${info.keyBits} (need ${keyBytes}). ` +
        `Use a hash algorithm with a larger digest size (e.g. SHA-512 ` +
        `for keyBits ≤ 512).`
    );
  }
  return final.slice(0, keyBytes);
}

/** Block keys as defined in MS-OFFCRYPTO 2.3.4.13. */
export const AGILE_BLOCK_KEYS = {
  /** For encrypting verifier hash input. */
  verifierHashInput: new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  /** For encrypting verifier hash value. */
  verifierHashValue: new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  /** For encrypting the main key. */
  encryptedKey: new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  /** For data integrity HMAC key. */
  dataIntegrityKey: new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  /** For data integrity HMAC value. */
  dataIntegrityValue: new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])
} as const;

/**
 * Verify a password against the encryption info.
 *
 * @returns True if password is correct.
 */
export async function verifyPassword(
  password: string,
  info: AgileEncryptionInfo
): Promise<boolean> {
  try {
    // Derive verifier hash input key
    const verifierInputKey = await deriveEncryptionKey(
      password,
      info,
      AGILE_BLOCK_KEYS.verifierHashInput
    );

    // Decrypt the verifier hash input (PKCS#7 padded by aesCbcEncrypt during write)
    const verifierInput = await aesCbcDecryptPkcs7(
      info.encryptedVerifierHashInput,
      verifierInputKey,
      info.keySalt
    );

    // Hash the verifier input
    const hashAlg = mapHashName(info.hashAlgorithm);
    const computedHash = await sha(hashAlg, verifierInput);

    // Derive verifier hash value key
    const verifierValueKey = await deriveEncryptionKey(
      password,
      info,
      AGILE_BLOCK_KEYS.verifierHashValue
    );

    // Decrypt the verifier hash value (PKCS#7 padded by aesCbcEncrypt during write)
    const expectedHash = await aesCbcDecryptPkcs7(
      info.encryptedVerifierHashValue,
      verifierValueKey,
      info.keySalt
    );

    // Compare (truncate to hashSize in case of padding)
    return bytesEqual(computedHash.slice(0, info.hashSize), expectedHash.slice(0, info.hashSize));
  } catch {
    return false;
  }
}

/**
 * Decrypt the package data using the password.
 *
 * @param encryptedPackage - The EncryptedPackage stream from CFB.
 * @param info - Agile encryption parameters.
 * @param password - The user password.
 * @param maxDecryptedSize - Optional upper bound on the decrypted size; if the
 *   `totalSize` header claims a larger value, throw rather than allocate.
 *   Defaults to 512 MiB so adversarial files cannot trigger arbitrary-size
 *   allocations before security policy is applied at the unzip layer.
 * @returns The decrypted (unencrypted) package bytes.
 */
export async function decryptPackage(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  password: string,
  maxDecryptedSize: number = 512 * 1024 * 1024
): Promise<Uint8Array> {
  // Derive key encryption key
  const keyEncryptionKey = await deriveEncryptionKey(password, info, AGILE_BLOCK_KEYS.encryptedKey);

  // Decrypt the actual package key (PKCS#7 padded by aesCbcEncrypt during write)
  const packageKey = await aesCbcDecryptPkcs7(
    info.encryptedKeyValue,
    keyEncryptionKey,
    info.keySalt
  );

  if (encryptedPackage.length < 8) {
    throw new DocxDecryptionError("EncryptedPackage too small (missing 8-byte size header)");
  }

  // First 8 bytes are the total decrypted size (uint64 LE)
  const totalSizeView = new DataView(encryptedPackage.buffer, encryptedPackage.byteOffset, 8);
  const totalSize = Number(totalSizeView.getBigUint64(0, true));

  // Sanity check the declared size: it cannot legally exceed the encrypted
  // body length, and we refuse to allocate above the configured bound.
  const encryptedBodyLength = encryptedPackage.length - 8;
  if (
    !Number.isFinite(totalSize) ||
    totalSize < 0 ||
    totalSize > encryptedBodyLength ||
    totalSize > maxDecryptedSize
  ) {
    throw new DocxDecryptionError(
      `EncryptedPackage declared size (${totalSize}) is invalid or exceeds the maximum allowed (${maxDecryptedSize})`
    );
  }

  // Decrypt in 4096-byte segments
  const segmentSize = 4096;
  const encData = encryptedPackage.slice(8);
  const segments: Uint8Array[] = [];
  const segCount = Math.ceil(encData.length / segmentSize);

  for (let i = 0; i < segCount; i++) {
    const segStart = i * segmentSize;
    const segEnd = Math.min(segStart + segmentSize, encData.length);
    const segData = encData.slice(segStart, segEnd);

    // Segment IV is hash(salt + segment_index_LE)
    const idxBytes = new Uint8Array(4);
    new DataView(idxBytes.buffer).setUint32(0, i, true);
    const segIv = (
      await sha(mapHashName(info.hashAlgorithm), concat(info.keySalt, idxBytes))
    ).slice(0, info.blockSize);

    // Package segments use zero padding (MS-OFFCRYPTO §2.3.4.15). Web Crypto
    // would silently mis-strip valid-looking PKCS#7 tails, so use the
    // zero-pad-safe path here.
    const decSeg = await aesCbcDecryptZeroPad(segData, packageKey, segIv);
    segments.push(decSeg);
  }

  // Concatenate and truncate to total size
  const result = new Uint8Array(totalSize);
  let offset = 0;
  for (const seg of segments) {
    const copyLen = Math.min(seg.length, totalSize - offset);
    result.set(seg.slice(0, copyLen), offset);
    offset += copyLen;
    if (offset >= totalSize) {
      break;
    }
  }
  return result;
}

// =============================================================================
// Internal Crypto Helpers (use Web Crypto API)
// =============================================================================

function mapHashName(name: string): string {
  const map: Record<string, string> = {
    SHA1: "SHA-1",
    SHA256: "SHA-256",
    SHA384: "SHA-384",
    SHA512: "SHA-512"
  };
  return map[name] ?? "SHA-512";
}

async function sha(algorithm: string, data: Uint8Array): Promise<Uint8Array> {
  const buf = await crypto.subtle.digest(algorithm, toBuffer(data));
  return new Uint8Array(buf);
}

/**
 * Decrypt AES-CBC ciphertext that uses **PKCS#7 padding**.
 *
 * Used for the encryptedKeyValue, encryptedVerifierHashInput and
 * encryptedVerifierHashValue blobs. Web Crypto strips PKCS#7 padding
 * automatically.
 */
async function aesCbcDecryptPkcs7(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    toBuffer(key),
    { name: "AES-CBC", length: key.length * 8 },
    false,
    ["decrypt"]
  );
  const iv16 = new Uint8Array(16);
  iv16.set(iv.slice(0, 16));
  const result = await crypto.subtle.decrypt(
    { name: "AES-CBC", iv: toBuffer(iv16) },
    cryptoKey,
    toBuffer(data)
  );
  return new Uint8Array(result);
}

/**
 * Decrypt AES-CBC ciphertext that uses **zero padding** (Agile package
 * segments).
 *
 * Web Crypto only supports PKCS#7. If the trailing zero-padded plaintext
 * coincidentally satisfies a valid PKCS#7 pattern (last byte 0x01..0x10 with
 * the matching tail), Web Crypto will silently strip those bytes and we'd
 * lose data. To avoid that we always go through the manual path which
 * appends a synthetic PKCS#7 padding block, lets Web Crypto strip it, then
 * slices back to the original ciphertext length — mathematically equivalent
 * to a true zero-pad CBC decryption.
 */
async function aesCbcDecryptZeroPad(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const iv16 = new Uint8Array(16);
  iv16.set(iv.slice(0, 16));
  return manualAesCbcDecrypt(data, key, iv16);
}

async function manualAesCbcDecrypt(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const blockSize = 16;
  if (data.length % blockSize !== 0) {
    throw new DocxDecryptionError("Data length must be a multiple of block size");
  }

  // Strategy: append a crafted ciphertext block that will produce valid PKCS#7
  // padding (\x10 * 16) when decrypted, allowing Web Crypto to accept the input.
  //
  // For CBC decryption, the last plaintext block P_n = AES_Dec(C_n) XOR C_{n-1}.
  // We want P_n = 0x10 repeated 16 times (full PKCS#7 padding block).
  // So we need C_n such that AES_Dec(C_n) = C_{n-1} XOR (0x10 * 16).
  // Equivalently, C_n = AES_Enc(C_{n-1} XOR (0x10 * 16)).
  //
  // We compute this using AES-CBC encrypt with IV=0 on the XOR'd block.
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    toBuffer(key),
    { name: "AES-CBC", length: key.length * 8 },
    false,
    ["encrypt", "decrypt"]
  );

  const lastCipherBlock = data.slice(data.length - blockSize);
  const plainForPad = new Uint8Array(blockSize);
  for (let j = 0; j < blockSize; j++) {
    plainForPad[j] = lastCipherBlock[j] ^ 0x10;
  }

  // Encrypt to get the padding ciphertext block (use zero IV for ECB-like behavior)
  const encResult = await crypto.subtle.encrypt(
    { name: "AES-CBC", iv: toBuffer(new Uint8Array(16)) },
    cryptoKey,
    toBuffer(plainForPad)
  );
  // encResult = [AES_Enc(plainForPad XOR 0), PKCS_padding_block] = 32 bytes
  const padCipherBlock = new Uint8Array(encResult).slice(0, blockSize);

  // Append the crafted block to the ciphertext
  const augmented = concat(data, padCipherBlock);

  // Now decrypt with Web Crypto — the last block will be valid PKCS#7
  const decResult = await crypto.subtle.decrypt(
    { name: "AES-CBC", iv: toBuffer(iv) },
    cryptoKey,
    toBuffer(augmented)
  );

  // Web Crypto strips the padding, so we get exactly data.length bytes
  return new Uint8Array(decResult).slice(0, data.length);
}

/** Convert Uint8Array to a fresh ArrayBuffer (for strict crypto.subtle typing). */
function toBuffer(arr: Uint8Array): ArrayBuffer {
  const buf = new ArrayBuffer(arr.length);
  new Uint8Array(buf).set(arr);
  return buf;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const result = new Uint8Array(a.length + b.length);
  result.set(a, 0);
  result.set(b, a.length);
  return result;
}

function stringToUtf16LE(s: string): Uint8Array {
  const buf = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    buf[i * 2] = code & 0xff;
    buf[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return buf;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a[i] ^ b[i];
  }
  return diff === 0;
}

// =============================================================================
// EncryptionInfo XML Parser
// =============================================================================

/**
 * Parse an EncryptionInfo XML stream from a CFB-wrapped encrypted DOCX.
 *
 * This is the UTF-8 XML part after the 8-byte version/flags header.
 *
 * @param xmlStr - The EncryptionInfo XML content.
 * @returns Parsed agile encryption info.
 */
export function parseEncryptionInfoXml(xmlStr: string): AgileEncryptionInfo {
  // Simple regex-based extraction for the key <keyEncryptors><keyEncryptor>... element
  const keyDataMatch = /<keyData\s([\s\S]*?)\/>/.exec(xmlStr);
  const pwdEncryptorMatch = /<p:encryptedKey\s([\s\S]*?)\/>/.exec(xmlStr);

  if (!keyDataMatch || !pwdEncryptorMatch) {
    throw new DocxDecryptionError("Invalid EncryptionInfo XML - missing keyData or encryptedKey");
  }

  const pwdData = parseAttrs(pwdEncryptorMatch[1]);

  // Required cryptographic fields — empty/missing values cannot decrypt and
  // produce confusing CryptoOperation errors downstream. Fail fast with a
  // clear message instead.
  const requireField = (name: string): string => {
    const v = pwdData[name];
    if (v === undefined || v === "") {
      throw new DocxDecryptionError(
        `Invalid EncryptionInfo XML - missing or empty required field: ${name}`
      );
    }
    return v;
  };

  const keyBitsRaw = parseInt(pwdData.keyBits ?? "256", 10);
  if (keyBitsRaw !== 128 && keyBitsRaw !== 192 && keyBitsRaw !== 256) {
    throw new DocxDecryptionError(
      `Unsupported keyBits in EncryptionInfo: ${pwdData.keyBits} (expected 128, 192, or 256)`
    );
  }
  const hashAlgRaw = pwdData.hashAlgorithm ?? "SHA512";
  if (
    hashAlgRaw !== "SHA1" &&
    hashAlgRaw !== "SHA256" &&
    hashAlgRaw !== "SHA384" &&
    hashAlgRaw !== "SHA512"
  ) {
    throw new DocxDecryptionError(
      `Unsupported hashAlgorithm in EncryptionInfo: ${hashAlgRaw} (expected SHA1/SHA256/SHA384/SHA512)`
    );
  }

  // Validate numeric fields against sane ranges. EncryptionInfo XML is
  // parsed straight from the CFB stream of an arbitrary input file, so
  // every numeric attribute is attacker-controlled. Without bounds
  // checks:
  //   - spinCount: a hostile file specifying spinCount=10⁹ would freeze
  //     the verifier-derivation loop in deriveEncryptionKey.
  //   - hashSize/blockSize: used as slice lengths and AES IV sizes.
  //     Outsized values silently produce wrong keys / IVs and obscure
  //     decryption errors; negative or NaN crashes downstream.
  // Limits are deliberately wider than what real Office files emit so we
  // don't reject legitimate documents.
  const spinCount = parseInt(pwdData.spinCount ?? "100000", 10);
  if (!Number.isFinite(spinCount) || spinCount < 0 || spinCount > 10_000_000) {
    throw new DocxDecryptionError(
      `Invalid spinCount in EncryptionInfo: ${pwdData.spinCount} (expected 0..10_000_000)`
    );
  }
  const hashSize = parseInt(pwdData.hashSize ?? "64", 10);
  // Largest legitimate hash here is SHA-512 → 64 bytes. Cap at 128 to
  // accommodate hypothetical extensions while still rejecting crazy values.
  if (!Number.isFinite(hashSize) || hashSize < 1 || hashSize > 128) {
    throw new DocxDecryptionError(
      `Invalid hashSize in EncryptionInfo: ${pwdData.hashSize} (expected 1..128)`
    );
  }
  const blockSize = parseInt(pwdData.blockSize ?? "16", 10);
  // AES is fixed at 16; some tooling emits 8 for legacy ciphers. Anything
  // outside [8, 64] is bogus.
  if (!Number.isFinite(blockSize) || blockSize < 8 || blockSize > 64) {
    throw new DocxDecryptionError(
      `Invalid blockSize in EncryptionInfo: ${pwdData.blockSize} (expected 8..64)`
    );
  }

  return {
    cipherAlgorithm: "AES",
    cipherChaining: "ChainingModeCBC",
    keyBits: keyBitsRaw,
    hashAlgorithm: hashAlgRaw,
    hashSize,
    spinCount,
    keySalt: base64ToBytes(requireField("saltValue")),
    encryptedVerifierHashInput: base64ToBytes(requireField("encryptedVerifierHashInput")),
    encryptedVerifierHashValue: base64ToBytes(requireField("encryptedVerifierHashValue")),
    encryptedKeyValue: base64ToBytes(requireField("encryptedKeyValue")),
    blockSize
  };
}

function parseAttrs(str: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /(\w+)="([^"]*)"/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(str)) !== null) {
    attrs[match[1]] = match[2];
  }
  return attrs;
}

// =============================================================================
// High-level Encrypted DOCX Reader
// =============================================================================

/**
 * Read an encrypted DOCX file by providing a password.
 *
 * Handles the full pipeline:
 * 1. Parse CFB container
 * 2. Extract EncryptionInfo and EncryptedPackage streams
 * 3. Verify password and decrypt
 * 4. Return the decrypted DOCX ZIP bytes
 *
 * @param buffer - The encrypted DOCX file (CFB format).
 * @param password - The password to decrypt with.
 * @param maxDecryptedSize - Optional cap on the decrypted size (defaults to 512 MiB).
 * @returns The decrypted DOCX ZIP bytes (can be passed to readDocx).
 * @throws Error if the file is not encrypted, password is wrong, or decryption fails.
 */
export async function decryptDocx(
  buffer: Uint8Array,
  password: string,
  maxDecryptedSize?: number
): Promise<Uint8Array> {
  if (!isEncryptedDocx(buffer)) {
    throw new DocxDecryptionError("Not an encrypted DOCX file (CFB signature not found)");
  }

  // Parse CFB container
  const entries = readCfb(buffer);

  // Find EncryptionInfo stream
  const _encInfoEntry = entries.find(
    e => e.name === "EncryptionInfo" || e.name === "\x06DataSpaces"
  );
  // The actual EncryptionInfo stream
  const encInfoStream = entries.find(e => e.name === "EncryptionInfo");
  if (!encInfoStream) {
    throw new DocxDecryptionError("CFB: EncryptionInfo stream not found");
  }

  // Find EncryptedPackage stream
  const encPkgEntry = entries.find(e => e.name === "EncryptedPackage");
  if (!encPkgEntry) {
    throw new DocxDecryptionError("CFB: EncryptedPackage stream not found");
  }

  // Parse EncryptionInfo — first 4 bytes are version major + minor + flags.
  const infoData = encInfoStream.data;
  const versionView = new DataView(infoData.buffer, infoData.byteOffset, 4);
  const versionMajor = versionView.getUint16(0, true);
  const versionMinor = versionView.getUint16(2, true);

  // MS-OFFCRYPTO §2.3 distinguishes encryption families primarily by the
  // (major, minor) pair: Agile is 4.4, Extensible is 4.3, ECMA-376 Standard
  // Encryption is 4.2, RC4 CryptoAPI is 4.x with minor < 2, etc. We
  // currently implement only Agile.
  if (versionMajor === 4 && versionMinor === 4) {
    // Agile encryption — XML follows after 8 bytes
    const xmlStr = utf8Decoder.decode(infoData.slice(8));
    const info = parseEncryptionInfoXml(xmlStr);

    // Verify password
    const valid = await verifyPassword(password, info);
    if (!valid) {
      throw new DocxDecryptionError("Incorrect password");
    }

    // Decrypt package
    return decryptPackage(encPkgEntry.data, info, password, maxDecryptedSize);
  }

  throw new DocxDecryptionError(
    `Unsupported encryption version: ${versionMajor}.${versionMinor}. ` +
      `Only Agile Encryption (4.4) is supported.`
  );
}

// =============================================================================
// Encryption Options & AES-CBC Encrypt
// =============================================================================

/** Options for DOCX encryption. */
export interface EncryptOptions {
  /** Key length in bits. Default: 256. */
  readonly keyBits?: 128 | 192 | 256;
  /** Hash algorithm. Default: SHA512. */
  readonly hashAlgorithm?: "SHA1" | "SHA256" | "SHA384" | "SHA512";
  /** Number of KDF iterations (spin count). Default: 100000. */
  readonly spinCount?: number;
}

/**
 * AES-CBC encryption without PKCS#7 padding (raw block encryption).
 *
 * Data MUST already be padded to a multiple of 16 bytes.
 * Used for package segment encryption where zero-padding is used.
 */
async function aesCbcEncryptNoPadding(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    toBuffer(key),
    { name: "AES-CBC", length: key.length * 8 },
    false,
    ["encrypt"]
  );
  // Ensure IV is 16 bytes
  const iv16 = new Uint8Array(16);
  iv16.set(iv.slice(0, 16));

  const result = await crypto.subtle.encrypt(
    { name: "AES-CBC", iv: toBuffer(iv16) },
    cryptoKey,
    toBuffer(data)
  );
  // Web Crypto always adds PKCS#7 padding (extra 16 bytes when input is
  // already block-aligned). Strip the trailing padding block.
  return new Uint8Array(result, 0, data.length);
}

/**
 * AES-CBC encryption with PKCS#7 padding (standard).
 *
 * Used for encrypting verifier and key values per MS-OFFCRYPTO.
 */
async function aesCbcEncrypt(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    toBuffer(key),
    { name: "AES-CBC", length: key.length * 8 },
    false,
    ["encrypt"]
  );
  // Ensure IV is 16 bytes
  const iv16 = new Uint8Array(16);
  iv16.set(iv.slice(0, 16));

  const result = await crypto.subtle.encrypt(
    { name: "AES-CBC", iv: toBuffer(iv16) },
    cryptoKey,
    toBuffer(data)
  );
  return new Uint8Array(result);
}

// =============================================================================
// Encrypt DOCX
// =============================================================================

/**
 * Encrypt a DOCX ZIP file with a password using Agile Encryption (MS-OFFCRYPTO).
 *
 * Produces a CFB (OLE2 Compound File) containing:
 *   - EncryptionInfo stream (Agile encryption XML)
 *   - EncryptedPackage stream (segment-encrypted ZIP data)
 *
 * The output can be opened by Microsoft Word with the given password.
 *
 * @param zipBytes - The unencrypted DOCX content (ZIP format).
 * @param password - The password to encrypt with.
 * @param options - Encryption options (key length, hash algorithm, spin count).
 * @returns The encrypted DOCX as a CFB-formatted Uint8Array.
 */
export async function encryptDocx(
  zipBytes: Uint8Array,
  password: string,
  options?: EncryptOptions
): Promise<Uint8Array> {
  const keyBits = options?.keyBits ?? 256;
  const hashAlgorithm = options?.hashAlgorithm ?? "SHA512";
  const spinCount = options?.spinCount ?? 100000;
  const hashName = mapHashName(hashAlgorithm);
  const hashSize = await getHashSize(hashName);
  const keyBytes = keyBits / 8;
  const blockSize = 16;

  // 1. Generate random salt (16 bytes) and package key
  const keySalt = randomBytes(16);
  const packageKey = randomBytes(keyBytes);

  // 2. Generate key encryption key (for encrypting the package key)
  const keyEncryptionKey = await deriveEncryptionKey(
    password,
    { keySalt, spinCount, hashAlgorithm, keyBits },
    AGILE_BLOCK_KEYS.encryptedKey
  );

  // 3. Encrypt the package key
  const encryptedKeyValue = await aesCbcEncrypt(packageKey, keyEncryptionKey, keySalt);

  // 4. Generate verifier: random 16 bytes, hash them, encrypt both
  const verifierInput = randomBytes(16);
  const verifierHash = await sha(hashName, verifierInput);

  const verifierInputKey = await deriveEncryptionKey(
    password,
    { keySalt, spinCount, hashAlgorithm, keyBits },
    AGILE_BLOCK_KEYS.verifierHashInput
  );
  const encryptedVerifierHashInput = await aesCbcEncrypt(verifierInput, verifierInputKey, keySalt);

  const verifierValueKey = await deriveEncryptionKey(
    password,
    { keySalt, spinCount, hashAlgorithm, keyBits },
    AGILE_BLOCK_KEYS.verifierHashValue
  );
  const encryptedVerifierHashValue = await aesCbcEncrypt(verifierHash, verifierValueKey, keySalt);

  // 5. Encrypt the ZIP data in 4096-byte segments
  const encryptedPackage = await encryptPackageData(
    zipBytes,
    packageKey,
    keySalt,
    hashName,
    blockSize
  );

  // 6. Generate EncryptionInfo XML
  const encInfoXml = buildEncryptionInfoXml({
    keyBits,
    hashAlgorithm,
    hashSize,
    spinCount,
    blockSize,
    keySalt,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue
  });

  // Prepend 8-byte version header: version 4.4 + flags 0x40
  const xmlBytes = utf8Encoder.encode(encInfoXml);
  const encInfoStream = new Uint8Array(8 + xmlBytes.length);
  const encInfoView = new DataView(encInfoStream.buffer);
  encInfoView.setUint16(0, 4, true); // version major
  encInfoView.setUint16(2, 4, true); // version minor
  encInfoView.setUint32(4, 0x40, true); // flags (agile)
  encInfoStream.set(xmlBytes, 8);

  // 7. Package into CFB
  const cfbBytes = writeCfb([
    { name: "EncryptionInfo", data: encInfoStream },
    { name: "EncryptedPackage", data: encryptedPackage }
  ]);

  return cfbBytes;
}

// =============================================================================
// Encrypt Helpers
// =============================================================================

/**
 * Encrypt the package data in 4096-byte segments.
 *
 * Format: [8 bytes: total size LE uint64] [encrypted segment 0] [encrypted segment 1] ...
 * Each segment IV = hash(salt + segment_index_LE) truncated to 16 bytes.
 */
async function encryptPackageData(
  data: Uint8Array,
  packageKey: Uint8Array,
  keySalt: Uint8Array,
  hashName: string,
  blockSize: number
): Promise<Uint8Array> {
  const segmentSize = 4096;
  const segCount = Math.ceil(data.length / segmentSize);
  const encSegments: Uint8Array[] = [];

  for (let i = 0; i < segCount; i++) {
    const segStart = i * segmentSize;
    const segEnd = Math.min(segStart + segmentSize, data.length);
    let segData = data.slice(segStart, segEnd);

    // Pad last segment to multiple of blockSize if needed
    if (segData.length % blockSize !== 0) {
      const padded = new Uint8Array(Math.ceil(segData.length / blockSize) * blockSize);
      padded.set(segData);
      segData = padded;
    }

    // Segment IV = hash(salt + segment_index_LE) truncated to blockSize
    const idxBytes = new Uint8Array(4);
    new DataView(idxBytes.buffer).setUint32(0, i, true);
    const segIv = (await sha(hashName, concat(keySalt, idxBytes))).slice(0, blockSize);

    const encSeg = await aesCbcEncryptNoPadding(segData, packageKey, segIv);
    encSegments.push(encSeg);
  }

  // Total size (8 bytes LE uint64) + concatenated encrypted segments
  const totalEncLen = encSegments.reduce((sum, s) => sum + s.length, 0);
  const result = new Uint8Array(8 + totalEncLen);
  const sizeView = new DataView(result.buffer);
  sizeView.setBigUint64(0, BigInt(data.length), true);

  let offset = 8;
  for (const seg of encSegments) {
    result.set(seg, offset);
    offset += seg.length;
  }
  return result;
}

/** Get hash output size in bytes for a given Web Crypto hash name. */
async function getHashSize(hashName: string): Promise<number> {
  const test = await sha(hashName, new Uint8Array(0));
  return test.length;
}

/** Build the Agile EncryptionInfo XML document. */
function buildEncryptionInfoXml(params: {
  keyBits: number;
  hashAlgorithm: string;
  hashSize: number;
  spinCount: number;
  blockSize: number;
  keySalt: Uint8Array;
  encryptedVerifierHashInput: Uint8Array;
  encryptedVerifierHashValue: Uint8Array;
  encryptedKeyValue: Uint8Array;
}): string {
  const {
    keyBits,
    hashAlgorithm,
    hashSize,
    spinCount,
    blockSize,
    keySalt,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue
  } = params;

  const saltB64 = bytesToBase64(keySalt);
  const vhiB64 = bytesToBase64(encryptedVerifierHashInput);
  const vhvB64 = bytesToBase64(encryptedVerifierHashValue);
  const ekvB64 = bytesToBase64(encryptedKeyValue);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
    'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ' +
    'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">\r\n' +
    `<keyData saltSize="16" blockSize="${blockSize}" keyBits="${keyBits}" ` +
    `hashSize="${hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" ` +
    `hashAlgorithm="${hashAlgorithm}" saltValue="${saltB64}"/>\r\n` +
    '<dataIntegrity encryptedHmacKey="" encryptedHmacValue=""/>\r\n' +
    "<keyEncryptors>\r\n" +
    '<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">\r\n' +
    `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="${blockSize}" ` +
    `keyBits="${keyBits}" hashSize="${hashSize}" cipherAlgorithm="AES" ` +
    `cipherChaining="ChainingModeCBC" hashAlgorithm="${hashAlgorithm}" ` +
    `saltValue="${saltB64}" ` +
    `encryptedVerifierHashInput="${vhiB64}" ` +
    `encryptedVerifierHashValue="${vhvB64}" ` +
    `encryptedKeyValue="${ekvB64}"/>\r\n` +
    "</keyEncryptor>\r\n" +
    "</keyEncryptors>\r\n" +
    "</encryption>"
  );
}
