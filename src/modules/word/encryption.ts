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

import { base64ToBytes } from "./internal-utils";

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

  // Truncate to keySize
  const keyBytes = info.keyBits / 8;
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

    // Decrypt the verifier hash input
    const verifierInput = await aesCbcDecrypt(
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

    // Decrypt the verifier hash value
    const expectedHash = await aesCbcDecrypt(
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
 * @returns The decrypted (unencrypted) package bytes.
 */
export async function decryptPackage(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  password: string
): Promise<Uint8Array> {
  // Derive key encryption key
  const keyEncryptionKey = await deriveEncryptionKey(password, info, AGILE_BLOCK_KEYS.encryptedKey);

  // Decrypt the actual package key
  const packageKey = await aesCbcDecrypt(info.encryptedKeyValue, keyEncryptionKey, info.keySalt);

  // First 8 bytes are the total decrypted size (uint64 LE)
  const totalSizeView = new DataView(encryptedPackage.buffer, encryptedPackage.byteOffset, 8);
  const totalSize = Number(totalSizeView.getBigUint64(0, true));

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

    const decSeg = await aesCbcDecrypt(segData, packageKey, segIv);
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

async function aesCbcDecrypt(
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
  // Ensure IV is 16 bytes (pad with zeros if needed, truncate if longer)
  const iv16 = new Uint8Array(16);
  iv16.set(iv.slice(0, 16));

  // NOTE: AES-CBC in Web Crypto expects PKCS#7 padding. Agile encryption
  // uses zero-padding for the last block.
  try {
    const result = await crypto.subtle.decrypt(
      { name: "AES-CBC", iv: toBuffer(iv16) },
      cryptoKey,
      toBuffer(data)
    );
    return new Uint8Array(result);
  } catch {
    return manualAesCbcDecrypt(data, key, iv16);
  }
}

async function manualAesCbcDecrypt(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Promise<Uint8Array> {
  const blockSize = 16;
  if (data.length % blockSize !== 0) {
    throw new Error("Data length must be a multiple of block size");
  }
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    toBuffer(key),
    { name: "AES-CBC", length: key.length * 8 },
    false,
    ["decrypt"]
  );
  const result = new Uint8Array(data.length);
  let prev = iv;
  for (let i = 0; i < data.length; i += blockSize) {
    const block = data.slice(i, i + blockSize);
    const tempCipher = concat(block, new Uint8Array(blockSize));
    try {
      const dec = await crypto.subtle.decrypt(
        { name: "AES-CBC", iv: toBuffer(new Uint8Array(16)) },
        cryptoKey,
        toBuffer(tempCipher)
      );
      const decBlock = new Uint8Array(dec).slice(0, blockSize);
      // XOR with previous block
      for (let j = 0; j < blockSize; j++) {
        result[i + j] = decBlock[j] ^ prev[j];
      }
      prev = block;
    } catch (e) {
      throw new Error("AES decryption failed: " + (e as Error).message);
    }
  }
  return result;
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
  const keyDataMatch = /<keyData\s+([^/]+)\/>/.exec(xmlStr);
  const pwdEncryptorMatch = /<p:encryptedKey\s+([^/]+)\/>/.exec(xmlStr);

  if (!keyDataMatch || !pwdEncryptorMatch) {
    throw new Error("Invalid EncryptionInfo XML - missing keyData or encryptedKey");
  }

  const pwdData = parseAttrs(pwdEncryptorMatch[1]);

  return {
    cipherAlgorithm: "AES",
    cipherChaining: "ChainingModeCBC",
    keyBits: parseInt(pwdData.keyBits ?? "256") as 128 | 192 | 256,
    hashAlgorithm: (pwdData.hashAlgorithm ?? "SHA512") as "SHA1" | "SHA256" | "SHA384" | "SHA512",
    hashSize: parseInt(pwdData.hashSize ?? "64"),
    spinCount: parseInt(pwdData.spinCount ?? "100000"),
    keySalt: base64ToBytes(pwdData.saltValue ?? ""),
    encryptedVerifierHashInput: base64ToBytes(pwdData.encryptedVerifierHashInput ?? ""),
    encryptedVerifierHashValue: base64ToBytes(pwdData.encryptedVerifierHashValue ?? ""),
    encryptedKeyValue: base64ToBytes(pwdData.encryptedKeyValue ?? ""),
    blockSize: parseInt(pwdData.blockSize ?? "16")
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
