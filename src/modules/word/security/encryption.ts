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
 * All AES and SHA primitives are delegated to `@utils/crypto`, which uses
 * synchronous `node:crypto` APIs in Node and Web Crypto in the browser.
 * The KDF inner loop runs ~100,000 hashes per derived key; on Node we go
 * through the synchronous fast path so a round-trip completes in well
 * under a second instead of the multi-second microtask churn that the
 * old `await crypto.subtle.digest()` loop produced.
 *
 * References:
 *   - MS-OFFCRYPTO: Office Document Cryptography Structure
 *   - ECMA-376 Part 3: Markup Compatibility and Extensibility
 */

import {
  aesCbcDecryptRaw as aesCbcDecryptRawSync,
  aesCbcEncryptRaw as aesCbcEncryptRawSync,
  hash as hashSyncMaybe,
  hashAsync
} from "@utils/crypto";
import {
  base64ToBytes,
  bytesToBase64,
  randomBytes,
  utf8Decoder,
  utf8Encoder
} from "@word/core/internal-utils";
import { DocxDecryptionError } from "@word/errors";
import { readCfb, writeCfb } from "@word/security/cfb-reader";
import type { CfbEntry } from "@word/security/cfb-reader";

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
 * Detect whether the platform's `@utils/crypto.hash` supports the requested
 * algorithm synchronously. Node's implementation accepts every OpenSSL
 * digest (SHA-1/256/384/512, MD5); the browser implementation only ships
 * synchronous SHA-256 and MD5 — anything else throws.
 *
 * The detection result is cached per algorithm so we only pay the probe
 * cost once.
 */
const _syncHashCache = new Map<string, boolean>();
function canHashSync(algorithm: string): boolean {
  const cached = _syncHashCache.get(algorithm);
  if (cached !== undefined) {
    return cached;
  }
  let ok = false;
  try {
    hashSyncMaybe(algorithm, new Uint8Array(0));
    ok = true;
  } catch {
    ok = false;
  }
  _syncHashCache.set(algorithm, ok);
  return ok;
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
  const hashName = mapHashName(info.hashAlgorithm);

  // Reusable 4-byte iterator + scratch buffer to avoid per-iteration
  // allocations inside the spin loop. The scratch holds [iter || H_{i-1}]
  // and is rebuilt in place.
  const iterAndPrev = new Uint8Array(4 + getHashSizeFor(hashName));
  const iterView = new DataView(iterAndPrev.buffer, 0, 4);

  let h: Uint8Array;
  if (canHashSync(hashName)) {
    // Fast path: ~100,000 synchronous digests, no microtask churn.
    h = hashSyncMaybe(hashName, concat(info.keySalt, pwdBytes));
    for (let i = 0; i < info.spinCount; i++) {
      iterView.setUint32(0, i, true);
      iterAndPrev.set(h, 4);
      h = hashSyncMaybe(hashName, iterAndPrev);
    }
    h = hashSyncMaybe(hashName, concat(h, blockKey));
  } else {
    // Browser fallback: Web Crypto SHA-1/384/512 are async-only.
    h = await hashAsync(hashName, concat(info.keySalt, pwdBytes));
    for (let i = 0; i < info.spinCount; i++) {
      iterView.setUint32(0, i, true);
      iterAndPrev.set(h, 4);
      h = await hashAsync(hashName, iterAndPrev);
    }
    h = await hashAsync(hashName, concat(h, blockKey));
  }

  // Truncate to keySize. The hash size MUST be at least the requested
  // key length — otherwise we'd hand the AES layer a short key buffer
  // that either fails import or produces a key the counterparty can't
  // validate. Reject misconfigured EncryptionInfo up-front with a clear
  // error.
  const keyBytes = info.keyBits / 8;
  if (h.length < keyBytes) {
    throw new Error(
      `deriveEncryptionKey: hash output of ${h.length} bytes is too ` +
        `short for keyBits=${info.keyBits} (need ${keyBytes}). ` +
        `Use a hash algorithm with a larger digest size (e.g. SHA-512 ` +
        `for keyBits ≤ 512).`
    );
  }
  return h.slice(0, keyBytes);
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

    // Decrypt the verifier hash input. MS-OFFCRYPTO §2.3.4.13 specifies
    // these blobs are AES-CBC with the plaintext zero-padded to the block
    // boundary — NOT PKCS#7. Using a PKCS#7 decrypt would mis-strip bytes
    // and break interop with Word / msoffcrypto.
    const verifierInput = aesCbcRawDecrypt(
      info.encryptedVerifierHashInput,
      verifierInputKey,
      info.keySalt
    );

    // Hash the verifier input. The plaintext is exactly 16 bytes (saltSize),
    // so trim any zero padding before hashing.
    const hashAlg = mapHashName(info.hashAlgorithm);
    const computedHash = await hashAsync(hashAlg, verifierInput.slice(0, info.blockSize));

    // Derive verifier hash value key
    const verifierValueKey = await deriveEncryptionKey(
      password,
      info,
      AGILE_BLOCK_KEYS.verifierHashValue
    );

    // Decrypt the verifier hash value (zero-padded AES-CBC, see above).
    const expectedHash = aesCbcRawDecrypt(
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

  // Decrypt the actual package key. Zero-padded AES-CBC per MS-OFFCRYPTO
  // §2.3.4.13 (NOT PKCS#7); the key is exactly keyBits/8 bytes.
  const packageKey = aesCbcRawDecrypt(info.encryptedKeyValue, keyEncryptionKey, info.keySalt).slice(
    0,
    info.keyBits / 8
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
  const hashName = mapHashName(info.hashAlgorithm);
  const idxBytes = new Uint8Array(4);
  const idxView = new DataView(idxBytes.buffer);

  for (let i = 0; i < segCount; i++) {
    const segStart = i * segmentSize;
    const segEnd = Math.min(segStart + segmentSize, encData.length);
    const segData = encData.slice(segStart, segEnd);

    // Segment IV is hash(salt + segment_index_LE), truncated to block size.
    idxView.setUint32(0, i, true);
    const segIv = (await hashAsync(hashName, concat(info.keySalt, idxBytes))).slice(
      0,
      info.blockSize
    );

    // Package segments use zero padding (MS-OFFCRYPTO §2.3.4.15). Use the
    // raw (no-padding) AES-CBC decrypt — never the PKCS#7 path, which
    // would silently mis-strip valid-looking trailing zeros.
    segments.push(aesCbcDecryptRawSync(segData, packageKey, ivToBlockSize(segIv)));
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
// Internal Crypto Helpers
//
// All AES and SHA work is delegated to `@utils/crypto`. Node uses
// synchronous `node:crypto`, the browser variant uses Web Crypto / pure
// JS — same API on both sides.
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

/** Static lookup so the KDF spin loop doesn't pay a per-call cost. */
function getHashSizeFor(hashName: string): number {
  switch (hashName) {
    case "SHA-1":
      return 20;
    case "SHA-256":
      return 32;
    case "SHA-384":
      return 48;
    case "SHA-512":
      return 64;
    default:
      // Conservative upper bound: any reasonable hash output fits in 64 bytes.
      // Callers that go through this path are misconfigured anyway and will
      // be rejected by the keyBits sanity check after the first digest.
      return 64;
  }
}

/**
 * Decrypt an AES-CBC blob written with zero padding (no PKCS#7). Used for
 * the encryptedKeyValue / encryptedVerifierHashInput / encryptedVerifierHashValue
 * blobs, per MS-OFFCRYPTO §2.3.4.13. The IV is truncated/extended to 16 bytes.
 */
function aesCbcRawDecrypt(data: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  return aesCbcDecryptRawSync(data, key, ivToBlockSize(iv));
}

/**
 * Encrypt with AES-CBC and zero-padding (no PKCS#7). Used by package
 * segment encryption and the verifier / key blobs: data is already padded
 * to a 16-byte boundary by the caller, and the on-disk format does not
 * include a PKCS#7 trailer.
 */
function aesCbcZeroPadEncrypt(data: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  return aesCbcEncryptRawSync(data, key, ivToBlockSize(iv));
}

/** Truncate or right-pad an IV to the AES block size (16 bytes). */
function ivToBlockSize(iv: Uint8Array): Uint8Array {
  if (iv.length === 16) {
    return iv;
  }
  const out = new Uint8Array(16);
  out.set(iv.slice(0, 16));
  return out;
}

/** Right-pad data with zeros so its length is a multiple of `blockSize`. */
function padToBlock(data: Uint8Array, blockSize: number): Uint8Array {
  if (data.length % blockSize === 0) {
    return data;
  }
  const out = new Uint8Array(Math.ceil(data.length / blockSize) * blockSize);
  out.set(data);
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const result = new Uint8Array(a.length + b.length);
  result.set(a, 0);
  result.set(b, a.length);
  return result;
}

/** HMAC block size (in bytes) for the supported hash algorithms. */
function hmacBlockSize(hashName: string): number {
  // SHA-1 / SHA-256 use a 512-bit (64-byte) block; SHA-384 / SHA-512 use a
  // 1024-bit (128-byte) block.
  return hashName === "SHA-384" || hashName === "SHA-512" ? 128 : 64;
}

/**
 * Compute HMAC(hashName, key, message) using the generic hash primitive.
 * Implemented here (rather than in @utils/crypto, which only ships
 * hmacSha256) so agile encryption can use SHA-512 etc. for the data
 * integrity HMAC that Word verifies on open.
 */
async function hmac(hashName: string, key: Uint8Array, message: Uint8Array): Promise<Uint8Array> {
  const blockSize = hmacBlockSize(hashName);
  // Keys longer than the block size are hashed down first.
  let k = key.length > blockSize ? await hashAsync(hashName, key) : key;
  if (k.length < blockSize) {
    const padded = new Uint8Array(blockSize);
    padded.set(k);
    k = padded;
  }
  const ipad = new Uint8Array(blockSize);
  const opad = new Uint8Array(blockSize);
  for (let i = 0; i < blockSize; i++) {
    ipad[i] = k[i] ^ 0x36;
    opad[i] = k[i] ^ 0x5c;
  }
  const inner = await hashAsync(hashName, concat(ipad, message));
  return hashAsync(hashName, concat(opad, inner));
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
  // Locate the `<keyData ... />` and `<p:encryptedKey ... />` elements
  // with a linear scan. Using regular expressions with lazy `[\s\S]*?`
  // quantifiers triggers CodeQL's polynomial-regex warning because the
  // input is attacker-controlled (a hostile EncryptionInfo XML stream
  // with very long unterminated tags caused catastrophic backtracking).
  const keyDataAttrs = extractSelfClosingTagAttrs(xmlStr, "keyData");
  const pwdAttrs = extractSelfClosingTagAttrs(xmlStr, "p:encryptedKey");

  if (!keyDataAttrs || !pwdAttrs) {
    throw new DocxDecryptionError("Invalid EncryptionInfo XML - missing keyData or encryptedKey");
  }

  const pwdData = pwdAttrs;

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

/**
 * Find a `<tagName ... />` element and return its parsed attributes, or
 * `null` if no such self-closing element exists.
 *
 * Uses a linear scan instead of a regex with `[\s\S]*?` to avoid
 * catastrophic backtracking on adversarial EncryptionInfo XML.
 */
function extractSelfClosingTagAttrs(xml: string, tagName: string): Record<string, string> | null {
  const needle = `<${tagName}`;
  let from = 0;
  while (from <= xml.length) {
    const start = xml.indexOf(needle, from);
    if (start < 0) {
      return null;
    }
    const after = start + needle.length;
    const ch = xml.charCodeAt(after);
    // Require a whitespace, '/' or '>' after the tag name so `<keyDataExtra`
    // does not match `<keyData`.
    if (ch !== 0x20 && ch !== 0x09 && ch !== 0x0a && ch !== 0x0d && ch !== 0x2f && ch !== 0x3e) {
      from = after;
      continue;
    }
    // Find the closing '>' from `start`. Bail out if there isn't one.
    const close = xml.indexOf(">", after);
    if (close < 0) {
      return null;
    }
    // The element must be self-closing: the char before '>' is '/'.
    if (xml.charCodeAt(close - 1) !== 0x2f) {
      from = close + 1;
      continue;
    }
    const inner = xml.slice(after, close - 1);
    return parseAttrs(inner);
  }
  return null;
}

/**
 * Parse XML-style attributes (`name="value"`) from a fragment. Implemented
 * as a single linear scan rather than a global regex so attacker-controlled
 * input cannot trigger polynomial-time backtracking (CodeQL js/polynomial-redos).
 */
function parseAttrs(str: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const n = str.length;
  let i = 0;
  while (i < n) {
    // Skip whitespace.
    while (i < n) {
      const c = str.charCodeAt(i);
      if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
        break;
      }
      i++;
    }
    if (i >= n) {
      break;
    }
    // Read attribute name (\w+ equivalent: [A-Za-z0-9_]+).
    const nameStart = i;
    while (i < n) {
      const c = str.charCodeAt(i);
      const isWord =
        (c >= 0x30 && c <= 0x39) || // 0-9
        (c >= 0x41 && c <= 0x5a) || // A-Z
        (c >= 0x61 && c <= 0x7a) || // a-z
        c === 0x5f; // _
      if (!isWord) {
        break;
      }
      i++;
    }
    if (i === nameStart) {
      // Not at an attribute — advance one char so we make progress.
      i++;
      continue;
    }
    const name = str.slice(nameStart, i);
    // Expect `="` exactly. Anything else means we resync to the next
    // whitespace and try again — robust to malformed input.
    if (i + 1 >= n || str.charCodeAt(i) !== 0x3d || str.charCodeAt(i + 1) !== 0x22) {
      // Skip to next whitespace.
      while (i < n) {
        const c = str.charCodeAt(i);
        if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) {
          break;
        }
        i++;
      }
      continue;
    }
    i += 2; // past `="`
    // Read until next `"`.
    const valStart = i;
    const valEnd = str.indexOf('"', i);
    if (valEnd < 0) {
      // Unterminated value — store what we have and stop.
      attrs[name] = str.slice(valStart);
      break;
    }
    attrs[name] = str.slice(valStart, valEnd);
    i = valEnd + 1;
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
  const hashSize = getHashSizeFor(hashName);
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

  // 3. Encrypt the package key. MS-OFFCRYPTO §2.3.4.13: these blobs use
  //    zero-padded AES-CBC (no PKCS#7). packageKey is already block-aligned.
  const encryptedKeyValue = aesCbcZeroPadEncrypt(
    padToBlock(packageKey, blockSize),
    keyEncryptionKey,
    keySalt
  );

  // 4. Generate verifier: random 16 bytes, hash them, encrypt both
  const verifierInput = randomBytes(16);
  const verifierHash = await hashAsync(hashName, verifierInput);

  const verifierInputKey = await deriveEncryptionKey(
    password,
    { keySalt, spinCount, hashAlgorithm, keyBits },
    AGILE_BLOCK_KEYS.verifierHashInput
  );
  const encryptedVerifierHashInput = aesCbcZeroPadEncrypt(
    padToBlock(verifierInput, blockSize),
    verifierInputKey,
    keySalt
  );

  const verifierValueKey = await deriveEncryptionKey(
    password,
    { keySalt, spinCount, hashAlgorithm, keyBits },
    AGILE_BLOCK_KEYS.verifierHashValue
  );
  const encryptedVerifierHashValue = aesCbcZeroPadEncrypt(
    padToBlock(verifierHash, blockSize),
    verifierValueKey,
    keySalt
  );

  // 5. Encrypt the ZIP data in 4096-byte segments
  const encryptedPackage = await encryptPackageData(
    zipBytes,
    packageKey,
    keySalt,
    hashName,
    blockSize
  );

  // 5b. Data integrity (MS-OFFCRYPTO §2.3.4.14). Word verifies this HMAC on
  //     open; an empty or missing value makes it report the file as corrupt.
  //       - hmacKey: random, hashSize bytes (padded to block size).
  //       - encryptedHmacKey  = AES-CBC(packageKey, IV0, hmacKey)
  //       - hmacValue = HMAC(hashAlg, hmacKey, encryptedPackage)  (entire
  //         stream including the 8-byte size prefix)
  //       - encryptedHmacValue = AES-CBC(packageKey, IV1, hmacValue)
  //     IV0 = H(keySalt + blockKeyDataIntegrityKey)[:blockSize]
  //     IV1 = H(keySalt + blockKeyDataIntegrityValue)[:blockSize]
  const hmacKey = randomBytes(hashSize);
  const ivHmacKey = (
    await hashAsync(hashName, concat(keySalt, AGILE_BLOCK_KEYS.dataIntegrityKey))
  ).slice(0, blockSize);
  const ivHmacValue = (
    await hashAsync(hashName, concat(keySalt, AGILE_BLOCK_KEYS.dataIntegrityValue))
  ).slice(0, blockSize);

  const encryptedHmacKey = aesCbcZeroPadEncrypt(
    padToBlock(hmacKey, blockSize),
    packageKey,
    ivHmacKey
  );
  const hmacValue = await hmac(hashName, hmacKey, encryptedPackage);
  const encryptedHmacValue = aesCbcZeroPadEncrypt(
    padToBlock(hmacValue, blockSize),
    packageKey,
    ivHmacValue
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
    encryptedKeyValue,
    encryptedHmacKey,
    encryptedHmacValue
  });

  // Prepend 8-byte version header: version 4.4 + flags 0x40
  const xmlBytes = utf8Encoder.encode(encInfoXml);
  const encInfoStream = new Uint8Array(8 + xmlBytes.length);
  const encInfoView = new DataView(encInfoStream.buffer);
  encInfoView.setUint16(0, 4, true); // version major
  encInfoView.setUint16(2, 4, true); // version minor
  encInfoView.setUint32(4, 0x40, true); // flags (agile)
  encInfoStream.set(xmlBytes, 8);

  // 7. Package into CFB.
  //    Office requires the \x06DataSpaces structure (MS-OFFCRYPTO §2.3.2)
  //    in addition to EncryptionInfo + EncryptedPackage; without it Word
  //    rejects the file as corrupt even when the password is correct.
  const cfbBytes = writeCfb([
    ...buildDataSpacesStreams(),
    { name: "EncryptionInfo", data: encInfoStream },
    { name: "EncryptedPackage", data: encryptedPackage }
  ]);

  return cfbBytes;
}

// =============================================================================
// Encrypt Helpers
// =============================================================================

// -----------------------------------------------------------------------------
// \x06DataSpaces structure (MS-OFFCRYPTO §2.3.2)
//
// Office encrypted documents wrap the EncryptedPackage in a DataSpaces map so
// the consumer knows which transform (StrongEncryption) was applied. The four
// streams below are byte-for-byte what Office writes for password-based agile
// encryption. Word validates this structure on open; omitting it makes the
// file "corrupt" even with the correct password.
// -----------------------------------------------------------------------------

/** Encode a UTF-8 length-prefixed unicode string (UNICODE-LP-P4):
 *  [4-byte LE byte length of UTF-16LE payload][UTF-16LE chars][pad to 4-byte boundary]. */
function lengthPrefixedUtf16(str: string): Uint8Array {
  const chars = stringToUtf16LE(str);
  const padded = Math.ceil(chars.length / 4) * 4;
  const out = new Uint8Array(4 + padded);
  new DataView(out.buffer).setUint32(0, chars.length, true);
  out.set(chars, 4);
  return out;
}

/** Concatenate several byte arrays. */
function concatAll(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** Build the four \x06DataSpaces streams Office requires for agile encryption. */
function buildDataSpacesStreams(): CfbEntry[] {
  const DATASPACES = "\u0006DataSpaces";

  // --- Version stream (DataSpaceVersionInfo) ---
  // FeatureIdentifier "Microsoft.Container.DataSpaces" + reader/updater/writer
  // version (each major=1, minor=0).
  const versionStream = concatAll(
    lengthPrefixedUtf16("Microsoft.Container.DataSpaces"),
    u16le(1),
    u16le(0), // reader version
    u16le(1),
    u16le(0), // updater version
    u16le(1),
    u16le(0) // writer version
  );

  // --- DataSpaceMap stream ---
  // Header: HeaderLength(8) + EntryCount(1) followed by one MapEntry.
  // MapEntry: EntryLength + ReferenceComponentCount(1) +
  //           [ReferenceComponent: type(0=stream) + LP name "EncryptedPackage"] +
  //           LP DataSpaceName "StrongEncryptionDataSpace".
  const refName = lengthPrefixedUtf16("EncryptedPackage");
  const dsName = lengthPrefixedUtf16("StrongEncryptionDataSpace");
  const refComponent = concatAll(u32le(0), refName); // 0 = stream component
  const mapEntryBody = concatAll(u32le(1), refComponent, dsName); // 1 reference component
  const entryLength = 4 + mapEntryBody.length; // include the EntryLength field itself
  const mapEntry = concatAll(u32le(entryLength), mapEntryBody);
  const dataSpaceMap = concatAll(u32le(8), u32le(1), mapEntry); // headerLen=8, entryCount=1

  // --- DataSpaceInfo/StrongEncryptionDataSpace stream (DataSpaceDefinition) ---
  // HeaderLength(8) + TransformReferenceCount(1) + LP transform name.
  const transformName = lengthPrefixedUtf16("StrongEncryptionTransform");
  const dataSpaceDefinition = concatAll(u32le(8), u32le(1), transformName);

  // --- TransformInfo/StrongEncryptionTransform/\x06Primary stream ---
  // TransformInfoHeader:
  //   TransformLength + TransformType(1) + LP TransformId +
  //   LP TransformName + reader/updater/writer versions.
  // Followed by EncryptionTransformInfo:
  //   LP EncryptionName + EncryptionBlockSize(4) + CipherMode(4).
  const transformId = lengthPrefixedUtf16("{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
  const transformNamePrimary = lengthPrefixedUtf16("Microsoft.Container.EncryptionTransform");
  const headerBody = concatAll(
    u32le(1), // TransformType = 1
    transformId,
    transformNamePrimary,
    u16le(1),
    u16le(0), // reader version
    u16le(1),
    u16le(0), // updater version
    u16le(1),
    u16le(0) // writer version
  );
  const transformLength = 4 + headerBody.length; // include the length field itself
  const transformHeader = concatAll(u32le(transformLength), headerBody);
  const encryptionTransformInfo = concatAll(
    lengthPrefixedUtf16(""), // EncryptionName (empty for agile)
    u32le(0), // EncryptionBlockSize
    u32le(0) // CipherMode
  );
  const primary = concatAll(transformHeader, encryptionTransformInfo);

  return [
    { name: "Version", path: [DATASPACES], data: versionStream },
    { name: "DataSpaceMap", path: [DATASPACES], data: dataSpaceMap },
    {
      name: "StrongEncryptionDataSpace",
      path: [DATASPACES, "DataSpaceInfo"],
      data: dataSpaceDefinition
    },
    {
      name: "\u0006Primary",
      path: [DATASPACES, "TransformInfo", "StrongEncryptionTransform"],
      data: primary
    }
  ];
}

function u16le(n: number): Uint8Array {
  const b = new Uint8Array(2);
  new DataView(b.buffer).setUint16(0, n, true);
  return b;
}

function u32le(n: number): Uint8Array {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
}

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
  const idxBytes = new Uint8Array(4);
  const idxView = new DataView(idxBytes.buffer);

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
    idxView.setUint32(0, i, true);
    const segIv = (await hashAsync(hashName, concat(keySalt, idxBytes))).slice(0, blockSize);

    encSegments.push(aesCbcZeroPadEncrypt(segData, packageKey, segIv));
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
  encryptedHmacKey: Uint8Array;
  encryptedHmacValue: Uint8Array;
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
    encryptedKeyValue,
    encryptedHmacKey,
    encryptedHmacValue
  } = params;

  const saltB64 = bytesToBase64(keySalt);
  const vhiB64 = bytesToBase64(encryptedVerifierHashInput);
  const vhvB64 = bytesToBase64(encryptedVerifierHashValue);
  const ekvB64 = bytesToBase64(encryptedKeyValue);
  const hmacKeyB64 = bytesToBase64(encryptedHmacKey);
  const hmacValB64 = bytesToBase64(encryptedHmacValue);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
    'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ' +
    'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">\r\n' +
    `<keyData saltSize="16" blockSize="${blockSize}" keyBits="${keyBits}" ` +
    `hashSize="${hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" ` +
    `hashAlgorithm="${hashAlgorithm}" saltValue="${saltB64}"/>\r\n` +
    `<dataIntegrity encryptedHmacKey="${hmacKeyB64}" encryptedHmacValue="${hmacValB64}"/>\r\n` +
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
