/**
 * Office file encryption — ECMA-376 Agile Encryption (MS-OFFCRYPTO §2.3.4.10–2.3.4.15), what Word, Excel and
 * PowerPoint write for "Encrypt with Password": the OOXML package (a ZIP) encrypted with AES in 4096-byte segments,
 * stored as the `EncryptedPackage` stream of a CFB container beside its `EncryptionInfo` and the `\x06DataSpaces`
 * structure Office requires.
 *
 * One implementation for every module that reads or writes an Office package: Word (`Document.read`/`Document.toBuffer` and the
 * published `WordCrypto`) and Excel (`Workbook.read`/`toBuffer` and their siblings). Internal — each module checks its
 * own arguments, loads this module only when a password is involved, and reports {@link OfficeCryptoError} and
 * `CfbError` as its own errors.
 *
 * Only Agile Encryption (version 4.4) is implemented. Standard Encryption (3.2/4.2, Office 2007) and the RC4
 * schemes of the binary formats are refused with an `OfficeCryptoError` naming the scheme.
 *
 * Two salts take part, as the specification has it: the password key encryptor's (`p:encryptedKey/@saltValue`)
 * derives the keys that unwrap the verifier and the package key, and the key data's (`keyData/@saltValue`) makes
 * the IVs of the package segments and of the data-integrity HMAC. Office writes different values for the two.
 *
 * AES and SHA come from `@utils/crypto` — synchronous `node:crypto` on Node, Web Crypto / pure JS in a browser.
 */

import { CfbError, isCfb, readCfb, writeCfb } from "@archive/cfb/cfb";
import type { CfbEntry } from "@archive/cfb/cfb";
import { ArchiveError } from "@archive/core/errors";
import { textDecoder, textEncoder } from "@utils/binary";
import {
  aesCbcDecryptRaw as aesCbcDecryptRawSync,
  aesCbcEncryptRaw as aesCbcEncryptRawSync,
  hash,
  randomBytes
} from "@utils/crypto";
import { canonicalHashAlgorithm, compactHashAlgorithm } from "@utils/hash-names";
import type { CompactHashAlgorithm } from "@utils/hash-names";
import { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils";
import { createYieldBudget, PACED_SLICE_LENGTH } from "@utils/yield";
import type { YieldPoint } from "@utils/yield";

/**
 * Why an encrypted package could not be decrypted (or a key derived): the password is wrong; the file uses an
 * encryption this library does not implement; its encryption metadata or package is malformed; its data-integrity
 * check failed (the file was altered or damaged); or it is larger than the read allows.
 */
export type OfficeCryptoFailure = "password" | "unsupported" | "malformed" | "integrity" | "limit";

/** An encrypted Office package that cannot be decrypted; `reason` says why. */
export class OfficeCryptoError extends ArchiveError {
  override name = "OfficeCryptoError";
  constructor(
    readonly reason: OfficeCryptoFailure,
    message: string
  ) {
    super(message);
  }
}

export { CfbError };

/** Agile encryption parameters. */
export interface AgileEncryptionInfo {
  /** Cipher algorithm (AES). */
  readonly cipherAlgorithm: "AES";
  /** Cipher chaining mode. */
  readonly cipherChaining: "ChainingModeCBC";
  /** Key bit length (128, 192, 256). */
  readonly keyBits: 128 | 192 | 256;
  /** The password key encryptor's hash algorithm. */
  readonly hashAlgorithm: "SHA1" | "SHA256" | "SHA384" | "SHA512";
  /** Hash output size in bytes. */
  readonly hashSize: number;
  /** Number of iterations (spin count). */
  readonly spinCount: number;
  /** The password key encryptor's salt (`p:encryptedKey/@saltValue`), 16 bytes: key derivation, and the IV of its blobs. */
  readonly keySalt: Uint8Array;
  /** Encrypted verifier hash input. */
  readonly encryptedVerifierHashInput: Uint8Array;
  /** Encrypted verifier hash value. */
  readonly encryptedVerifierHashValue: Uint8Array;
  /** Encrypted key value. */
  readonly encryptedKeyValue: Uint8Array;
  /** Block size (16 for AES). */
  readonly blockSize: number;
  /**
   * The data-integrity HMAC key, as stored: `dataIntegrity/@encryptedHmacKey`, encrypted with the package key.
   * Present on a file that carries a data-integrity check (Office always writes one); decryption verifies it.
   */
  readonly dataIntegritySalt?: Uint8Array;
  /** The data-integrity HMAC, as stored: `dataIntegrity/@encryptedHmacValue`, encrypted with the package key. */
  readonly dataIntegrityHmac?: Uint8Array;
  /**
   * The `keyData` element's parameters — the package key's salt, hash and block size, which the package segments
   * and the data-integrity check use. Absent on info built by hand: the key encryptor's (`keySalt`,
   * `hashAlgorithm`, `blockSize`) are used, which is what a file holds when both carry the same values.
   */
  readonly keyData?: {
    readonly salt: Uint8Array;
    readonly hashAlgorithm: "SHA1" | "SHA256" | "SHA384" | "SHA512";
    readonly hashSize: number;
    readonly blockSize: number;
    readonly keyBits: 128 | 192 | 256;
  };
}

/** Options for encrypting a package (`WordCrypto.encrypt`, and the `encryption` write option of Word and Excel). */
export interface EncryptOptions {
  /** Key length in bits. Default: 256. */
  readonly keyBits?: 128 | 192 | 256;
  /**
   * Hash algorithm. Default: `"SHA-512"`. Write it `"SHA-256"`, the spelling Word's `Security.protect` takes and
   * `Security.getState` reports; the encryption format's own spelling (`"SHA256"`) is accepted too, and is what the
   * file holds either way.
   */
  readonly hashAlgorithm?: "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512" | CompactHashAlgorithm;
  /** Number of KDF iterations (spin count), 0 to 10,000,000. Default: 100000, as Office writes. */
  readonly spinCount?: number;
}

/** Whether bytes start with the CFB signature — the container every encrypted Office package is stored in. */
export { isCfb };

/** What {@link deriveAgileKey} derives a key from. */
interface KeyDerivationInfo {
  readonly keySalt: Uint8Array;
  readonly spinCount: number;
  readonly hashAlgorithm: string;
  readonly keyBits: number;
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
export async function deriveAgileKey(
  password: string,
  info: KeyDerivationInfo,
  blockKey: Uint8Array
): Promise<Uint8Array> {
  const owned = ownKeyInfo(info);
  const key = ownBytes(blockKey);
  return agileKeyFromSpun(await spinAgilePassword(password, owned), owned, key);
}

/**
 * `H_spinCount` of {@link deriveAgileKey}: the part of the derivation every block key shares, so a caller deriving
 * several keys from one password — the verifier's two and the package key's — spins once rather than once per key.
 *
 * The spin is the cost of the derivation, 100,000 digests by default, and ran as one synchronous loop: it blocked the
 * event loop for the whole of it, three times over per encrypted file read or written. It yields once it has worked for
 * the budget, as every long loop in this package does.
 */
async function spinAgilePassword(
  password: string,
  info: { keySalt: Uint8Array; spinCount: number; hashAlgorithm: string },
  yieldIfDue: YieldPoint = createYieldBudget()
): Promise<Uint8Array> {
  const pwdBytes = stringToUtf16LE(password);
  const hashName = mapHashName(info.hashAlgorithm);

  // Reusable 4-byte iterator + scratch buffer to avoid per-iteration
  // allocations inside the spin loop. The scratch holds [iter || H_{i-1}]
  // and is rebuilt in place.
  const iterAndPrev = new Uint8Array(4 + getHashSizeFor(hashName));
  const iterView = new DataView(iterAndPrev.buffer, 0, 4);

  // Synchronous digests — every algorithm the format names has one on every platform (see `@utils/crypto`'s `hash`),
  // which in a browser is five times faster than awaiting one Web Crypto digest per step. The clock is read once per
  // batch of them. `yieldIfDue` is the operation's: work done before the spin counts against the same budget.
  let h = hash(hashName, concat(info.keySalt, pwdBytes));
  for (let i = 0; i < info.spinCount; i++) {
    iterView.setUint32(0, i, true);
    iterAndPrev.set(h, 4);
    h = hash(hashName, iterAndPrev);
    if ((i & SPIN_BATCH_MASK) === SPIN_BATCH_MASK) {
      await yieldIfDue();
    }
  }
  return h;
}

/** Digests the spin loop runs between two looks at the clock, less one: a mask, as the loop counts in powers of two. */
const SPIN_BATCH_MASK = 1023;

/** The key {@link deriveAgileKey} derives for `blockKey`, from the password's {@link spinAgilePassword}. */
function agileKeyFromSpun(
  spun: Uint8Array,
  info: { hashAlgorithm: string; keyBits: number },
  blockKey: Uint8Array
): Uint8Array {
  const h = hash(mapHashName(info.hashAlgorithm), concat(spun, blockKey));

  // Truncate to keySize. The hash size MUST be at least the requested
  // key length — otherwise we'd hand the AES layer a short key buffer
  // that either fails import or produces a key the counterparty can't
  // validate. Reject misconfigured EncryptionInfo up-front with a clear
  // error.
  const keyBytes = info.keyBits / 8;
  if (h.length < keyBytes) {
    throw new OfficeCryptoError(
      "malformed",
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
 * Whether `password` opens a package with this encryption info, checked against its password verifier
 * (MS-OFFCRYPTO §2.3.4.13) without decrypting the package.
 *
 * @throws {OfficeCryptoError} `reason: "malformed"` when the verifier is not whole AES blocks: a damaged file is not a
 * wrong password, and every error used to be answered `false`.
 */
export async function verifyAgilePassword(
  password: string,
  info: AgileEncryptionInfo
): Promise<boolean> {
  const owned = ownInfo(info);
  return verifySpun(await spinAgilePassword(password, owned), owned);
}

/**
 * {@link verifyAgilePassword}, from the password's {@link spinAgilePassword}. A verifier that is not whole AES
 * blocks is malformed and refused as such, not answered as a wrong password.
 */
function verifySpun(spun: Uint8Array, info: AgileEncryptionInfo): boolean {
  const verifierInputKey = agileKeyFromSpun(spun, info, AGILE_BLOCK_KEYS.verifierHashInput);
  // MS-OFFCRYPTO §2.3.4.13: these blobs are AES-CBC with the plaintext zero-padded to the block boundary — NOT
  // PKCS#7 — and the key encryptor's salt as the IV.
  const verifierInput = aesCbcRawDecrypt(
    info.encryptedVerifierHashInput,
    verifierInputKey,
    info.keySalt,
    "encryptedVerifierHashInput"
  );
  // The plaintext is exactly saltSize (16) bytes; trim the zero padding before hashing.
  const computedHash = hash(
    mapHashName(info.hashAlgorithm),
    verifierInput.slice(0, info.blockSize)
  );
  const verifierValueKey = agileKeyFromSpun(spun, info, AGILE_BLOCK_KEYS.verifierHashValue);
  const expectedHash = aesCbcRawDecrypt(
    info.encryptedVerifierHashValue,
    verifierValueKey,
    info.keySalt,
    "encryptedVerifierHashValue"
  );
  return bytesEqual(computedHash.slice(0, info.hashSize), expectedHash.slice(0, info.hashSize));
}

/** The parameters the package segments and the data-integrity check use: `keyData`'s, else the key encryptor's. */
function keyDataOf(info: AgileEncryptionInfo): {
  salt: Uint8Array;
  hashName: string;
  blockSize: number;
} {
  return info.keyData
    ? {
        salt: info.keyData.salt,
        hashName: mapHashName(info.keyData.hashAlgorithm),
        blockSize: info.keyData.blockSize
      }
    : { salt: info.keySalt, hashName: mapHashName(info.hashAlgorithm), blockSize: info.blockSize };
}

/** `H(salt + blockKey)` truncated to the block size — the IV MS-OFFCRYPTO §2.3.4.12 gives each keyed blob. */
function blockIv(
  hashName: string,
  salt: Uint8Array,
  blockKey: Uint8Array,
  blockSize: number
): Uint8Array {
  return hash(hashName, concat(salt, blockKey)).slice(0, blockSize);
}

/**
 * Decrypt an `EncryptedPackage` stream with its parsed info and the password, checking the data-integrity HMAC
 * (MS-OFFCRYPTO §2.3.4.14) when the info carries one.
 *
 * @param maxDecryptedSize - Upper bound on the decrypted size; a `totalSize` header claiming more fails with
 *   `reason: "limit"` before anything is allocated.
 */
export async function decryptAgilePackage(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  password: string,
  maxDecryptedSize: number = 512 * 1024 * 1024
): Promise<Uint8Array> {
  // A package too large is refused before the password is spun for it.
  packageSize(encryptedPackage, maxDecryptedSize);
  const length = encryptedPackage.length;
  const owned = ownInfo(info);
  // One budget for the whole decrypt: each phase starting a budget of its own let the work between two yields add up.
  const yieldIfDue = createYieldBudget();
  const spun = await spinAgilePassword(password, owned, yieldIfDue);
  return decryptSpun(
    await ownPackage(encryptedPackage, length, maxDecryptedSize, yieldIfDue),
    owned,
    spun,
    maxDecryptedSize,
    yieldIfDue
  );
}

function checkUnchanged(encryptedPackage: Uint8Array, length: number): void {
  if (encryptedPackage.length !== length) {
    throw new OfficeCryptoError(
      "malformed",
      `EncryptedPackage changed size while it was read (${length} bytes, then ${encryptedPackage.length})`
    );
  }
}

/**
 * The `EncryptedPackage` stream copied into bytes nothing outside this module holds, a slice at a time so a large one
 * does not block. The integrity check and the decryption must read the same bytes, and the decrypt yields between them:
 * a caller's buffer written to meanwhile was verified as it was and decrypted as it had become. All of the stream — the
 * HMAC covers it as stored — so the copy is as long as the input, which the readers hold to their package-size limit.
 *
 * `length` is the stream's length when the decrypt began: one that changes size before or while it is copied (a
 * resizable `ArrayBuffer` shrunk while the password was spun) is refused, where it was read as a package of another
 * size. A package declaring more than `maxDecryptedSize` is refused before anything is allocated for it.
 */
async function ownPackage(
  encryptedPackage: Uint8Array,
  length: number,
  maxDecryptedSize: number,
  yieldIfDue: YieldPoint
): Promise<Uint8Array> {
  checkUnchanged(encryptedPackage, length);
  packageSize(encryptedPackage, maxDecryptedSize);
  const owned = new Uint8Array(length);
  for (let start = 0; start < length; start += PACED_SLICE_LENGTH) {
    checkUnchanged(encryptedPackage, length);
    owned.set(
      encryptedPackage.subarray(start, Math.min(start + PACED_SLICE_LENGTH, length)),
      start
    );
    await yieldIfDue();
  }
  return owned;
}

/**
 * The fields of `info` an operation reads, copied — for an operation that reads them on either side of a yield: a
 * caller's object changed in between (its `spinCount`, a salt, the integrity fields) would have the derivation and the
 * check run with values that were never the file's.
 *
 * Only these fields, by name: the info is a structural type, so a caller's object may carry others, which are neither
 * read nor copied (cloning the whole object refused one holding a function). Each byte array is copied with the
 * constructor, which copies whatever view it is given — a `Buffer`'s `slice` is a view, and a view of a
 * `SharedArrayBuffer` keeps sharing its memory under `structuredClone`.
 */
function ownInfo(info: AgileEncryptionInfo): AgileEncryptionInfo {
  return {
    keySalt: ownBytes(info.keySalt),
    spinCount: info.spinCount,
    cipherAlgorithm: info.cipherAlgorithm,
    cipherChaining: info.cipherChaining,
    hashAlgorithm: info.hashAlgorithm,
    keyBits: info.keyBits,
    hashSize: info.hashSize,
    encryptedVerifierHashInput: ownBytes(info.encryptedVerifierHashInput),
    encryptedVerifierHashValue: ownBytes(info.encryptedVerifierHashValue),
    encryptedKeyValue: ownBytes(info.encryptedKeyValue),
    blockSize: info.blockSize,
    ...(info.dataIntegritySalt === undefined
      ? {}
      : { dataIntegritySalt: ownBytes(info.dataIntegritySalt) }),
    ...(info.dataIntegrityHmac === undefined
      ? {}
      : { dataIntegrityHmac: ownBytes(info.dataIntegrityHmac) }),
    ...(info.keyData === undefined
      ? {}
      : {
          keyData: {
            salt: ownBytes(info.keyData.salt),
            hashAlgorithm: info.keyData.hashAlgorithm,
            hashSize: info.keyData.hashSize,
            blockSize: info.keyData.blockSize,
            keyBits: info.keyData.keyBits
          }
        })
  };
}

/** The fields of `info` a key derivation reads, copied — see {@link ownInfo}. */
function ownKeyInfo(info: KeyDerivationInfo): KeyDerivationInfo {
  return {
    keySalt: ownBytes(info.keySalt),
    spinCount: info.spinCount,
    hashAlgorithm: info.hashAlgorithm,
    keyBits: info.keyBits
  };
}

/** `bytes` copied into memory nothing else holds — see {@link ownInfo}. */
function ownBytes(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes);
}

/** {@link decryptAgilePackage}, from the password's {@link spinAgilePassword}. */
async function decryptSpun(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  spun: Uint8Array,
  maxDecryptedSize: number,
  yieldIfDue: YieldPoint
): Promise<Uint8Array> {
  // `encryptedPackage` is this module's own copy (`ownPackage`).
  const totalSize = packageSize(encryptedPackage, maxDecryptedSize);
  const keyEncryptionKey = agileKeyFromSpun(spun, info, AGILE_BLOCK_KEYS.encryptedKey);
  // Zero-padded AES-CBC per MS-OFFCRYPTO §2.3.4.13 (NOT PKCS#7); the key is exactly keyBits/8 bytes.
  const packageKey = aesCbcRawDecrypt(
    info.encryptedKeyValue,
    keyEncryptionKey,
    info.keySalt,
    "encryptedKeyValue"
  ).slice(0, (info.keyData?.keyBits ?? info.keyBits) / 8);

  const keyData = keyDataOf(info);
  if (info.dataIntegritySalt && info.dataIntegrityHmac) {
    await checkDataIntegrity(encryptedPackage, info, packageKey, keyData);
  }

  // Decrypt in 4096-byte segments, each with IV = H(keyData salt + segment index LE). Package segments use zero
  // padding (MS-OFFCRYPTO §2.3.4.15): the raw (no-padding) decrypt, never PKCS#7, which would mis-strip zeros.
  const segmentSize = 4096;
  const encData = encryptedPackage.subarray(8);
  const result = new Uint8Array(totalSize);
  const idxBytes = new Uint8Array(4);
  const idxView = new DataView(idxBytes.buffer);
  // One segment is microseconds of work and the package can be hundreds of megabytes of them.
  for (let i = 0, offset = 0; offset < totalSize; i++) {
    const segStart = i * segmentSize;
    const segData = encData.subarray(segStart, Math.min(segStart + segmentSize, encData.length));
    idxView.setUint32(0, i, true);
    const segIv = blockIv(keyData.hashName, keyData.salt, idxBytes, keyData.blockSize);
    const plain = aesCbcRawDecrypt(segData, packageKey, segIv, `EncryptedPackage segment ${i}`);
    const copyLen = Math.min(plain.length, totalSize - offset);
    result.set(plain.subarray(0, copyLen), offset);
    offset += copyLen;
    await yieldIfDue();
  }
  return result;
}

/**
 * The decrypted size an `EncryptedPackage` stream declares, checked: the first 8 bytes (uint64 LE) are the size, which
 * cannot legally exceed the encrypted body, and a size past `maxDecryptedSize` is a `limit` failure.
 */
function packageSize(encryptedPackage: Uint8Array, maxDecryptedSize: number): number {
  if (encryptedPackage.length < 8) {
    throw new OfficeCryptoError(
      "malformed",
      "EncryptedPackage too small (missing 8-byte size header)"
    );
  }
  const totalSize = Number(
    new DataView(encryptedPackage.buffer, encryptedPackage.byteOffset, 8).getBigUint64(0, true)
  );
  if (!Number.isFinite(totalSize) || totalSize < 0 || totalSize > encryptedPackage.length - 8) {
    throw new OfficeCryptoError(
      "malformed",
      `EncryptedPackage declared size (${totalSize}) is invalid or exceeds the maximum allowed (${maxDecryptedSize})`
    );
  }
  if (totalSize > maxDecryptedSize) {
    throw new OfficeCryptoError(
      "limit",
      `EncryptedPackage declared size (${totalSize}) is invalid or exceeds the maximum allowed (${maxDecryptedSize})`
    );
  }
  return totalSize;
}

/** MS-OFFCRYPTO §2.3.4.14: the HMAC of the whole `EncryptedPackage` stream must match the one the file stores. */
async function checkDataIntegrity(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  packageKey: Uint8Array,
  keyData: { salt: Uint8Array; hashName: string; blockSize: number }
): Promise<void> {
  const hashSize = info.keyData?.hashSize ?? info.hashSize;
  const ivKey = blockIv(
    keyData.hashName,
    keyData.salt,
    AGILE_BLOCK_KEYS.dataIntegrityKey,
    keyData.blockSize
  );
  const ivValue = blockIv(
    keyData.hashName,
    keyData.salt,
    AGILE_BLOCK_KEYS.dataIntegrityValue,
    keyData.blockSize
  );
  const hmacKey = aesCbcRawDecrypt(
    info.dataIntegritySalt!,
    packageKey,
    ivKey,
    "encryptedHmacKey"
  ).slice(0, hashSize);
  const expected = aesCbcRawDecrypt(
    info.dataIntegrityHmac!,
    packageKey,
    ivValue,
    "encryptedHmacValue"
  ).slice(0, hashSize);
  // The key is never empty: its blob is whole AES blocks (`aesCbcRawDecrypt`) and `hashSize` a digest's size.
  const actual = await hmac(keyData.hashName, hmacKey, encryptedPackage);
  if (!bytesEqual(actual.slice(0, hashSize), expected)) {
    throw new OfficeCryptoError(
      "integrity",
      "The encrypted package failed its data-integrity check: the file was altered or damaged"
    );
  }
}

/**
 * The WebCrypto name of an algorithm in either spelling. Every name reaching
 * here was checked (`EncryptOptions`, `parseEncryptionInfo`, `Agile.deriveKey`),
 * so the SHA-512 fallback is for the unchecked deprecated forwards only.
 */
function mapHashName(name: string): string {
  return canonicalHashAlgorithm(name) ?? "SHA-512";
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
 * Decrypt an AES-CBC blob written with zero padding (no PKCS#7): the key and verifier blobs (MS-OFFCRYPTO §2.3.4.13)
 * and the package's segments (§2.3.4.15). The IV is truncated/extended to 16 bytes.
 *
 * A blob that is not whole blocks is malformed — no writer pads to less — and is refused as such, `what` naming it:
 * handed to the platform it failed with its own error (OpenSSL's `wrong final block length`), not this module's.
 */
function aesCbcRawDecrypt(
  data: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array,
  what: string
): Uint8Array {
  if (data.length === 0 || data.length % 16 !== 0) {
    throw new OfficeCryptoError(
      "malformed",
      `${what} is ${data.length} byte(s), not a whole number of AES blocks`
    );
  }
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

/**
 * HMAC(hashName, key, message), by Web Crypto on every platform.
 *
 * Native, and computed off the main thread: the message is the whole `EncryptedPackage` stream, and the hand-built
 * HMAC this replaced hashed it in one synchronous call on Node — a copy of the package and a blocked event loop for as
 * long as the digest took, which grows with the file. Web Crypto hashes a key longer than the block size down first,
 * as RFC 2104 has it.
 */
async function hmac(hashName: string, key: Uint8Array, message: Uint8Array): Promise<Uint8Array> {
  const subtle = globalThis.crypto.subtle;
  const cryptoKey = await subtle.importKey(
    "raw",
    key as Uint8Array<ArrayBuffer>,
    { name: "HMAC", hash: hashName },
    false,
    ["sign"]
  );
  return new Uint8Array(await subtle.sign("HMAC", cryptoKey, message as Uint8Array<ArrayBuffer>));
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
// EncryptionInfo
// =============================================================================

/** A hash name as `EncryptionInfo` spells it, or a malformed-info error. */
function infoHashAlgorithm(raw: string): "SHA1" | "SHA256" | "SHA384" | "SHA512" {
  if (raw !== "SHA1" && raw !== "SHA256" && raw !== "SHA384" && raw !== "SHA512") {
    throw new OfficeCryptoError(
      "unsupported",
      `Unsupported hashAlgorithm in EncryptionInfo: ${raw} (expected SHA1/SHA256/SHA384/SHA512)`
    );
  }
  return raw;
}

/**
 * A `hashSize`, which must be what `algorithm` outputs (MS-OFFCRYPTO §2.3.4.10, §2.3.4.11). It is also how many
 * bytes of the verifier and of the data-integrity HMAC are compared, so a file free to give less would have the
 * integrity check compare one byte — a forged package passing it once in 256 tries.
 */
function infoHashSize(raw: string | undefined, algorithm: string): number {
  const expected = getHashSizeFor(mapHashName(algorithm));
  const value = raw === undefined ? expected : Number(raw);
  if (value !== expected) {
    throw new OfficeCryptoError(
      "malformed",
      `Invalid hashSize in EncryptionInfo: ${raw} (${algorithm} outputs ${expected} bytes)`
    );
  }
  return value;
}

function infoKeyBits(raw: string | undefined): 128 | 192 | 256 {
  const keyBits = parseInt(raw ?? "256", 10);
  if (keyBits !== 128 && keyBits !== 192 && keyBits !== 256) {
    throw new OfficeCryptoError(
      "unsupported",
      `Unsupported keyBits in EncryptionInfo: ${raw} (expected 128, 192, or 256)`
    );
  }
  return keyBits;
}

/**
 * Bounded numeric fields. Every attribute of an `EncryptionInfo` read from a file is attacker-controlled: a spin
 * count of 10⁹ would freeze key derivation, and hash/block sizes are slice lengths and IV sizes. The bounds are
 * wider than anything Office writes.
 */
function infoNumber(
  raw: string | undefined,
  fallback: string,
  name: string,
  min: number,
  max: number
): number {
  const value = parseInt(raw ?? fallback, 10);
  if (!Number.isFinite(value) || value < min || value > max) {
    throw new OfficeCryptoError(
      "malformed",
      `Invalid ${name} in EncryptionInfo: ${raw} (expected ${min}..${max === 10_000_000 ? "10_000_000" : max})`
    );
  }
  return value;
}

/**
 * Parse the XML of an Agile `EncryptionInfo` stream (the part after its 8-byte version header).
 */
export function parseAgileEncryptionInfo(xmlStr: string): AgileEncryptionInfo {
  // Located with a linear scan: a regular expression with lazy quantifiers backtracks catastrophically on a hostile
  // stream with very long unterminated tags.
  const keyDataAttrs = extractSelfClosingTagAttrs(xmlStr, "keyData");
  const pwdAttrs = extractSelfClosingTagAttrs(xmlStr, "p:encryptedKey");
  if (!keyDataAttrs || !pwdAttrs) {
    throw new OfficeCryptoError(
      "malformed",
      "Invalid EncryptionInfo XML - missing keyData or encryptedKey"
    );
  }
  const required = (attrs: Record<string, string>, name: string): Uint8Array => {
    const v = attrs[name];
    if (v === undefined || v === "") {
      throw new OfficeCryptoError(
        "malformed",
        `Invalid EncryptionInfo XML - missing or empty required field: ${name}`
      );
    }
    return new Uint8Array(base64ToUint8Array(v));
  };
  const integrity = extractSelfClosingTagAttrs(xmlStr, "dataIntegrity");
  const hmacKey = integrity?.encryptedHmacKey;
  const hmacValue = integrity?.encryptedHmacValue;
  const keyDataSalt = keyDataAttrs.saltValue;
  const hashAlgorithm = infoHashAlgorithm(pwdAttrs.hashAlgorithm ?? "SHA512");
  const keyDataHashAlgorithm = infoHashAlgorithm(keyDataAttrs.hashAlgorithm ?? "SHA512");
  return {
    cipherAlgorithm: "AES",
    cipherChaining: "ChainingModeCBC",
    keyBits: infoKeyBits(pwdAttrs.keyBits),
    hashAlgorithm,
    hashSize: infoHashSize(pwdAttrs.hashSize, hashAlgorithm),
    spinCount: infoNumber(pwdAttrs.spinCount, "100000", "spinCount", 0, 10_000_000),
    keySalt: required(pwdAttrs, "saltValue"),
    encryptedVerifierHashInput: required(pwdAttrs, "encryptedVerifierHashInput"),
    encryptedVerifierHashValue: required(pwdAttrs, "encryptedVerifierHashValue"),
    encryptedKeyValue: required(pwdAttrs, "encryptedKeyValue"),
    blockSize: infoNumber(pwdAttrs.blockSize, "16", "blockSize", 8, 64),
    ...(hmacKey && hmacValue
      ? {
          dataIntegritySalt: new Uint8Array(base64ToUint8Array(hmacKey)),
          dataIntegrityHmac: new Uint8Array(base64ToUint8Array(hmacValue))
        }
      : {}),
    ...(keyDataSalt
      ? {
          keyData: {
            salt: new Uint8Array(base64ToUint8Array(keyDataSalt)),
            hashAlgorithm: keyDataHashAlgorithm,
            hashSize: infoHashSize(keyDataAttrs.hashSize, keyDataHashAlgorithm),
            blockSize: infoNumber(keyDataAttrs.blockSize, "16", "blockSize", 8, 64),
            keyBits: infoKeyBits(keyDataAttrs.keyBits)
          }
        }
      : {})
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
// Encrypted packages
// =============================================================================

/**
 * The name MS-OFFCRYPTO §2.3 gives an `EncryptionInfo`'s encryption, from its version and flags: Standard is x.2
 * with `fCryptoAPI` and `fAES`, RC4 CryptoAPI is x.2 with `fCryptoAPI` alone, Extensible is x.3 with `fExternal`.
 */
function encryptionName(major: number, minor: number, flags: number): string {
  const cryptoApi = (flags & 0x04) !== 0;
  const aes = (flags & 0x20) !== 0;
  if (minor === 2 && major >= 2 && major <= 4 && cryptoApi) {
    return aes ? "Standard Encryption" : "RC4 CryptoAPI Encryption";
  }
  if (minor === 3 && (major === 3 || major === 4)) {
    return "Extensible Encryption";
  }
  return "an unknown encryption";
}

/**
 * The streams of a compound file, as views of `bytes` where they lie in consecutive sectors: they are only read here,
 * and the package they hold — up to the whole file — is not copied before it is decrypted.
 */
function cfbStreams(bytes: Uint8Array, maxStreamSize?: number): CfbEntry[] {
  return readCfb(
    bytes,
    maxStreamSize === undefined ? { views: true } : { maxStreamSize, views: true }
  );
}

/**
 * Whether `bytes` are an encrypted OOXML package: a compound file holding the `EncryptionInfo` and
 * `EncryptedPackage` streams. A compound file without them — a legacy `.xls` or `.doc` — is not.
 */
export function isEncryptedPackage(bytes: Uint8Array): boolean {
  if (!isCfb(bytes)) {
    return false;
  }
  try {
    const names = new Set(cfbStreams(bytes).map(e => e.name));
    return names.has("EncryptionInfo") && names.has("EncryptedPackage");
  } catch {
    return false;
  }
}

/**
 * The Agile encryption parameters and the encrypted package stream of an encrypted package. The stream may be a view
 * of `bytes` (see `cfbStreams`): what decrypts it copies it first, so a change to `bytes` cannot reach the result.
 *
 * @throws {CfbError} If the bytes are not a readable compound file.
 * @throws {OfficeCryptoError} If it is not an encrypted package, or not Agile-encrypted.
 */
export function readEncryptedPackage(
  bytes: Uint8Array,
  maxStreamSize?: number
): { info: AgileEncryptionInfo; encryptedPackage: Uint8Array } {
  if (!isCfb(bytes)) {
    throw new OfficeCryptoError(
      "malformed",
      "Not an encrypted Office package (CFB signature not found)"
    );
  }
  const entries = cfbStreams(bytes, maxStreamSize);
  const encInfoStream = entries.find(e => e.name === "EncryptionInfo");
  if (!encInfoStream) {
    throw new OfficeCryptoError("malformed", "CFB: EncryptionInfo stream not found");
  }
  const encPkgEntry = entries.find(e => e.name === "EncryptedPackage");
  if (!encPkgEntry) {
    throw new OfficeCryptoError("malformed", "CFB: EncryptedPackage stream not found");
  }
  const infoData = encInfoStream.data;
  if (infoData.length < 8) {
    throw new OfficeCryptoError("malformed", "EncryptionInfo too small (missing version header)");
  }
  const versionView = new DataView(infoData.buffer, infoData.byteOffset, 8);
  const versionMajor = versionView.getUint16(0, true);
  const versionMinor = versionView.getUint16(2, true);
  const flags = versionView.getUint32(4, true);
  // MS-OFFCRYPTO §2.3 tells the families apart by (major, minor): Agile is 4.4. Only Agile is implemented.
  if (versionMajor === 4 && versionMinor === 4) {
    const xmlStr = textDecoder.decode(infoData.subarray(8));
    return { info: parseAgileEncryptionInfo(xmlStr), encryptedPackage: encPkgEntry.data };
  }
  throw new OfficeCryptoError(
    "unsupported",
    `Unsupported encryption version: ${versionMajor}.${versionMinor} (${encryptionName(versionMajor, versionMinor, flags)}). ` +
      `Only Agile Encryption (4.4) is supported.`
  );
}

/**
 * Decrypt an encrypted package with `password`: the OOXML package (ZIP) bytes inside.
 *
 * @param maxDecryptedSize - Upper bound on the package's size. Default: 512 MiB.
 * @throws {OfficeCryptoError} `reason: "password"` for a wrong password; see {@link OfficeCryptoFailure}.
 * @throws {CfbError} If the container is malformed.
 */
export async function decryptPackage(
  bytes: Uint8Array,
  password: string,
  maxDecryptedSize?: number
): Promise<Uint8Array> {
  const limit = maxDecryptedSize ?? 512 * 1024 * 1024;
  // One budget for the whole decrypt, the container read below included — see `decryptAgilePackage`.
  const yieldIfDue = createYieldBudget();
  // The container's streams are bounded by the input, already in memory; the package's own size is checked against
  // `limit` before it is decrypted, which is what reports a package too large for the read.
  // `info` is parsed from the file here, so it is this call's own; the package is refused, if too large, before the
  // password is spun for it.
  const { info, encryptedPackage } = readEncryptedPackage(bytes, Infinity);
  packageSize(encryptedPackage, limit);
  const length = encryptedPackage.length;
  // Spun once for the verifier and the package key alike — see `spinAgilePassword`.
  const spun = await spinAgilePassword(password, info, yieldIfDue);
  if (!verifySpun(spun, info)) {
    throw new OfficeCryptoError("password", "Incorrect password");
  }
  return decryptSpun(
    await ownPackage(encryptedPackage, length, limit, yieldIfDue),
    info,
    spun,
    limit,
    yieldIfDue
  );
}

/**
 * Whether `password` opens the encrypted package `bytes`, checked against its password verifier.
 *
 * @throws {OfficeCryptoError} If it is not an Agile-encrypted package.
 */
export async function verifyPackagePassword(bytes: Uint8Array, password: string): Promise<boolean> {
  return verifyAgilePassword(password, readEncryptedPackage(bytes).info);
}

/**
 * Encrypt an OOXML package (ZIP bytes) with `password`: the CFB container Office opens with that password.
 * Arguments are the caller's to check.
 */
export async function encryptPackage(
  zip: Uint8Array,
  password: string,
  options?: EncryptOptions
): Promise<Uint8Array> {
  const keyBits = options?.keyBits ?? 256;
  // `EncryptionInfo` names the hash in the format's own spelling (MS-OFFCRYPTO §2.3.4.10).
  const hashAlgorithm = compactHashAlgorithm(
    canonicalHashAlgorithm(options?.hashAlgorithm ?? "SHA-512")!
  );
  const spinCount = options?.spinCount ?? 100000;
  const hashName = mapHashName(hashAlgorithm);
  const hashSize = getHashSizeFor(hashName);
  const keyBytes = keyBits / 8;
  const blockSize = 16;

  // The key data and the password key encryptor each get their own salt, as Office writes them.
  const keyDataSalt = randomBytes(16);
  const keySalt = randomBytes(16);
  const packageKey = randomBytes(keyBytes);
  const keyInfo = { keySalt, spinCount, hashAlgorithm, keyBits };

  // One budget for the whole encrypt — see `decryptAgilePackage`. Spun once for the three keys below — see
  // `spinAgilePassword`.
  const yieldIfDue = createYieldBudget();
  const spun = await spinAgilePassword(password, keyInfo, yieldIfDue);
  // The package key, under a key derived from the password (zero-padded AES-CBC, MS-OFFCRYPTO §2.3.4.13).
  const keyEncryptionKey = agileKeyFromSpun(spun, keyInfo, AGILE_BLOCK_KEYS.encryptedKey);
  const encryptedKeyValue = aesCbcZeroPadEncrypt(
    padToBlock(packageKey, blockSize),
    keyEncryptionKey,
    keySalt
  );

  // The verifier: 16 random bytes and their hash, each under its own derived key.
  const verifierInput = randomBytes(16);
  const verifierHash = hash(hashName, verifierInput);
  const verifierInputKey = agileKeyFromSpun(spun, keyInfo, AGILE_BLOCK_KEYS.verifierHashInput);
  const encryptedVerifierHashInput = aesCbcZeroPadEncrypt(
    padToBlock(verifierInput, blockSize),
    verifierInputKey,
    keySalt
  );
  const verifierValueKey = agileKeyFromSpun(spun, keyInfo, AGILE_BLOCK_KEYS.verifierHashValue);
  const encryptedVerifierHashValue = aesCbcZeroPadEncrypt(
    padToBlock(verifierHash, blockSize),
    verifierValueKey,
    keySalt
  );

  // The package, in 4096-byte segments keyed by the key data's salt.
  const encryptedPackage = await encryptPackageData(
    zip,
    packageKey,
    keyDataSalt,
    hashName,
    blockSize,
    yieldIfDue
  );

  // Data integrity (MS-OFFCRYPTO §2.3.4.14), which Office verifies on open: a random HMAC key, the HMAC of the
  // whole EncryptedPackage stream (size prefix included), each encrypted with the package key and an IV of
  // H(keyData salt + block key).
  const hmacKey = randomBytes(hashSize);
  const ivHmacKey = blockIv(hashName, keyDataSalt, AGILE_BLOCK_KEYS.dataIntegrityKey, blockSize);
  const ivHmacValue = blockIv(
    hashName,
    keyDataSalt,
    AGILE_BLOCK_KEYS.dataIntegrityValue,
    blockSize
  );
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

  const encInfoXml = buildEncryptionInfoXml({
    keyBits,
    hashAlgorithm,
    hashSize,
    spinCount,
    blockSize,
    keyDataSalt,
    keySalt,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue,
    encryptedHmacKey,
    encryptedHmacValue
  });

  // 8-byte version header: version 4.4, flags 0x40 (agile).
  const xmlBytes = textEncoder.encode(encInfoXml);
  const encInfoStream = new Uint8Array(8 + xmlBytes.length);
  const encInfoView = new DataView(encInfoStream.buffer);
  encInfoView.setUint16(0, 4, true);
  encInfoView.setUint16(2, 4, true);
  encInfoView.setUint32(4, 0x40, true);
  encInfoStream.set(xmlBytes, 8);

  // Office requires the \x06DataSpaces structure (MS-OFFCRYPTO §2.3.2) beside EncryptionInfo and EncryptedPackage;
  // without it Word reports the file as corrupt even with the right password.
  return writeCfb([
    ...buildDataSpacesStreams(),
    { name: "EncryptionInfo", data: encInfoStream },
    { name: "EncryptedPackage", data: encryptedPackage }
  ]);
}

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
  blockSize: number,
  yieldIfDue: YieldPoint
): Promise<Uint8Array> {
  const segmentSize = 4096;
  // The length the package is encrypted at, read once: the size header and every segment agree on it, and a package
  // that changes size while it is read — a resizable buffer — is refused, as a decrypt refuses one.
  const length = data.length;
  // The size (8 bytes LE uint64), then each segment encrypted into its place: the stream is allocated once, at its
  // size, where collecting the segments and joining them copied the package again in one call.
  const result = new Uint8Array(8 + Math.ceil(length / blockSize) * blockSize);
  new DataView(result.buffer).setBigUint64(0, BigInt(length), true);
  const idxBytes = new Uint8Array(4);
  const idxView = new DataView(idxBytes.buffer);
  for (let i = 0, segStart = 0; segStart < length; i++, segStart += segmentSize) {
    if (data.length !== length) {
      throw new OfficeCryptoError(
        "malformed",
        `the package changed size while it was encrypted (${length} bytes, then ${data.length})`
      );
    }
    let segData = data.subarray(segStart, Math.min(segStart + segmentSize, length));
    // Pad the last segment to a whole number of blocks.
    if (segData.length % blockSize !== 0) {
      const padded = new Uint8Array(Math.ceil(segData.length / blockSize) * blockSize);
      padded.set(segData);
      segData = padded;
    }
    // Segment IV = hash(salt + segment_index_LE) truncated to blockSize
    idxView.setUint32(0, i, true);
    const segIv = blockIv(hashName, keySalt, idxBytes, blockSize);
    result.set(aesCbcZeroPadEncrypt(segData, packageKey, segIv), 8 + segStart);
    // As in `decryptSpun`: one segment is microseconds, the package can be hundreds of megabytes.
    await yieldIfDue();
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
  keyDataSalt: Uint8Array;
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
    keyDataSalt,
    keySalt,
    encryptedVerifierHashInput,
    encryptedVerifierHashValue,
    encryptedKeyValue,
    encryptedHmacKey,
    encryptedHmacValue
  } = params;

  const keyDataSaltB64 = uint8ArrayToBase64(keyDataSalt);
  const saltB64 = uint8ArrayToBase64(keySalt);
  const vhiB64 = uint8ArrayToBase64(encryptedVerifierHashInput);
  const vhvB64 = uint8ArrayToBase64(encryptedVerifierHashValue);
  const ekvB64 = uint8ArrayToBase64(encryptedKeyValue);
  const hmacKeyB64 = uint8ArrayToBase64(encryptedHmacKey);
  const hmacValB64 = uint8ArrayToBase64(encryptedHmacValue);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ' +
    'xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ' +
    'xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">\r\n' +
    `<keyData saltSize="16" blockSize="${blockSize}" keyBits="${keyBits}" ` +
    `hashSize="${hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" ` +
    `hashAlgorithm="${hashAlgorithm}" saltValue="${keyDataSaltB64}"/>\r\n` +
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
