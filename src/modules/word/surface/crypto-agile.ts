import { asPlainBytes } from "@utils/binary";
import {
  checkBytes,
  checkInteger,
  checkLimit,
  checkOneOf,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
/**
 * `WordCrypto.Agile` — the ECMA-376 Agile Encryption primitives behind
 * `WordCrypto.decrypt` / `WordCrypto.verifyPassword`, for callers that work
 * with the `EncryptionInfo` and `EncryptedPackage` streams themselves.
 *
 * `WordCrypto.Agile.parseEncryptionInfo(xml)`,
 * `WordCrypto.Agile.deriveKey(...)`, `WordCrypto.Agile.verifyPassword(info, pw)`,
 * `WordCrypto.Agile.decryptPackage(...)`, `WordCrypto.Agile.BLOCK_KEYS`.
 */
import type { AgileEncryptionInfo, DecryptOptions } from "@word/security/encryption";
import {
  decryptPackage as decryptAgilePackage,
  deriveEncryptionKey,
  parseEncryptionInfoXml,
  verifyPassword as verifyAgilePassword
} from "@word/security/encryption";
import { HASH_ALGORITHM_NAMES } from "@word/security/hash-names";
import { DEFAULT_SECURITY_POLICY } from "@word/security/policy";

export { AGILE_BLOCK_KEYS as BLOCK_KEYS } from "@word/security/encryption";

const KEY_BITS: ReadonlySet<number> = new Set([128, 192, 256]);

/** The fields of an `AgileEncryptionInfo` a key is derived from. */
function checkKeyInfo(info: unknown, field: string, where: string): void {
  checkRecord(info, field, where);
  const i = info as Record<string, unknown>;
  checkBytes(i.keySalt, `${field}.keySalt`, where);
  // MS-OFFCRYPTO: the spin count is at most 10,000,000.
  checkInteger(i.spinCount, `${field}.spinCount`, where, { min: 0, max: 10_000_000 });
  checkOneOf(i.hashAlgorithm, HASH_ALGORITHM_NAMES, `${field}.hashAlgorithm`, where);
  if (!KEY_BITS.has(i.keyBits as number)) {
    invalidArgument(where, `${field}.keyBits`, "one of 128, 192, 256", i.keyBits);
  }
  // The key data's parameters, when given, key the package segments and the data-integrity check.
  if (i.keyData !== undefined) {
    const at = `${field}.keyData`;
    checkRecord(i.keyData, at, where);
    const k = i.keyData as Record<string, unknown>;
    checkBytes(k.salt, `${at}.salt`, where);
    checkOneOf(k.hashAlgorithm, HASH_ALGORITHM_NAMES, `${at}.hashAlgorithm`, where);
    checkInteger(k.hashSize, `${at}.hashSize`, where, { min: 1, max: 128 });
    checkInteger(k.blockSize, `${at}.blockSize`, where, { min: 8, max: 64 });
    if (!KEY_BITS.has(k.keyBits as number)) {
      invalidArgument(where, `${at}.keyBits`, "one of 128, 192, 256", k.keyBits);
    }
  }
}

/**
 * Parse the XML of an Agile `EncryptionInfo` stream — the stream's bytes after
 * its 8-byte version header, decoded as text.
 *
 * @throws {DocxDecryptionError} If the XML is not Agile encryption info.
 */
export function parseEncryptionInfo(xml: string): AgileEncryptionInfo {
  checkString(xml, "xml", "WordCrypto.Agile.parseEncryptionInfo");
  return parseEncryptionInfoXml(xml);
}

/**
 * Derive an Agile encryption key from the info's salt, spin count, hash and key
 * size, `password`, and a block key (one of {@link BLOCK_KEYS}). The info comes
 * first, as in {@link verifyPassword} and {@link decryptPackage}.
 *
 * @throws {DocxArgumentError} If an argument would derive a key other than the
 *   one asked for: a hash other than SHA-1, SHA-256, SHA-384 or SHA-512 (either
 *   spelling, `"SHA-256"` or `"SHA256"`), a
 *   spin count that is not an integer from 0 to 10,000,000, a key size other
 *   than 128, 192 or 256 bits.
 */
export async function deriveKey(
  info: {
    keySalt: Uint8Array;
    spinCount: number;
    hashAlgorithm: string;
    keyBits: number;
  },
  password: string,
  blockKey: Uint8Array
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "WordCrypto.Agile.deriveKey";
  checkKeyInfo(info, "info", where);
  checkString(password, "password", where);
  const block = checkBytes(blockKey, "blockKey", where);
  return deriveEncryptionKey(
    password,
    { ...info, keySalt: checkBytes(info.keySalt, "info.keySalt", where) },
    block
  );
}

/**
 * Decrypt an `EncryptedPackage` stream with its parsed info and the password.
 * `options.maxDecryptedSize` caps the package, as in `WordCrypto.decrypt`
 * (default: the default security policy's `maxPackageSize`; `Infinity` for none).
 */
export async function decryptPackage(
  encryptedPackage: Uint8Array | ArrayBuffer | ArrayBufferView,
  info: AgileEncryptionInfo,
  password: string,
  options?: DecryptOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "WordCrypto.Agile.decryptPackage";
  const bytes = checkBytes(encryptedPackage, "encryptedPackage", where);
  checkKeyInfo(info, "info", where);
  checkString(password, "password", where);
  checkRecord(options, "options", where, { optional: true });
  checkLimit(options?.maxDecryptedSize, "options.maxDecryptedSize", where, { optional: true });
  return asPlainBytes(
    await decryptAgilePackage(
      bytes,
      info,
      password,
      options?.maxDecryptedSize ?? DEFAULT_SECURITY_POLICY.maxPackageSize
    )
  );
}

/**
 * Whether `password` opens a package with this encryption info, checked
 * against the info's password verifier without decrypting the package.
 * `WordCrypto.verifyPassword(bytes, password)` is the same check starting
 * from the encrypted file; this form starts from already-parsed info.
 *
 * @throws {DocxDecryptionError} `reason: "malformed"` when the info's verifier is damaged — not answered as a wrong
 * password.
 */
export async function verifyPassword(
  info: AgileEncryptionInfo,
  password: string
): Promise<boolean> {
  const where = "WordCrypto.Agile.verifyPassword";
  checkKeyInfo(info, "info", where);
  checkString(password, "password", where);
  return verifyAgilePassword(password, info);
}
