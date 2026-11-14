/**
 * DOCX Agile Encryption (MS-OFFCRYPTO) — Word's side of the shared implementation.
 *
 * The format-agnostic work — the CFB container, key derivation, the password verifier, package encryption and
 * decryption, the data-integrity HMAC, `EncryptionInfo` — lives in `@archive/office-crypto/agile` (and
 * `@archive/cfb/cfb`), where Excel uses it too. This module keeps the names `documonster/word/crypto` has always
 * published, checks their arguments with Word's checks, and reports the shared errors as Word's: a malformed
 * container is a `DocxParseError`, anything that stops decryption (a wrong password, an unsupported encryption, a
 * damaged package) a `DocxDecryptionError` — each with the shared error as its `cause`.
 *
 * It is loaded only when a password is involved: the readers test for a compound file with `@archive/cfb/signature`
 * (see `security/decrypt.ts`) and `Document.toBuffer` imports this module for `{ password }` alone.
 */

import { CfbError } from "@archive/cfb/cfb";
import { isCfb } from "@archive/cfb/signature";
import {
  AGILE_BLOCK_KEYS as SHARED_BLOCK_KEYS,
  decryptAgilePackage,
  decryptPackage as decryptSharedPackage,
  deriveAgileKey,
  encryptPackage,
  OfficeCryptoError,
  parseAgileEncryptionInfo,
  verifyAgilePassword,
  verifyPackagePassword
} from "@archive/office-crypto/agile";
import type { AgileEncryptionInfo, EncryptOptions } from "@archive/office-crypto/agile";
import { asPlainBytes } from "@utils/binary";
import { checkBytes, checkString, invalidArgument } from "@word/core/argument-checks";
import { DocxDecryptionError, DocxParseError } from "@word/errors";
import { checkEncryptOptions } from "@word/security/package-encryption";

export type { AgileEncryptionInfo, EncryptOptions } from "@archive/office-crypto/agile";

/**
 * The shared code's errors as Word reports them: a malformed container is a `DocxParseError`, a failure to decrypt
 * a `DocxDecryptionError`, each with the same message and the shared error as its `cause`.
 */
export function docxCryptoError(error: unknown): unknown {
  if (error instanceof CfbError) {
    return new DocxParseError(error.message, { cause: error });
  }
  if (error instanceof OfficeCryptoError) {
    return new DocxDecryptionError(error.message, { cause: error, reason: error.reason });
  }
  return error;
}

async function translated<T>(work: () => Promise<T>): Promise<T> {
  try {
    return await work();
  } catch (error) {
    throw docxCryptoError(error);
  }
}

/** Detect whether a buffer is an encrypted Office document (CFB format). */
export function isEncryptedDocx(buffer: Uint8Array | ArrayBuffer | ArrayBufferView): boolean {
  return isCfb(checkBytes(buffer, "buffer", "WordCrypto.isEncrypted"));
}

/**
 * Derive an encryption key from a password using the agile encryption KDF (MS-OFFCRYPTO §2.3.4.11).
 */
export function deriveEncryptionKey(
  password: string,
  info: {
    keySalt: Uint8Array;
    spinCount: number;
    hashAlgorithm: string;
    keyBits: number;
  },
  blockKey: Uint8Array
): Promise<Uint8Array<ArrayBuffer>> {
  return translated(() => deriveAgileKey(password, info, blockKey)).then(asPlainBytes);
}

/** Block keys as defined in MS-OFFCRYPTO 2.3.4.13. */
export const AGILE_BLOCK_KEYS = SHARED_BLOCK_KEYS;

/**
 * Verify a password against the encryption info.
 *
 * @returns True if password is correct.
 * @throws {DocxDecryptionError} `reason: "malformed"` when the info's verifier is damaged.
 */
export function verifyPassword(password: string, info: AgileEncryptionInfo): Promise<boolean> {
  return translated(() => verifyAgilePassword(password, info));
}

/**
 * Decrypt the package data using the password.
 *
 * @param maxDecryptedSize - Optional upper bound on the decrypted size; if the `totalSize` header claims a larger
 *   value, throw rather than allocate. Defaults to 512 MiB.
 * @returns The decrypted (unencrypted) package bytes.
 */
export function decryptPackage(
  encryptedPackage: Uint8Array,
  info: AgileEncryptionInfo,
  password: string,
  maxDecryptedSize?: number
): Promise<Uint8Array> {
  return translated(() => decryptAgilePackage(encryptedPackage, info, password, maxDecryptedSize));
}

/**
 * Parse an EncryptionInfo XML stream from a CFB-wrapped encrypted DOCX (the UTF-8 XML part after the 8-byte
 * version/flags header).
 *
 * @throws {DocxDecryptionError} If the XML is not Agile encryption info this library can use.
 */
export function parseEncryptionInfoXml(xmlStr: string): AgileEncryptionInfo {
  try {
    return parseAgileEncryptionInfo(xmlStr);
  } catch (error) {
    throw docxCryptoError(error);
  }
}

/**
 * Read an encrypted DOCX file by providing a password: parse the CFB container, check the password, decrypt.
 *
 * @param maxDecryptedSize - Optional cap on the decrypted size (defaults to 512 MiB).
 * @returns The decrypted DOCX ZIP bytes (can be passed to readDocx).
 * @throws {DocxDecryptionError} If the file is not encrypted, the password is wrong, or decryption fails.
 */
export function decryptDocx(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  password: string,
  maxDecryptedSize?: number
): Promise<Uint8Array<ArrayBuffer>> {
  const bytes = checkBytes(buffer, "buffer", "WordCrypto.decrypt");
  return translated(() => decryptSharedPackage(bytes, password, maxDecryptedSize)).then(
    asPlainBytes
  );
}

/**
 * Check a password against an encrypted DOCX without decrypting the package — the same check {@link decryptDocx}
 * runs before it decrypts, so a `true` here means `decryptDocx` will accept the password.
 *
 * @returns `true` if the password opens the file, `false` if it does not.
 * @throws {DocxDecryptionError} If the buffer is not an encrypted DOCX or uses an unsupported encryption version.
 */
export function verifyDocxPassword(
  buffer: Uint8Array | ArrayBuffer | ArrayBufferView,
  password: string
): Promise<boolean> {
  const bytes = checkBytes(buffer, "buffer", "WordCrypto.verifyPassword");
  return translated(() => verifyPackagePassword(bytes, password));
}

/** Options for `WordCrypto.decrypt` and `WordCrypto.Agile.decryptPackage`. */
export interface DecryptOptions {
  /**
   * Cap on the decrypted package size in bytes, `Infinity` for none. Default: the
   * default security policy's `maxPackageSize` (500 MiB), the cap `Document.read` applies.
   */
  readonly maxDecryptedSize?: number;
}

/**
 * Encrypt a DOCX ZIP file with a password using Agile Encryption (MS-OFFCRYPTO).
 *
 * Produces a CFB (OLE2 Compound File) holding the `EncryptionInfo` and `EncryptedPackage` streams and the
 * `\x06DataSpaces` structure Office requires. The output can be opened by Microsoft Word with the given password.
 *
 * @param zipBytes - The unencrypted DOCX content (ZIP format).
 * @param password - The password to encrypt with.
 * @param options - Encryption options (key length, hash algorithm, spin count).
 * @returns The encrypted DOCX as a CFB-formatted Uint8Array.
 */
export async function encryptDocx(
  zipBytes: Uint8Array | ArrayBuffer | ArrayBufferView,
  password: string,
  options?: EncryptOptions
): Promise<Uint8Array<ArrayBuffer>> {
  const where = "WordCrypto.encrypt";
  const zip = checkBytes(zipBytes, "zipBytes", where);
  checkString(password, "password", where);
  checkEncryptOptions(options, "options", where);
  // The package is encrypted as given; a file that is not a ZIP would decrypt to bytes no reader opens. A ZIP begins
  // with a local file header — or, empty, with the end-of-central-directory record.
  const signature =
    zip.byteLength >= 4 ? new DataView(zip.buffer, zip.byteOffset, 4).getUint32(0, true) : 0;
  if (signature !== 0x04034b50 && signature !== 0x06054b50) {
    invalidArgument(where, "zipBytes", "a ZIP package (a .docx file's bytes)", zip);
  }
  return asPlainBytes(await translated(() => encryptPackage(zip, password, options)));
}
