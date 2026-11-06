/**
 * The encrypted packages of Word's readers and writers, kept apart from `security/encryption.ts` so a reader or
 * writer carries only the signature test and the option checks: the CFB container and the encryption code load
 * through `import()` on the first encrypted file read, or the first write given a password.
 *
 * - Reading: an encrypted package (a compound file holding `EncryptionInfo` and `EncryptedPackage`) is decrypted with
 *   the read's `password`; anything else is read as it is.
 * - Writing: `Document.toBuffer` / `toBase64` / `writeFile` take `password` (and `encryption`), and encrypt the package.
 */

import { hasEncryptedPackageStreams, isCfb } from "@archive/cfb/signature";
import type { EncryptOptions } from "@archive/office-crypto/agile";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import { invalidArgument } from "@word/core/argument-checks";
import { checkSchema, w } from "@word/core/option-schema";
import { DocxEncryptedError, DocxParseError } from "@word/errors";
import { HASH_ALGORITHM_NAMES } from "@word/security/hash-names";

/**
 * The package inside `buffer`, as every reader opens it: an encrypted file is decrypted with `password`, its size
 * capped at `maxPackageSize`; one without a password is a `DocxEncryptedError`. A compound file that is not an
 * encrypted package — a legacy `.doc` — is a `DocxParseError` saying so. Anything else is returned as it is.
 */
export async function decryptIfEncrypted(
  buffer: Uint8Array,
  password: string | undefined,
  maxPackageSize: number
): Promise<Uint8Array> {
  if (!isCfb(buffer)) {
    return buffer;
  }
  if (!hasEncryptedPackageStreams(buffer)) {
    throw new DocxParseError(
      "Not a DOCX package — an OLE compound file: a legacy .doc document (Word 97–2003) or another Office " +
        "binary format, which this library does not read"
    );
  }
  if (password == null) {
    throw new DocxEncryptedError();
  }
  // The policy's package-size cap bounds the decrypted size, so a hostile container cannot claim a multi-GiB
  // package and force a huge allocation before the unzip stage applies its own limits.
  const { decryptDocx } = await import("@word/security/encryption");
  return decryptDocx(buffer, password, maxPackageSize);
}

const encryptOptionsSchema = w.record({
  keyBits: s.oneOf([128, 192, 256], { expected: "one of 128, 192, 256" }).optional(),
  hashAlgorithm: s
    .member<NonNullable<EncryptOptions["hashAlgorithm"]>>(HASH_ALGORITHM_NAMES)
    .optional(),
  // MS-OFFCRYPTO: the spin count is at most 10,000,000.
  spinCount: w.integer({ min: 0, max: 10_000_000 }).optional()
});
const _encryptOptionsMatch: SchemaMatches<typeof encryptOptionsSchema, EncryptOptions> = true;
void _encryptOptionsMatch;

/**
 * The `EncryptOptions` of `where` (`keyBits`, `hashAlgorithm`, `spinCount`), checked; `field` is their path
 * (`options` for `WordCrypto.encrypt`, `options.encryption` for a write option).
 */
export function checkEncryptOptions(options: unknown, field: string, where: string): void {
  checkSchema(encryptOptionsSchema.optional(), options, where, field);
}

/** What a write that encrypts takes: the password, and how to encrypt with it. */
interface EncryptingWriteOptions {
  readonly password?: string;
  readonly encryption?: EncryptOptions;
}

/**
 * A write's `encryption`, once the options are known to be an object whose `password` is a string of 1 to 255
 * characters: it comes with a password, and holds `EncryptOptions`.
 */
export function checkWriteEncryption(
  options: EncryptingWriteOptions | undefined,
  where: string
): void {
  const encryption = options?.encryption;
  if (encryption !== undefined && options?.password === undefined) {
    invalidArgument(where, "options.encryption", "set only with options.password", encryption);
  }
  checkEncryptOptions(encryption, "options.encryption", where);
}

/** `zip`, encrypted with `options.password` when one is given; the encryption code loads only then. */
export async function encryptIfRequested(
  zip: Uint8Array,
  options: EncryptingWriteOptions | undefined
): Promise<Uint8Array> {
  if (options?.password === undefined) {
    return zip;
  }
  const { encryptDocx } = await import("@word/security/encryption");
  return encryptDocx(zip, options.password, options.encryption);
}
