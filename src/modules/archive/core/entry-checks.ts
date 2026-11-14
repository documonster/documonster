/**
 * Checks for what an archive writer is asked to write: entry names, compression levels,
 * dates and modes. Each refusal is an {@link ArchiveArgumentError} naming the call and the field,
 * thrown when the entry is added rather than when the archive is streamed.
 */

import {
  OPTIONS,
  checkArchiveArgument,
  dateSchema,
  levelSchema,
  limitSchema,
  uintSchema
} from "@archive/core/arguments";
import { ArchiveArgumentError } from "@archive/core/errors";
import { resolveZipStringCodec } from "@archive/core/text";
import { isArchiveSource } from "@archive/io/archive-source";
import { zipModTimeProblem } from "@archive/zip-spec/timestamps";
import { isUint8Array } from "@utils/binary";
import { describeValue as describe } from "@utils/describe-value";
import { s } from "@utils/schema";

const nonEmptyString = s.string({ nonEmpty: true });

/** A `Uint8Array` from any realm: one made in a worker, an iframe or a test DOM fails `instanceof`. */
const uint8ArraySchema = s.custom<Uint8Array>(isUint8Array, "a Uint8Array");

export const archiveSourceSchema = s.custom(
  isArchiveSource,
  "a Uint8Array, ArrayBuffer, string, Blob, ReadableStream or async iterable"
);

/**
 * An entry name another tool can extract without writing outside its destination.
 *
 * An absolute name (`/etc/passwd`), a drive prefix (`C:x`) or one whose `..` segments climb
 * above the archive root (`../evil`, `a/../../evil`) is the zip-slip pattern: nothing a
 * document toolkit produces needs one, and an extractor that trusts names writes wherever it
 * says. Backslashes count as separators, as Windows extractors treat them. A `..` that stays
 * inside (`a/../b`) is kept as written.
 *
 * Writers that take a `path` option skip this when one is given: an explicit naming mode
 * (`"legacy"`, `"posix"`, `"safe"`) decides instead.
 */
export function checkEntryName(where: string, name: unknown): void {
  checkArchiveArgument(nonEmptyString, name, where, "name");
  const slashed = name.replace(/\\/g, "/");
  if (slashed.startsWith("/") || /^[a-zA-Z]:/.test(slashed)) {
    throw new ArchiveArgumentError(
      "name",
      `${where}: name ${describe(name)} is an absolute path; entry names are relative to the archive root`
    );
  }
  let depth = 0;
  for (const segment of slashed.split("/")) {
    if (segment === "..") {
      depth--;
      if (depth < 0) {
        throw new ArchiveArgumentError(
          "name",
          `${where}: name ${describe(name)} climbs above the archive root with ".."`
        );
      }
    } else if (segment !== "" && segment !== ".") {
      depth++;
    }
  }
}

/**
 * The name a reader is asked to look an entry up by: a non-empty string. A number or `undefined`
 * answered "no such entry" from some readers and failed with a raw `TypeError` in others.
 */
export function checkEntryPath(where: string, value: unknown, field = "path"): void {
  checkArchiveArgument(nonEmptyString, value, where, field);
}

/** Entry content held in memory: a `Uint8Array`. */
export function checkBytes(where: string, field: string, value: unknown): void {
  checkArchiveArgument(uint8ArraySchema, value, where, field);
}

/** A string; `undefined` is accepted when `optional`. */
export function checkString(where: string, field: string, value: unknown, optional = false): void {
  checkArchiveArgument(optional ? s.string().optional() : s.string(), value, where, field);
}

/** Something an archive can read entry content from: an `ArchiveSource`. */
export function checkArchiveSource(where: string, field: string, value: unknown): void {
  checkArchiveArgument(archiveSourceSchema, value, where, field);
}

/** A compression level: an integer from 0 (store) to 9. */
export function checkLevel(where: string, field: string, value: unknown): void {
  checkArchiveArgument(levelSchema.optional(), value, where, field);
}

/** A valid `Date`. */
export function checkDate(where: string, field: string, value: unknown): void {
  checkArchiveArgument(dateSchema.optional(), value, where, field);
}

/** An integer from 0 to `max`. */
export function checkUint(where: string, field: string, value: unknown, max: number): void {
  checkArchiveArgument(uintSchema(max).optional(), value, where, field);
}

/** One of `allowed`. */
export function checkOneOf(
  where: string,
  field: string,
  value: unknown,
  allowed: readonly unknown[]
): void {
  checkArchiveArgument(s.oneOf(allowed).optional(), value, where, field);
}

const TIMESTAMP_MODES = ["dos", "dos+utc", "dos+utc+ntfs"] as const;
const ZIP64_MODES = [true, false, "auto"] as const;

/** A ZIP name encoding: `"utf-8"`, `"cp437"` or a codec. */
export function checkZipEncoding(where: string, value: unknown, field = "options.encoding"): void {
  if (value === undefined) {
    return;
  }
  try {
    resolveZipStringCodec(value as Parameters<typeof resolveZipStringCodec>[0]);
  } catch (cause) {
    throw new ArchiveArgumentError(field, `${where}: ${(cause as Error).message}`, { cause });
  }
}

/**
 * An entry modification time a ZIP can store exactly (see `zipModTimeProblem`): refused
 * rather than wrapped or clamped, since the caller chose it.
 */
export function checkZipModTime(where: string, field: string, value: unknown): void {
  checkDate(where, field, value);
  if (value !== undefined) {
    const problem = zipModTimeProblem(value as Date);
    if (problem) {
      throw new ArchiveArgumentError(field, `${where}: ${field} ${problem}`);
    }
  }
}

/** A size or count limit: a number from 0, `Infinity` meaning none. */
export function checkLimit(where: string, field: string, value: unknown): void {
  checkArchiveArgument(limitSchema.optional(), value, where, field);
}

/** The size and count limits a reader takes, by format. */
export const ZIP_READ_LIMITS = ["maxEntrySize", "maxEntries", "maxTotalUncompressedSize"] as const;
export const TAR_READ_LIMITS = ["maxEntrySize", "maxFileSize"] as const;

/**
 * A reader's options: an object or `undefined`, with each of `limits` it sets checked by
 * {@link checkLimit}. Every reader runs this where it is called, so a mistyped limit is refused
 * rather than read as no limit at all.
 */
export function checkReadOptions(where: string, options: unknown, limits: readonly string[]): void {
  checkOptionsObject(where, options);
  if (options === undefined) {
    return;
  }
  for (const limit of limits) {
    checkLimit(where, `options.${limit}`, (options as Record<string, unknown>)[limit]);
  }
}

/** An options argument: an object or `undefined`. */
export function checkOptionsObject(where: string, options: unknown, field = "options"): void {
  checkArchiveArgument(s.object({}, OPTIONS).optional(), options, where, field);
}

/** Archive-level ZIP options shared by `createZip`, `packZip` and `editZip`. */
export function checkZipOptions(
  where: string,
  options:
    | {
        level?: unknown;
        modTime?: unknown;
        timestamps?: unknown;
        zip64?: unknown;
        comment?: unknown;
        encoding?: unknown;
      }
    | undefined
): void {
  checkOptionsObject(where, options);
  if (options === undefined) {
    return;
  }
  checkLevel(where, "options.level", options.level);
  checkZipEncoding(where, options.encoding);
  checkOneOf(where, "options.timestamps", options.timestamps, TIMESTAMP_MODES);
  checkZipModTime(where, "options.modTime", options.modTime);
  checkOneOf(where, "options.zip64", options.zip64, ZIP64_MODES);
  checkString(where, "options.comment", options.comment, true);
}

const ZIP_ENCRYPTION_METHODS = ["none", "zipcrypto", "aes-128", "aes-192", "aes-256"] as const;

const zipPasswordSchema = s.union(
  [nonEmptyString, uint8ArraySchema.refine(v => v.length > 0, "a non-empty Uint8Array")],
  { expected: "a non-empty string or Uint8Array" }
);

/**
 * The encryption a ZIP builder is asked for (`Archive.createZip`, `Archive.packZip`): a method it
 * knows, and a non-empty password exactly when the method is not `"none"`. A password with no
 * method is refused too — it encrypted nothing, so an archive the caller meant to protect was
 * written in the clear.
 */
export function checkZipEncryptionOptions(
  where: string,
  options: { encryptionMethod?: unknown; password?: unknown } | undefined
): void {
  if (options === undefined) {
    return;
  }
  checkOneOf(where, "options.encryptionMethod", options.encryptionMethod, ZIP_ENCRYPTION_METHODS);
  const method = options.encryptionMethod ?? "none";
  const password = options.password;
  if (password !== undefined) {
    checkArchiveArgument(zipPasswordSchema, password, where, "options.password");
  }
  if (method === "none" && password !== undefined) {
    throw new ArchiveArgumentError(
      "options.password",
      `${where}: options.password encrypts nothing without options.encryptionMethod ("zipcrypto", "aes-128", "aes-192" or "aes-256")`
    );
  }
  if (method !== "none" && password === undefined) {
    throw new ArchiveArgumentError(
      "options.password",
      `${where}: options.encryptionMethod ${describe(method)} needs options.password`
    );
  }
}

/** Per-entry ZIP options. */
export function checkZipEntryOptions(
  where: string,
  options:
    | {
        level?: unknown;
        modTime?: unknown;
        atime?: unknown;
        ctime?: unknown;
        birthTime?: unknown;
        mode?: unknown;
        zip64?: unknown;
        comment?: unknown;
        encoding?: unknown;
      }
    | undefined
): void {
  checkOptionsObject(where, options);
  if (options === undefined) {
    return;
  }
  checkLevel(where, "options.level", options.level);
  checkZipEncoding(where, options.encoding);
  for (const field of ["atime", "ctime", "birthTime"] as const) {
    checkDate(where, `options.${field}`, options[field]);
  }
  checkZipModTime(where, "options.modTime", options.modTime);
  // A Unix `st_mode`, type bits included, lives in the high 16 bits of the external attributes.
  checkUint(where, "options.mode", options.mode, 0xffff);
  checkOneOf(where, "options.zip64", options.zip64, ZIP64_MODES);
  checkString(where, "options.comment", options.comment, true);
}
