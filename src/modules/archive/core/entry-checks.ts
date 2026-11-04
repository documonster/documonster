/**
 * Checks for what an archive writer is asked to write: entry names, compression levels,
 * dates and modes. Each refusal is an {@link ArchiveArgumentError} naming the call and the field,
 * thrown when the entry is added rather than when the archive is streamed.
 */

import { ArchiveArgumentError } from "@archive/core/errors";
import { resolveZipStringCodec } from "@archive/core/text";
import { isArchiveSource } from "@archive/io/archive-source";
import { zipModTimeProblem } from "@archive/zip-spec/timestamps";
import { describeValue as describe } from "@utils/describe-value";

function fail(where: string, field: string, expected: string, value: unknown): never {
  throw new ArchiveArgumentError(
    field,
    `${where}: ${field} must be ${expected}, got ${describe(value)}`
  );
}

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
  if (typeof name !== "string" || name === "") {
    fail(where, "name", "a non-empty string", name);
  }
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
  if (typeof value !== "string" || value === "") {
    fail(where, field, "a non-empty string", value);
  }
}

/** Entry content held in memory: a `Uint8Array`. */
export function checkBytes(where: string, field: string, value: unknown): void {
  if (!(value instanceof Uint8Array)) {
    fail(where, field, "a Uint8Array", value);
  }
}

/** A string; `undefined` is accepted when `optional`. */
export function checkString(where: string, field: string, value: unknown, optional = false): void {
  if (typeof value !== "string" && !(optional && value === undefined)) {
    fail(where, field, "a string", value);
  }
}

/** Something an archive can read entry content from: an `ArchiveSource`. */
export function checkArchiveSource(where: string, field: string, value: unknown): void {
  if (!isArchiveSource(value)) {
    fail(
      where,
      field,
      "a Uint8Array, ArrayBuffer, string, Blob, ReadableStream or async iterable",
      value
    );
  }
}

/** A compression level: an integer from 0 (store) to 9. */
export function checkLevel(where: string, field: string, value: unknown): void {
  if (
    value !== undefined &&
    (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > 9)
  ) {
    fail(where, field, "an integer from 0 to 9", value);
  }
}

/** A valid `Date`. */
export function checkDate(where: string, field: string, value: unknown): void {
  if (value !== undefined && (!(value instanceof Date) || Number.isNaN(value.getTime()))) {
    fail(where, field, "a valid Date", value);
  }
}

/** An integer from 0 to `max`. */
export function checkUint(where: string, field: string, value: unknown, max: number): void {
  if (
    value !== undefined &&
    (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > max)
  ) {
    fail(where, field, `an integer from 0 to ${max}`, value);
  }
}

/** One of `allowed`. */
export function checkOneOf(
  where: string,
  field: string,
  value: unknown,
  allowed: readonly unknown[]
): void {
  if (value !== undefined && !allowed.includes(value)) {
    fail(where, field, `one of ${allowed.map(describe).join(", ")}`, value);
  }
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
  if (
    value !== undefined &&
    (typeof value !== "number" ||
      Number.isNaN(value) ||
      value < 0 ||
      (Number.isFinite(value) && !Number.isInteger(value)))
  ) {
    fail(where, field, "an integer from 0, or Infinity", value);
  }
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
  if (options !== undefined && (typeof options !== "object" || options === null)) {
    fail(where, field, "an options object", options);
  }
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
  if (options.comment !== undefined && typeof options.comment !== "string") {
    fail(where, "options.comment", "a string", options.comment);
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
  if (options.comment !== undefined && typeof options.comment !== "string") {
    fail(where, "options.comment", "a string", options.comment);
  }
}
