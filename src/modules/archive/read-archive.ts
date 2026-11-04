import { checkReadOptions } from "@archive/core/entry-checks";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import { checkSourceArgument } from "@archive/io/archive-source";
import { TarReader } from "@archive/tar/tar-archive";
import type { UnzipOptions, UnzipOptionsTar, UnzipOptionsZip } from "@archive/unzip/zip-reader";
import { ZipReader } from "@archive/unzip/zip-reader";

/**
 * Open a ZIP or TAR archive for reading, the format chosen by `options.format`. Kept for
 * the deprecated `Archive.unzip`; `Archive.openZip` and `Archive.openTar` name the format.
 */
export function unzip(source: ArchiveSource, options: UnzipOptionsTar): TarReader;
export function unzip(source: ArchiveSource, options?: UnzipOptionsZip): ZipReader;
export function unzip(
  source: ArchiveSource,
  options: UnzipOptionsTar | UnzipOptionsZip = {}
): ZipReader | TarReader {
  if (options.format === "tar") {
    return new TarReader(source, {
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    });
  }

  return new ZipReader(source, { ...options, format: "zip" });
}

// `openZip` has no total-size limit: it reads one entry at a time
const OPEN_ZIP_LIMITS = ["maxEntrySize", "maxEntries"] as const;

/**
 * Open a ZIP archive for reading. Published as `Archive.openZip`.
 *
 * Over a one-shot source (a `ReadableStream`, an async iterable or a Node stream) the first call
 * decides how it is read. `list()`, `get()` or `bytes()` reads it into memory once, and every
 * later call, `entries()` included, uses that. Iterating `entries()`, `entriesStream()` or
 * `operation().iterable` streams it instead and spends it: any later call throws an
 * `ArchiveError`. Bytes, a string or a `Blob` are unaffected. See `ZipReader`.
 *
 * @example
 * ```ts
 * import { Archive } from "documonster/archive";
 *
 * const reader = Archive.openZip(zipBytes);
 * for await (const entry of reader.entries()) {
 *   console.log(entry.path);
 * }
 * ```
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function openZip(
  source: ArchiveSource | ArchiveBytes,
  options: Omit<UnzipOptions, "format"> = {}
): ZipReader {
  const where = "Archive.openZip";
  checkSourceArgument(where, source);
  checkReadOptions(where, options, OPEN_ZIP_LIMITS);
  return new ZipReader(source, { ...options, format: "zip" });
}
