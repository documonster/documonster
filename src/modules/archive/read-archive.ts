import type { ArchiveSource } from "@archive/io/archive-source";
import { TarReader } from "@archive/tar/tar-archive";
import type { UnzipOptions, UnzipOptionsTar, UnzipOptionsZip } from "@archive/unzip/zip-reader";
import { ZipReader } from "@archive/unzip/zip-reader";

/**
 * Open an archive for reading
 *
 * @param source - Archive data source
 * @param options - Options including format
 * @returns ZipReader or TarReader depending on format option
 *
 * @example
 * ```ts
 * // Read ZIP archive (default)
 * const zipReader = unzip(zipBytes);
 * for await (const entry of zipReader.entries()) {
 *   console.log(entry.path);
 * }
 *
 * // Read TAR archive
 * const tarReader = unzip(tarBytes, { format: "tar" });
 * for await (const entry of tarReader.entries()) {
 *   console.log(entry.path);
 * }
 * ```
 */
export function unzip(source: ArchiveSource, options: UnzipOptionsTar): TarReader;
export function unzip(source: ArchiveSource, options?: UnzipOptionsZip): ZipReader;
export function unzip(source: ArchiveSource, options: UnzipOptions = {}): ZipReader | TarReader {
  if (options.format === "tar") {
    return new TarReader(source, {
      signal: options.signal,
      onProgress: options.onProgress as any,
      progressIntervalMs: options.progressIntervalMs
    });
  }

  return new ZipReader(source, { ...options, format: "zip" });
}
