import { TarArchive } from "@archive/tar/tar-archive";
import type { ZipOptionsTar, ZipOptionsZip } from "@archive/zip/zip-archive";
import { ZipArchive } from "@archive/zip/zip-archive";

/**
 * Create a new archive
 *
 * @param options - Archive options including format
 * @returns ZipArchive or TarArchive depending on format option
 *
 * @example
 * ```ts
 * // Create ZIP archive (default)
 * const zipArchive = zip();
 * zipArchive.add("file.txt", "content");
 * const zipBytes = await zipArchive.bytes();
 *
 * // Create TAR archive
 * const tarArchive = zip({ format: "tar" });
 * tarArchive.add("file.txt", "content");
 * const tarBytes = await tarArchive.bytes();
 * ```
 */
export function zip(options: ZipOptionsTar): TarArchive;
export function zip(options?: ZipOptionsZip): ZipArchive;
export function zip(options: ZipOptionsTar | ZipOptionsZip = {}): ZipArchive | TarArchive {
  if (options.format === "tar") {
    return new TarArchive({
      modTime: options.modTime,
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    });
  }

  return new ZipArchive({ ...options, format: "zip" });
}
