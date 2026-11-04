import { TarArchive } from "@archive/tar/tar-archive";
import type { ZipOptions, ZipOptionsTar, ZipOptionsZip } from "@archive/zip/zip-archive";
import { ZipArchive } from "@archive/zip/zip-archive";

/**
 * Create a ZIP or TAR builder, the format chosen by `options.format`. Kept for the
 * deprecated `Archive.zip`; `Archive.createZip` and `Archive.createTar` name the format.
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

/**
 * Create a ZIP builder. Published as `Archive.createZip`.
 *
 * @example
 * ```ts
 * import { Archive } from "documonster/archive";
 *
 * const archive = Archive.createZip({ level: 9 });
 * archive.add("file.txt", "content");
 * const bytes = await archive.toBuffer();
 * ```
 */
export function createZip(options: Omit<ZipOptions, "format"> = {}): ZipArchive {
  return new ZipArchive({ ...options, format: "zip" });
}
