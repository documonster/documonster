import { checkZipEncryptionOptions, checkZipOptions } from "@archive/core/entry-checks";
import { ArchiveArgumentError } from "@archive/core/errors";
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
    // A TAR cannot be encrypted. The ZIP options it otherwise ignores change only how a ZIP is
    // written, but ignoring these would write in the clear an archive the caller asked to protect.
    const { encryptionMethod, password } = options as {
      encryptionMethod?: unknown;
      password?: unknown;
    };
    if (encryptionMethod !== undefined || password !== undefined) {
      throw new ArchiveArgumentError(
        encryptionMethod !== undefined ? "options.encryptionMethod" : "options.password",
        'Archive.zip: a TAR archive cannot be encrypted; encryption needs format "zip"'
      );
    }
    return new TarArchive({
      modTime: options.modTime,
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    });
  }

  return new ZipArchive(options);
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
export function createZip(options?: Omit<ZipOptions, "format">): ZipArchive {
  // Checked here too, so a refusal names the call that was made rather than the constructor.
  checkZipOptions("Archive.createZip", options);
  checkZipEncryptionOptions("Archive.createZip", options);
  return new ZipArchive(options);
}
