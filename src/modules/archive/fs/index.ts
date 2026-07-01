/**
 * Node.js file system convenience layer for archive operations.
 *
 * This module provides a high-level API for working with ZIP and TAR files
 * on the file system.
 *
 * @example Create a ZIP from files and directories
 * ```ts
 * import { ArchiveFile } from "@archive/fs";
 *
 * const zip = new ArchiveFile();
 * zip.addFile("./readme.md");
 * zip.addDirectory("./src");
 * zip.addGlob("**\/*.json", { cwd: "./config" });
 * await zip.writeToFile("./output.zip");
 * ```
 *
 * @example Create a TAR archive
 * ```ts
 * import { ArchiveFile } from "@archive/fs";
 *
 * const tar = new ArchiveFile({ format: "tar" });
 * tar.addFile("./readme.md");
 * tar.addDirectory("./src");
 * await tar.writeToFile("./output.tar");
 * ```
 *
 * @example Create a gzipped TAR archive
 * ```ts
 * import { ArchiveFile } from "@archive/fs";
 *
 * const tar = new ArchiveFile({ format: "tar", gzip: true });
 * tar.addFile("./readme.md");
 * await tar.writeToFile("./output.tar.gz");
 * ```
 *
 * @example Extract a ZIP file
 * ```ts
 * import { ArchiveFile } from "@archive/fs";
 *
 * const zip = await ArchiveFile.fromFile("./archive.zip");
 * await zip.extractTo("./output", { overwrite: "newer" });
 * ```
 *
 * @example Read archive contents
 * ```ts
 * import { ArchiveFile } from "@archive/fs";
 *
 * const archive = await ArchiveFile.fromFile("./archive.zip");
 * for (const entry of archive.getEntries()) {
 *   console.log(entry.path, entry.size);
 * }
 * const content = await archive.readAsText("readme.txt");
 * ```
 *
 * @module
 */

// Main unified class
export { ArchiveFile } from "@archive/fs/archive-file";

// Types
export type {
  // Common types
  OverwriteStrategy,
  ArchiveFormat,
  ArchiveEntryInfo,
  // ZIP types
  AddFileOptions,
  AddDirectoryOptions,
  AddGlobOptions,
  ExtractToOptions,
  ExtractProgress,
  ZipFileOptions,
  OpenZipOptions,
  WriteZipOptions,
  WriteArchiveOptions,
  ZipEntryInfo,
  // TAR types
  TarFileOptions,
  OpenTarOptions,
  AddTarFileOptions,
  AddTarDirectoryOptions,
  AddTarGlobOptions,
  // Unified options (with format discrimination)
  ArchiveFileOptions,
  ArchiveFileOptionsZip,
  ArchiveFileOptionsTar,
  OpenArchiveOptions,
  OpenArchiveOptionsZip,
  OpenArchiveOptionsTar,
  // Streaming types
  ArchiveStreamOptions,
  ArchiveStreamProgress,
  ArchiveStreamOperation,
  ArchiveStreamPhase,
  // Warning types
  ArchiveWarning,
  // Transform function types
  TransformFunction,
  TransformEntryData,
  TransformResult
} from "@archive/fs/types";

// Re-export ArchiveSink for pipeTo() usage
export type { ArchiveSink } from "@archive/io/archive-sink";

// File system utilities (for advanced users)
export {
  traverseDirectory,
  traverseDirectorySync,
  glob,
  globSync,
  globToRegex,
  matchGlob,
  matchGlobAny,
  ensureDir,
  ensureDirSync,
  fileExists,
  fileExistsSync,
  readFileBytes,
  readFileBytesSync,
  writeFileBytes,
  writeFileBytesSync,
  setFileTime,
  setFileTimeSync,
  safeStats,
  safeStatsSync,
  readFileText,
  readFileTextSync,
  writeFileText,
  writeFileTextSync,
  remove,
  removeSync,
  copyFile,
  copyFileSync,
  createReadStream,
  createWriteStream,
  createTempDir,
  createTempDirSync,
  type FileEntry,
  type TraverseOptions,
  type GlobOptions,
  type ReadStreamOptions,
  type WriteStreamOptions
} from "@utils/fs";
