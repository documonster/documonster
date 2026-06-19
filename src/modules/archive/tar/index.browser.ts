/**
 * TAR Archive Module (Browser)
 *
 * Browser-compatible TAR support (excludes gzip which requires Node.js zlib).
 * Use CompressionStream API for gzip in browsers if needed.
 */

// Constants
export { TAR_BLOCK_SIZE, TAR_TYPE, type TarType } from "./tar-constants";

// Entry info
export {
  type TarEntryInfo,
  isFile as isTarFile,
  isDirectory as isTarDirectory,
  isSymlink as isTarSymlink,
  isHardLink as isTarHardLink,
  isDataEntry as isTarDataEntry
} from "./tar-entry-info";

// Header utilities
export {
  encodeHeader,
  decodeHeader,
  validateChecksum,
  isZeroBlock,
  calculatePadding,
  createPadding,
  createEndOfArchive,
  type TarHeaderOptions
} from "./tar-header";

// Parser (low-level)
export { parseTar, parseTarStream, untar, type TarEntry, type TarParseOptions } from "./tar-parser";

// Unified archive API (high-level, ZIP-compatible interface)
export {
  TarArchive,
  TarReader,
  createTarReaderEntry,
  type TarReaderEntry,
  createTarArchive,
  createTarReader,
  tar,
  tarSync,
  type TarArchiveOptions,
  type TarArchiveEntryOptions,
  type TarArchiveProgress,
  type TarArchiveStreamOptions,
  type TarArchiveOperation,
  type TarReaderOptions,
  type TarReaderProgress,
  type TarReaderStreamOptions,
  type TarReaderOperation
} from "./tar-archive";

// Note: Gzip support (TarGzArchive, targz, etc.) is NOT available in browser
// because it depends on Node.js zlib module.
