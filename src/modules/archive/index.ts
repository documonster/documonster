/**
 * Archive module — Node.js entry.
 *
 * Namespaces: `Archive` (ZIP / TAR / tar.gz, CRC-32, encryption, PNG), `Compression`
 * (DEFLATE / gzip / zlib, streams, worker pool) and `Fs` (Node file-system helpers).
 * Constants, error classes and types are flat.
 */

import {
  compress as compressImpl,
  compressSync as compressSyncImpl,
  decompress as decompressImpl,
  decompressSync as decompressSyncImpl,
  gzip as gzipImpl,
  gunzip as gunzipImpl,
  gzipSync as gzipSyncImpl,
  gunzipSync as gunzipSyncImpl,
  isGzipData as isGzipDataImpl,
  hasCompressionStream as hasCompressionStreamImpl,
  hasWorkerSupport as hasWorkerSupportImpl,
  hasGzipCompressionStream as hasGzipCompressionStreamImpl,
  hasGzipDecompressionStream as hasGzipDecompressionStreamImpl,
  zlib as zlibImpl,
  unzlib as unzlibImpl,
  zlibSync as zlibSyncImpl,
  unzlibSync as unzlibSyncImpl,
  isZlibData as isZlibDataImpl,
  decompressAuto as decompressAutoImpl,
  decompressAutoSync as decompressAutoSyncImpl,
  detectCompressionFormat as detectCompressionFormatImpl
} from "@archive/compression/compress";
import {
  crc32 as crc32Impl,
  crc32Update as crc32UpdateImpl,
  crc32Finalize as crc32FinalizeImpl
} from "@archive/compression/crc32";
import {
  createDeflateStream as createDeflateStreamImpl,
  createInflateStream as createInflateStreamImpl,
  createGzipStream as createGzipStreamImpl,
  createGunzipStream as createGunzipStreamImpl,
  createZlibStream as createZlibStreamImpl,
  createUnzlibStream as createUnzlibStreamImpl,
  hasDeflateRaw as hasDeflateRawImpl
} from "@archive/compression/streaming-compress";
import type {
  TaskOptions as WorkerTaskOptionsImpl,
  TaskResult as WorkerTaskResultImpl
} from "@archive/compression/worker-pool";
import {
  WorkerPool as WorkerPoolImpl,
  getDefaultWorkerPool as getDefaultWorkerPoolImpl,
  terminateDefaultWorkerPool as terminateDefaultWorkerPoolImpl,
  deflateWithPool as deflateWithPoolImpl,
  inflateWithPool as inflateWithPoolImpl,
  deflateBatchWithPool as deflateBatchWithPoolImpl,
  inflateBatchWithPool as inflateBatchWithPoolImpl
} from "@archive/compression/worker-pool";
import type { ArchiveFormat } from "@archive/core/types";
import { ArchiveFile as ArchiveFileImpl } from "@archive/fs/archive-file";
import { toNodeReadable as toNodeReadableImpl } from "@archive/io/archive-source.node";
import {
  TarGzArchive as TarGzArchiveImpl,
  targz as targzImpl,
  untargz as untargzImpl,
  parseTarGz as parseTarGzImpl,
  parseTarGzStream as parseTarGzStreamImpl
} from "@archive/tar/tar-gzip";
import type {
  FileEntry as FsFileEntryImpl,
  TraverseOptions as FsTraverseOptionsImpl,
  GlobOptions as FsGlobOptionsImpl,
  ReadStreamOptions as FsReadStreamOptionsImpl,
  WriteStreamOptions as FsWriteStreamOptionsImpl
} from "@utils/fs";
import {
  traverseDirectory as traverseDirectoryImpl,
  traverseDirectorySync as traverseDirectorySyncImpl,
  glob as globImpl,
  globSync as globSyncImpl,
  globToRegex as globToRegexImpl,
  matchGlob as matchGlobImpl,
  matchGlobAny as matchGlobAnyImpl,
  ensureDir as ensureDirImpl,
  ensureDirSync as ensureDirSyncImpl,
  fileExists as fileExistsImpl,
  fileExistsSync as fileExistsSyncImpl,
  readFileBytes as readFileBytesImpl,
  readFileBytesSync as readFileBytesSyncImpl,
  writeFileBytes as writeFileBytesImpl,
  writeFileBytesSync as writeFileBytesSyncImpl,
  setFileTime as setFileTimeImpl,
  setFileTimeSync as setFileTimeSyncImpl,
  safeStats as safeStatsImpl,
  safeStatsSync as safeStatsSyncImpl,
  readFileText as readFileTextImpl,
  readFileTextSync as readFileTextSyncImpl,
  writeFileText as writeFileTextImpl,
  writeFileTextSync as writeFileTextSyncImpl,
  remove as removeImpl,
  removeSync as removeSyncImpl,
  copyFile as copyFileImpl,
  copyFileSync as copyFileSyncImpl,
  createReadStream as createReadStreamImpl,
  createWriteStream as createWriteStreamImpl,
  createTempDir as createTempDirImpl,
  createTempDirSync as createTempDirSyncImpl
} from "@utils/fs";

export * from "@archive/index.base";

export * as Archive from "@archive/surface/archive-node";
export * as Compression from "@archive/surface/compression";
export * as Fs from "@archive/surface/fs";

// GZIP / zlib constants (RFC 1952 / RFC 1950)
export {
  GZIP_ID1,
  GZIP_ID2,
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE,
  type CompressOptions
} from "@archive/compression/compress";
export type {
  StreamCompressOptions,
  StreamingCodec,
  DeflateStream,
  InflateStream,
  GzipStream,
  GunzipStream,
  ZlibStream,
  UnzlibStream
} from "@archive/compression/streaming-compress";
export type {
  WorkerPoolOptions,
  WorkerPoolStats,
  WorkerPoolStream,
  TaskOptions as WorkerTaskOptions,
  TaskResult as WorkerTaskResult,
  WorkerTaskType
} from "@archive/compression/worker-pool";

// Node file-system layer types
export type {
  OverwriteStrategy,
  ArchiveEntryInfo,
  AddFileOptions,
  AddDirectoryOptions,
  AddGlobOptions,
  ExtractToOptions,
  ExtractProgress,
  ZipFileOptions,
  WriteZipOptions,
  WriteArchiveOptions,
  ZipEntryInfo,
  TarFileOptions,
  AddTarFileOptions,
  AddTarDirectoryOptions,
  AddTarGlobOptions,
  ArchiveFileOptions,
  ArchiveFileOptionsZip,
  ArchiveFileOptionsTar,
  ArchiveFileOpenOptions,
  ArchiveFileReadOptions,
  ArchiveFileReadTextOptions,
  ArchiveFileOpenOptionsZip,
  ArchiveFileOpenOptionsTar,
  ArchiveStreamOptions,
  ArchiveStreamProgress,
  ArchiveStreamOperation,
  ArchiveStreamPhase,
  ArchiveWarning,
  ArchiveTransformFunction,
  ArchiveTransformEntryData,
  ArchiveTransformResult,
  OpenZipOptions,
  OpenTarOptions,
  OpenArchiveOptions,
  OpenArchiveOptionsZip,
  OpenArchiveOptionsTar,
  TransformFunction,
  TransformEntryData,
  TransformResult
} from "@archive/fs/types";
export type {
  FileEntry as FsFileEntry,
  TraverseOptions as FsTraverseOptions,
  GlobOptions as FsGlobOptions,
  ReadStreamOptions as FsReadStreamOptions,
  WriteStreamOptions as FsWriteStreamOptions
} from "@utils/fs";
/** @deprecated Use {@link FsFileEntry}. Removed in 1.0. */
export type FileEntry = FsFileEntryImpl;
/** @deprecated Use {@link FsTraverseOptions}. Removed in 1.0. */
export type TraverseOptions = FsTraverseOptionsImpl;
/** @deprecated Use {@link FsGlobOptions}. Removed in 1.0. */
export type GlobOptions = FsGlobOptionsImpl;
/** @deprecated Use {@link FsReadStreamOptions}. Removed in 1.0. */
export type ReadStreamOptions = FsReadStreamOptionsImpl;
/** @deprecated Use {@link FsWriteStreamOptions}. Removed in 1.0. */
export type WriteStreamOptions = FsWriteStreamOptionsImpl;

/** @deprecated Use {@link WorkerTaskOptions}. Removed in 1.0. */
export type TaskOptions = WorkerTaskOptionsImpl;
/** @deprecated Use {@link WorkerTaskResult}. Removed in 1.0. */
export type TaskResult = WorkerTaskResultImpl;

// =============================================================================
// Deprecated flat forwards — use the `Archive`, `Compression` and `Fs` namespaces
// =============================================================================

/** @deprecated Use {@link Archive.crc32}. Removed in 1.0. */
export const crc32 = crc32Impl;
/** @deprecated Use {@link Archive.crc32Update}. Removed in 1.0. */
export const crc32Update = crc32UpdateImpl;
/** @deprecated Use {@link Archive.crc32Finalize}. Removed in 1.0. */
export const crc32Finalize = crc32FinalizeImpl;
/** @deprecated Use {@link Compression.deflate}. Removed in 1.0. */
export const compress = compressImpl;
/** @deprecated Use {@link Compression.deflateSync}. Removed in 1.0. */
export const compressSync = compressSyncImpl;
/** @deprecated Use {@link Compression.inflate}. Removed in 1.0. */
export const decompress = decompressImpl;
/** @deprecated Use {@link Compression.inflateSync}. Removed in 1.0. */
export const decompressSync = decompressSyncImpl;
/** @deprecated Use {@link Compression.gzip}. Removed in 1.0. */
export const gzip = gzipImpl;
/** @deprecated Use {@link Compression.gunzip}. Removed in 1.0. */
export const gunzip = gunzipImpl;
/** @deprecated Use {@link Compression.gzipSync}. Removed in 1.0. */
export const gzipSync = gzipSyncImpl;
/** @deprecated Use {@link Compression.gunzipSync}. Removed in 1.0. */
export const gunzipSync = gunzipSyncImpl;
/** @deprecated Use {@link Compression.isGzipData}. Removed in 1.0. */
export const isGzipData = isGzipDataImpl;
/** @deprecated Use {@link Compression.hasCompressionStream}. Removed in 1.0. */
export const hasCompressionStream = hasCompressionStreamImpl;
/** @deprecated Use {@link Compression.hasWorkerSupport}. Removed in 1.0. */
export const hasWorkerSupport = hasWorkerSupportImpl;
/** @deprecated Use {@link Compression.hasGzipCompressionStream}. Removed in 1.0. */
export const hasGzipCompressionStream = hasGzipCompressionStreamImpl;
/** @deprecated Use {@link Compression.hasGzipDecompressionStream}. Removed in 1.0. */
export const hasGzipDecompressionStream = hasGzipDecompressionStreamImpl;
/** @deprecated Use {@link Compression.zlib}. Removed in 1.0. */
export const zlib = zlibImpl;
/** @deprecated Use {@link Compression.unzlib}. Removed in 1.0. */
export const unzlib = unzlibImpl;
/** @deprecated Use {@link Compression.zlibSync}. Removed in 1.0. */
export const zlibSync = zlibSyncImpl;
/** @deprecated Use {@link Compression.unzlibSync}. Removed in 1.0. */
export const unzlibSync = unzlibSyncImpl;
/** @deprecated Use {@link Compression.isZlibData}. Removed in 1.0. */
export const isZlibData = isZlibDataImpl;
/** @deprecated Use {@link Compression.decompressAuto}. Removed in 1.0. */
export const decompressAuto = decompressAutoImpl;
/** @deprecated Use {@link Compression.decompressAutoSync}. Removed in 1.0. */
export const decompressAutoSync = decompressAutoSyncImpl;
/** @deprecated Use {@link Compression.detectFormat}. Removed in 1.0. */
export const detectCompressionFormat = detectCompressionFormatImpl;
/** @deprecated Use {@link Compression.createDeflateStream}. Removed in 1.0. */
export const createDeflateStream = createDeflateStreamImpl;
/** @deprecated Use {@link Compression.createInflateStream}. Removed in 1.0. */
export const createInflateStream = createInflateStreamImpl;
/** @deprecated Use {@link Compression.createGzipStream}. Removed in 1.0. */
export const createGzipStream = createGzipStreamImpl;
/** @deprecated Use {@link Compression.createGunzipStream}. Removed in 1.0. */
export const createGunzipStream = createGunzipStreamImpl;
/** @deprecated Use {@link Compression.createZlibStream}. Removed in 1.0. */
export const createZlibStream = createZlibStreamImpl;
/** @deprecated Use {@link Compression.createUnzlibStream}. Removed in 1.0. */
export const createUnzlibStream = createUnzlibStreamImpl;
/** @deprecated Use {@link Compression.hasDeflateRaw}. Removed in 1.0. */
export const hasDeflateRaw = hasDeflateRawImpl;
/** @deprecated Use {@link Compression.Workers.WorkerPool}. Removed in 1.0. */
export const WorkerPool = WorkerPoolImpl;
/** @deprecated Use {@link Compression.Workers.WorkerPool}. Removed in 1.0. */
export type WorkerPool = WorkerPoolImpl;
/** @deprecated Use {@link Compression.Workers.getDefaultPool}. Removed in 1.0. */
export const getDefaultWorkerPool = getDefaultWorkerPoolImpl;
/** @deprecated Use {@link Compression.Workers.terminateDefaultPool}. Removed in 1.0. */
export const terminateDefaultWorkerPool = terminateDefaultWorkerPoolImpl;
/** @deprecated Use {@link Compression.Workers.deflate}. Removed in 1.0. */
export const deflateWithPool = deflateWithPoolImpl;
/** @deprecated Use {@link Compression.Workers.inflate}. Removed in 1.0. */
export const inflateWithPool = inflateWithPoolImpl;
/** @deprecated Use {@link Compression.Workers.deflateBatch}. Removed in 1.0. */
export const deflateBatchWithPool = deflateBatchWithPoolImpl;
/** @deprecated Use {@link Compression.Workers.inflateBatch}. Removed in 1.0. */
export const inflateBatchWithPool = inflateBatchWithPoolImpl;
/** @deprecated Use {@link Fs.traverseDirectory}. Removed in 1.0. */
export const traverseDirectory = traverseDirectoryImpl;
/** @deprecated Use {@link Fs.traverseDirectorySync}. Removed in 1.0. */
export const traverseDirectorySync = traverseDirectorySyncImpl;
/** @deprecated Use {@link Fs.glob}. Removed in 1.0. */
export const glob = globImpl;
/** @deprecated Use {@link Fs.globSync}. Removed in 1.0. */
export const globSync = globSyncImpl;
/** @deprecated Use {@link Fs.globToRegex}. Removed in 1.0. */
export const globToRegex = globToRegexImpl;
/** @deprecated Use {@link Fs.matchGlob}. Removed in 1.0. */
export const matchGlob = matchGlobImpl;
/** @deprecated Use {@link Fs.matchGlobAny}. Removed in 1.0. */
export const matchGlobAny = matchGlobAnyImpl;
/** @deprecated Use {@link Fs.ensureDir}. Removed in 1.0. */
export const ensureDir = ensureDirImpl;
/** @deprecated Use {@link Fs.ensureDirSync}. Removed in 1.0. */
export const ensureDirSync = ensureDirSyncImpl;
/** @deprecated Use {@link Fs.fileExists}. Removed in 1.0. */
export const fileExists = fileExistsImpl;
/** @deprecated Use {@link Fs.fileExistsSync}. Removed in 1.0. */
export const fileExistsSync = fileExistsSyncImpl;
/** @deprecated Use {@link Fs.readFileBytes}. Removed in 1.0. */
export const readFileBytes = readFileBytesImpl;
/** @deprecated Use {@link Fs.readFileBytesSync}. Removed in 1.0. */
export const readFileBytesSync = readFileBytesSyncImpl;
/** @deprecated Use {@link Fs.writeFileBytes}. Removed in 1.0. */
export const writeFileBytes = writeFileBytesImpl;
/** @deprecated Use {@link Fs.writeFileBytesSync}. Removed in 1.0. */
export const writeFileBytesSync = writeFileBytesSyncImpl;
/** @deprecated Use {@link Fs.setFileTime}. Removed in 1.0. */
export const setFileTime = setFileTimeImpl;
/** @deprecated Use {@link Fs.setFileTimeSync}. Removed in 1.0. */
export const setFileTimeSync = setFileTimeSyncImpl;
/** @deprecated Use {@link Fs.safeStats}. Removed in 1.0. */
export const safeStats = safeStatsImpl;
/** @deprecated Use {@link Fs.safeStatsSync}. Removed in 1.0. */
export const safeStatsSync = safeStatsSyncImpl;
/** @deprecated Use {@link Fs.readFileText}. Removed in 1.0. */
export const readFileText = readFileTextImpl;
/** @deprecated Use {@link Fs.readFileTextSync}. Removed in 1.0. */
export const readFileTextSync = readFileTextSyncImpl;
/** @deprecated Use {@link Fs.writeFileText}. Removed in 1.0. */
export const writeFileText = writeFileTextImpl;
/** @deprecated Use {@link Fs.writeFileTextSync}. Removed in 1.0. */
export const writeFileTextSync = writeFileTextSyncImpl;
/** @deprecated Use {@link Fs.remove}. Removed in 1.0. */
export const remove = removeImpl;
/** @deprecated Use {@link Fs.removeSync}. Removed in 1.0. */
export const removeSync = removeSyncImpl;
/** @deprecated Use {@link Fs.copyFile}. Removed in 1.0. */
export const copyFile = copyFileImpl;
/** @deprecated Use {@link Fs.copyFileSync}. Removed in 1.0. */
export const copyFileSync = copyFileSyncImpl;
/** @deprecated Use {@link Fs.createReadStream}. Removed in 1.0. */
export const createReadStream = createReadStreamImpl;
/** @deprecated Use {@link Fs.createWriteStream}. Removed in 1.0. */
export const createWriteStream = createWriteStreamImpl;
/** @deprecated Use {@link Fs.createTempDir}. Removed in 1.0. */
export const createTempDir = createTempDirImpl;
/** @deprecated Use {@link Fs.createTempDirSync}. Removed in 1.0. */
export const createTempDirSync = createTempDirSyncImpl;
/**
 * @deprecated Use {@link Archive.ArchiveFile}. Removed in 1.0.
 */
export const ArchiveFile = ArchiveFileImpl;
/** @deprecated Use {@link Archive.ArchiveFile}. Removed in 1.0. */
export type ArchiveFile<F extends ArchiveFormat = "zip"> = ArchiveFileImpl<F>;
/** @deprecated Use {@link Archive.toNodeReadable}. Removed in 1.0. */
export const toNodeReadable = toNodeReadableImpl;
/**
 * @deprecated Use {@link Archive.TarGzArchive} (or build one with
 * {@link Archive.createTarGz}). Removed in 1.0.
 */
export const TarGzArchive = TarGzArchiveImpl;
/** @deprecated Use {@link Archive.TarGzArchive}. Removed in 1.0. */
export type TarGzArchive = TarGzArchiveImpl;
/** @deprecated Use {@link Archive.packTarGz}. Removed in 1.0. */
export const targz = targzImpl;
/** @deprecated Use {@link Archive.unpackTarGz}. Removed in 1.0. */
export const untargz = untargzImpl;
/** @deprecated Use {@link Archive.parseTarGz}. Removed in 1.0. */
export const parseTarGz = parseTarGzImpl;
/**
 * @deprecated Use {@link Archive.parseTarGzStream}, which decompresses and parses together
 * in bounded memory; this one decompresses the whole archive before the first entry. For
 * that all-or-nothing behaviour iterate {@link Archive.parseTarGz}. Removed in 1.0.
 */
export const parseTarGzStream = parseTarGzStreamImpl;
