/**
 * Archive module — browser entry.
 *
 * Mirrors `index.ts` without the Node-only `Fs` namespace and the Node-only
 * `Archive.ArchiveFile` / `Archive.toNodeReadable` members.
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
} from "@archive/compression/compress.browser";
import {
  crc32 as crc32Impl,
  crc32Update as crc32UpdateImpl,
  crc32Finalize as crc32FinalizeImpl
} from "@archive/compression/crc32.browser";
import {
  createDeflateStream as createDeflateStreamImpl,
  createInflateStream as createInflateStreamImpl,
  createGzipStream as createGzipStreamImpl,
  createGunzipStream as createGunzipStreamImpl,
  createZlibStream as createZlibStreamImpl,
  createUnzlibStream as createUnzlibStreamImpl,
  hasDeflateRaw as hasDeflateRawImpl
} from "@archive/compression/streaming-compress.browser";
import type {
  TaskOptions as WorkerTaskOptionsImpl,
  TaskResult as WorkerTaskResultImpl
} from "@archive/compression/worker-pool/index.browser";
import {
  WorkerPool as WorkerPoolImpl,
  getDefaultWorkerPool as getDefaultWorkerPoolImpl,
  terminateDefaultWorkerPool as terminateDefaultWorkerPoolImpl,
  deflateWithPool as deflateWithPoolImpl,
  inflateWithPool as inflateWithPoolImpl,
  deflateBatchWithPool as deflateBatchWithPoolImpl,
  inflateBatchWithPool as inflateBatchWithPoolImpl
} from "@archive/compression/worker-pool/index.browser";

export * from "@archive/index.base";

export * as Archive from "@archive/surface/archive";
export * as Compression from "@archive/surface/compression";

// GZIP / zlib constants (RFC 1952 / RFC 1950)
export {
  GZIP_ID1,
  GZIP_ID2,
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE,
  type CompressOptions
} from "@archive/compression/compress.browser";
export type {
  StreamCompressOptions,
  StreamingCodec,
  DeflateStream,
  InflateStream,
  GzipStream,
  GunzipStream,
  ZlibStream,
  UnzlibStream
} from "@archive/compression/streaming-compress.browser";
export type {
  WorkerPoolOptions,
  WorkerPoolStats,
  WorkerPoolStream,
  TaskOptions as WorkerTaskOptions,
  TaskResult as WorkerTaskResult,
  WorkerTaskType
} from "@archive/compression/worker-pool/index.browser";

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
