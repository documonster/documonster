/**
 * Browser compression utilities
 *
 * Supports multiple formats:
 * - deflate-raw: Raw DEFLATE (for ZIP files)
 * - gzip: GZIP format (for tar.gz, HTTP compression)
 *
 * Primary: CompressionStream API (Chrome 103+, Firefox 113+, Safari 16.4+)
 * Fallback: Pure JS DEFLATE implementation for older browsers
 *
 * Worker Pool: Optional off-main-thread compression/decompression
 * to prevent UI blocking for large files.
 */

import type { CompressOptions } from "@archive/compression/compress.base";
import {
  compressWithStream,
  decompressWithStream,
  transformWithStream,
  hasCompressionStream,
  hasDeflateRawCompressionStream,
  hasDeflateRawDecompressionStream,
  GZIP_ID1,
  GZIP_ID2,
  GZIP_CM_DEFLATE,
  GZIP_FLAG_FEXTRA,
  GZIP_FLAG_FNAME,
  GZIP_FLAG_FCOMMENT,
  GZIP_FLAG_FHCRC,
  GZIP_MIN_SIZE,
  hasGzipCompressionStream,
  hasGzipDecompressionStream,
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE,
  isZlibData,
  detectCompressionFormat,
  adler32,
  hasDeflateCompressionStream,
  hasDeflateDecompressionStream,
  getZlibHeader,
  buildZlibTrailer,
  parseZlibHeader,
  readZlibTrailer,
  verifyAdler32
} from "@archive/compression/compress.base";
import { crc32 } from "@archive/compression/crc32.browser";
import {
  inflateRaw,
  deflateRawCompressed,
  deflateRawStore
} from "@archive/compression/deflate-fallback";
import {
  deflateWithPool,
  inflateWithPool,
  hasWorkerSupport
} from "@archive/compression/worker-pool/index.browser";
import { DEFAULT_COMPRESS_LEVEL } from "@archive/core/defaults";
import { ArchiveError, createAbortError, isAbortError, throwIfAborted } from "@archive/core/errors";
import { readUint32LE } from "@archive/zip-spec/binary";
import { concatUint8Arrays } from "@utils/binary";

// Re-export shared types and GZIP utilities
export { type CompressOptions };
export {
  hasCompressionStream,
  hasWorkerSupport,
  // GZIP
  GZIP_ID1,
  GZIP_ID2,
  GZIP_CM_DEFLATE,
  GZIP_MIN_SIZE,
  hasGzipCompressionStream,
  hasGzipDecompressionStream,
  // Zlib
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE,
  isZlibData,
  detectCompressionFormat,
  // Zlib native stream detection (for API parity)
  hasDeflateCompressionStream,
  hasDeflateDecompressionStream
};
export { isGzipData } from "@archive/compression/compress.base";

/**
 * Default threshold (1MB) above which compression automatically uses workers.
 * Set to 0 to disable auto-worker, or Infinity to always use main thread.
 */
const DEFAULT_AUTO_WORKER_THRESHOLD = 1024 * 1024;

export { DEFAULT_AUTO_WORKER_THRESHOLD };

/**
 * Decide whether to use worker based on options and data size
 */
function shouldUseWorker(data: Uint8Array, options: CompressOptions): boolean {
  const workerSupported = hasWorkerSupport();
  if (options.useWorker === true) {
    return workerSupported;
  }
  if (options.useWorker === false) {
    return false;
  }

  const threshold = options.autoWorkerThreshold ?? DEFAULT_AUTO_WORKER_THRESHOLD;
  return workerSupported && data.length >= threshold;
}

/**
 * Check if an error or signal indicates an abort. Rethrow as AbortError if so.
 */
function rethrowIfAborted(err: unknown, signal?: AbortSignal): void {
  if (signal?.aborted || isAbortError(err)) {
    throw createAbortError((signal as any)?.reason ?? err);
  }
}

// =============================================================================
// Unified Codec Strategy
// =============================================================================

interface CodecStrategy {
  hasNative: () => boolean;
  native: (data: Uint8Array) => Promise<Uint8Array>;
  worker: (
    data: Uint8Array,
    opts: { level?: number; signal?: AbortSignal; allowTransfer?: boolean }
  ) => Promise<Uint8Array>;
  jsFallback: (data: Uint8Array, level?: number) => Uint8Array;
}

const deflateStrategy: CodecStrategy = {
  hasNative: hasDeflateRawCompressionStream,
  native: compressWithStream,
  worker: deflateWithPool,
  jsFallback: deflateRawCompressed
};

const inflateStrategy: CodecStrategy = {
  hasNative: hasDeflateRawDecompressionStream,
  native: decompressWithStream,
  worker: inflateWithPool,
  jsFallback: inflateRaw
};

/**
 * Unified compression/decompression with automatic strategy selection.
 */
async function processWithStrategy(
  strategy: CodecStrategy,
  data: Uint8Array,
  options: CompressOptions
): Promise<Uint8Array> {
  const canUseNative = strategy.hasNative();
  const workerSupported = hasWorkerSupport();
  const useWorker = options.useWorker;

  // If the user explicitly requested workers, honor it.
  if (useWorker === true && workerSupported) {
    try {
      return await strategy.worker(data, {
        level: options.level,
        signal: options.signal,
        allowTransfer: options.allowTransfer
      });
    } catch (err) {
      // If the user aborts, do NOT fall back to main-thread work.
      rethrowIfAborted(err, options.signal);
      // Fall through to best available in-process path.
    }
  }

  // Default: use native stream if supported (fastest, no worker overhead).
  if (canUseNative) {
    try {
      return await strategy.native(data);
    } catch (err) {
      // Respect aborts — never silently retry an aborted operation.
      rethrowIfAborted(err, options.signal);
      // Native CompressionStream / DecompressionStream can intermittently
      // reject input that is in fact valid (observed in Chromium under heavy
      // concurrent stream creation: a `DecompressionStream` rejects a deflate
      // payload that the pure-JS inflater — and a fresh native stream — decode
      // correctly). Rather than surface a spurious "invalid literal/lengths
      // set" / corruption error, fall back to the deterministic pure-JS
      // implementation. If the data is genuinely corrupt the JS path throws
      // too, so this never masks a real error.
      return strategy.jsFallback(data, options.level);
    }
  }

  // Use worker in fallback environments (no native deflate-raw) when appropriate.
  if (useWorker !== true && shouldUseWorker(data, options)) {
    return strategy.worker(data, {
      level: options.level,
      signal: options.signal,
      allowTransfer: options.allowTransfer
    });
  }

  // Fallback to pure JS implementation.
  return strategy.jsFallback(data, options.level);
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Compress data using browser's native CompressionStream or JS fallback
 */
export async function compress(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;

  // Level 0 means no compression
  if (level === 0) {
    return data;
  }

  const processOptions = options.level === undefined ? { ...options, level } : options;
  return processWithStrategy(deflateStrategy, data, processOptions);
}

/**
 * Compress data synchronously using pure JS implementation
 */
export function compressSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  if (level === 0) {
    return data;
  }
  return deflateRawCompressed(data, level);
}

/**
 * Decompress data using browser's native DecompressionStream or JS fallback
 */
export async function decompress(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array> {
  return processWithStrategy(inflateStrategy, data, options);
}

/**
 * Decompress data synchronously using pure JS implementation
 */
export function decompressSync(data: Uint8Array): Uint8Array {
  return inflateRaw(data);
}

// =============================================================================
// GZIP API
// =============================================================================

// Cached GZIP header: ID1, ID2, CM=DEFLATE, FLG=0, MTIME=0, XFL=0, OS=unknown
const GZIP_HEADER = new Uint8Array([GZIP_ID1, GZIP_ID2, GZIP_CM_DEFLATE, 0, 0, 0, 0, 0, 0, 255]);

function buildGzipTrailer(crcValue: number, size: number): Uint8Array {
  const trailer = new Uint8Array(8);
  const view = new DataView(trailer.buffer, trailer.byteOffset, trailer.byteLength);
  view.setUint32(0, crcValue >>> 0, true);
  view.setUint32(4, size >>> 0, true);
  return trailer;
}

function parseGzipPayload(data: Uint8Array): {
  deflateData: Uint8Array;
  expectedCrc32: number;
  expectedSize: number;
} {
  if (data.length < GZIP_MIN_SIZE) {
    throw new ArchiveError("Invalid gzip data (too small)");
  }
  if (data[0] !== GZIP_ID1 || data[1] !== GZIP_ID2) {
    throw new ArchiveError("Invalid gzip header (magic mismatch)");
  }
  if (data[2] !== GZIP_CM_DEFLATE) {
    throw new ArchiveError("Unsupported gzip compression method");
  }

  const flags = data[3];
  let offset = 10;

  if (flags & GZIP_FLAG_FEXTRA) {
    if (offset + 2 > data.length) {
      throw new ArchiveError("Invalid gzip extra field");
    }
    const extraLen = data[offset] | (data[offset + 1] << 8);
    offset += 2 + extraLen;
  }

  // Skip null-terminated strings
  const skipNullTerminated = () => {
    while (offset < data.length && data[offset] !== 0) {
      offset++;
    }
    offset++;
  };

  if (flags & GZIP_FLAG_FNAME) {
    skipNullTerminated();
  }
  if (flags & GZIP_FLAG_FCOMMENT) {
    skipNullTerminated();
  }

  if (flags & GZIP_FLAG_FHCRC) {
    offset += 2;
  }

  if (offset > data.length - 8) {
    throw new ArchiveError("Invalid gzip data (truncated payload)");
  }

  const trailerOffset = data.length - 8;
  const expectedCrc32 = readUint32LE(data, trailerOffset);
  const expectedSize = readUint32LE(data, trailerOffset + 4);
  const deflateData = data.subarray(offset, trailerOffset);

  return { deflateData, expectedCrc32, expectedSize };
}

/**
 * Verify decompressed data against GZIP trailer CRC32 and ISIZE
 */
function verifyGzipOutput(out: Uint8Array, expectedCrc32: number, expectedSize: number): void {
  const actualCrc32 = crc32(out) >>> 0;
  const actualSize = out.length >>> 0;

  if (actualCrc32 !== expectedCrc32) {
    throw new ArchiveError("Invalid gzip data (CRC32 mismatch)");
  }
  if (actualSize !== expectedSize) {
    throw new ArchiveError("Invalid gzip data (ISIZE mismatch)");
  }
}

function wrapGzip(deflated: Uint8Array, original: Uint8Array): Uint8Array {
  const trailer = buildGzipTrailer(crc32(original), original.length);
  return concatUint8Arrays([GZIP_HEADER, deflated, trailer]);
}

/**
 * Gzip-compress data in the browser.
 *
 * Strategy:
 * 1. Native CompressionStream("gzip") when available
 * 2. Fallback: compress (deflate-raw) + manual GZIP wrapper
 *    - Inherits Worker Pool support from compress() for large files
 */
export async function gzip(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  throwIfAborted(options.signal);

  if (hasGzipCompressionStream()) {
    const cs = new CompressionStream("gzip");
    const out = await transformWithStream(data, cs);
    throwIfAborted(options.signal);
    return out;
  }

  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const compressOptions = options.level === undefined ? { ...options, level } : options;
  const deflated = level === 0 ? deflateRawStore(data) : await compress(data, compressOptions);
  return wrapGzip(deflated, data);
}

/**
 * Gzip-compress data synchronously using the JS fallback.
 */
export function gzipSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const deflated = level === 0 ? deflateRawStore(data) : deflateRawCompressed(data, level);
  return wrapGzip(deflated, data);
}

/**
 * Gunzip data in the browser.
 *
 * Strategy:
 * 1. Native DecompressionStream("gzip") when available
 * 2. Fallback: parse header + decompress (inflate-raw) + verify CRC32
 *    - Inherits Worker Pool support from decompress() for large files
 */
export async function gunzip(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  throwIfAborted(options.signal);

  if (hasGzipDecompressionStream()) {
    const ds = new DecompressionStream("gzip");
    const out = await transformWithStream(data, ds);
    throwIfAborted(options.signal);
    return out;
  }

  const { deflateData, expectedCrc32, expectedSize } = parseGzipPayload(data);
  const out = await decompress(deflateData, options);
  verifyGzipOutput(out, expectedCrc32, expectedSize);
  return out;
}

/**
 * Gunzip data synchronously using the JS fallback.
 */
export function gunzipSync(data: Uint8Array): Uint8Array {
  const { deflateData, expectedCrc32, expectedSize } = parseGzipPayload(data);
  const out = inflateRaw(deflateData);
  verifyGzipOutput(out, expectedCrc32, expectedSize);
  return out;
}

// =============================================================================
// ZLIB API (RFC 1950: DEFLATE with zlib header/trailer + Adler-32)
// =============================================================================

/**
 * Wrap raw DEFLATE data in Zlib format
 */
function wrapZlib(deflated: Uint8Array, original: Uint8Array, level: number): Uint8Array {
  return concatUint8Arrays([getZlibHeader(level), deflated, buildZlibTrailer(adler32(original))]);
}

/**
 * Parse Zlib data and extract the raw DEFLATE payload
 */
function parseZlibPayload(data: Uint8Array): {
  deflateData: Uint8Array;
  expectedAdler32: number;
} {
  const offset = parseZlibHeader(data);
  const expectedAdler32 = readZlibTrailer(data);
  return { deflateData: data.subarray(offset, data.length - 4), expectedAdler32 };
}

/**
 * Compress data with Zlib wrapper (RFC 1950)
 *
 * Strategy:
 * 1. Native CompressionStream("deflate") when available
 * 2. Fallback: compress (deflate-raw) + manual Zlib wrapper
 */
export async function zlib(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  throwIfAborted(options.signal);

  // Native "deflate" format is Zlib
  if (hasDeflateCompressionStream()) {
    const cs = new CompressionStream("deflate");
    const out = await transformWithStream(data, cs);
    throwIfAborted(options.signal);
    return out;
  }

  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const compressOptions = options.level === undefined ? { ...options, level } : options;
  const deflated = level === 0 ? deflateRawStore(data) : await compress(data, compressOptions);
  return wrapZlib(deflated, data, level);
}

/**
 * Compress data with Zlib wrapper (sync)
 */
export function zlibSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const deflated = level === 0 ? deflateRawStore(data) : deflateRawCompressed(data, level);
  return wrapZlib(deflated, data, level);
}

/**
 * Decompress Zlib data (RFC 1950)
 *
 * Strategy:
 * 1. Native DecompressionStream("deflate") when available
 * 2. Fallback: parse header + decompress (inflate-raw) + verify Adler-32
 */
export async function unzlib(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  throwIfAborted(options.signal);

  // Native "deflate" format is Zlib
  if (hasDeflateDecompressionStream()) {
    const ds = new DecompressionStream("deflate");
    const out = await transformWithStream(data, ds);
    throwIfAborted(options.signal);
    return out;
  }

  const { deflateData, expectedAdler32 } = parseZlibPayload(data);
  const out = await decompress(deflateData, options);
  verifyAdler32(out, expectedAdler32);
  return out;
}

/**
 * Decompress Zlib data (sync)
 */
export function unzlibSync(data: Uint8Array): Uint8Array {
  const { deflateData, expectedAdler32 } = parseZlibPayload(data);
  const out = inflateRaw(deflateData);
  verifyAdler32(out, expectedAdler32);
  return out;
}

// =============================================================================
// AUTO-DETECT DECOMPRESSION
// =============================================================================

/**
 * Decompress data, automatically detecting the format (GZIP, Zlib, or raw DEFLATE)
 *
 * Detection order:
 * 1. GZIP: magic bytes 0x1f 0x8b
 * 2. Zlib: valid CMF/FLG header with checksum
 * 3. Raw DEFLATE: fallback
 *
 * @example
 * ```ts
 * // Works with any format
 * const data = await decompressAuto(compressed);
 * ```
 */
export async function decompressAuto(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array> {
  const format = detectCompressionFormat(data);

  switch (format) {
    case "gzip":
      return gunzip(data, options);
    case "zlib":
      return unzlib(data, options);
    case "deflate-raw":
      return decompress(data, options);
  }
}

/**
 * Decompress data synchronously, automatically detecting the format
 */
export function decompressAutoSync(data: Uint8Array): Uint8Array {
  const format = detectCompressionFormat(data);

  switch (format) {
    case "gzip":
      return gunzipSync(data);
    case "zlib":
      return unzlibSync(data);
    case "deflate-raw":
      return decompressSync(data);
  }
}
