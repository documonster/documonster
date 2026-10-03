/**
 * Node.js compression utilities using native zlib
 *
 * Supports multiple formats:
 * - deflate-raw: Raw DEFLATE (for ZIP files)
 * - gzip: GZIP format (for tar.gz, HTTP compression)
 */

import { kMaxLength } from "buffer";
import { promisify } from "util";
import * as nodeZlib from "zlib";

import { DEFAULT_COMPRESS_LEVEL } from "@archive/core/defaults";
import { ArchiveLimitError } from "@archive/core/errors";
import { uint8ArrayToNodeBufferView } from "@utils/binary";

// Re-export shared types and utilities
export {
  type CompressOptions,
  hasCompressionStream,
  // GZIP constants and utilities
  GZIP_ID1,
  GZIP_ID2,
  GZIP_CM_DEFLATE,
  GZIP_FLAG_FTEXT,
  GZIP_FLAG_FHCRC,
  GZIP_FLAG_FEXTRA,
  GZIP_FLAG_FNAME,
  GZIP_FLAG_FCOMMENT,
  GZIP_MIN_SIZE,
  isGzipData,
  hasGzipCompressionStream,
  hasGzipDecompressionStream,
  // Zlib constants and utilities
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE,
  isZlibData,
  detectCompressionFormat,
  // Zlib native stream detection (also available in Node.js for API parity)
  hasDeflateCompressionStream,
  hasDeflateDecompressionStream
} from "@archive/compression/compress.base";

import type { CompressOptions } from "@archive/compression/compress.base";

/**
 * Check if Web Workers are available.
 * Always returns false in Node.js (zlib uses native thread pool).
 */
export function hasWorkerSupport(): boolean {
  return false;
}

/** Convert Uint8Array to Node.js Buffer (zero-copy view) */
export function uint8ArrayToBuffer(data: Uint8Array): Buffer {
  return uint8ArrayToNodeBufferView(data) as Buffer;
}

/** Convert Node.js Buffer to Uint8Array (zero-copy view) */
export function bufferToUint8Array(buffer: Buffer): Uint8Array {
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

const deflateRawAsync = promisify(nodeZlib.deflateRaw) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

const inflateRawAsync = promisify(nodeZlib.inflateRaw) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

import {
  resolveCompressThresholdBytes,
  detectCompressionFormat
} from "@archive/compression/compress.base";
import { assertLimitOption } from "@utils/limits";

/**
 * Compress data using Node.js native zlib
 *
 * @param data - Data to compress
 * @param options - Compression options
 * @returns Compressed data
 *
 * @example
 * ```ts
 * const data = new TextEncoder().encode("Hello, World!");
 * const compressed = await compress(data, { level: 6 });
 * ```
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

  const thresholdBytes = resolveCompressThresholdBytes(options);
  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };

  // Small-input fast path: avoid threadpool overhead.
  if (data.byteLength <= thresholdBytes) {
    return bufferToUint8Array(nodeZlib.deflateRawSync(input, zlibOptions));
  }

  return bufferToUint8Array(await deflateRawAsync(input, zlibOptions));
}

/**
 * Compress data synchronously using Node.js zlib
 *
 * @param data - Data to compress
 * @param options - Compression options
 * @returns Compressed data
 */
export function compressSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;

  if (level === 0) {
    return data;
  }

  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };
  return bufferToUint8Array(nodeZlib.deflateRawSync(input, zlibOptions));
}

/**
 * zlib options for an optional output bound. zlib rejects a bound below 1, so
 * a zero bound is requested as 1 and {@link inflateBounded} rejects any output.
 */
function boundOptions(maxOutputLength: number | undefined): nodeZlib.ZlibOptions | undefined {
  // Every decompressing entry point resolves its bound here first.
  assertLimitOption("maxOutputLength", maxOutputLength);
  return maxOutputLength === undefined
    ? undefined
    : { maxOutputLength: Math.max(1, Math.min(maxOutputLength, kMaxLength)) };
}

/**
 * Settle a zlib inflate run with {@link boundOptions}: zlib stops as soon as
 * the inflater crosses the bound and throws `ERR_BUFFER_TOO_LARGE`, which
 * becomes the archive's own error.
 */
function inflateBounded(inflate: () => Buffer, maxOutputLength: number | undefined): Uint8Array {
  let out: Buffer;
  try {
    out = inflate();
  } catch (err) {
    if (
      maxOutputLength !== undefined &&
      (err as NodeJS.ErrnoException).code === "ERR_BUFFER_TOO_LARGE"
    ) {
      throw new ArchiveLimitError("maxOutputLength", maxOutputLength, { cause: err });
    }
    throw err;
  }
  if (maxOutputLength !== undefined && out.byteLength > maxOutputLength) {
    throw new ArchiveLimitError("maxOutputLength", maxOutputLength);
  }
  return bufferToUint8Array(out);
}

/** {@link inflateBounded} for a run settled on the thread pool. */
async function inflateBoundedAsync(
  pending: Promise<Buffer>,
  maxOutputLength: number | undefined
): Promise<Uint8Array> {
  const settled = await pending.then(
    out => () => out,
    (err: unknown) => (): Buffer => {
      throw err;
    }
  );
  return inflateBounded(settled, maxOutputLength);
}

/**
 * Decompress data using Node.js native zlib
 *
 * @param data - Compressed data (deflate-raw format)
 * @param options - `maxOutputLength` bounds the output; inflation stops early
 *   with an `ArchiveLimitError` when it is exceeded
 * @returns Decompressed data
 */
export async function decompress(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array> {
  const input = uint8ArrayToBuffer(data);
  const { maxOutputLength } = options;
  const zlibOptions = boundOptions(maxOutputLength);

  // Small-input fast path: avoid threadpool overhead.
  if (data.byteLength <= resolveCompressThresholdBytes(options)) {
    return inflateBounded(() => nodeZlib.inflateRawSync(input, zlibOptions), maxOutputLength);
  }
  return inflateBoundedAsync(inflateRawAsync(input, zlibOptions), maxOutputLength);
}

/**
 * Decompress data synchronously using Node.js zlib
 *
 * @param data - Compressed data (deflate-raw format)
 * @param options - `maxOutputLength` bounds the output (see {@link decompress})
 * @returns Decompressed data
 */
export function decompressSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const input = uint8ArrayToBuffer(data);
  const { maxOutputLength } = options;
  return inflateBounded(
    () => nodeZlib.inflateRawSync(input, boundOptions(maxOutputLength)),
    maxOutputLength
  );
}

// =============================================================================
// GZIP API
// =============================================================================

const gzipAsync = promisify(nodeZlib.gzip) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

const gunzipAsync = promisify(nodeZlib.gunzip) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

/**
 * Compress data with gzip
 */
export async function gzip(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const thresholdBytes = resolveCompressThresholdBytes(options);
  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };

  // Small-input fast path
  if (data.byteLength <= thresholdBytes) {
    return bufferToUint8Array(nodeZlib.gzipSync(input, zlibOptions));
  }

  return bufferToUint8Array(await gzipAsync(input, zlibOptions));
}

/**
 * Decompress gzip data
 */
export async function gunzip(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  const thresholdBytes = resolveCompressThresholdBytes(options);
  const input = uint8ArrayToBuffer(data);

  const { maxOutputLength } = options;
  const zlibOptions = boundOptions(maxOutputLength);

  // Small-input fast path
  if (data.byteLength <= thresholdBytes) {
    return inflateBounded(() => nodeZlib.gunzipSync(input, zlibOptions), maxOutputLength);
  }

  return inflateBoundedAsync(gunzipAsync(input, zlibOptions), maxOutputLength);
}

/**
 * Compress data with gzip (sync)
 */
export function gzipSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };
  return bufferToUint8Array(nodeZlib.gzipSync(input, zlibOptions));
}

/**
 * Decompress gzip data (sync)
 */
export function gunzipSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const input = uint8ArrayToBuffer(data);
  const { maxOutputLength } = options;
  return inflateBounded(
    () => nodeZlib.gunzipSync(input, boundOptions(maxOutputLength)),
    maxOutputLength
  );
}

// =============================================================================
// ZLIB API (RFC 1950: DEFLATE with zlib header/trailer + Adler-32)
// =============================================================================

const zlibAsync = promisify(nodeZlib.deflate) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

const unzlibAsync = promisify(nodeZlib.inflate) as (
  input: nodeZlib.InputType,
  options?: nodeZlib.ZlibOptions
) => Promise<Buffer>;

/**
 * Compress data with Zlib wrapper (RFC 1950)
 *
 * Zlib format: CMF + FLG header + DEFLATE data + Adler-32 checksum
 * Used for: PNG, HTTP deflate encoding, general-purpose compression
 */
export async function zlib(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const thresholdBytes = resolveCompressThresholdBytes(options);
  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };

  // Small-input fast path
  if (data.byteLength <= thresholdBytes) {
    return bufferToUint8Array(nodeZlib.deflateSync(input, zlibOptions));
  }

  return bufferToUint8Array(await zlibAsync(input, zlibOptions));
}

/**
 * Decompress Zlib data (RFC 1950)
 */
export async function unzlib(data: Uint8Array, options: CompressOptions = {}): Promise<Uint8Array> {
  const thresholdBytes = resolveCompressThresholdBytes(options);
  const input = uint8ArrayToBuffer(data);

  const { maxOutputLength } = options;
  const zlibOptions = boundOptions(maxOutputLength);

  // Small-input fast path
  if (data.byteLength <= thresholdBytes) {
    return inflateBounded(() => nodeZlib.inflateSync(input, zlibOptions), maxOutputLength);
  }

  return inflateBoundedAsync(unzlibAsync(input, zlibOptions), maxOutputLength);
}

/**
 * Compress data with Zlib wrapper (sync)
 */
export function zlibSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const input = uint8ArrayToBuffer(data);
  const zlibOptions = { level };
  return bufferToUint8Array(nodeZlib.deflateSync(input, zlibOptions));
}

/**
 * Decompress Zlib data (sync)
 *
 * @param options - `maxOutputLength` bounds the output (see {@link decompress})
 */
export function unzlibSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const input = uint8ArrayToBuffer(data);
  const { maxOutputLength } = options;
  return inflateBounded(
    () => nodeZlib.inflateSync(input, boundOptions(maxOutputLength)),
    maxOutputLength
  );
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
export function decompressAutoSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const format = detectCompressionFormat(data);

  switch (format) {
    case "gzip":
      return gunzipSync(data, options);
    case "zlib":
      return unzlibSync(data, options);
    case "deflate-raw":
      return decompressSync(data, options);
  }
}
