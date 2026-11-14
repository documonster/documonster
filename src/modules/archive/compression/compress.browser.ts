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
  hasDeflateRawWebStreams,
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
  inflateRawPrefix,
  deflateRawCompressed,
  deflateRawStore
} from "@archive/compression/deflate-fallback";
import { hasWorkerSupport } from "@archive/compression/worker-pool/pool.base";
import { DEFAULT_COMPRESS_LEVEL } from "@archive/core/defaults";
import {
  ArchiveError,
  ArchiveLimitExceededError,
  createAbortError,
  isAbortError,
  throwIfAborted
} from "@archive/core/errors";
import { readUint32LE } from "@archive/zip-spec/binary";
import { concatUint8Arrays, asPlainBytes } from "@utils/binary";
import { toError } from "@utils/errors";
import { assertLimitOption } from "@utils/limits";

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
 * Size (1 MB) from which a call with `useWorker` unset runs on a worker. `0` sends every call to
 * one; `Infinity` keeps every call on this thread.
 */
const DEFAULT_AUTO_WORKER_THRESHOLD = 1024 * 1024;

export { DEFAULT_AUTO_WORKER_THRESHOLD };

/**
 * Whether this call goes to the worker pool.
 *
 * The worker codes with the platform's native deflate-raw streams, so it is useful only where they
 * exist; without them every call stays on this thread with the pure-JS codec. A native stream runs
 * on the thread that calls it and holds that thread for the whole input — measured at 150 ms for
 * 8 MB of deflate in Chromium and Firefox, 312 ms in Safari, against about 5 ms on a worker for the
 * same total time within 10 % — which is what moving a large call off this thread buys.
 */
function shouldUseWorker(data: Uint8Array, options: CompressOptions): boolean {
  if (options.useWorker === false || !hasWorkerSupport() || !hasDeflateRawWebStreams()) {
    return false;
  }
  if (options.useWorker === true) {
    return true;
  }
  return data.length >= (options.autoWorkerThreshold ?? DEFAULT_AUTO_WORKER_THRESHOLD);
}

/**
 * Check if an error or signal indicates an abort. Rethrow as AbortError if so.
 */
function rethrowIfAborted(err: unknown, signal?: AbortSignal): void {
  if (signal?.aborted || isAbortError(err)) {
    throw createAbortError(signal?.reason ?? err);
  }
}

// =============================================================================
// Unified Codec Strategy
// =============================================================================

/**
 * The worker pool, loaded the first time a call goes to a worker. It is reached only from an asynchronous path, and
 * most callers never take it — below the auto-worker threshold, or with `useWorker: false` — so a static import put
 * its scheduler and worker script in every bundle that compresses anything.
 */
function loadWorkerPool() {
  return import("@archive/compression/worker-pool/pool.browser");
}

interface CodecStrategy {
  /** What the codec does, for a message that names it. */
  operation: "compress" | "decompress";
  hasNative: () => boolean;
  native: (data: Uint8Array, options: CompressOptions) => Promise<Uint8Array>;
  worker: (data: Uint8Array, options: CompressOptions) => Promise<Uint8Array>;
  jsFallback: (data: Uint8Array, options: CompressOptions) => Uint8Array;
}

const deflateStrategy: CodecStrategy = {
  operation: "compress",
  hasNative: hasDeflateRawCompressionStream,
  native: data => compressWithStream(data),
  worker: async (data, { level, signal, allowTransfer }) =>
    (await loadWorkerPool()).deflateWithPool(data, { level, signal, allowTransfer }),
  jsFallback: (data, options) => deflateRawCompressed(data, options.level)
};

const inflateStrategy: CodecStrategy = {
  operation: "decompress",
  hasNative: hasDeflateRawDecompressionStream,
  native: (data, options) => decompressWithStream(data, options.maxOutputLength),
  // The worker enforces the bound while it inflates and reports a hit as an
  // ArchiveLimitExceededError, exactly like the in-process paths.
  worker: async (data, { signal, allowTransfer, maxOutputLength }) =>
    (await loadWorkerPool()).inflateWithPool(data, { signal, allowTransfer, maxOutputLength }),
  jsFallback: (data, options) => inflateRaw(data, options.maxOutputLength)
};

/**
 * Rethrow `err` when no other codec should be tried: the caller aborted, or a limit was hit — a
 * verdict on the data that every codec would reach again.
 */
function rethrowIfFinal(err: unknown, signal?: AbortSignal): void {
  rethrowIfAborted(err, signal);
  if (err instanceof ArchiveLimitExceededError) {
    throw err;
  }
}

/**
 * Run a native stream codec, and the pure-JS one when it fails.
 *
 * Native CompressionStream / DecompressionStream were seen to reject input that is in fact valid
 * (Chromium, June 2026, under heavy concurrent stream creation: about 2 % of reads, on payloads the
 * pure-JS inflater — and a fresh native stream — decode correctly). It has not been reproduced
 * since in Chromium, Firefox or WebKit, so the cause is unknown; the fallback stays because it
 * costs nothing when the native codec succeeds. Every format runs on the same native zlib, so raw
 * DEFLATE, gzip and zlib all go through here. If the data is genuinely corrupt the JS path throws
 * too, so this never masks a real error — and that error is the archive's own.
 */
async function withJsFallback(
  native: () => Promise<Uint8Array>,
  jsFallback: () => Uint8Array,
  signal: AbortSignal | undefined
): Promise<Uint8Array> {
  try {
    return await native();
  } catch (err) {
    rethrowIfFinal(err, signal);
    return jsFallback();
  }
}

/**
 * Run on the worker pool, answering `undefined` when the work should be done on this thread instead.
 *
 * A worker that could not start — its script refused by a Content Security Policy, the pool's chunk
 * not loaded, a script that would not boot — ran nothing, so the work is still to do. A worker's
 * verdict on the data arrives as a message that has lost its class, so the in-process codec runs it
 * again and reports it typed; corrupt data fails there too, so nothing is masked. The one failure
 * that cannot be retried is one after `allowTransfer` gave the worker the input's buffer: the bytes
 * are gone from this thread, and coding the empty view left behind would return a wrong answer.
 */
async function tryWorker(
  strategy: CodecStrategy,
  data: Uint8Array,
  options: CompressOptions
): Promise<Uint8Array | undefined> {
  const byteLength = data.byteLength;
  try {
    return await strategy.worker(data, options);
  } catch (err) {
    rethrowIfFinal(err, options.signal);
    if (data.byteLength !== byteLength) {
      throw new ArchiveError(
        `Worker ${strategy.operation} failed after allowTransfer had moved the input to it, so it ` +
          `cannot be run on this thread instead: ${toError(err).message}`,
        { cause: err }
      );
    }
    return undefined;
  }
}

/**
 * Unified compression/decompression with automatic strategy selection: a worker for a large call
 * (or when asked), else the native stream, else the pure-JS codec — each falling back to the next.
 */
async function processWithStrategy(
  strategy: CodecStrategy,
  data: Uint8Array,
  options: CompressOptions
): Promise<Uint8Array> {
  if (shouldUseWorker(data, options)) {
    const out = await tryWorker(strategy, data, options);
    if (out !== undefined) {
      return out;
    }
  }

  if (strategy.hasNative()) {
    return withJsFallback(
      () => strategy.native(data, options),
      () => strategy.jsFallback(data, options),
      options.signal
    );
  }

  return strategy.jsFallback(data, options);
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
): Promise<Uint8Array<ArrayBuffer>> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;

  // Level 0 is DEFLATE's stored blocks: still a DEFLATE stream.
  if (level === 0) {
    return asPlainBytes(deflateRawStore(data));
  }

  const processOptions = options.level === undefined ? { ...options, level } : options;
  return asPlainBytes(await processWithStrategy(deflateStrategy, data, processOptions));
}

/**
 * Compress data synchronously using pure JS implementation
 */
export function compressSync(
  data: Uint8Array,
  options: CompressOptions = {}
): Uint8Array<ArrayBuffer> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  if (level === 0) {
    return asPlainBytes(deflateRawStore(data));
  }
  return asPlainBytes(deflateRawCompressed(data, level));
}

/**
 * Decompress data using browser's native DecompressionStream or JS fallback
 */
export async function decompress(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  return asPlainBytes(await processWithStrategy(inflateStrategy, data, options));
}

/**
 * Decompress data synchronously using pure JS implementation
 */
export function decompressSync(
  data: Uint8Array,
  options: CompressOptions = {}
): Uint8Array<ArrayBuffer> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  return asPlainBytes(inflateRaw(data, options.maxOutputLength));
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

/**
 * Where the DEFLATE data of the gzip member at `offset` starts: just past its header (RFC 1952 §2.3).
 */
function skipGzipHeader(data: Uint8Array, offset: number): number {
  if (data.length - offset < GZIP_MIN_SIZE) {
    throw new ArchiveError("Invalid gzip data (too small)");
  }
  if (data[offset] !== GZIP_ID1 || data[offset + 1] !== GZIP_ID2) {
    throw new ArchiveError("Invalid gzip header (magic mismatch)");
  }
  if (data[offset + 2] !== GZIP_CM_DEFLATE) {
    throw new ArchiveError("Unsupported gzip compression method");
  }

  const flags = data[offset + 3];
  let pos = offset + 10;

  if (flags & GZIP_FLAG_FEXTRA) {
    if (pos + 2 > data.length) {
      throw new ArchiveError("Invalid gzip extra field");
    }
    const extraLen = data[pos] | (data[pos + 1] << 8);
    pos += 2 + extraLen;
  }

  // Skip null-terminated strings
  const skipNullTerminated = () => {
    while (pos < data.length && data[pos] !== 0) {
      pos++;
    }
    pos++;
  };

  if (flags & GZIP_FLAG_FNAME) {
    skipNullTerminated();
  }
  if (flags & GZIP_FLAG_FCOMMENT) {
    skipNullTerminated();
  }

  if (flags & GZIP_FLAG_FHCRC) {
    pos += 2;
  }

  if (pos > data.length - 8) {
    throw new ArchiveError("Invalid gzip data (truncated payload)");
  }
  return pos;
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

/**
 * Inflate the DEFLATE stream at `start`, holding the output to what is left of `maxOutputLength`
 * after `produced` bytes, and reporting a hit against the caller's own bound.
 */
function inflateStreamAt(
  data: Uint8Array,
  start: number,
  maxOutputLength: number | undefined,
  produced: number
): { output: Uint8Array; end: number } {
  try {
    const { output, consumed } = inflateRawPrefix(
      data.subarray(start),
      maxOutputLength === undefined ? undefined : maxOutputLength - produced
    );
    return { output, end: start + consumed };
  } catch (err) {
    if (err instanceof ArchiveLimitExceededError && maxOutputLength !== undefined) {
      throw new ArchiveLimitExceededError("maxOutputLength", maxOutputLength);
    }
    throw err;
  }
}

/**
 * Every member of a gzip file, decoded and concatenated, the way zlib — and so Node — reads one.
 *
 * A gzip file is a series of members (RFC 1952 §2.2), which is what `cat a.gz b.gz` produces, and
 * zero bytes after the last one are padding (tape blocks, fixed-size records). Anything else after
 * a member must be another member. Every browser's `DecompressionStream("gzip")` refuses both
 * shapes; decoding them here is what makes a browser read the same files Node does.
 */
function gunzipMembers(data: Uint8Array, maxOutputLength: number | undefined): Uint8Array {
  const parts: Uint8Array[] = [];
  let produced = 0;
  let offset = 0;
  do {
    const start = skipGzipHeader(data, offset);
    const { output, end } = inflateStreamAt(data, start, maxOutputLength, produced);
    if (end + 8 > data.length) {
      throw new ArchiveError("Invalid gzip data (truncated payload)");
    }
    verifyGzipOutput(output, readUint32LE(data, end), readUint32LE(data, end + 4));
    parts.push(output);
    produced += output.length;
    offset = end + 8;
  } while (offset < data.length && data[offset] !== 0);
  return parts.length === 1 ? parts[0]! : concatUint8Arrays(parts);
}

/**
 * `deflated` as a gzip member whose trailer says `crc` and `size`. They are taken before the input
 * is compressed rather than from it after: `allowTransfer` may have handed its buffer to a worker.
 */
function wrapGzip(deflated: Uint8Array, crc: number, size: number): Uint8Array {
  return concatUint8Arrays([GZIP_HEADER, deflated, buildGzipTrailer(crc, size)]);
}

/**
 * Gzip-compress data in the browser.
 *
 * Strategy:
 * 1. A large call (see `CompressOptions.useWorker`): raw DEFLATE through {@link compress}, which
 *    runs it on a worker, wrapped here
 * 2. Native CompressionStream("gzip") when available, retried with the pure-JS codec when it
 *    fails (see `withJsFallback`)
 * 3. Fallback: the pure-JS codec + manual GZIP wrapper
 */
export async function gzip(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  throwIfAborted(options.signal);
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;

  if (level !== 0 && shouldUseWorker(data, options)) {
    const crc = crc32(data);
    const size = data.length;
    const deflated = await compress(
      data,
      options.level === undefined ? { ...options, level } : options
    );
    return asPlainBytes(wrapGzip(deflated, crc, size));
  }

  if (hasGzipCompressionStream()) {
    const out = await withJsFallback(
      () => transformWithStream(data, new CompressionStream("gzip")),
      () => gzipSync(data, options),
      options.signal
    );
    throwIfAborted(options.signal);
    return asPlainBytes(out);
  }

  return gzipSync(data, options);
}

/**
 * Gzip-compress data synchronously using the JS fallback.
 */
export function gzipSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array<ArrayBuffer> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const deflated = level === 0 ? deflateRawStore(data) : deflateRawCompressed(data, level);
  return asPlainBytes(wrapGzip(deflated, crc32(data), data.length));
}

/**
 * Gunzip data in the browser.
 *
 * Strategy:
 * 1. Native DecompressionStream("gzip") when available; when it fails — on a file of several
 *    members or with padding after the last, which every browser's refuses, or spuriously (see
 *    `withJsFallback`) — the pure-JS decoder, which reads those as Node does
 * 2. Fallback: the pure-JS decoder
 *
 * It stays on this thread whatever `useWorker` says: the worker decodes only raw DEFLATE, and where
 * a member ends is known only once it has been decoded. Inflating is several times faster than
 * deflating — 8 MB in about 27 ms.
 */
export async function gunzip(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  throwIfAborted(options.signal);

  if (hasGzipDecompressionStream()) {
    const out = await withJsFallback(
      () => transformWithStream(data, new DecompressionStream("gzip"), options.maxOutputLength),
      () => gunzipSync(data, options),
      options.signal
    );
    throwIfAborted(options.signal);
    return asPlainBytes(out);
  }

  return gunzipSync(data, options);
}

/**
 * Gunzip data synchronously using the JS fallback. Reads every member, as {@link gunzip} does.
 */
export function gunzipSync(
  data: Uint8Array,
  options: CompressOptions = {}
): Uint8Array<ArrayBuffer> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  return asPlainBytes(gunzipMembers(data, options.maxOutputLength));
}

// =============================================================================
// ZLIB API (RFC 1950: DEFLATE with zlib header/trailer + Adler-32)
// =============================================================================

/**
 * Wrap raw DEFLATE data in Zlib format
 */
/** `deflated` as a zlib stream whose trailer says `checksum` — taken before compressing, as in {@link wrapGzip}. */
function wrapZlib(deflated: Uint8Array, checksum: number, level: number): Uint8Array {
  return concatUint8Arrays([getZlibHeader(level), deflated, buildZlibTrailer(checksum)]);
}

/**
 * Compress data with Zlib wrapper (RFC 1950)
 *
 * Strategy:
 * 1. A large call (see `CompressOptions.useWorker`): raw DEFLATE through {@link compress}, which
 *    runs it on a worker, wrapped here
 * 2. Native CompressionStream("deflate") when available, retried with the pure-JS codec when it
 *    fails (see `withJsFallback`)
 * 3. Fallback: the pure-JS codec + manual Zlib wrapper
 */
export async function zlib(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  throwIfAborted(options.signal);
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;

  if (level !== 0 && shouldUseWorker(data, options)) {
    const checksum = adler32(data);
    const deflated = await compress(
      data,
      options.level === undefined ? { ...options, level } : options
    );
    return asPlainBytes(wrapZlib(deflated, checksum, level));
  }

  // Native "deflate" format is Zlib
  if (hasDeflateCompressionStream()) {
    const out = await withJsFallback(
      () => transformWithStream(data, new CompressionStream("deflate")),
      () => zlibSync(data, options),
      options.signal
    );
    throwIfAborted(options.signal);
    return asPlainBytes(out);
  }

  return zlibSync(data, options);
}

/**
 * Compress data with Zlib wrapper (sync)
 */
export function zlibSync(data: Uint8Array, options: CompressOptions = {}): Uint8Array<ArrayBuffer> {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  const deflated = level === 0 ? deflateRawStore(data) : deflateRawCompressed(data, level);
  return asPlainBytes(wrapZlib(deflated, adler32(data), level));
}

/**
 * Decompress Zlib data (RFC 1950)
 *
 * Strategy:
 * 1. Native DecompressionStream("deflate") when available; when it fails — on bytes after the
 *    Adler-32, which every browser's refuses and zlib ignores, or spuriously (see `withJsFallback`)
 *    — the pure-JS decoder, which reads them as Node does
 * 2. Fallback: the pure-JS decoder
 *
 * Like {@link gunzip}, it stays on this thread whatever `useWorker` says.
 */
export async function unzlib(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  throwIfAborted(options.signal);

  // Native "deflate" format is Zlib
  if (hasDeflateDecompressionStream()) {
    const out = await withJsFallback(
      () => transformWithStream(data, new DecompressionStream("deflate"), options.maxOutputLength),
      () => unzlibSync(data, options),
      options.signal
    );
    throwIfAborted(options.signal);
    return asPlainBytes(out);
  }

  return unzlibSync(data, options);
}

/**
 * Decompress Zlib data (sync)
 *
 * The Adler-32 is the four bytes after the DEFLATE stream, wherever it ends; anything after them
 * is not part of the stream and is ignored, as zlib — and so Node — ignores it.
 *
 * @param options - `maxOutputLength` bounds the output; decoding stops with a
 *   `ArchiveLimitExceededError` as soon as it is exceeded
 */
export function unzlibSync(
  data: Uint8Array,
  options: CompressOptions = {}
): Uint8Array<ArrayBuffer> {
  assertLimitOption("maxOutputLength", options.maxOutputLength);
  const start = parseZlibHeader(data);
  const { output, end } = inflateStreamAt(data, start, options.maxOutputLength, 0);
  if (end + 4 > data.length) {
    throw new ArchiveError("Invalid zlib data (truncated payload)");
  }
  verifyAdler32(output, readZlibTrailer(data.subarray(0, end + 4)));
  return asPlainBytes(output);
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
 * const data = await Compression.decompressAuto(compressed);
 * ```
 */
export async function decompressAuto(
  data: Uint8Array,
  options: CompressOptions = {}
): Promise<Uint8Array<ArrayBuffer>> {
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
export function decompressAutoSync(
  data: Uint8Array,
  options: CompressOptions = {}
): Uint8Array<ArrayBuffer> {
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
