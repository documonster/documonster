/**
 * `Compression` namespace surface — raw DEFLATE, gzip and zlib, one-shot and
 * streaming, plus the browser worker pool under `Compression.Workers`.
 *
 * `import { Compression } from "documonster/archive"` →
 *   `Compression.deflate(bytes)` / `Compression.inflate(bytes)`,
 *   `Compression.gzip(bytes)`, `Compression.createDeflateStream()`,
 *   `Compression.Workers.deflate(bytes)`, …
 *
 * One verb per codec direction throughout: `deflate` / `inflate` (raw DEFLATE),
 * `gzip` / `gunzip`, `zlib` / `unzlib`, each with a `Sync` and a `create…Stream` form;
 * DEFLATE also runs on the worker pool through `Compression.Workers`. Identical on both
 * platforms through the build's platform variants.
 */
import {
  compress,
  compressSync,
  decompress,
  decompressSync,
  gzip as gzipImpl,
  gunzip as gunzipImpl,
  gzipSync as gzipSyncImpl,
  gunzipSync as gunzipSyncImpl,
  isGzipData as isGzipDataImpl,
  zlib as zlibImpl,
  unzlib as unzlibImpl,
  zlibSync as zlibSyncImpl,
  unzlibSync as unzlibSyncImpl,
  isZlibData as isZlibDataImpl,
  decompressAuto as decompressAutoImpl,
  decompressAutoSync as decompressAutoSyncImpl,
  detectCompressionFormat
} from "@archive/compression/compress";
import type { CompressOptions } from "@archive/compression/compress";
import {
  createDeflateStream as createDeflateStreamImpl,
  createGzipStream as createGzipStreamImpl,
  createZlibStream as createZlibStreamImpl,
  createInflateStream as createInflateStreamImpl,
  createGunzipStream as createGunzipStreamImpl,
  createUnzlibStream as createUnzlibStreamImpl
} from "@archive/compression/streaming-compress";
import type {
  DeflateStream,
  GunzipStream,
  GzipStream,
  InflateStream,
  StreamCompressOptions,
  UnzlibStream,
  ZlibStream
} from "@archive/compression/streaming-compress";
import { checkArchiveArgument } from "@archive/core/arguments";
import { checkBytes, checkLevel } from "@archive/core/entry-checks";
import { ArchiveError, isAbortError, toDecompressionError } from "@archive/core/errors";
import { s } from "@utils/schema";

// Every member checks its arguments as the archive builders do and reports damaged input
// as an `ArchiveDecompressionError` — the codecs below throw their own errors (zlib's on Node,
// the JS inflater's or `DecompressionStream`'s in a browser), which differ by platform. A
// compressor's own failure says nothing about its input's integrity, so it is a plain
// `ArchiveError`.

type Codec = "deflate" | "gzip" | "zlib" | "compressed";

/** What a member does with its codec: decoding fails on damaged input, encoding does not. */
type Direction = "decode" | "encode";

function checkArguments(where: string, data: unknown, options: CompressOptions | undefined): void {
  checkBytes(`Compression.${where}`, "data", data);
  checkLevel(`Compression.${where}`, "options.level", options?.level);
}

function toCodecError(codec: Codec, direction: Direction, error: unknown): unknown {
  if (error instanceof ArchiveError || isAbortError(error) || !(error instanceof Error)) {
    return error;
  }
  return direction === "decode"
    ? toDecompressionError(error, codec)
    : new ArchiveError(`${codec} compression failed: ${error.message}`, { cause: error });
}

function wrapSync<R>(
  where: string,
  codec: Codec,
  direction: Direction,
  run: (data: Uint8Array, options?: CompressOptions) => R
): (data: Uint8Array, options?: CompressOptions) => R {
  return (data, options) => {
    checkArguments(where, data, options);
    try {
      return run(data, options);
    } catch (error) {
      throw toCodecError(codec, direction, error);
    }
  };
}

function wrapAsync<R>(
  where: string,
  codec: Codec,
  direction: Direction,
  run: (data: Uint8Array, options?: CompressOptions) => Promise<R>
): (data: Uint8Array, options?: CompressOptions) => Promise<R> {
  return async (data, options) => {
    checkArguments(where, data, options);
    try {
      return await run(data, options);
    } catch (error) {
      throw toCodecError(codec, direction, error);
    }
  };
}

/** Raw DEFLATE (RFC 1951), as a ZIP entry holds it. */
export const deflate = /* @__PURE__ */ wrapAsync("deflate", "deflate", "encode", compress);
/** {@link deflate}, synchronously. */
export const deflateSync = /* @__PURE__ */ wrapSync(
  "deflateSync",
  "deflate",
  "encode",
  compressSync
);
/** Inflate raw DEFLATE data; damaged data is an `ArchiveDecompressionError`. */
export const inflate = /* @__PURE__ */ wrapAsync("inflate", "deflate", "decode", decompress);
/** {@link inflate}, synchronously. */
export const inflateSync = /* @__PURE__ */ wrapSync(
  "inflateSync",
  "deflate",
  "decode",
  decompressSync
);
/** gzip (RFC 1952). */
export const gzip = /* @__PURE__ */ wrapAsync("gzip", "gzip", "encode", gzipImpl);
/** Decompress gzip data; damaged data is an `ArchiveDecompressionError`. */
export const gunzip = /* @__PURE__ */ wrapAsync("gunzip", "gzip", "decode", gunzipImpl);
/** {@link gzip}, synchronously. */
export const gzipSync = /* @__PURE__ */ wrapSync("gzipSync", "gzip", "encode", gzipSyncImpl);
/** {@link gunzip}, synchronously. */
export const gunzipSync = /* @__PURE__ */ wrapSync("gunzipSync", "gzip", "decode", gunzipSyncImpl);
/** zlib (RFC 1950). */
export const zlib = /* @__PURE__ */ wrapAsync("zlib", "zlib", "encode", zlibImpl);
/** Decompress zlib data; damaged data is an `ArchiveDecompressionError`. */
export const unzlib = /* @__PURE__ */ wrapAsync("unzlib", "zlib", "decode", unzlibImpl);
/** {@link zlib}, synchronously. */
export const zlibSync = /* @__PURE__ */ wrapSync("zlibSync", "zlib", "encode", zlibSyncImpl);
/** {@link unzlib}, synchronously. */
export const unzlibSync = /* @__PURE__ */ wrapSync("unzlibSync", "zlib", "decode", unzlibSyncImpl);
/** Decompress gzip, zlib or raw DEFLATE, whichever {@link detectFormat} finds. */
export const decompressAuto = /* @__PURE__ */ wrapAsync(
  "decompressAuto",
  "compressed",
  "decode",
  decompressAutoImpl
);
/** {@link decompressAuto}, synchronously. */
export const decompressAutoSync = /* @__PURE__ */ wrapSync(
  "decompressAutoSync",
  "compressed",
  "decode",
  decompressAutoSyncImpl
);

/** Whether `data` starts with a gzip header. */
export function isGzipData(data: Uint8Array): boolean {
  checkBytes("Compression.isGzipData", "data", data);
  return isGzipDataImpl(data);
}

/** Whether `data` starts with a valid zlib header. */
export function isZlibData(data: Uint8Array): boolean {
  checkBytes("Compression.isZlibData", "data", data);
  return isZlibDataImpl(data);
}

/** `"gzip"` or `"zlib"` from the header; anything else is taken to be raw DEFLATE. */
export function detectFormat(data: Uint8Array): "zlib" | "gzip" | "deflate-raw" {
  checkBytes("Compression.detectFormat", "data", data);
  return detectCompressionFormat(data);
}

/** Stream options: an object or `undefined`; a compressor's `level` 0–9. */
function checkStreamOptions(where: string, options: unknown, compressor: boolean): void {
  checkArchiveArgument(
    s.object({}, { expected: "an object" }).optional(),
    options,
    `Compression.${where}`,
    "options"
  );
  if (compressor) {
    checkLevel(
      `Compression.${where}`,
      "options.level",
      (options as StreamCompressOptions | undefined)?.level
    );
  }
}

/** A raw DEFLATE compressor that emits output after every write. */
export function createDeflateStream(options?: StreamCompressOptions): DeflateStream {
  checkStreamOptions("createDeflateStream", options, true);
  return createDeflateStreamImpl(options);
}

/** A gzip compressor that emits output after every write. */
export function createGzipStream(options?: StreamCompressOptions): GzipStream {
  checkStreamOptions("createGzipStream", options, true);
  return createGzipStreamImpl(options);
}

/** A zlib compressor that emits output after every write. */
export function createZlibStream(options?: StreamCompressOptions): ZlibStream {
  checkStreamOptions("createZlibStream", options, true);
  return createZlibStreamImpl(options);
}

/** A raw DEFLATE decompressor. `level` means nothing to it and is ignored. */
export function createInflateStream(options?: StreamCompressOptions): InflateStream {
  checkStreamOptions("createInflateStream", options, false);
  return createInflateStreamImpl(options);
}

/** A gzip decompressor. `level` means nothing to it and is ignored. */
export function createGunzipStream(options?: StreamCompressOptions): GunzipStream {
  checkStreamOptions("createGunzipStream", options, false);
  return createGunzipStreamImpl(options);
}

/** A zlib decompressor. `level` means nothing to it and is ignored. */
export function createUnzlibStream(options?: StreamCompressOptions): UnzlibStream {
  checkStreamOptions("createUnzlibStream", options, false);
  return createUnzlibStreamImpl(options);
}

export {
  hasCompressionStream,
  hasWorkerSupport,
  hasGzipCompressionStream,
  hasGzipDecompressionStream
} from "@archive/compression/compress";

export { hasDeflateRaw } from "@archive/compression/streaming-compress";

/** The browser Web Worker pool: `Compression.Workers.deflate(bytes)`, `new Compression.Workers.Pool(options)`, … */
export * as Workers from "@archive/surface/compression-workers";
