/**
 * Stream Module (Node.js entry)
 *
 * Public entrypoint for stream utilities and classes.
 *
 * Notes:
 * - This file is intentionally export-only (tree-shaking friendly).
 * - Browser builds should import from `@stream/index.browser`.
 *
 * @example
 * ```ts
 * import { pipeline, createTransform, createCollector } from "@stream";
 *
 * const upper = createTransform<Uint8Array, Uint8Array>(chunk => chunk);
 * const out = createCollector<Uint8Array>();
 * await pipeline(source, upper, out);
 * ```
 */

// Shared type + platform-independent exports
export * from "@stream/index.base";

// Core stream classes (native Node.js)
import { Readable, Transform, Duplex, PassThrough } from "stream";
export { Readable, Transform, Duplex, PassThrough };

// Writable (extended with browser-compatible API)
export { Writable, toWritable } from "@stream/node/writable";
export type { WritableOptions } from "@stream/node/writable";

// Collector
export { Collector, createCollector } from "@stream/node/collector";

// Factory functions + re-exported helpers
export {
  PullStream,
  BufferedStream,
  StringChunk,
  ByteChunk,
  createReadable,
  createReadableFromAsyncIterable,
  createReadableFromArray,
  createWritable,
  createTransform,
  createPassThrough,
  createPullStream,
  createBufferedStream,
  createDuplex,
  createReadableFromGenerator,
  createReadableFromPromise,
  createEmptyReadable,
  createNullWritable
} from "@stream/node/factories";

// Pipeline & Finished
export { pipeline, finished, finishedAll } from "@stream/node/pipeline";

// Compose
export { compose } from "@stream/node/compose";

// Utilities
export {
  streamToPromise,
  streamToUint8Array,
  streamToBuffer,
  streamToString,
  drainStream,
  copyStream,
  isTransform,
  isDuplex,
  isStream,
  addAbortSignal,
  isDisturbed,
  isReadable,
  isWritable,
  duplexPair,
  consumers,
  promises
} from "@stream/node/utils";
