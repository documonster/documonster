/**
 * Stream Module (browser entry)
 *
 * Mirrors the public surface of `./index.ts`, but exports the browser
 * implementation from `./streams.browser`.
 *
 * This file is intentionally export-only (tree-shaking friendly).
 */

// Shared type + platform-independent exports
export * from "@stream/index.base";

// Core stream classes (browser implementations)
import { Readable } from "@stream/browser/readable";
export { Writable, toWritable } from "@stream/browser/writable";
export type { WritableOptions } from "@stream/browser/writable";
export { Transform } from "@stream/browser/transform";
import { Duplex } from "@stream/browser/duplex";
export { PassThrough } from "@stream/browser/passthrough";
export { Collector } from "@stream/browser/collector";
export { Readable, Duplex };

// Late-binding injection: break circular Readable ↔ Duplex
import { registerDuplexFrom } from "@stream/browser/_lazy";

registerDuplexFrom(source => Duplex.from(source));

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
  createCollector,
  createPassThrough,
  createPullStream,
  createBufferedStream,
  createDuplex,
  createReadableFromGenerator,
  createReadableFromPromise,
  createEmptyReadable,
  createNullWritable
} from "@stream/browser/factories";

// Pipeline & Finished
export { pipeline, finished, finishedAll } from "@stream/browser/pipeline";

// Compose
export { compose } from "@stream/browser/compose";

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
} from "@stream/browser/utils";
