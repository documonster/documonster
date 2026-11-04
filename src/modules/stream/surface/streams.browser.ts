/**
 * `Streams` namespace surface — the helpers this module adds around Node's
 * stream API (browser build).
 *
 * `import { Streams } from "documonster/stream"` →
 *   `Streams.createTransform(fn)`, `Streams.fromString(s)`, `Streams.text(stream)`,
 *   `Streams.fromWebReadable(web)`, …
 *
 * `Streams.bytes` resolves to a `Uint8Array` on every platform, never a Node `Buffer`.
 * Adapters for event emitters, arbitrary sinks and callbacks live one level down, on
 * `Streams.LowLevel`.
 *
 * Everything that mirrors `node:stream` (`Readable`, `pipeline`, `finished`, …) stays a
 * flat export under Node's own name; everything Node does not have lives here.
 */

// Factories
export {
  createReadable,
  createReadableFromAsyncIterable,
  createReadableFromArray,
  createReadableFromGenerator,
  createReadableFromPromise,
  createEmptyReadable,
  createWritable,
  createNullWritable,
  createTransform,
  createPassThrough,
  createDuplex,
  createPullStream,
  createBufferedStream,
  createStringChunk,
  createByteChunk,
  fromBytes,
  fromString,
  fromJSON as fromJson,
  filter
} from "@stream/browser/factories";
export { createCollector } from "@stream/browser/collector";

// Classes behind the factories — public for `instanceof` checks and subclassing
export { PullStream, BufferedStream } from "@stream/browser/factories";
export { Collector } from "@stream/browser/collector";
export { ChunkedBuilder, TransactionalChunkedBuilder } from "@stream/chunked-builder";

// Adapters for event emitters, sinks and callbacks — a labelled low-level corner
export * as LowLevel from "@stream/surface/streams-low-level";

// Consuming a whole stream. `bytes` / `text` / `json` are nouns on purpose: they mirror the
// body readers of `Response` and `Blob` (`await res.text()`) and of `node:stream/consumers`,
// so `await Streams.text(stream)` reads the way every Web and Node caller already writes it.
export {
  streamToUint8Array as bytes,
  readText as text,
  readJson as json,
  collect,
  drainStream as discard,
  readableStreamToAsyncIterable
} from "@stream/core/stream-helpers";

// Completion of several streams at once
export { finishedAll as waitForFinished } from "@stream/browser/pipeline";

// Web Streams conversions
export {
  readableFromWeb as fromWebReadable,
  readableToWeb as toWebReadable,
  writableFromWeb as fromWebWritable,
  writableToWeb as toWebWritable,
  toWritable
} from "@stream/browser/web";

// Type guards beyond node:stream's
export { isTransform, isDuplex, isStream } from "@stream/browser/utils";
export {
  isReadableStream,
  isWritableStream,
  isTransformStream,
  isAsyncIterable
} from "@stream/core/type-guards";
