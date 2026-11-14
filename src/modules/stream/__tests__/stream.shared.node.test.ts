/**
 * Stream Module Node Tests
 *
 * Runs the shared stream test suite against the Node.js implementation.
 * The same suite is also executed in browser tests to ensure behavior and
 * data integrity stay consistent across environments.
 */

import {
  Readable,
  Writable,
  Transform,
  Duplex,
  PassThrough,
  isReadable,
  isWritable,
  isDestroyed,
  isDisturbed,
  isErrored,
  pipeline,
  finished,
  addAbortSignal,
  compose,
  getDefaultHighWaterMark,
  setDefaultHighWaterMark,
  duplexPair,
  consumers,
  promises,
  Streams
} from "@stream";
import type { StreamModuleImports } from "@stream/__tests__/stream.shared";
import { runStreamTests } from "@stream/__tests__/stream.shared";
import {
  stringToUint8Array,
  uint8ArrayToString,
  uint8ArrayEquals,
  uint8ArrayIndexOf,
  concatUint8Arrays
} from "@utils/binary";
import { EventEmitter } from "@utils/event-emitter";
import { describe } from "vitest";

// What the running Node does, probed rather than read off its version: the behaviour changed inside a major line —
// Node 24.11 still wraps `compose`'s result through `Readable.from()` and leaves a `duplexPair` peer open, which later
// releases do not — so a version number named the wrong one.
const nativeComposeReturnsDuplex = ((): boolean => {
  const composed = Readable.from([]).compose(new PassThrough());
  // Destroyed unfinished, it emits an `AbortError` nothing else would listen for.
  composed.on("error", () => {});
  composed.destroy();
  return typeof (composed as { write?: unknown }).write === "function";
})();
const duplexPairDestroysPeerOnError = await new Promise<boolean>(resolve => {
  const [side, peer] = duplexPair();
  side.on("error", () => {});
  peer.on("error", () => {});
  side.destroy(new Error("probe"));
  setTimeout(() => {
    resolve(peer.destroyed === true);
    peer.destroy();
  }, 20);
});

describe("Stream Module - Shared Tests (Node)", () => {
  const imports: StreamModuleImports = {
    EventEmitter,
    Readable,
    Writable,
    Transform,
    Duplex,
    PassThrough,
    BufferedStream: Streams.BufferedStream,
    PullStream: Streams.PullStream,
    createStringChunk: Streams.createStringChunk,
    createByteChunk: Streams.createByteChunk,
    ChunkedBuilder: Streams.ChunkedBuilder,
    TransactionalChunkedBuilder: Streams.TransactionalChunkedBuilder,
    createReadable: Streams.createReadable,
    createWritable: Streams.createWritable,
    createPassThrough: Streams.createPassThrough,
    createTransform: Streams.createTransform,
    createCollector: (options?: ConstructorParameters<typeof Streams.Collector>[0]) =>
      new Streams.Collector(options),
    createDuplex: Streams.createDuplex,
    createReadableFromArray: Streams.createReadableFromArray,
    createReadableFromAsyncIterable: Streams.createReadableFromAsyncIterable,
    createReadableFromGenerator: Streams.createReadableFromGenerator,
    createReadableFromPromise: Streams.createReadableFromPromise,
    createEmptyReadable: Streams.createEmptyReadable,
    createNullWritable: Streams.createNullWritable,
    duplexPair,
    pipeline,
    finished,
    streamToUint8Array: Streams.bytes,
    streamToString: (stream, encoding) => Streams.text(stream, { encoding }),
    drainStream: Streams.discard,
    copyStream: pipeline,
    concatUint8Arrays,
    addAbortSignal,
    compose,
    finishedAll: Streams.waitForFinished,
    promisify: Streams.LowLevel.fromCallback,
    isReadable,
    isWritable,
    isTransform: Streams.isTransform,
    isDuplex: Streams.isDuplex,
    isStream: Streams.isStream,
    isDestroyed,
    isDisturbed,
    isErrored,
    getDefaultHighWaterMark,
    setDefaultHighWaterMark,
    consumers,
    promises,
    stringToUint8Array,
    uint8ArrayToString,
    uint8ArrayEquals,
    uint8ArrayIndexOf,
    nativeComposeReturnsDuplex,
    duplexPairDestroysPeerOnError
  };

  runStreamTests(imports);
});
