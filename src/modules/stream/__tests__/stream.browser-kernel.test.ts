/**
 * Stream Module Browser Tests
 *
 * Runs the shared stream test suite against the Browser implementation.
 * The same suite is also executed in Node.js tests to ensure behavior and
 * data integrity stay consistent across environments.
 */

import type { StreamModuleImports } from "@stream/__tests__/stream.shared";
import { runStreamTests } from "@stream/__tests__/stream.shared";
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
} from "@stream/index.browser";
import {
  stringToUint8Array,
  uint8ArrayToString,
  uint8ArrayEquals,
  uint8ArrayIndexOf,
  concatUint8Arrays
} from "@utils/binary";
import { EventEmitter } from "@utils/event-emitter";
import { describe } from "vitest";

// =============================================================================
// Run Shared Tests (Identical behavior for Node.js and Browser)
// =============================================================================
describe("Stream Module - Shared Tests (Browser kernel on Node)", () => {
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
    createCollector: Streams.createCollector,
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
    // Browser uses the project's own compose implementation which always returns a Duplex.
    nativeComposeReturnsDuplex: true,
    duplexPairDestroysPeerOnError: true
  };

  runStreamTests(imports);
});
