/**
 * Stream Module - Common Consumers Factory
 *
 * Creates the `consumers` object both entries publish. Its `bytes`, `json` and `text` are
 * the `Streams` readers, so the two surfaces cannot disagree; only `buffer()`'s result type
 * is the platform's.
 */

// =============================================================================
// Types
// =============================================================================

import { readJson, readTextAt, streamToUint8Array } from "@stream/core/stream-helpers";

/** What `node:stream/consumers` reads: an async iterable, a Web stream or a plain iterable. */
type StreamInput =
  | AsyncIterable<Uint8Array>
  | ReadableStream<Uint8Array>
  | Iterable<Uint8Array | string>;

// Use the `Blob` constructor's own options type (available from both
// `@types/node` 18+ and the DOM lib) instead of the DOM-only global
// `BlobPropertyBag`, so the emitted `.d.ts` type-checks in Node-only projects
// without `lib: ["DOM"]`.
type BlobOptions = NonNullable<ConstructorParameters<typeof Blob>[1]>;

export interface StreamConsumers {
  arrayBuffer(stream: StreamInput): Promise<ArrayBuffer>;
  blob(stream: StreamInput, options?: BlobOptions): Promise<Blob>;
  /** The whole stream; a Node `Buffer` on Node, as `node:stream/consumers` returns. */
  buffer(stream: StreamInput): Promise<Uint8Array>;
  /** The whole stream as a plain `Uint8Array`, on every platform. */
  bytes(stream: StreamInput): Promise<Uint8Array>;
  json(stream: StreamInput): Promise<unknown>;
  /**
   * The whole stream as text, as `node:stream/consumers.text`: a leading BOM is stripped.
   * `encoding` (a `TextDecoder` label, default `"utf-8"`) is this module's addition.
   */
  text(stream: StreamInput, encoding?: string): Promise<string>;
}

interface StreamConverters {
  /** The platform's `buffer()` result for collected bytes; the bytes themselves if omitted. */
  toBuffer?(bytes: Uint8Array): Uint8Array;
}

// =============================================================================
// Factory
// =============================================================================

/**
 * Create a `consumers` object bound to platform-specific stream converters.
 *
 * Both Node.js and browser implementations delegate to their own
 * `streamToUint8Array` / `streamToString`; everything else is identical.
 */
export function createConsumers(converters: StreamConverters = {}): StreamConsumers {
  const { toBuffer = (bytes: Uint8Array) => bytes } = converters;

  return {
    async arrayBuffer(stream: StreamInput): Promise<ArrayBuffer> {
      const bytes = await streamToUint8Array(stream);
      if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
        return bytes.buffer as ArrayBuffer;
      }
      return bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength
      ) as ArrayBuffer;
    },

    async blob(stream: StreamInput, options?: BlobOptions): Promise<Blob> {
      const bytes = await streamToUint8Array(stream);
      return new Blob([bytes as BlobPart], options);
    },

    async buffer(stream: StreamInput): Promise<Uint8Array> {
      return toBuffer(await streamToUint8Array(stream));
    },

    // The `Streams` readers themselves, so there is one implementation of each.
    bytes: streamToUint8Array,
    json: readJson,

    text(stream: StreamInput, encoding?: string): Promise<string> {
      return readTextAt("consumers.text", stream, { encoding });
    }
  };
}
