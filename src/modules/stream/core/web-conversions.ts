/**
 * Web Streams conversions with the module's guarantees, on both entry points.
 *
 * `Readable.fromWeb` / `.toWeb` and friends stay exactly what each platform's
 * classes provide — on the Node entry, Node's own, with Node's known defects
 * (see the README). These functions run the module's adapters instead, with
 * that platform's classes, so the same call behaves the same on both:
 *
 * - no chunk is lost: a write made while `'drain'` is pending still reaches the
 *   Node-style stream, and a `null` read from a Web stream fails it rather than
 *   ending it early;
 * - strings from `setEncoding()` can be queued;
 * - the lock on a Web stream is handed back once the adapted stream ends,
 *   errors or is destroyed, and is never taken at all when adapting fails;
 * - the adapters remove their own listeners when they are done.
 */

import { isReadableStream, isWritableStream } from "@stream/core/type-guards";
import { StreamArgumentError, StreamUnsupportedTypeError } from "@stream/errors";
import { isEncoding } from "@stream/kernel/state";
import {
  readableFromWeb as adaptReadableFromWeb,
  readableToWeb as adaptReadableToWeb,
  writableFromWeb as adaptWritableFromWeb,
  writableToWeb as adaptWritableToWeb
} from "@stream/kernel/web-adapters";
import type { IReadable, IWritable, WritableLike } from "@stream/types";

/** `stream` if it can be written to as a Node-style writable (`write` and `end`); else a typed error. */
export function checkWritableLike(stream: unknown): WritableLike {
  const s = stream as { write?: unknown; end?: unknown } | null;
  if (
    typeof s !== "object" ||
    s === null ||
    typeof s.write !== "function" ||
    typeof s.end !== "function"
  ) {
    throw new StreamUnsupportedTypeError(
      "Streams.toWritable",
      stream === null ? "null" : typeof stream,
      {
        field: "stream"
      }
    );
  }
  return stream as WritableLike;
}

/** Options for adapting a Web `ReadableStream` into a Node-style Readable. */
export interface WebReadableOptions {
  objectMode?: boolean;
  highWaterMark?: number;
  encoding?: string;
  signal?: AbortSignal;
}

/** Options for adapting a Web `WritableStream` into a Node-style Writable. */
export interface WebWritableOptions {
  objectMode?: boolean;
  highWaterMark?: number;
  decodeStrings?: boolean;
  signal?: AbortSignal;
}

type StreamCtor = new (options?: any) => any;

/**
 * A Web `ReadableStream` as a Node-style Readable of the platform's class,
 * typed by what it yields. The default is byte mode, as in Node: every chunk
 * reads back as a `Uint8Array`, and only strings and binary views can be
 * carried. `objectMode: true` passes chunks through; an `encoding` decodes
 * them to strings. A mode known only at run time gets every possibility.
 */
export interface ReadableFromWeb {
  (
    stream: ReadableStream<string | ArrayBufferView>,
    options: WebReadableOptions & { encoding: string }
  ): IReadable<string>;
  <T>(
    stream: ReadableStream<T>,
    options: WebReadableOptions & { objectMode: true; encoding?: undefined }
  ): IReadable<T>;
  (
    stream: ReadableStream<string | ArrayBufferView>,
    options?: WebReadableOptions & { objectMode?: false; encoding?: undefined }
  ): IReadable<Uint8Array>;
  /**
   * A mode known only at run time — a `boolean`, or options held in a
   * variable. It may be bytes, so the stream must carry what byte mode can.
   */
  <T extends string | ArrayBufferView>(
    stream: ReadableStream<T>,
    options?: WebReadableOptions
  ): IReadable<T | Uint8Array | string>;
}

/** A Web `WritableStream` as a Node-style Writable of the platform's class. */
export type WritableFromWeb = <T = any>(
  stream: WritableStream<T>,
  options?: WebWritableOptions
) => IWritable<T>;

function describeValue(value: unknown): string {
  return value === null ? "null" : typeof value;
}

type OptionKind = "boolean" | "highWaterMark" | "encoding" | "signal" | "strategy";

function optionProblem(kind: OptionKind, value: unknown): string | undefined {
  switch (kind) {
    case "boolean":
      return typeof value === "boolean" ? undefined : "a boolean";
    case "highWaterMark":
      return typeof value === "number" && Number.isInteger(value) && value >= 0
        ? undefined
        : "a non-negative integer highWaterMark";
    case "encoding":
      return isEncoding(value) ? undefined : "a Buffer encoding name";
    case "signal": {
      const signal = value as { aborted?: unknown; addEventListener?: unknown } | null;
      return typeof signal === "object" &&
        signal !== null &&
        typeof signal.aborted === "boolean" &&
        typeof signal.addEventListener === "function"
        ? undefined
        : "an AbortSignal";
    }
    case "strategy":
      return typeof value === "object" && value !== null ? undefined : "a QueuingStrategy";
  }
}

/**
 * Refuse an options argument the adapter would hand to a stream constructor unchecked: there it
 * failed with the platform's own `TypeError`, naming no public call.
 */
function checkOptions(
  where: string,
  options: unknown,
  fields: Readonly<Record<string, OptionKind>>
): void {
  if (options === undefined) {
    return;
  }
  if (typeof options !== "object" || options === null || Array.isArray(options)) {
    throw new StreamArgumentError("an options object", describeValue(options), {
      field: "options",
      where
    });
  }
  for (const [name, kind] of Object.entries(fields)) {
    const value = (options as Record<string, unknown>)[name];
    const expected = value === undefined ? undefined : optionProblem(kind, value);
    if (expected !== undefined) {
      throw new StreamArgumentError(expected, describeValue(value), {
        field: `options.${name}`,
        where
      });
    }
  }
}

const READABLE_FROM_WEB_OPTIONS = {
  objectMode: "boolean",
  highWaterMark: "highWaterMark",
  encoding: "encoding",
  signal: "signal"
} as const;

const WRITABLE_FROM_WEB_OPTIONS = {
  objectMode: "boolean",
  highWaterMark: "highWaterMark",
  decodeStrings: "boolean",
  signal: "signal"
} as const;

export function makeReadableFromWeb(Readable: StreamCtor): ReadableFromWeb {
  return (stream: ReadableStream, options?: WebReadableOptions) => {
    if (!isReadableStream(stream)) {
      throw new StreamUnsupportedTypeError("Streams.fromWebReadable", describeValue(stream), {
        field: "stream"
      });
    }
    checkOptions("Streams.fromWebReadable", options, READABLE_FROM_WEB_OPTIONS);
    return adaptReadableFromWeb(Readable, stream, options);
  };
}

export function makeWritableFromWeb(Writable: StreamCtor): WritableFromWeb {
  return (stream, options) => {
    if (!isWritableStream(stream)) {
      throw new StreamUnsupportedTypeError("Streams.fromWebWritable", describeValue(stream), {
        field: "stream"
      });
    }
    checkOptions("Streams.fromWebWritable", options, WRITABLE_FROM_WEB_OPTIONS);
    return adaptWritableFromWeb(Writable, stream, options);
  };
}

/** A Node-style Readable as a Web `ReadableStream`. */
export function readableToWeb<T = any>(
  stream: IReadable<T>,
  options?: { strategy?: QueuingStrategy<T> }
): ReadableStream<T> {
  const s = stream as { on?: unknown; read?: unknown } | null;
  if (
    typeof s !== "object" ||
    s === null ||
    typeof s.on !== "function" ||
    typeof s.read !== "function"
  ) {
    throw new StreamUnsupportedTypeError("Streams.toWebReadable", describeValue(stream), {
      field: "stream"
    });
  }
  checkOptions("Streams.toWebReadable", options, { strategy: "strategy" });
  return adaptReadableToWeb<T>(stream, options);
}

/** A Node-style Writable as a Web `WritableStream`. */
export function writableToWeb<T = any>(stream: IWritable<T>): WritableStream<T> {
  const s = stream as { on?: unknown; write?: unknown } | null;
  if (
    typeof s !== "object" ||
    s === null ||
    typeof s.on !== "function" ||
    typeof s.write !== "function"
  ) {
    throw new StreamUnsupportedTypeError("Streams.toWebWritable", describeValue(stream), {
      field: "stream"
    });
  }
  return adaptWritableToWeb<T>(stream);
}
