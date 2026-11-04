/**
 * Async CSV Parser
 *
 * Provides async CSV parsing supporting:
 * - String input (delegates to sync parser)
 * - AsyncIterable<string|Uint8Array> inputs
 * - ReadableStream inputs (WHATWG streams)
 *
 * API Semantics:
 * - parseCsvAsync(): Collects all rows into memory, returns full CsvParseResult.
 *   Best for small-to-medium files where you need complete data at once.
 *
 * - parseCsvRows(): True streaming async generator that yields rows one at a time.
 *   Best for large files or when processing rows incrementally.
 *   Memory-efficient as it doesn't buffer the entire result.
 *
 * Note: Both functions are async but have different memory characteristics.
 * parseCsvAsync buffers the entire input; parseCsvRows streams progressively.
 */

import { getUtf8ByteLength } from "@csv/constants";
import { CsvArgumentError } from "@csv/errors";
import { parseCsvAt, recordErrorToThrow } from "@csv/parse/sync";
import { checkParserStreamOptions, CsvParserStream } from "@csv/stream/parser";
import type {
  CsvParseOptions,
  CsvParseArrayOptions,
  CsvParseObjectOptions,
  CsvParseResult,
  CsvRecordError,
  CsvParseResultWithObjname,
  RecordWithInfo
} from "@csv/types";
import { createCsvTextDecoder, describeValue } from "@csv/utils/check-options";
import { isReadableStreamLike, readableStreamToAsyncIterable } from "@stream/core/stream-helpers";
import type { ReadableStreamLike } from "@stream/core/stream-helpers";
import { toError } from "@utils/errors";

/**
 * Rows buffered for the generator's consumer before the parser is paused.
 *
 * Small enough that a slow consumer cannot accumulate the whole input, large enough that a
 * fast one is never throttled between pulls.
 */
const ROW_QUEUE_HIGH_WATER_MARK = 1024;
const ROW_QUEUE_LOW_WATER_MARK = 256;

type AsyncInput = AsyncIterable<string | Uint8Array>;

/**
 * Everything the async / streaming CSV parsers accept as input: a whole CSV
 * string, the whole input in bytes (any `ArrayBufferView`, an `ArrayBuffer`, or
 * a `Blob` / `File`, decoded like the same bytes streamed), any async iterable of
 * text or bytes (a Node `Readable`, an async generator, …), or a WHATWG
 * `ReadableStream` of bytes.
 *
 * Narrower than the runtime guard (`isReadableStreamLike` accepts any object
 * with a `getReader`), because these are the chunk shapes `collectText` can
 * actually decode.
 */
export type CsvAsyncInput =
  | string
  | ArrayBufferView
  | ArrayBuffer
  | Blob
  | AsyncIterable<string | Uint8Array>
  | ReadableStream<Uint8Array>;

/**
 * A single row produced by the streaming parser.
 *
 * The object / array / with-info variants are the normal (object-mode) output.
 * The `string` arm only occurs with `objectMode: false`, where the parser emits
 * each row as a JSON string.
 */
export type ParsedRow =
  | Record<string, unknown>
  | string[]
  | RecordWithInfo<Record<string, unknown>>
  | RecordWithInfo<string[]>
  | string;

function isAsyncIterable(value: unknown): value is AsyncInput {
  return Boolean(
    value &&
    typeof (value as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] === "function"
  );
}

/** Bytes held whole: any `ArrayBufferView` or an `ArrayBuffer` (from any realm). */
function isWholeBytes(value: unknown): value is ArrayBufferView | ArrayBuffer {
  return (
    ArrayBuffer.isView(value) ||
    value instanceof ArrayBuffer ||
    Object.prototype.toString.call(value) === "[object ArrayBuffer]"
  );
}

function isBlob(value: unknown): value is Blob {
  return typeof Blob !== "undefined" && value instanceof Blob;
}

async function* oneChunk(value: ArrayBufferView | ArrayBuffer): AsyncGenerator<Uint8Array> {
  yield ArrayBuffer.isView(value)
    ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
    : new Uint8Array(value);
}

function normalizeAsyncInput(input: unknown, where: string): AsyncInput {
  // Whole bytes are the byte stream they would be as one chunk, so they decode alike
  if (isWholeBytes(input)) {
    return oneChunk(input);
  }
  if (isBlob(input)) {
    return readableStreamToAsyncIterable(input.stream() as ReadableStream<Uint8Array>);
  }
  if (isAsyncIterable(input)) {
    return input;
  }
  if (isReadableStreamLike(input)) {
    // The documented input: a Web stream of text or bytes, like any other source here.
    return readableStreamToAsyncIterable(input as ReadableStreamLike<string | Uint8Array>);
  }
  throw new CsvArgumentError(
    "input",
    `${where}: input must be a string, bytes, a Blob, an AsyncIterable or a ReadableStream, got ${describeValue(input)}`
  );
}

/** Refuse an input none of the async parsers can read, before anything else happens. */
function checkAsyncInput(input: unknown, where: string): void {
  if (
    typeof input !== "string" &&
    !isWholeBytes(input) &&
    !isBlob(input) &&
    !isAsyncIterable(input) &&
    !isReadableStreamLike(input)
  ) {
    throw new CsvArgumentError(
      "input",
      `${where}: input must be a string, bytes, a Blob, an AsyncIterable or a ReadableStream, got ${describeValue(input)}`
    );
  }
}

/** Result from collectText with optional byte tracking */
interface CollectResult {
  content: string;
  totalBytes: number;
}

/**
 * Collect AsyncIterable<string|Uint8Array> to a complete string.
 * Unified helper for both collectAsyncInput and parseCsvWithProgress.
 *
 * @param input - Async iterable of chunks
 * @param encoding - `TextDecoder` label (default: utf-8)
 * @param onChunk - Optional callback for each chunk (for progress reporting)
 */
async function collectText(
  where: string,
  input: AsyncIterable<string | Uint8Array>,
  encoding?: string,
  onChunk?: (bytesProcessed: number) => void
): Promise<CollectResult> {
  const chunks: string[] = [];
  const decoder = createCsvTextDecoder(encoding, where);
  let totalBytes = 0;

  for await (const chunk of input) {
    if (typeof chunk === "string") {
      chunks.push(chunk);
      // Always track bytes for consistent semantics
      totalBytes += getUtf8ByteLength(chunk);
    } else {
      chunks.push(decoder.decode(chunk, { stream: true }));
      totalBytes += chunk.length;
    }

    if (onChunk) {
      onChunk(totalBytes);
    }
  }

  // Flush decoder
  const final = decoder.decode();
  if (final) {
    chunks.push(final);
  }

  return { content: chunks.join(""), totalBytes };
}

/**
 * Progress reported by `Csv.parseAsync` through `onProgress`.
 *
 * Parsing runs once the whole input has been collected, so progress comes in two phases:
 * while a streamed input is read, each call carries `bytesProcessed` so far and
 * `rowsProcessed: 0`; then exactly one final call carries the number of rows in the result
 * (data rows — a header row is not counted; with `objname`, the number of distinct keys)
 * and the total byte count. A string input skips the first phase.
 */
export interface CsvParseProgress {
  /** `0` while input is being read; the result's row count on the final call. */
  rowsProcessed: number;
  /** Input bytes read so far (UTF-8 length for string chunks); on the final call, the total. */
  bytesProcessed?: number;
}

/**
 * Options only the asynchronous parser understands. Accepted by every `Csv.parseAsync`
 * overload, since progress reporting does not change the shape of the result.
 */
export interface CsvParseAsyncOptions {
  /** Progress callback — see {@link CsvParseProgress} for when it is called. */
  onProgress?: (info: CsvParseProgress) => void;
}

/**
 * Parse CSV asynchronously.
 *
 * For string input, this simply wraps the sync parser.
 * For AsyncIterable input (streams), collects chunks and parses.
 *
 * @example
 * ```ts
 * // From string
 * const result = await Csv.parseAsync("a,b\n1,2", { headers: true });
 *
 * // From fetch response
 * const response = await fetch("data.csv");
 * const result = await Csv.parseAsync(response.body, { headers: true });
 *
 * // From file stream (Node.js)
 * import { createReadStream } from "fs";
 * const result = await Csv.parseAsync(createReadStream("data.csv"), { headers: true });
 * ```
 */

// =============================================================================
// Function Overloads for Better Type Inference
// =============================================================================

/**
 * Parse CSV async - returns string[][] when no options provided.
 */
export function parseCsvAsync(input: CsvAsyncInput): Promise<string[][]>;

/**
 * Parse CSV async - with `validate` (array mode) the result object is returned.
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseArrayOptions &
    CsvParseAsyncOptions & { validate: NonNullable<CsvParseOptions["validate"]> }
): Promise<CsvParseResult<string[] | RecordWithInfo<string[]>>>;

/**
 * Parse CSV async - returns string[][] when headers is false/undefined and no info option.
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseArrayOptions & CsvParseAsyncOptions & { info?: false }
): Promise<string[][]>;

/**
 * Parse CSV async - returns CsvParseResult with RecordWithInfo when info: true (array mode).
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseArrayOptions & CsvParseAsyncOptions & { info: true }
): Promise<CsvParseResult<RecordWithInfo<string[]>>>;

/**
 * Parse CSV async - with `objname`, rows are an object keyed by that column's value.
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseObjectOptions & CsvParseAsyncOptions & { objname: string }
): Promise<CsvParseResultWithObjname<Record<string, unknown>>>;

/**
 * Parse CSV async - returns CsvParseResult when headers are enabled.
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseObjectOptions & CsvParseAsyncOptions & { info?: false }
): Promise<CsvParseResult<Record<string, unknown>>>;

/**
 * Parse CSV async - returns CsvParseResult with RecordWithInfo when info: true (object mode).
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseObjectOptions & CsvParseAsyncOptions & { info: true }
): Promise<CsvParseResult<RecordWithInfo<Record<string, unknown>>>>;

/**
 * Parse CSV async - general overload for backward compatibility.
 */
export function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseOptions & CsvParseAsyncOptions
): Promise<
  | string[][]
  | CsvParseResult<Record<string, string>>
  | CsvParseResult<Record<string, unknown>>
  | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
  | CsvParseResult<RecordWithInfo<string[]>>
>;

/**
 * Parse CSV asynchronously (implementation).
 */
export async function parseCsvAsync(
  input: CsvAsyncInput,
  options: CsvParseOptions & CsvParseAsyncOptions = {}
): Promise<
  | string[][]
  | CsvParseResult<Record<string, string>>
  | CsvParseResult<Record<string, unknown>>
  | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
  | CsvParseResult<RecordWithInfo<string[]>>
  | CsvParseResult<string[] | unknown[]>
  | CsvParseResult<string[] | RecordWithInfo<string[]>>
  | CsvParseResultWithObjname<Record<string, unknown>>
> {
  const where = "Csv.parseAsync";
  checkAsyncInput(input, where);
  if (typeof options !== "object" || options === null) {
    throw new CsvArgumentError(
      "options",
      `${where}: options must be an options object, got ${describeValue(options)}`
    );
  }
  const { onProgress, ...parseOptions } = options;
  if (onProgress !== undefined && typeof onProgress !== "function") {
    throw new CsvArgumentError(
      "options.onProgress",
      `${where}: onProgress must be a function, got ${describeValue(onProgress)}`
    );
  }
  if (onProgress) {
    return parseWithProgressAt(where, input, parseOptions, onProgress);
  }
  // If input is a string, use sync parser directly
  if (typeof input === "string") {
    return parseCsvAt(where, input, parseOptions);
  }

  const asyncInput = normalizeAsyncInput(input, where);
  const { content } = await collectText(where, asyncInput, parseOptions.encoding);
  return parseCsvAt(where, content, parseOptions);
}

/**
 * Parse CSV as an async generator, yielding rows as they are parsed.
 * This is the true streaming version that yields rows one at a time.
 *
 * @example
 * ```ts
 * // Process large file row by row
 * for await (const row of Csv.parseRows(fileStream, { headers: true })) {
 *   console.log(row);
 * }
 *
 * // With validation
 * for await (const row of Csv.parseRows(input, {
 *   headers: true,
 *   validate: (row) => row.id !== ""
 * })) {
 *   // Only valid rows
 * }
 * ```
 */
export async function* parseCsvRows(
  input: CsvAsyncInput,
  options: CsvParseOptions = {}
): AsyncGenerator<ParsedRow, void, unknown> {
  const where = "Csv.parseRows";
  checkAsyncInput(input, where);
  checkParserStreamOptions(options, where);
  // A record left out because of an error is reported to `onSkip`; without one, a generator
  // has nowhere to put it, so it is thrown rather than dropped.
  let recordError: CsvRecordError | undefined;
  const streamOptions =
    typeof options === "object" && options !== null && options.onSkip === undefined
      ? {
          ...options,
          onSkip: (error: CsvRecordError) => {
            recordError ??= error;
          }
        }
      : options;
  const parser = new CsvParserStream(streamOptions);

  type StreamEvent =
    | { type: "data"; value: ParsedRow }
    | { type: "end" }
    | { type: "error"; error: unknown };

  const queue: StreamEvent[] = [];
  let pendingResolve: ((ev: StreamEvent) => void) | null = null;
  let ended = false;
  let streamError: unknown = null;
  let aborted = false;
  let paused = false;

  const pushEvent = (ev: StreamEvent): void => {
    if (pendingResolve) {
      const resolve = pendingResolve;
      pendingResolve = null;
      resolve(ev);
      return;
    }
    queue.push(ev);
  };

  // A `data` listener puts the parser in flowing mode, so without this the parser would read
  // the whole input into `queue` no matter how slowly the consumer of this generator pulls.
  // Pausing hands that backpressure through to the writer loop below, which already waits for
  // `drain`.
  const applyBackpressure = (): void => {
    if (!paused && queue.length >= ROW_QUEUE_HIGH_WATER_MARK) {
      paused = true;
      parser.pause();
    }
  };

  const releaseBackpressure = (): void => {
    if (paused && queue.length <= ROW_QUEUE_LOW_WATER_MARK) {
      paused = false;
      parser.resume();
    }
  };

  const onData = (value: ParsedRow): void => {
    pushEvent({ type: "data", value });
    applyBackpressure();
  };
  const onEnd = (): void => {
    ended = true;
    pushEvent({ type: "end" });
  };
  const onError = (error: unknown): void => {
    streamError = error;
    pushEvent({ type: "error", error });
  };

  parser.on("data", onData);
  parser.once("end", onEnd);
  parser.once("error", onError);

  const writePromise = (async (): Promise<void> => {
    try {
      if (typeof input === "string") {
        parser.end(input);
        return;
      }

      const asyncInput = normalizeAsyncInput(input, where);

      for await (const chunk of asyncInput) {
        if (aborted) {
          break;
        }
        const canContinue = typeof chunk === "string" ? parser.write(chunk) : parser.write(chunk);
        if (!canContinue) {
          await new Promise<void>(resolve => parser.once("drain", resolve));
        }
      }

      if (!aborted) {
        parser.end();
      } else {
        parser.destroy();
      }
    } catch (e) {
      parser.destroy(e as Error);
    }
  })();

  try {
    while (true) {
      if (queue.length > 0) {
        const ev = queue.shift()!;
        releaseBackpressure();
        if (recordError) {
          throw recordErrorToThrow(where, [recordError]);
        }
        if (ev.type === "data") {
          yield ev.value;
          continue;
        }
        if (ev.type === "error") {
          throw toError(ev.error);
        }
        // end
        break;
      }

      if (streamError) {
        throw toError(streamError);
      }
      if (ended) {
        break;
      }

      const ev = await new Promise<StreamEvent>(resolve => {
        pendingResolve = resolve;
      });

      if (recordError) {
        throw recordErrorToThrow(where, [recordError]);
      }
      if (ev.type === "data") {
        yield ev.value;
      } else if (ev.type === "error") {
        throw toError(ev.error);
      } else {
        break;
      }
    }
  } finally {
    aborted = true;
    if (paused) {
      // Let the writer loop finish rather than leaving it blocked behind a paused parser.
      paused = false;
      parser.resume();
    }
    // Ensure stream stops as soon as possible.
    parser.destroy();
    // Release the writer if it is waiting for drain (destroy does not emit drain).
    parser.emit("drain");
    parser.off("data", onData);
    parser.off("end", onEnd);
    parser.off("error", onError);
    // Avoid unhandled rejections from the writer task.
    await writePromise.catch(() => undefined);
  }
}

/**
 * Parse CSV with progress callback for large files.
 *
 * @param input - CSV string or async iterable
 * @param options - Parse options
 * @param onProgress - Called periodically with progress info
 */
export async function parseCsvWithProgress(
  input: CsvAsyncInput,
  options: CsvParseOptions = {},
  onProgress?: (info: CsvParseProgress) => void
): Promise<
  | string[][]
  | CsvParseResult<Record<string, string>>
  | CsvParseResult<Record<string, unknown>>
  | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
  | CsvParseResult<RecordWithInfo<string[]>>
> {
  return parseWithProgressAt("Csv.parseWithProgress", input, options, onProgress) as Promise<
    | string[][]
    | CsvParseResult<Record<string, string>>
    | CsvParseResult<Record<string, unknown>>
    | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
    | CsvParseResult<RecordWithInfo<string[]>>
  >;
}

async function parseWithProgressAt(
  where: string,
  input: CsvAsyncInput,
  options: CsvParseOptions,
  onProgress: ((info: CsvParseProgress) => void) | undefined
): Promise<ReturnType<typeof parseCsvAt>> {
  // Collect input and track bytes
  checkAsyncInput(input, where);
  let content: string;
  let totalBytes: number;

  if (typeof input === "string") {
    content = input;
    totalBytes = getUtf8ByteLength(content);
  } else {
    const asyncInput = normalizeAsyncInput(input, where);
    const result = await collectText(
      where,
      asyncInput,
      options.encoding,
      onProgress ? bytes => onProgress({ rowsProcessed: 0, bytesProcessed: bytes }) : undefined
    );
    content = result.content;
    totalBytes = result.totalBytes;
  }

  // Parse
  const result = parseCsvAt(where, content, options);

  // Report final progress
  if (onProgress) {
    let rowCount: number;
    if (Array.isArray(result)) {
      rowCount = result.length;
    } else if (typeof result.rows === "object" && !Array.isArray(result.rows)) {
      // objname mode: rows is a plain object
      rowCount = Object.keys(result.rows).length;
    } else if (Array.isArray(result.rows)) {
      rowCount = result.rows.length;
    } else {
      rowCount = 0;
    }
    onProgress({ rowsProcessed: rowCount, bytesProcessed: totalBytes });
  }

  return result;
}
