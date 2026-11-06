import type { Workbook } from "@excel/core/workbook.browser";
import { decryptionOf, openWorkbookPackage } from "@excel/utils/container-errors";
import { readXlsxBytesInto } from "@excel/xlsx/read/package";
import type { XlsxParseStream, XlsxReadOptions } from "@excel/xlsx/types";
import { streamToUint8Array } from "@stream/core/stream-helpers";
import { concatUint8Arrays } from "@utils/binary";

/** What `Workbook.readStream` accepts: a Web `ReadableStream`, an (async) iterable of chunks, or an event stream. */
type StreamInput =
  | XlsxParseStream
  | ReadableStream<Uint8Array>
  | AsyncIterable<Uint8Array | string>
  | Iterable<Uint8Array>;

/**
 * Read a workbook from a stream — browser variant.
 *
 * There is no streaming ZIP parser on this platform, so the stream is collected and read as `Workbook.read` reads
 * bytes: an encrypted workbook decrypted with `options.password` first, then the package read with the same
 * container limits. It accepts every input `Workbook.readStream` declares — it used to take only an event stream
 * (`.on("data")`), so a Web `ReadableStream` or an async iterable failed with a raw `TypeError`, and a password was
 * ignored. The Node variant (`read/stream.ts`) parses entries as they arrive instead.
 */
export async function readXlsxStreamInto(
  workbook: Workbook,
  stream: StreamInput,
  options?: XlsxReadOptions
): Promise<Workbook> {
  const collected = await collect(stream);
  const bytes = await openWorkbookPackage(collected, decryptionOf(options));
  return readXlsxBytesInto(workbook, bytes, options);
}

function isIterableInput(
  stream: StreamInput
): stream is
  | ReadableStream<Uint8Array>
  | AsyncIterable<Uint8Array | string>
  | Iterable<Uint8Array> {
  const candidate = stream as {
    getReader?: unknown;
    [Symbol.asyncIterator]?: unknown;
    [Symbol.iterator]?: unknown;
  };
  return (
    typeof candidate.getReader === "function" ||
    typeof candidate[Symbol.asyncIterator] === "function" ||
    typeof candidate[Symbol.iterator] === "function"
  );
}

async function collect(stream: StreamInput): Promise<Uint8Array> {
  if (isIterableInput(stream)) {
    return streamToUint8Array(stream as Parameters<typeof streamToUint8Array>[0]);
  }
  const chunks: Uint8Array[] = [];
  await new Promise<void>((resolve, reject) => {
    const onData = (chunk: Uint8Array) => {
      chunks.push(chunk);
    };
    const detach = () => {
      stream.off("data", onData);
      stream.off("end", onEnd);
      stream.off("error", onError);
    };
    const onEnd = () => {
      detach();
      resolve();
    };
    const onError = (err: Error) => {
      detach();
      reject(err);
    };
    stream.on("data", onData);
    stream.on("end", onEnd);
    stream.on("error", onError);
  });
  return concatUint8Arrays(chunks);
}
