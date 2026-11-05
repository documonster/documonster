/**
 * `Stream` namespace surface — streaming DOCX writer + reader.
 *
 * `import { Stream } from "documonster/word"` →
 *   `new Stream.DocumentWriter(opts)`, `new Stream.DocumentReader(bytes)` —
 *   tree-shaken via `export * as Stream`. Named after the primary object, as
 *   Excel's streaming workbook writer and reader are.
 */
import { DocumentReader } from "@word/reader/streaming-reader";
import type { StreamingReadOptions } from "@word/reader/streaming-reader";
import { DocumentWriter } from "@word/writer/streaming-writer";
import type { StreamingDocxOptions } from "@word/writer/streaming-writer";

export { DocumentWriter, DocumentReader };

/** @deprecated Use {@link DocumentWriter}. Removed in 1.0. */
export const StreamingDocxWriter = DocumentWriter;
/** @deprecated Use {@link DocumentWriter}. Removed in 1.0. */
export type StreamingDocxWriter = DocumentWriter;
/** @deprecated Use {@link DocumentReader}. Removed in 1.0. */
export const StreamingDocxReader = DocumentReader;
/** @deprecated Use {@link DocumentReader}. Removed in 1.0. */
export type StreamingDocxReader = DocumentReader;

/** @deprecated Use `new Stream.DocumentWriter(options)`. Removed in 1.0. */
export function createWriter(options?: StreamingDocxOptions): DocumentWriter {
  return new DocumentWriter(options);
}

/** @deprecated Use `new Stream.DocumentReader(source, options)`. Removed in 1.0. */
export function createReader(
  source: Uint8Array | ArrayBuffer | ArrayBufferView,
  options?: StreamingReadOptions
): DocumentReader {
  return new DocumentReader(source, options);
}

/** @deprecated Use `new Stream.DocumentWriter(options)`. Removed in 1.0. */
export const createDocxStream = createWriter;
/** @deprecated Use `new Stream.DocumentReader(source, options)`. Removed in 1.0. */
export const createDocxStreamReader = createReader;
