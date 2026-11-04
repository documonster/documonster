import { createDocxStreamReader as createDocxStreamReaderImpl } from "@word/reader/streaming-reader";
/**
 * `Streaming` namespace surface — streaming DOCX writer + reader.
 *
 * `import { Streaming } from "documonster/word"` →
 *   `Streaming.createWriter(opts)`, `Streaming.createReader(bytes)`
 *   (or `new Streaming.StreamingDocxWriter(...)` /
 *   `new Streaming.StreamingDocxReader(...)`) — tree-shaken via
 *   `export * as Streaming`.
 */
import { createDocxStream as createDocxStreamImpl } from "@word/writer/streaming-writer";

export {
  StreamingDocxWriter,
  createDocxStream as createWriter
} from "@word/writer/streaming-writer";
export {
  StreamingDocxReader,
  createDocxStreamReader as createReader
} from "@word/reader/streaming-reader";

/** @deprecated Use {@link createWriter}. Removed in 1.0. */
export const createDocxStream = createDocxStreamImpl;
/** @deprecated Use {@link createReader}. Removed in 1.0. */
export const createDocxStreamReader = createDocxStreamReaderImpl;
