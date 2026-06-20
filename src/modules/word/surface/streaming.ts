/**
 * `Streaming` namespace surface — streaming DOCX writer + reader.
 *
 * `import { Streaming } from "documonster/word"` →
 *   `Streaming.createDocxStream(opts)`, `new Streaming.StreamingDocxWriter(...)`,
 *   `Streaming.createDocxStreamReader(bytes)`,
 *   `new Streaming.StreamingDocxReader(...)`
 *   — tree-shaken via `export * as Streaming`.
 */
export { StreamingDocxWriter, createDocxStream } from "@word/writer/streaming-writer";
export { StreamingDocxReader, createDocxStreamReader } from "@word/reader/streaming-reader";
