/**
 * `Streaming` namespace surface — streaming DOCX writer.
 *
 * `import { Streaming } from "documonster/word"` →
 *   `Streaming.createDocxStream(opts)`, `new Streaming.StreamingDocxWriter(...)`
 *   — tree-shaken via `export * as Streaming`.
 */
export { StreamingDocxWriter, createDocxStream } from "@word/writer/streaming-writer";
