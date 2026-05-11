/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/writer"` */

export { packageDocx } from "./docx-packager";
export { StreamingDocxWriter, createDocxStream } from "./streaming-writer";
export type { StreamingDocxOptions, StreamingProgressCallback } from "./streaming-writer";
