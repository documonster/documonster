/**
 * DOCX Module - Markdown Converters (Subpath Export)
 *
 * Import separately to avoid pulling renderer code into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { renderToMarkdown, markdownToDocx, markdownToDocxBody } from "excelts/word/markdown";
 * ```
 */

// Primary API — Export (DOCX → Markdown)
export { renderToMarkdown } from "./markdown-renderer";
export type { MarkdownRenderOptions } from "./markdown-renderer";

// Primary API — Import (Markdown → DOCX)
export { markdownToDocx, markdownToDocxBody } from "./markdown-import";
export type { MarkdownImportOptions, MarkdownImageData } from "./markdown-import";

// Legacy alias (deprecated — use renderToMarkdown)
export { docxToMarkdown } from "./markdown-converter";
export type { MarkdownOptions } from "./markdown-converter";
