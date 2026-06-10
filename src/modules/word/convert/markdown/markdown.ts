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

// Export (DOCX → Markdown)
export { renderToMarkdown } from "./markdown-renderer";
export type { MarkdownRenderOptions } from "./markdown-renderer";

// Import (Markdown → DOCX)
export { markdownToDocx, markdownToDocxBody } from "./markdown-import";
export type {
  MarkdownImportOptions,
  MarkdownImageData,
  MarkdownBodyResult
} from "./markdown-import";
