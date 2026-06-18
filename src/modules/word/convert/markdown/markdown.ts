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
export { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
export type { MarkdownRenderOptions } from "@word/convert/markdown/markdown-renderer";

// Import (Markdown → DOCX)
export { markdownToDocx, markdownToDocxBody } from "@word/convert/markdown/markdown-import";
export type {
  MarkdownImportOptions,
  MarkdownImageData,
  MarkdownBodyResult
} from "@word/convert/markdown/markdown-import";
