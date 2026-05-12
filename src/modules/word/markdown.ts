/**
 * DOCX Module - Markdown Converters (Subpath Entry)
 *
 * Re-exports the Markdown converter API at `excelts/word/markdown`. This file
 * is referenced by `package.json#exports["./word/markdown"]`; it forwards to
 * the implementation under `./convert/markdown`.
 *
 * @example
 * ```ts
 * import { renderToMarkdown, markdownToDocx, markdownToDocxBody } from "excelts/word/markdown";
 * ```
 */

export { renderToMarkdown, markdownToDocx, markdownToDocxBody } from "./convert/markdown/markdown";
export type {
  MarkdownRenderOptions,
  MarkdownImportOptions,
  MarkdownImageData
} from "./convert/markdown/markdown";
