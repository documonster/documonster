/**
 * DOCX Module - Markdown Converters (Subpath Entry)
 *
 * Re-exports the Markdown converter API at `documonster/word/markdown`. This file
 * is referenced by `package.json#exports["./word/markdown"]`; it forwards to
 * the implementation under `./convert/markdown`.
 *
 * @example
 * ```ts
 * import { renderToMarkdown, markdownToDocx, markdownToDocxBody } from "documonster/word/markdown";
 * ```
 */

export {
  renderToMarkdown,
  markdownToDocx,
  markdownToDocxBody
} from "@word/convert/markdown/markdown";
export type {
  MarkdownRenderOptions,
  MarkdownImportOptions,
  MarkdownImageData,
  MarkdownBodyResult
} from "@word/convert/markdown/markdown";
