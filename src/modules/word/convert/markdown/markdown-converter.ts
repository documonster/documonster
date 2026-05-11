/**
 * DOCX Module - DOCX to Markdown Converter (Legacy Alias)
 *
 * This module re-exports `renderToMarkdown` as `docxToMarkdown` for backward
 * compatibility. New code should use `renderToMarkdown` from "excelts/word/markdown".
 *
 * @example
 * ```ts
 * import { renderToMarkdown } from "excelts/word/markdown";
 *
 * const doc = await readDocx(buffer);
 * const md = renderToMarkdown(doc);
 * ```
 */

import type { DocxDocument } from "../../types";
import { renderToMarkdown } from "./markdown-renderer";
import type { MarkdownRenderOptions } from "./markdown-renderer";

/** Options for DOCX to Markdown conversion. */
export interface MarkdownOptions {
  /** Include images as markdown references (default: true). */
  readonly includeImages?: boolean;
}

/**
 * Convert a DocxDocument to a GFM-compatible Markdown string.
 *
 * @deprecated Use `renderToMarkdown` from "excelts/word/markdown" instead.
 * @param doc - The document model to convert.
 * @param options - Optional conversion settings.
 * @returns Markdown string.
 */
export function docxToMarkdown(doc: DocxDocument, options?: MarkdownOptions): string {
  const renderOpts: MarkdownRenderOptions = {
    includeImages: options?.includeImages
  };
  return renderToMarkdown(doc, renderOpts);
}
