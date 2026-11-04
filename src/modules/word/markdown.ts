/**
 * DOCX Module - Markdown Converters (Subpath Entry)
 *
 * Public API at `documonster/word/markdown`. Referenced by
 * `package.json#exports["./word/markdown"]`.
 *
 * @example
 * ```ts
 * import { WordMarkdown } from "documonster/word/markdown";
 *
 * const doc = await WordMarkdown.toDocx("# Title");
 * const md = WordMarkdown.fromDocx(doc);
 * ```
 */
import {
  markdownToDocx as markdownToDocxImpl,
  markdownToDocxBody as markdownToDocxBodyImpl
} from "@word/convert/markdown/markdown-import";
import { renderToMarkdown as renderToMarkdownImpl } from "@word/convert/markdown/markdown-renderer";

export * as WordMarkdown from "@word/surface/markdown";

// The errors this entry's functions throw (also exported from `documonster/word`, the same
// class objects): a wrong argument → `DocxArgumentError`, a picture `resolveImage` returns
// with a media type Word cannot hold → `DocxInvalidImageError`.
export { DocxArgumentError, DocxError, DocxInvalidImageError, isDocxError } from "@word/errors";
// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
export type {
  MarkdownRenderOptions,
  MarkdownImportOptions,
  MarkdownImageData,
  MarkdownBodyResult
} from "@word/convert/markdown/markdown";

/** @deprecated Use {@link WordMarkdown.fromDocx}. Removed in 1.0. */
export const renderToMarkdown = renderToMarkdownImpl;
/** @deprecated Use {@link WordMarkdown.toDocx}. Removed in 1.0. */
export const markdownToDocx = markdownToDocxImpl;
/** @deprecated Use {@link WordMarkdown.toDocxBody}. Removed in 1.0. */
export const markdownToDocxBody = markdownToDocxBodyImpl;
