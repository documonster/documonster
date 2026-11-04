/**
 * DOCX Module - HTML Converters (Subpath Entry)
 *
 * Public API at `documonster/word/html`. Referenced by
 * `package.json#exports["./word/html"]`.
 *
 * @example
 * ```ts
 * import { WordHtml } from "documonster/word/html";
 *
 * const { html } = WordHtml.fromDocx(doc);
 * const doc2 = WordHtml.toDocx("<p>Hello</p>");
 * const { body, images, styles } = WordHtml.toDocxBody("<p>Hello</p>");
 * ```
 */
import {
  htmlToDocx as htmlToDocxImpl,
  htmlToDocxBody as htmlToDocxBodyImpl
} from "@word/convert/html/html-import";
import { renderToHtml as renderToHtmlImpl } from "@word/convert/html/html-renderer";

export * as WordHtml from "@word/surface/html";

// The errors this entry's functions throw (also exported from `documonster/word`, the same
// class objects): a wrong argument → `DocxArgumentError`.
export { DocxArgumentError, DocxError, isDocxError } from "@word/errors";
// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
export { HTML_BULLET_LIST_NUM_ID, HTML_NUMBERED_LIST_NUM_ID } from "@word/convert/html/html-import";
export type {
  HtmlRenderOptions,
  HtmlRenderResult,
  HtmlImportOptions,
  HtmlToDocxResult
} from "@word/convert/html/html";
export type { HtmlBodyResult } from "@word/convert/html/html-document";

/** @deprecated Use {@link WordHtml.fromDocx}. Removed in 1.0. */
export const renderToHtml = renderToHtmlImpl;
/**
 * Body content and images only.
 *
 * @deprecated Use {@link WordHtml.toDocxBody}, which also returns the styles
 * and list numbering the body refers to, or {@link WordHtml.toDocx} for a
 * complete document. Removed in 1.0.
 */
export const htmlToDocx = htmlToDocxImpl;
/**
 * Body content only; pictures from `data:` URLs are not decoded.
 *
 * @deprecated Use {@link WordHtml.toDocxBody}`(html).body`. Removed in 1.0.
 */
export const htmlToDocxBody = htmlToDocxBodyImpl;
