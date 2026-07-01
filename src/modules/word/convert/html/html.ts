/**
 * DOCX Module - HTML Converters (Subpath Export)
 *
 * Import separately to avoid pulling converter code into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { renderToHtml, htmlToDocxBody } from "documonster/word/html";
 * ```
 */

// HTML → render (DocxDocument → HTML output)
export { renderToHtml } from "@word/convert/html/html-renderer";
export type { HtmlRenderOptions, HtmlRenderResult } from "@word/convert/html/html-renderer";

// HTML → DOCX import (HTML string → BodyContent[])
export { htmlToDocxBody, htmlToDocx } from "@word/convert/html/html-import";
export type { HtmlImportOptions, HtmlToDocxResult } from "@word/convert/html/html-import";
