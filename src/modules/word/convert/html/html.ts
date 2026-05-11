/**
 * DOCX Module - HTML Converters (Subpath Export)
 *
 * Import separately to avoid pulling converter code into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { renderToHtml, htmlToDocxBody } from "excelts/word/html";
 * ```
 */

// HTML → render (DocxDocument → HTML output)
export { renderToHtml } from "./html-renderer";
export type { HtmlRenderOptions, HtmlRenderResult } from "./html-renderer";

// HTML → DOCX import (HTML string → BodyContent[])
export { htmlToDocxBody } from "./html-import";
export type { HtmlImportOptions } from "./html-import";
