/**
 * DOCX Module - HTML Renderer (Subpath Export)
 *
 * Import separately to avoid pulling html-renderer into the bundle
 * when only core document building is needed.
 *
 * @example
 * ```ts
 * import { renderToHtml } from "excelts/word/html";
 * ```
 */

export { renderToHtml } from "./html-renderer";
export type { HtmlRenderOptions, HtmlRenderResult } from "./html-renderer";
