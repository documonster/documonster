/**
 * DOCX Module - HTML Converters (Subpath Entry)
 *
 * Re-exports the HTML converter API at `excelts/word/html`. This file is
 * referenced by `package.json#exports["./word/html"]`; it forwards to the
 * implementation under `./convert/html`.
 *
 * @example
 * ```ts
 * import { renderToHtml, htmlToDocxBody } from "excelts/word/html";
 * ```
 */

export { renderToHtml, htmlToDocxBody, htmlToDocx } from "@word/convert/html/html";
export type {
  HtmlRenderOptions,
  HtmlRenderResult,
  HtmlImportOptions,
  HtmlToDocxResult
} from "@word/convert/html/html";
