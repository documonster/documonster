/**
 * `WordHtml` namespace surface — DOCX ↔ HTML, at `documonster/word/html`.
 *
 * `WordHtml.fromDocx(doc)` renders a document to HTML; `WordHtml.toDocx(html)`
 * returns a complete document, and `WordHtml.toDocxBody(html)` the body with
 * the styles, list numbering and images it refers to — as `WordMarkdown.toDocx`
 * and `WordMarkdown.toDocxBody` do for Markdown.
 */
export { renderToHtml as fromDocx } from "@word/convert/html/html-renderer";
export {
  htmlToDocument as toDocx,
  htmlToDocxFragment as toDocxBody
} from "@word/convert/html/html-document";
