/**
 * `WordMarkdown` namespace surface — DOCX ↔ Markdown, at
 * `documonster/word/markdown`.
 *
 * `WordMarkdown.fromDocx(doc)` renders a document to GFM; `WordMarkdown.toDocx(md)`
 * and `WordMarkdown.toDocxBody(md)` import GFM as a document or body content.
 */
export { renderToMarkdown as fromDocx } from "@word/convert/markdown/markdown-renderer";
export {
  markdownToDocx as toDocx,
  markdownToDocxBody as toDocxBody
} from "@word/convert/markdown/markdown-import";
