/**
 * `Markdown` namespace surface — GFM table parsing / formatting.
 *
 * `import { Markdown } from "documonster/markdown"` →
 *   `Markdown.parse(text)`, `Markdown.parseAll(text)`, `Markdown.format(rows)`.
 *
 * Single flat namespace (markdown is a single-purpose module). Re-exported
 * via `export * as Markdown`, tree-shaken per-member on rolldown / rspack.
 */
export { parseMarkdown as parse, parseMarkdownAll as parseAll } from "../parse/index";
export { formatMarkdown as format } from "../format/index";
