/**
 * `Layout` namespace surface — document layout and SVG rendering.
 *
 * `import { Layout } from "@cj-tech-master/excelts/word"` →
 *   `Layout.document(doc)`, `Layout.documentFull(doc)`,
 *   `Layout.renderDocumentToSvg(doc)`, … — tree-shaken via `export * as Layout`.
 */
export { layoutDocument as document } from "@word/layout/layout";
export { layoutDocumentFull as documentFull } from "@word/layout/layout-full";
export {
  renderPageToSvg,
  renderDocumentToSvg,
  renderPageFromLayout
} from "@word/layout/render-page";
