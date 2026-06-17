/**
 * `Layout` namespace surface — document layout and SVG rendering.
 *
 * `import { Layout } from "documonster/word"` →
 *   `Layout.document(doc)`, `Layout.documentFull(doc)`,
 *   `Layout.renderDocumentToSvg(doc)`, … — tree-shaken via `export * as Layout`.
 */
export { layoutDocument as document } from "../layout/layout";
export { layoutDocumentFull as documentFull } from "../layout/layout-full";
export { renderPageToSvg, renderDocumentToSvg, renderPageFromLayout } from "../layout/render-page";
