/**
 * `Layout` namespace surface — document layout and SVG rendering.
 *
 * `import { Layout } from "documonster/word"` →
 *   `Layout.getPageMap(doc)` (which page each body item lands on, page counts,
 *   bookmark pages), `Layout.compute(doc)` (fully positioned pages),
 *   `Layout.toSvg(doc)` (every page), `Layout.pageToSvg(doc, 1)` (one page),
 *   `Layout.computedPageToSvg(layout, 1)` (one page of a layout from `compute`,
 *   without laying the document out again) — tree-shaken via
 *   `export * as Layout`.
 */
import { layoutDocumentFull, paginateDocument } from "@word/layout/layout-full";
import {
  renderDocumentToSvg as renderDocumentToSvgImpl,
  renderPageFromLayout as renderPageFromLayoutImpl,
  renderPageToSvg as renderPageToSvgImpl
} from "@word/layout/render-page";

export { paginateDocument as getPageMap } from "@word/layout/layout-full";
export { layoutDocumentFull as compute } from "@word/layout/layout-full";
export {
  renderDocumentToSvg as toSvg,
  renderPageToSvg as pageToSvg,
  renderPageFromLayout as computedPageToSvg
} from "@word/layout/render-page";

/** @deprecated Use {@link getPageMap}. Removed in 1.0. */
export const document = paginateDocument;
/** @deprecated Use {@link compute}. Removed in 1.0. */
export const documentFull = layoutDocumentFull;
/** @deprecated Use {@link toSvg}. Removed in 1.0. */
export const renderDocumentToSvg = renderDocumentToSvgImpl;
/** @deprecated Use {@link pageToSvg}. Removed in 1.0. */
export const renderPageToSvg = renderPageToSvgImpl;
/** @deprecated Use {@link computedPageToSvg}. Removed in 1.0. */
export const renderPageFromLayout = renderPageFromLayoutImpl;
