/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/layout"` */

export { layoutDocument } from "./layout";
export type { LayoutResult, LayoutOptions } from "./layout";
export { layoutDocumentFull } from "./layout-full";
export type { FullLayoutOptions } from "./layout-full";
export { renderPageToSvg, renderDocumentToSvg, renderPageFromLayout } from "./render-page";
export type { RenderOptions } from "./render-page";

export type {
  LayoutPoint,
  LayoutRect,
  PositionedRun,
  LineBox,
  LayoutParagraph,
  LayoutTable,
  LayoutTableCell,
  LayoutImage,
  LayoutFloat,
  PageContent,
  PageGeometry,
  LayoutPage,
  LayoutDocument
} from "./layout-model";
