/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/convert"` */

export { renderToHtml } from "./html/html-renderer";
export type { HtmlRenderOptions, HtmlRenderResult } from "./html/html-renderer";
export { htmlToDocxBody } from "./html/html-import";
export type { HtmlImportOptions } from "./html/html-import";
export { renderToMarkdown } from "./markdown/markdown-renderer";
export type { MarkdownRenderOptions } from "./markdown/markdown-renderer";
export { markdownToDocx, markdownToDocxBody } from "./markdown/markdown-import";
export type { MarkdownImportOptions, MarkdownImageData } from "./markdown/markdown-import";
export { docxToMarkdown } from "./markdown/markdown-converter";
export type { MarkdownOptions } from "./markdown/markdown-converter";
export { readOdt, writeOdt } from "./odt/odt";
export { parseFlatOpc, toFlatOpc, isFlatOpc } from "./flat-opc";
export type { FlatOpcPart } from "./flat-opc";
export { docxToSemantic } from "./docx-to-semantic";
export type { DocxToSemanticOptions } from "./docx-to-semantic";
export { createConversionContext } from "./conversion-ir";
export type {
  ConversionSeverity,
  ConversionWarning,
  ConversionAsset,
  ResolvedFormatting,
  SemanticInline,
  SemanticBlock,
  SemanticParagraphStyle,
  SemanticListItem,
  SemanticTableRow,
  SemanticTableCell,
  SemanticNote,
  SemanticDocument,
  ConversionContext
} from "./conversion-ir";
