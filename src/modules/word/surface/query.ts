/**
 * `Query` namespace surface — read-only document queries (text, search,
 * fields, revisions, style resolution, compatibility, form fields, format
 * search, tree iteration). No `Query` function modifies the document it is
 * given.
 *
 * Edits live elsewhere, by ownership: `Document.*` edits the document you pass
 * in place (`Document.replaceText`, `Document.mailMerge`, the revision
 * functions, `Document.setCompatibilityMode`), and `Transform.*` returns a new
 * document (`Transform.fillFormFields`, `Transform.resolveDataBindings`). The
 * `Query` names of those functions remain as deprecated forwards.
 *
 * `import { Query } from "documonster/word"` →
 *   `Query.getText(doc)`, `Query.searchText(doc, /x/)`,
 *   `Query.listParagraphs(doc)`, … — tree-shaken via `export * as Query`.
 */
import {
  checkArray,
  checkHandle,
  checkNode,
  checkRecord,
  checkString
} from "@word/core/argument-checks";
import {
  collectParagraphs as collectParagraphsImpl,
  collectRuns as collectRunsImpl,
  collectTables as collectTablesImpl,
  walkBlocks as walkBlocksImpl,
  walkDocument as walkDocumentImpl
} from "@word/core/walker";
import type { DocxVisitor, WalkOptions, WalkPath } from "@word/core/walker";
import { setCompatibilityMode as setCompatibilityModeImpl } from "@word/query/compat";
import { resolveDataBindings as resolveDataBindingsImpl } from "@word/query/data-binding";
import {
  extractFormFields as extractFormFieldsImpl,
  fillFormFields as fillFormFieldsImpl
} from "@word/query/form-fields";
import {
  countByFormat as countByFormatImpl,
  getUsedFormats as getUsedFormatsImpl
} from "@word/query/format-search";
import { mailMerge as mailMergeImpl } from "@word/query/mail-merge";
import { replaceText as replaceTextImpl } from "@word/query/replace";
import {
  acceptAllRevisions as acceptAllRevisionsImpl,
  acceptRevision as acceptRevisionImpl,
  rejectAllRevisions as rejectAllRevisionsImpl,
  rejectRevision as rejectRevisionImpl
} from "@word/query/revisions";
import {
  extractText as extractTextImpl,
  findBookmark as findBookmarkImpl,
  findComment as findCommentImpl,
  getHeadings as getHeadingsImpl,
  countWords as countWordsImpl,
  paragraphCount as paragraphCountImpl,
  tableCount as tableCountImpl
} from "@word/query/search";
import {
  indexStyles as indexStylesImpl,
  resolveNumberingLevel as resolveNumberingLevelImpl,
  resolveRunStyle as resolveRunStyleImpl,
  resolveStyle as resolveStyleImpl,
  resolveTableStyle as resolveTableStyleImpl
} from "@word/query/style-resolve";
import type {
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  ResolvedTableStyle,
  StyleIndex,
  StyleResolveContext
} from "@word/query/style-resolve";
import type { BodyContent, DocxDocument, Paragraph, Run, RunProperties } from "@word/types";

export {
  paragraphCount as getParagraphCount,
  tableCount as getTableCount,
  countWords as getWordCount,
  extractText as getText,
  findBookmark as getBookmark,
  findComment as getComment,
  getHeadings as listHeadings,
  listImages,
  listTables,
  listHyperlinks,
  listSections,
  searchText
} from "@word/query/search";
export { listFields } from "@word/query/fields";
export { listRevisions } from "@word/query/revisions";
export {
  indexStyles as createStyleIndex,
  resolveNumberingLevel as getNumberingLevel
} from "@word/query/style-resolve";
export type { StyleIndex } from "@word/query/style-resolve";
export { getCompatibilityMode } from "@word/query/compat";
export { extractFormFields as listFormFields } from "@word/query/form-fields";
export {
  searchByFormat,
  countByFormat as getCountByFormat,
  getUsedFormats as listUsedFormats
} from "@word/query/format-search";
export { collectParagraphs as listParagraphs, collectRuns as listRuns } from "@word/core/walker";

/** Visit every block, paragraph and run of a document — see {@link DocxVisitor}. */
export function each(doc: DocxDocument, visitor: DocxVisitor, options?: WalkOptions): void {
  checkHandle(doc, "document", "doc", "Query.each");
  checkRecord(visitor, "visitor", "Query.each");
  checkRecord(options, "options", "Query.each", { optional: true });
  walkDocumentImpl(doc, visitor, options);
}

/** Visit `blocks` and everything in them — see {@link DocxVisitor}. */
export function eachBlock(
  blocks: readonly BodyContent[],
  visitor: DocxVisitor,
  basePath?: WalkPath
): "stop" | "continue" {
  checkArray(blocks, "blocks", "Query.eachBlock");
  checkRecord(visitor, "visitor", "Query.eachBlock");
  return walkBlocksImpl(blocks, visitor, basePath);
}

/** Options for {@link getParagraphStyle}. */
export interface ParagraphStyleOptions {
  /** Table position, for overlaying a table style's conditional formatting. */
  readonly context?: StyleResolveContext;
  /** Prebuilt {@link createStyleIndex} index, reused across calls. */
  readonly styles?: StyleIndex;
}

/** Options for {@link getRunStyle}. */
export interface RunStyleOptions {
  /**
   * Run properties inherited from the paragraph's resolved style — pass
   * `getParagraphStyle(doc, para).runProperties` to layer the paragraph style
   * on top of the document defaults.
   */
  readonly paragraphRunProperties?: RunProperties;
  /** Prebuilt {@link createStyleIndex} index, reused across calls. */
  readonly styles?: StyleIndex;
}

/** Options for {@link getTableStyle}. */
export interface TableStyleOptions {
  /** Prebuilt {@link createStyleIndex} index, reused across calls. */
  readonly styles?: StyleIndex;
}

/** Resolve a paragraph's effective style: its style chain merged over the document defaults. */
export function getParagraphStyle(
  doc: DocxDocument,
  para: Paragraph,
  options?: ParagraphStyleOptions
): ResolvedParagraphStyle {
  checkHandle(doc, "document", "doc", "Query.getParagraphStyle");
  checkNode(para, "paragraph", "para", "Query.getParagraphStyle");
  checkRecord(options, "options", "Query.getParagraphStyle", { optional: true });
  return resolveStyleImpl(doc, para, options?.context, options?.styles);
}

/** Resolve a run's effective style: its character style chain merged over inherited properties. */
export function getRunStyle(
  doc: DocxDocument,
  run: Run,
  options?: RunStyleOptions
): ResolvedRunStyle {
  checkHandle(doc, "document", "doc", "Query.getRunStyle");
  checkRecord(run, "run", "Query.getRunStyle");
  checkRecord(options, "options", "Query.getRunStyle", { optional: true });
  return resolveRunStyleImpl(doc, run, options?.paragraphRunProperties, options?.styles);
}

/** Resolve a table style by id, following its `basedOn` chain. */
export function getTableStyle(
  doc: DocxDocument,
  tableStyleId: string,
  options?: TableStyleOptions
): ResolvedTableStyle {
  checkHandle(doc, "document", "doc", "Query.getTableStyle");
  checkString(tableStyleId, "tableStyleId", "Query.getTableStyle");
  checkRecord(options, "options", "Query.getTableStyle", { optional: true });
  return resolveTableStyleImpl(doc, tableStyleId, options?.styles);
}

/** @deprecated Use {@link getParagraphCount}. Removed in 1.0. */
export const paragraphCount = paragraphCountImpl;
/** @deprecated Use {@link getTableCount}. Removed in 1.0. */
export const tableCount = tableCountImpl;
/** @deprecated Use {@link getWordCount}. Removed in 1.0. */
export const countWords = countWordsImpl;
/** @deprecated Use {@link getCountByFormat}. Removed in 1.0. */
export const countByFormat = countByFormatImpl;
/** @deprecated Use {@link createStyleIndex}. Removed in 1.0. */
export const indexStyles = indexStylesImpl;
/** @deprecated Use {@link getText}. Removed in 1.0. */
export const extractText = extractTextImpl;
/** @deprecated Use {@link getBookmark}. Removed in 1.0. */
export const findBookmark = findBookmarkImpl;
/** @deprecated Use {@link getComment}. Removed in 1.0. */
export const findComment = findCommentImpl;
/** @deprecated Use {@link listHeadings}. Removed in 1.0. */
export const getHeadings = getHeadingsImpl;
/** @deprecated Use {@link listFormFields}. Removed in 1.0. */
export const extractFormFields = extractFormFieldsImpl;
/** @deprecated Use {@link listUsedFormats}. Removed in 1.0. */
export const getUsedFormats = getUsedFormatsImpl;
/** @deprecated Use {@link each}. Removed in 1.0. */
export const walkDocument = each;
/** @deprecated Use {@link eachBlock}. Removed in 1.0. */
export const walkBlocks = eachBlock;
/** @deprecated Use {@link listParagraphs}. Removed in 1.0. */
export const collectParagraphs = collectParagraphsImpl;
/** @deprecated Use {@link listRuns}. Removed in 1.0. */
export const collectRuns = collectRunsImpl;
/**
 * @deprecated Use {@link listTables}, passing the parts to include — `listTables`
 * searches only the body unless asked. Removed in 1.0.
 */
export const collectTables = collectTablesImpl;
/** @deprecated Use {@link getParagraphStyle}. Removed in 1.0. */
export const resolveStyle = resolveStyleImpl;
/** @deprecated Use {@link getRunStyle}. Removed in 1.0. */
export const resolveRunStyle = resolveRunStyleImpl;
/** @deprecated Use {@link getTableStyle}. Removed in 1.0. */
export const resolveTableStyle = resolveTableStyleImpl;
/** @deprecated Use {@link getNumberingLevel}. Removed in 1.0. */
export const resolveNumberingLevel = resolveNumberingLevelImpl;
/** @deprecated Use `Document.replaceText`. Removed in 1.0. */
export const replaceText = replaceTextImpl;
/** @deprecated Use `Document.mailMerge`. Removed in 1.0. */
export const mailMerge = mailMergeImpl;
/** @deprecated Use `Document.acceptAllRevisions`. Removed in 1.0. */
export const acceptAllRevisions = acceptAllRevisionsImpl;
/** @deprecated Use `Document.rejectAllRevisions`. Removed in 1.0. */
export const rejectAllRevisions = rejectAllRevisionsImpl;
/** @deprecated Use `Document.acceptRevision`. Removed in 1.0. */
export const acceptRevision = acceptRevisionImpl;
/** @deprecated Use `Document.rejectRevision`. Removed in 1.0. */
export const rejectRevision = rejectRevisionImpl;
/** @deprecated Use `Document.setCompatibilityMode`. Removed in 1.0. */
export const setCompatibilityMode = setCompatibilityModeImpl;
/** @deprecated Use `Transform.fillFormFields`. Removed in 1.0. */
export const fillFormFields = fillFormFieldsImpl;
/** @deprecated Use `Transform.resolveDataBindings`. Removed in 1.0. */
export const resolveDataBindings = resolveDataBindingsImpl;
