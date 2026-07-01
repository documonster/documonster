/**
 * `Query` namespace surface — read-only document queries plus in-place
 * mutations (search, replace, mail-merge, revisions, style resolution,
 * compatibility, data binding, form fields, format search, tree walking).
 *
 * `import { Query } from "documonster/word"` →
 *   `Query.extractText(doc)`, `Query.searchText(doc, /x/)`,
 *   `Query.replaceText(...)`, `Query.collectParagraphs(doc)`, … — tree-shaken
 *   via `export * as Query`.
 */
export {
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  listSections,
  tableCount,
  extractText,
  searchText
} from "@word/query/search";
export { replaceText } from "@word/query/replace";
export { mailMerge } from "@word/query/mail-merge";
export {
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision
} from "@word/query/revisions";
export {
  resolveStyle,
  resolveRunStyle,
  resolveNumberingLevel,
  resolveTableStyle
} from "@word/query/style-resolve";
export { getCompatibilityMode, setCompatibilityMode } from "@word/query/compat";
export { resolveDataBindings } from "@word/query/data-binding";
export { extractFormFields, fillFormFields } from "@word/query/form-fields";
export { searchByFormat, countByFormat, getUsedFormats } from "@word/query/format-search";
export {
  walkDocument,
  walkBlocks,
  collectParagraphs,
  collectRuns,
  collectTables
} from "@word/core/walker";
