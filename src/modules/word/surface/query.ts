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
} from "../query/search";
export { replaceText } from "../query/replace";
export { mailMerge } from "../query/mail-merge";
export {
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision
} from "../query/revisions";
export {
  resolveStyle,
  resolveRunStyle,
  resolveNumberingLevel,
  resolveTableStyle
} from "../query/style-resolve";
export { getCompatibilityMode, setCompatibilityMode } from "../query/compat";
export { resolveDataBindings } from "../query/data-binding";
export { extractFormFields, fillFormFields } from "../query/form-fields";
export { searchByFormat, countByFormat, getUsedFormats } from "../query/format-search";
export {
  walkDocument,
  walkBlocks,
  collectParagraphs,
  collectRuns,
  collectTables
} from "../core/walker";
