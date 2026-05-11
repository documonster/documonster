/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/query"` */

export type { SearchResult, DocumentHeading, DocumentSection } from "./search";
export {
  paragraphText,
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  tableCount,
  listSections,
  extractText,
  searchText
} from "./search";

export { replaceText } from "./replace";

export { mailMerge } from "./mail-merge";

export type { MergeOptions } from "./merge";
export { mergeDocuments } from "./merge";

export type { SplitOptions } from "./split";
export { splitDocument } from "./split";

export type { RevisionEntry } from "./revisions";
export {
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision
} from "./revisions";

export type {
  StyleResolveContext,
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  ResolvedNumberingLevel
} from "./style-resolve";
export {
  resolveStyle,
  resolveRunStyle,
  resolveNumberingLevel,
  resolveTableStyle
} from "./style-resolve";

export type { CompatibilityMode } from "./compat";
export { getCompatibilityMode, setCompatibilityMode } from "./compat";

export { resolveDataBindings } from "./data-binding";

export type { FormFieldEntry } from "./form-fields";
export { extractFormFields, fillFormFields } from "./form-fields";

export type { FormatCriteria, FormatSearchResult } from "./format-search";
export { searchByFormat, countByFormat, getUsedFormats } from "./format-search";
