/**
 * DOCX Module - Document Builder & Query API
 *
 * This file re-exports the document construction and query APIs from their
 * dedicated sub-modules. Previously a 3947-line monolith, the implementation
 * is now cleanly separated into:
 *
 * - `builder/` — Run, paragraph, table, math, and Document handle builders
 * - `query/`   — Search, replace, merge, revisions, style resolution, etc.
 *
 * All public symbols continue to be exported from this file for backward
 * compatibility — existing `import { ... } from "./document"` works unchanged.
 */

// =============================================================================
// Builder API
// =============================================================================

export {
  // Run builders
  text,
  bold,
  italic,
  underline,
  strikethrough,
  pageBreak,
  lineBreak,
  columnBreak,
  tab,
  positionalTab,
  ruby,
  carriageReturn,
  noBreakHyphen,
  softHyphen,
  // Field builders
  field,
  pageNumberField,
  totalPagesField,
  sectionPagesField,
  sectionField,
  dateField,
  sequenceField,
  timeField,
  authorField,
  titleField,
  subjectField,
  keywordsField,
  fileNameField,
  fileSizeField,
  styleRefField,
  refField,
  pageRefField,
  noteRefField,
  hyperlinkField,
  quoteField,
  tocField,
  tcField,
  indexEntryField,
  indexField,
  ifField,
  includeTextField,
  includePictureField,
  // Form field builders
  formTextField,
  formCheckboxField,
  formDropdownField,
  // Math builders
  mathBlock,
  mathRun,
  mathFraction,
  mathSqrt,
  mathRoot,
  mathSum,
  mathIntegral,
  mathProduct,
  mathSuperScript,
  mathSubScript,
  mathSubSuperScript,
  mathPreSubSuperScript,
  mathPhantom,
  mathGroupChar,
  mathBorderBox,
  mathDelimiter,
  mathNary,
  mathFunction,
  mathLimit,
  mathMatrix,
  mathAccent,
  mathBar,
  mathBox,
  mathEquationArray,
  // Other run-level builders
  symbol,
  floatingImage,
  drawingShape,
  chart,
  structuredDocumentTag,
  checkBox
} from "./builder/run-builders";

export {
  // Paragraph builders
  paragraph,
  textParagraph,
  heading,
  hyperlink,
  bookmarkStart,
  bookmarkEnd,
  commentRangeStart,
  commentRangeEnd,
  commentReference,
  // Track changes
  insertedRun,
  deletedRun,
  movedFromRun,
  movedToRun,
  moveFromRangeStart,
  moveFromRangeEnd,
  moveToRangeStart,
  moveToRangeEnd
} from "./builder/paragraph-builders";

export { border, gridBorders, cell, row, table, simpleTable } from "./builder/table-builders";

export { Document } from "./builder/document-handle";
export type { DocumentHandle } from "./builder/document-handle";

// =============================================================================
// Query / Mutation API
// =============================================================================

export {
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
} from "./query/search";
export type { SearchResult, DocumentHeading, DocumentSection } from "./query/search";

export { replaceText } from "./query/replace";

export { mailMerge } from "./query/mail-merge";

export { mergeDocuments } from "./query/merge";

export type { SplitOptions } from "./query/split";
export { splitDocument } from "./query/split";
export type { MergeOptions } from "./query/merge";

export type { RevisionEntry } from "./query/revisions";
export {
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision
} from "./query/revisions";

export {
  resolveStyle,
  resolveRunStyle,
  resolveNumberingLevel,
  resolveTableStyle
} from "./query/style-resolve";
export type {
  StyleResolveContext,
  ResolvedParagraphStyle,
  ResolvedRunStyle,
  ResolvedNumberingLevel
} from "./query/style-resolve";

export { getCompatibilityMode, setCompatibilityMode } from "./query/compat";
export type { CompatibilityMode } from "./query/compat";

export { resolveDataBindings } from "./query/data-binding";

export { extractFormFields, fillFormFields } from "./query/form-fields";
export type { FormFieldEntry } from "./query/form-fields";

// =============================================================================
// Theme Color Resolution (re-export from color-utils for backward compat)
// =============================================================================

export { resolveThemeColor } from "./core/color-utils";

// =============================================================================
// Patcher / Template Fill API — types re-exported for backward compatibility
// =============================================================================

export type { PatchContent, PatchOperation, PatchOptions } from "./document-io";

// Incremental edit API (efficient open → modify → save)
export type { IncrementalEdit, IncrementalEditOptions } from "./incremental-edit";
export { editDocxIncremental, listDocxParts, readDocxPart } from "./incremental-edit";
