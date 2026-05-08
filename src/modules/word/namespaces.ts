/**
 * DOCX Module — Sub-namespace objects
 *
 * Groups flat builder helpers into logical namespaces for better
 * IDE discoverability. These are re-exports aggregated into objects;
 * since they reference the same underlying functions, tree-shaking
 * still applies at the individual function level for consumers who
 * import the flat named exports instead.
 *
 * @example
 * ```ts
 * import { Run, Paragraph, Table, Math, Field } from "excelts/word";
 *
 * const doc = Document.create();
 * Document.addParagraphElement(doc, Paragraph.create([
 *   Run.bold("Hello"),
 *   Run.text(" world")
 * ]));
 * ```
 */

import {
  text,
  bold,
  italic,
  pageBreak,
  lineBreak,
  columnBreak,
  tab,
  positionalTab,
  ruby,
  carriageReturn,
  noBreakHyphen,
  softHyphen,
  symbol,
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
  formTextField,
  formCheckboxField,
  formDropdownField,
  paragraph,
  textParagraph,
  heading,
  hyperlink,
  bookmarkStart,
  bookmarkEnd,
  commentRangeStart,
  commentRangeEnd,
  commentReference,
  insertedRun,
  deletedRun,
  movedFromRun,
  movedToRun,
  moveFromRangeStart,
  moveFromRangeEnd,
  moveToRangeStart,
  moveToRangeEnd,
  checkBox,
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
  floatingImage,
  drawingShape,
  chart,
  structuredDocumentTag,
  border,
  gridBorders,
  cell,
  row,
  table,
  simpleTable,
  searchText,
  replaceText,
  mailMerge,
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  tableCount,
  extractText
} from "./document";

// =============================================================================
// Run namespace — text run constructors
// =============================================================================

/** Namespace for creating text runs (inline content). */
export const Run = {
  text,
  bold,
  italic,
  pageBreak,
  lineBreak,
  columnBreak,
  tab,
  positionalTab,
  ruby,
  carriageReturn,
  noBreakHyphen,
  softHyphen,
  symbol
};

// =============================================================================
// Field namespace — field code constructors
// =============================================================================

/** Namespace for creating field codes. */
export const Field = {
  create: field,
  pageNumber: pageNumberField,
  totalPages: totalPagesField,
  sectionPages: sectionPagesField,
  section: sectionField,
  date: dateField,
  sequence: sequenceField,
  time: timeField,
  author: authorField,
  title: titleField,
  subject: subjectField,
  keywords: keywordsField,
  fileName: fileNameField,
  fileSize: fileSizeField,
  styleRef: styleRefField,
  ref: refField,
  pageRef: pageRefField,
  noteRef: noteRefField,
  hyperlink: hyperlinkField,
  quote: quoteField,
  toc: tocField,
  tc: tcField,
  indexEntry: indexEntryField,
  index: indexField,
  condition: ifField,
  includeText: includeTextField,
  includePicture: includePictureField,
  formText: formTextField,
  formCheckbox: formCheckboxField,
  formDropdown: formDropdownField
};

// =============================================================================
// Paragraph namespace — paragraph constructors
// =============================================================================

/** Namespace for creating paragraphs. */
export const Paragraph = {
  create: paragraph,
  text: textParagraph,
  heading,
  hyperlink,
  bookmarkStart,
  bookmarkEnd
};

// =============================================================================
// Comment namespace — comment markers
// =============================================================================

/** Namespace for comment-related markers. */
export const Comment = {
  rangeStart: commentRangeStart,
  rangeEnd: commentRangeEnd,
  reference: commentReference
};

// =============================================================================
// TrackChanges namespace — revision markers
// =============================================================================

/** Namespace for track-changes (revision) markers. */
export const TrackChanges = {
  insertedRun,
  deletedRun,
  movedFromRun,
  movedToRun,
  moveFromRangeStart,
  moveFromRangeEnd,
  moveToRangeStart,
  moveToRangeEnd
};

// =============================================================================
// Math namespace — Office Math constructors
// =============================================================================

/** Namespace for OMML (Office Math) content. */
export const Math = {
  block: mathBlock,
  run: mathRun,
  fraction: mathFraction,
  sqrt: mathSqrt,
  root: mathRoot,
  sum: mathSum,
  integral: mathIntegral,
  product: mathProduct,
  superScript: mathSuperScript,
  subScript: mathSubScript,
  subSuperScript: mathSubSuperScript,
  preSubSuperScript: mathPreSubSuperScript,
  phantom: mathPhantom,
  groupChar: mathGroupChar,
  borderBox: mathBorderBox,
  delimiter: mathDelimiter,
  nary: mathNary,
  func: mathFunction,
  limit: mathLimit,
  matrix: mathMatrix,
  accent: mathAccent,
  bar: mathBar,
  box: mathBox,
  equationArray: mathEquationArray
};

// =============================================================================
// Table namespace — table constructors
// =============================================================================

/** Namespace for creating tables. */
export const Table = {
  create: table,
  simple: simpleTable,
  row,
  cell,
  border,
  gridBorders
};

// =============================================================================
// Drawing namespace — images, shapes, charts
// =============================================================================

/** Namespace for drawings (images, shapes, charts). */
export const Drawing = {
  floatingImage,
  shape: drawingShape,
  chart
};

// =============================================================================
// Sdt namespace — structured document tags
// =============================================================================

/** Namespace for structured document tags (content controls). */
export const Sdt = {
  create: structuredDocumentTag,
  checkBox
};

// =============================================================================
// Query namespace — document query/search functions
// =============================================================================

/** Namespace for querying/searching document content. */
export const Query = {
  search: searchText,
  replace: replaceText,
  mailMerge,
  paragraphCount,
  countWords,
  getHeadings,
  findBookmark,
  findComment,
  listImages,
  listTables,
  listHyperlinks,
  tableCount,
  extractText
};
