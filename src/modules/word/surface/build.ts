/**
 * `Build` namespace surface — content-node builders (runs, paragraphs,
 * tables, shapes, fields, math, …).
 *
 * `import { Build } from "documonster/word"` →
 *   `Build.paragraph(...)`, `Build.text("hi")`, `Build.table(rows)`,
 *   `Build.heading("Title", 1)`, `Build.createShape(...)` — tree-shaken via
 *   `export * as Build`.
 *
 * Merged here (rather than separate `Run`/`Paragraph`/`Table` namespaces)
 * because those names are public data-model **types** (`Run`, `Paragraph`,
 * `Table`); the builders live under one factory namespace to avoid the
 * type/namespace name collision while staying discoverable.
 */
export {
  // run / inline builders
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
  symbol,
  floatingImage,
  drawingShape,
  chart,
  structuredDocumentTag
} from "@word/builder/run-builders";

export {
  // paragraph / block builders
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
  moveToRangeEnd
} from "@word/builder/paragraph-builders";

export {
  // table builders
  border,
  gridBorders,
  cell,
  row,
  table,
  simpleTable
} from "@word/builder/table-builders";

export {
  // drawing-shape builders
  createShape,
  createRect,
  createRoundRect,
  createEllipse,
  createLine,
  createArrow,
  createFlowchartShape,
  createCallout,
  createStar
} from "@word/advanced/drawing-shapes";
