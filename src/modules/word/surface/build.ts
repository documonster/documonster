/**
 * `Build` namespace surface — content-node builders (runs, paragraphs,
 * tables, shapes, fields, math, …).
 *
 * `import { Build } from "documonster/word"` →
 *   `Build.paragraph(...)`, `Build.text("hi")`, `Build.table(rows)`,
 *   `Build.heading("Title", 1)`, `Build.shape(...)` — tree-shaken via
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
  checkBox as checkbox,
  mathBlock,
  mathRun,
  mathFraction,
  mathSqrt,
  mathRoot,
  mathSum,
  mathIntegral,
  mathProduct,
  mathSuperScript as mathSuperscript,
  mathSubScript as mathSubscript,
  mathSubSuperScript as mathSubSuperscript,
  mathPreSubSuperScript as mathPreSubSuperscript,
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
export type {
  BorderOptions,
  GridBordersOptions,
  SimpleTableOptions,
  TableOptions
} from "@word/builder/table-builders";
export type {
  DateTimeFieldOptions,
  FieldOptions,
  MathBarOptions,
  MathNaryOptions
} from "@word/builder/run-builders";

import {
  createShape as createShapeImpl,
  createRect as createRectImpl,
  createRoundRect as createRoundRectImpl,
  createEllipse as createEllipseImpl,
  createLine as createLineImpl,
  createArrow as createArrowImpl,
  createFlowchartShape as createFlowchartShapeImpl,
  createCallout as createCalloutImpl,
  createStar as createStarImpl
} from "@word/advanced/drawing-shapes";
import {
  checkBox as checkBoxImpl,
  drawingShape as drawingShapeImpl,
  mathSuperScript as mathSuperScriptImpl,
  mathSubScript as mathSubScriptImpl,
  mathSubSuperScript as mathSubSuperScriptImpl,
  mathPreSubSuperScript as mathPreSubSuperScriptImpl
} from "@word/builder/run-builders";

export {
  // drawing-shape builders
  createShape as shape,
  createRect as rect,
  createRoundRect as roundRect,
  createEllipse as ellipse,
  createLine as line,
  createArrow as arrow,
  createFlowchartShape as flowchartShape,
  createCallout as callout,
  createStar as star
} from "@word/advanced/drawing-shapes";

/** @deprecated Use {@link shape}. Removed in 1.0. */
export const createShape = createShapeImpl;
/** @deprecated Use {@link rect}. Removed in 1.0. */
export const createRect = createRectImpl;
/** @deprecated Use {@link roundRect}. Removed in 1.0. */
export const createRoundRect = createRoundRectImpl;
/** @deprecated Use {@link ellipse}. Removed in 1.0. */
export const createEllipse = createEllipseImpl;
/** @deprecated Use {@link line}. Removed in 1.0. */
export const createLine = createLineImpl;
/** @deprecated Use {@link arrow}. Removed in 1.0. */
export const createArrow = createArrowImpl;
/** @deprecated Use {@link flowchartShape}. Removed in 1.0. */
export const createFlowchartShape = createFlowchartShapeImpl;
/** @deprecated Use {@link callout}. Removed in 1.0. */
export const createCallout = createCalloutImpl;
/** @deprecated Use {@link star}. Removed in 1.0. */
export const createStar = createStarImpl;
/**
 * @deprecated Use {@link shape}, which builds the same shape and more:
 * `fillColor` → `fill: { type: "solid", color }`, `noFill` → `fill: { type: "none" }`,
 * `outlineColor` / `outlineWidth` / `noOutline` → `outline: { color, width, noLine }`,
 * `textContent` → `textBody: { paragraphs }`. Removed in 1.0.
 */
export const drawingShape = drawingShapeImpl;
/** @deprecated Use {@link checkbox}. Removed in 1.0. */
export const checkBox = checkBoxImpl;
/** @deprecated Use {@link mathSuperscript}. Removed in 1.0. */
export const mathSuperScript = mathSuperScriptImpl;
/** @deprecated Use {@link mathSubscript}. Removed in 1.0. */
export const mathSubScript = mathSubScriptImpl;
/** @deprecated Use {@link mathSubSuperscript}. Removed in 1.0. */
export const mathSubSuperScript = mathSubSuperScriptImpl;
/** @deprecated Use {@link mathPreSubSuperscript}. Removed in 1.0. */
export const mathPreSubSuperScript = mathPreSubSuperScriptImpl;
