/**
 * documonster/excel — base exports (platform independent).
 *
 * Shared domain dot-namespaces + error classes for both the Node and browser
 * entries. The two platform-specific namespaces (`Workbook`, `Stream`) are
 * NOT here — each entry (`index.ts` / `index.browser.ts`) re-exports this base
 * and then adds its own platform variant of those two. Mirrors the word
 * module's `index.base.ts` structure.
 *
 *   import { Workbook, Worksheet, Cell, Chart } from "documonster/excel";
 *   const wb = Workbook.create();
 *   const ws = Workbook.addWorksheet(wb, "Sheet1");
 *   Cell.setValue(ws, "A1", 42);
 *   const buf = await Workbook.toBuffer(wb);
 *
 * Each namespace is an ESM namespace re-export over a physical `surface/*.ts`
 * module of flat functions, which tree-shakes per-member on rolldown / rspack.
 *
 * The public **types** are exported flat, under their declared names — see the
 * "Public types" section below and the "Types" section of the module README.
 */

import type { AddAoaOptions, AddJsonOptions, SheetToJsonOptions } from "@excel/core/worksheet";

// --- Domain namespaces (platform-independent) ---
export * as Worksheet from "@excel/surface/worksheet";
export * as Cell from "@excel/surface/cell";
export * as Row from "@excel/surface/row";
export * as Column from "@excel/surface/column";
export * as Range from "@excel/surface/range";
export * as Chart from "@excel/surface/chart";
export * as Table from "@excel/surface/table";
export * as Image from "@excel/surface/image";
export * as Pivot from "@excel/surface/pivot";
export * as Sparkline from "@excel/surface/sparkline";
export * as Form from "@excel/surface/form";
export * as Chartsheet from "@excel/surface/chartsheet";
export * as DataValidation from "@excel/surface/data-validation";
export * as DefinedNames from "@excel/surface/defined-names";
export * as Note from "@excel/surface/note";
export * as Address from "@excel/surface/address";
export * as Anchor from "@excel/surface/anchor";
export * as Watermark from "@excel/surface/watermark";
export * as HeaderFooterImage from "@excel/surface/header-footer-image";
export * as Format from "@excel/surface/format";
export * as ConditionalFormatting from "@excel/surface/conditional-formatting";

// ---------------------------------------------------------------------------
// Public types (platform-independent)
//
// Flat, declared-name re-exports — the same convention every other module uses
// (`word`, `csv`, `xml`, `archive`, `pdf`, `stream`): a type has exactly one
// public name, and it is the name TypeScript prints in errors and hovers.
//
// The one deliberate exception is handles. Where a namespace already publishes
// a `Handle` alias (`Worksheet.Handle`, `Workbook.Handle`, `Table.Handle`, …)
// that alias IS the public name, so the underlying `WorksheetData` /
// `WorkbookData` / … interfaces are not re-exported here — a second name for
// the same type would be worse than none. Cell / row / column handles have no
// such alias, so their declared names are public (see below).
//
// The workbook IO and streaming option types are listed here too: they are the
// same type on every platform, so only the `Workbook` and `Stream` namespaces
// differ between `index.ts` and `index.browser.ts`.
// ---------------------------------------------------------------------------

// Style vocabulary — `Cell.setStyle`, `Row.setStyle`, `Column.setStyle`,
// `Table` column styles and `Workbook.addCellStyle` all speak these.
export type {
  Style,
  Alignment,
  Border,
  BorderDiagonal,
  BorderStyle,
  Borders,
  Color,
  Fill,
  FillGradientAngle,
  FillGradientPath,
  FillPattern,
  FillPatterns,
  Font,
  GradientStop,
  NamedStyle,
  NumFmt,
  Protection,
  RichText
} from "@excel/types";
export type { NamedStyleEntry } from "@excel/core/workbook-core";
/** Recursively `readonly` view of a type — see {@link CellView}. */
export type { DeepReadonly } from "@utils/types";
export type { BuiltInCellStyle, BuiltinCellStyle } from "@excel/core/builtin-cell-styles";

// Cell values — what `Cell.getValue` returns and `Cell.setValue` accepts.
export type {
  CellValue,
  CellValueInput,
  CellArrayFormulaValue,
  CellCheckboxValue,
  CellErrorValue,
  CellFormulaHyperlinkValue,
  CellFormulaValue,
  CellHyperlinkValue,
  CellHyperlinkValueInput,
  CellRichTextValue,
  CellSharedFormulaValue
} from "@excel/types";
/**
 * Dates as calendar values rather than as instants.
 *
 * `ExcelDateTimeParts` is what `Cell.getDateParts` returns and `Cell.setDateParts` takes — timezone-free
 * fields, available on every runtime. The three `Plain*` aliases are `Temporal`'s types, derived structurally
 * from `globalThis` so that a consumer whose `lib` predates `esnext.temporal` gets `never` and the member
 * collapses, rather than an error in a declaration file they did not write.
 */
export type { ExcelDateTimeParts } from "@utils/excel-serial";
export type {
  PlainDate,
  PlainDateTime,
  PlainTime,
  TemporalKind,
  TemporalPlainValue
} from "@excel/core/temporal";
/**
 * What `Cell.getDateKind` reports. Wider than {@link TemporalKind}, because a format can also say
 * "a length of time" (`duration`) or say nothing at all (`unknown`) — and calling either of those
 * a date-time is how a `[h]:mm:ss` cell came back as a moment in 1899.
 */
export type { DateFormatKind } from "@excel/utils/cell-format";
export type { CellModel, FormulaResult } from "@excel/core/cell";
/** @deprecated Use {@link CellValue}. Removed in 1.0. */
export type { CellValueType } from "@excel/core/cell";
/** @deprecated Use {@link CellValueInput}. Removed in 1.0. */
export type { CellValueInputType } from "@excel/core/cell";
// The cell / row handles belong to the streaming API, where they are published as
// `Stream.CellHandle` and `Stream.RowHandle`. These flat names stay nameable while
// deprecated members still mention them.
/** @deprecated Use `Stream.CellHandle`. Removed in 1.0. */
export type { CellData } from "@excel/core/cell";
/** A read-only, allocation-free projection of one cell — what `Cell.getView` returns. */
export type { CellView } from "@excel/core/cell";
/**
 * The cell-value kinds (`Cell.getType` / `Cell.getEffectiveType`), the formula
 * kinds (`CellModel.formulaType`) and the Excel error strings. Each is a plain
 * constant lookup object that doubles as its own type, so `ValueType.Number`
 * works as a value and `ValueType` as a type — and each tree-shakes away when
 * unused (unlike a TS `enum`, which cannot be eliminated).
 */
export { ValueType, FormulaType, ErrorValue } from "@excel/core/enums";
/** Convenience OOXML paper-size codes for `PageSetup.paperSize`. */
export { PaperSize } from "@excel/types";

// Rows and columns.
export type { RowObject, RowValues } from "@excel/types";
export type { RowDimensions, RowModel } from "@excel/core/row";
/** @deprecated Use `Stream.RowHandle`. Removed in 1.0. */
export type { RowData } from "@excel/core/row";
export type { ColumnDefn, ColumnHeaderValue, ColumnModel } from "@excel/core/column";
/**
 * @deprecated Address columns by key, letter or number through the `Column` namespace; a streamed sheet's column is a
 * `Stream.ColumnHandle`. Removed in 1.0.
 */
export type { ColumnData } from "@excel/core/column";
/**
 * @deprecated Only the deprecated `Worksheet.columns` / `lastColumn` return it; every field
 * is read through the `Column` namespace (`Column.getDefinition`, `Column.getNumber`, …).
 * Removed in 1.0.
 */
export type { ColumnView } from "@excel/core/column";

// Addresses and ranges.
export type { DecodedAddress } from "@excel/types";
export type { CellAddress, Origin, SheetRange } from "@excel/utils/address";
export type { DecodedRange } from "@excel/utils/col-cache";
export type { RangeInput } from "@excel/core/range";

// Worksheet-level structure.
export type {
  AutoFilter,
  ColBreak,
  HeaderFooter,
  IgnoredError,
  Location,
  Margins,
  PageSetup,
  RowBreak,
  WorksheetProperties,
  WorksheetProtection,
  WorksheetState,
  WorksheetView,
  WorksheetViewCommon,
  WorksheetViewFrozen,
  WorksheetViewNormal,
  WorksheetViewSplit
} from "@excel/types";
export type {
  AutoFilterCriteria,
  FreezePanesOptions,
  FrozenPanes,
  SheetProtection,
  SplitPanes,
  SplitPanesOptions,
  WorksheetPanes
} from "@excel/core/worksheet-core";
export type {
  AddAoaOptions,
  AddJsonOptions,
  AutoFitColumnOptions,
  FillFormulaOptions,
  ProtectSheetOptions,
  SheetToJsonOptions,
  WorksheetModel
} from "@excel/core/worksheet";
/** @deprecated Use {@link AddAoaOptions}. Removed in 1.0. */
export type AddAOAOptions = AddAoaOptions;
/** @deprecated Use {@link AddJsonOptions}. Removed in 1.0. */
export type AddJSONOptions = AddJsonOptions;
/** @deprecated Use {@link SheetToJsonOptions}. Removed in 1.0. */
export type SheetToJSONOptions = SheetToJsonOptions;

// Conditional formatting — the `ConditionalFormatting` namespace.
export type {
  AboveAverageRuleType,
  CellIsOperators,
  CellIsRuleType,
  Cfvo,
  CfvoTypes,
  ColorScaleRuleType,
  ConditionalFormattingBaseRule,
  ConditionalFormattingOptions,
  ConditionalFormattingRule,
  ContainsTextOperators,
  ContainsTextRuleType,
  DataBarRuleType,
  ExpressionRuleType,
  IconSetRuleType,
  IconSetTypes,
  TimePeriodRuleType,
  TimePeriodTypes,
  Top10RuleType
} from "@excel/types";

// Data validation.
export type {
  DataValidationAny,
  DataValidationOperator,
  DataValidationRule,
  DataValidationWithFormulae
} from "@excel/types";
export type { DataValidationModel } from "@excel/core/data-validations";

// Notes and comments.
export type {
  Comment,
  CommentEditAs,
  CommentMargins,
  CommentProtection,
  ThreadedComment,
  ThreadedCommentMention,
  ThreadedCommentPerson
} from "@excel/types";
export type { NoteConfig, NoteModel, NoteText } from "@excel/core/cell";

// Tables.
export type { TableColumnProperties, TableProperties, TableStyleProperties } from "@excel/types";
export type { TableColumnView, TableModel, TableRowEditOptions } from "@excel/core/table";

// Images, shapes, drawing anchors, watermarks.
export type {
  AddImageRange,
  AddShapeOptions,
  HeaderFooterImagePosition,
  ImageAnchor,
  ImageData,
  ShapeModel,
  ShapeType,
  WatermarkMode,
  WatermarkOptions
} from "@excel/types";
export type {
  BackgroundImageModel,
  HeaderImageModel,
  ImageAbsolutePosition,
  ImageExtent,
  ImageHyperlinks,
  ImageModel,
  ImageModelInput,
  ImageRange,
  ImageRangeInput,
  ImageRangeModel,
  PlacedImageModel,
  WatermarkImageModel
} from "@excel/core/image";
export type { AnchorInput, AnchorModel } from "@excel/core/anchor";
export type { HeaderFooterImageEntry, HeaderFooterImageOptions } from "@excel/core/worksheet";

// Defined names, form controls, sparklines, chartsheets.
export type { DefinedNameModel, SyntaxProbe } from "@excel/core/defined-names";
export type {
  CheckboxState,
  FormCheckboxModel,
  FormCheckboxOptions,
  FormControlAnchor,
  FormControlRange
} from "@excel/core/form-control";
export type {
  AddSparklineGroupOptions,
  SparklineAxisType,
  SparklineColor,
  SparklineGroup,
  SparklineItem,
  SparklineType
} from "@excel/core/sparkline";
export type {
  AddChartsheetOptions,
  AddPivotChartsheetOptions,
  ChartsheetModel,
  ChartsheetOptions,
  ChartsheetViewOptions
} from "@excel/core/chartsheet";
export type { ChartHandle } from "@excel/core/worksheet-core";

// Workbook-level (platform-independent parts).
export type {
  AddWorksheetOptions,
  CalculationProperties,
  WorkbookProperties,
  WorkbookProtection,
  WorkbookView
} from "@excel/types";
export type { XlsxReadOptions, XlsxWriteOptions } from "@excel/core/xlsx-io";
export type { XlsxReadable, XlsxWritable } from "@excel/core/xlsx-io-types";
export type { WorkbookWritable } from "@excel/core/xlsx-sink";
// The browser build's deprecated `Workbook.readStream(wb, stream)` takes this event-emitter shape.
export type { XlsxParseStream } from "@excel/xlsx/types";
export type { XlsxStreamOptions } from "@excel/core/xlsx-stream";
// Format-aware IO. `WorkbookFormat` is the choice; the option bags are what carry it.
export type { WorkbookFormat } from "@excel/core/workbook-format";
export type {
  // The `encryption` write option: the same type Word's `Io.toBuffer` and `WordCrypto.encrypt` take (`DocxEncryptOptions`).
  // `Workbook`-prefixed like the option bag carrying it, because it applies to XLSX and XLSB alike.
  WorkbookEncryptOptions,
  UnrecognizedValue,
  WorkbookCreateOptions,
  WorkbookStreamInput
} from "@excel/core/workbook-io-types";
export type { DiscardUnrecognizedValuesOptions } from "@excel/core/workbook.browser";
export type {
  WorkbookDiagnosticReadOptions,
  WorkbookReadOptions,
  WorkbookReadReport,
  WorkbookWriteOptions
} from "@excel/core/xlsx-io";
export type {
  AddExternalLinkOptions,
  ExternalLinkCachedSheet,
  ExternalLinkModel,
  WorkbookMedia,
  ProtectWorkbookOptions,
  RegisterFunctionOptions,
  RegisterPersonOptions,
  WorkbookModel,
  WorkbookProtectionModel
} from "@excel/core/workbook.browser";
// Reachable from `WorkbookModel.opaqueParts` / `.opaqueDrops`, so a caller that
// inspects either needs to be able to name what it got back.
export type {
  OpaqueDrop,
  OpaqueDropReason,
  OpaquePart,
  OpaqueRelationship,
  OpaqueSourceRelationship
} from "@excel/core/opaque-part";
// Reachable from `WorksheetModel.styledBlankRanges`, for the same reason: the field is public on `getModel`/`setModel`, so
// a caller that reads or builds one needs to be able to name it. It was `readonly unknown[]`, which is why both writers
// carried their own cast and their shapes drifted apart.
export type { StyledBlankRangeModel } from "@excel/core/styled-blanks";

// --- Errors (extend BaseError; consistent with every other module's entry) ---
export {
  ExcelError,
  isExcelError,
  ExcelWorksheetNameError,
  ExcelInvalidAddressError,
  ExcelColumnOutOfBoundsError,
  ExcelRowOutOfBoundsError,
  ExcelMergeConflictError,
  ExcelInvalidValueTypeError,
  ExcelArgumentError,
  ExcelUnsupportedFeatureError,
  ExcelDecryptionError,
  ExcelEncryptedError,
  ExcelFileError,
  ExcelStreamStateError,
  ExcelDownloadError,
  ExcelPivotTableError,
  ExcelChartOptionsError,
  ExcelTableError,
  ExcelDefinedNameError,
  ExcelImageError,
  ExcelLimitExceededError,
  // Deprecated names (pre-prefix, or renamed to match the other modules); removed in 1.0.
  WorksheetNameError,
  InvalidAddressError,
  ColumnOutOfBoundsError,
  RowOutOfBoundsError,
  MergeConflictError,
  InvalidValueTypeError,
  PivotTableError,
  ChartOptionsError,
  TableError,
  DefinedNameError,
  ImageError,
  MaxItemsExceededError,
  ExcelNotSupportedError,
  // Thrown by the XLSB reader for a malformed record stream, and reachable now that
  // `Workbook.read` dispatches to it.
  XlsbParseError,
  // **Promised by `WorkbookReadOptions.formulas: "error"` and previously unreachable.**
  //
  // The option's documentation tells a caller to expect this type, and its own tests use `instanceof` — but the class was
  // only exported from the internal `@excel/errors` path, so a consumer of `documonster/excel` could not name it and had
  // to match on the message. An error a public option promises has to be part of the public surface.
  XlsbFormulaDecodeError,
  // The XLSX reader's counterpart of `XlsbParseError`, thrown for a malformed part — reachable through every
  // `Workbook.read*`, so a caller has to be able to name it.
  XlsxParseError
} from "@excel/errors";
// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
export type {
  ExcelFileErrorOptions,
  InvalidValueTypeErrorOptions,
  PositionErrorOptions
} from "@excel/errors";
