// =============================================================================
// Main Classes
// =============================================================================

export * from "@excel/workbook";
export * from "@excel/worksheet";
export * from "@excel/row";
export * from "@excel/column";
export * from "@excel/cell";
export * from "@excel/range";
export { imageClone, imageCreate, imageModel } from "@excel/image";
export type { ImageData, ImageModel } from "@excel/image";
export * from "@excel/anchor";
export {
  createTable,
  tableAddColumn,
  tableAddRow,
  tableCommit,
  tableGetColumn,
  tableModel,
  tableName,
  tableRemoveColumns,
  tableRemoveRows,
  tableSetModel,
  tableSetName,
  type TableColumnView,
  type TableData
} from "@excel/table";
export { noteCreate, noteFromModel, noteModel, isNoteData } from "@excel/note";
export type { NoteData } from "@excel/note";
export {
  createDataValidations,
  dataValidationAdd,
  dataValidationFind,
  dataValidationRemove
} from "@excel/data-validations";
export type { DataValidationsData } from "@excel/data-validations";
export {
  formCheckboxChecked,
  formCheckboxCreate,
  formCheckboxFromModel,
  formCheckboxLink,
  formCheckboxSetChecked,
  formCheckboxSetLink,
  formCheckboxSetText,
  formCheckboxText,
  formCheckboxVmlAnchor,
  formCheckboxVmlCheckedValue,
  formCheckboxVmlStyle,
  isFormCheckbox
} from "@excel/form-control";
export type { FormCheckboxData } from "@excel/form-control";
// Note: the formula engine lives at the `./formula` subpath so it stays
// out of bundles that only need to read / write XLSX files. Import
// `{ calculateFormulas }` from `@cj-tech-master/excelts/formula` and call
// `calculateFormulas(workbook)` to recompute; pass it as `excelToPdf`'s
// `recalculate` option for automatic recalculation during PDF export.

// =============================================================================
// Node.js Only: Streaming Classes
// These can also be accessed via Workbook.createStreamWriter/createStreamReader
// =============================================================================

export { WorkbookWriter } from "@excel/stream/workbook-writer";
export { WorkbookReader } from "@excel/stream/workbook-reader";
export { WorksheetWriter } from "@excel/stream/worksheet-writer";
export { WorksheetReader } from "@excel/stream/worksheet-reader";

// =============================================================================
// Enums
// =============================================================================

export * from "@excel/enums";

// =============================================================================
// Types
// =============================================================================

// Export all type definitions from types.ts
export * from "@excel/types";

// Watermark image generator utility
export { createTextWatermarkImage } from "@excel/utils/watermark-image";
export type { TextWatermarkImageOptions } from "@excel/utils/watermark-image";

// Pivot table types
export type {
  PivotTable,
  PivotTableModel,
  PivotTableValue,
  PivotTableSource,
  CacheField,
  SharedItemValue,
  DataField,
  PivotTableChartFormat,
  PivotTableSubtotal,
  RecordValue,
  ParsedCacheDefinition,
  ParsedCacheRecords
} from "@excel/pivot-table";

// Form control types
export type {
  FormCheckboxModel,
  FormCheckboxOptions,
  FormControlRange,
  FormControlAnchor
} from "@excel/form-control";

// Node.js Only: Streaming reader types
export type {
  WorkbookReaderOptions,
  ParseEvent,
  SharedStringEvent,
  WorksheetReadyEvent,
  HyperlinksEvent
} from "@excel/stream/workbook-reader";

export type {
  WorksheetReaderOptions,
  WorksheetEvent,
  RowEvent,
  HyperlinkEvent,
  WorksheetHyperlink
} from "@excel/stream/worksheet-reader";

// Node.js Only: Streaming writer types
export type {
  WorkbookWriterOptions,
  WorkbookZipOptions,
  ZlibOptions
} from "@excel/stream/workbook-writer";

// CSV types and stream classes
export type { CsvOptions, CsvInput } from "@excel/workbook";
export {
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "@csv/stream";

// Markdown types
export type { MarkdownOptions } from "@markdown/types";
export type {
  MarkdownAlignment,
  MarkdownParseResult,
  MarkdownParseOptions,
  MarkdownColumnConfig,
  MarkdownFormatOptions
} from "@markdown/types";

// =============================================================================
// Additional Classes & Types
// =============================================================================

export {
  createDefinedNames,
  definedNamesAdd,
  definedNamesAddFormula,
  definedNamesGetAllEntries,
  definedNamesGetNames,
  definedNamesModel,
  definedNamesRemove,
  definedNamesSetModel,
  type DefinedNameModel
} from "@excel/defined-names";
export type { DefinedNamesData } from "@excel/defined-names";
export {
  chartsheetChart,
  chartsheetChartExModel,
  chartsheetChartModel,
  chartsheetModel,
  chartsheetName,
  chartsheetSetName,
  createChartsheet,
  type ChartsheetData
} from "@excel/chartsheet";
export type {
  AddChartsheetOptions,
  AddPivotChartsheetOptions,
  ChartsheetOptions,
  ChartsheetViewOptions
} from "@excel/chartsheet";
export type { CheckboxState } from "@excel/form-control";
export type { ColumnDefn, ColumnHeaderValue } from "@excel/column";
export type { RangeInput } from "@excel/range";
export type {
  WorkbookModel,
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel,
  ExternalLinkCachedSheet
} from "@excel/workbook";
export type { NodeInput } from "@excel/stream/workbook-reader";

// =============================================================================
// Utilities
// =============================================================================

// Cell address encoding/decoding (0-indexed)
export {
  decodeCol,
  encodeCol,
  decodeRow,
  encodeRow,
  decodeCell,
  encodeCell,
  decodeRange,
  encodeRange
} from "@excel/utils/address";
export type { CellAddress, SheetRange, Origin } from "@excel/utils/address";

// Worksheet data conversion option types
export type { SheetToJSONOptions, AddJSONOptions, AddAOAOptions } from "@excel/worksheet";

// Chart programmatic surface (builders, renderers, presets, parsers,
// install function) lives at the `./chart` subpath so it stays out of
// bundles that only need to read / write XLSX files. Import
// `@cj-tech-master/excelts/chart` to use chart APIs — and call
// `installChartSupport()` once at startup to enable
// `worksheet.addChart()`, chart-cache population during write, and
// chart reconstruction during XLSX load.

// Cell display-text helpers (apply numFmt to produce an Excel-style string)
export { getCellDisplayText, formatCellValue, isDateDisplayFormat } from "@excel/utils/cell-format";

// Date conversion (Excel serial dates <-> JS Date)
export { dateToExcel, excelToDate } from "@utils/utils.base";

// Base64 utilities (cross-platform)
export { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils.base";

// XML utilities
export { xmlEncode, xmlDecode, xmlEncodeAttr, validateXmlName } from "@xml/encode";

// Date parsing/formatting (high-performance, zero-dep)
export { DateParser, DateFormatter, getSupportedFormats, type DateFormat } from "@utils/datetime";

// Error infrastructure
export {
  BaseError,
  type BaseErrorOptions,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@utils/errors";

// Binary utilities (cross-platform)
export {
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString
} from "@utils/binary";

// =============================================================================
// PDF Export
// =============================================================================

export {
  pdf,
  excelToPdf,
  PageSizes,
  PdfError,
  PdfRenderError,
  PdfFontError,
  PdfStructureError,
  isPdfError
} from "@pdf/index";
export type {
  PdfExportOptions,
  PdfPageSize,
  PdfOrientation,
  PdfMargins,
  PageSizeName,
  PdfColor
} from "@pdf/index";

// =============================================================================
// Errors
// =============================================================================

export {
  ExcelError,
  isExcelError,
  ExcelFileError,
  ExcelDownloadError,
  ExcelNotSupportedError,
  ExcelStreamStateError,
  InvalidAddressError,
  ColumnOutOfBoundsError,
  RowOutOfBoundsError,
  MergeConflictError,
  InvalidValueTypeError,
  XmlParseError,
  WorksheetNameError,
  PivotTableError,
  TableError,
  ImageError,
  MaxItemsExceededError
} from "@excel/errors";

// Markdown errors
export { MarkdownError, MarkdownParseError } from "@markdown/errors";
