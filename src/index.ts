// =============================================================================
// Main Classes
// =============================================================================

export { Workbook } from "@excel/workbook";
export { Worksheet } from "@excel/worksheet";
export { Row } from "@excel/row";
export { Column } from "@excel/column";
export { Cell } from "@excel/cell";
export { Range } from "@excel/range";
export { Image } from "@excel/image";
export * from "@excel/anchor";
export { Table } from "@excel/table";
export { DataValidations } from "@excel/data-validations";
export { FormCheckbox } from "@excel/form-control";

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

// Pivot table types
export type {
  PivotTable,
  PivotTableModel,
  PivotTableValue,
  PivotTableSource,
  CacheField,
  SharedItemValue,
  DataField,
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
  ZipOptions,
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

// =============================================================================
// Additional Classes & Types
// =============================================================================

export { DefinedNames, type DefinedNameModel } from "@excel/defined-names";
export type { CheckboxState } from "@excel/form-control";
export type { ColumnDefn, ColumnHeaderValue } from "@excel/column";
export type { RangeInput } from "@excel/range";
export type { WorkbookModel, WorkbookMedia } from "@excel/workbook";
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

// Date conversion (Excel serial dates <-> JS Date)
export { dateToExcel, excelToDate } from "@utils/utils.base";

// Base64 utilities (cross-platform)
export { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils.base";

// XML utilities
export { xmlEncode, xmlDecode } from "@utils/utils.base";

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
