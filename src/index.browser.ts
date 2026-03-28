/**
 * Browser entry point - No Node.js dependencies
 * This version is optimized for browser environments with minimal bundle size
 */

// =============================================================================
// Main Classes (Browser-compatible)
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
// Enums
// =============================================================================
export * from "@excel/enums";

// =============================================================================
// Types
// =============================================================================

// Export all type definitions from types.ts
export * from "@excel/types";

// Export pivot table types (type-only, no runtime dependency)
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

// =============================================================================
// Streaming Writer (Browser-compatible)
// Uses cross-platform base implementation without Node.js fs
// =============================================================================

export { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
export { WorkbookReader } from "@excel/stream/workbook-reader.browser";
export { WorksheetWriter } from "@excel/stream/worksheet-writer";
export { WorksheetReader } from "@excel/stream/worksheet-reader";

// =============================================================================
// NOTE: Node.js-only features not available in browser:
// - Reading from a file path is not supported (use Uint8Array/ArrayBuffer/Blob instead)
// - Writing to a file path is not supported (use writeBuffer() / stream output, then save as Blob/download)
// =============================================================================

// =============================================================================
// CSV types and stream classes
// =============================================================================
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
// PDF Export (Browser-compatible, zero external dependencies)
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
