/**
 * CSV Module - Public API
 *
 * Pure CSV parsing/formatting functionality with no Excel dependencies.
 * For CSV-Worksheet integration, use Workbook.readCsv/writeCsv methods instead.
 *
 * Design principles:
 * - Only export types and functions that are part of the PUBLIC API
 * - Internal utilities (like format helpers) are used internally but not exported
 * - This reduces bundle size and simplifies the public interface
 */

// =============================================================================
// Core Types (from types.ts)
// =============================================================================

export type {
  // Row types
  HeaderArray,
  RowHashArray,
  RowArray,
  RowMap,
  Row,

  // Transform types
  HeaderTransformFunction,
  RowTransformCallback,
  RowTransformFunction,
  RowValidateCallback,
  RowValidateFunction,
  TransformContext,
  FormattedValue,
  TransformResult,
  TypeTransformMap,

  // Dynamic typing
  DynamicTypingConfig,
  CastDateConfig,

  // Column config
  ColumnConfig,

  // Error types (unified)
  CsvErrorCode,
  CsvRecordError,
  OnSkipCallback,

  // Options (general)
  CsvParseOptions,
  CsvFormatOptions,

  // Options (type-safe variants for better inference)
  CsvParseArrayOptions,
  CsvParseObjectOptions,

  // Parse results
  ChunkMeta,
  CsvParseMeta,
  RecordInfo,
  RecordWithInfo,
  CsvParseResult,
  CsvParseResultWithObjname,

  // Column mismatch config
  ColumnMismatchConfig,
  ColumnMismatchLess,
  ColumnMismatchMore
} from "./types";

// Type guards and helpers from types.ts
export { isFormattedValue, quoted, unquoted } from "./types";

// =============================================================================
// Core Functions
// =============================================================================

// Synchronous parser (direct import to avoid barrel pulling in async -> stream)
export { parseCsv } from "./parse/sync";

// Async parser
export { parseCsvAsync, parseCsvRows, parseCsvWithProgress } from "./parse/async";

// Formatter
export { formatCsv } from "./format/index";

// =============================================================================
// Stream Classes
// =============================================================================

export {
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "./stream/index";

// =============================================================================
// Detection Utilities (commonly needed by users)
// =============================================================================

export { detectDelimiter, detectLinebreak, stripBom } from "./utils/detect";

// =============================================================================
// Row Utilities (commonly needed by users)
// =============================================================================

export {
  isRowHashArray,
  rowHashArrayToValues,
  rowHashArrayToHeaders,
  rowHashArrayMapByHeaders,
  processColumns,
  deduplicateHeaders,
  deduplicateHeadersWithRenames
} from "./utils/row";

// =============================================================================
// Extras (opt-in utilities - inline for tree-shaking)
// =============================================================================

// Dynamic Typing utilities
export { applyDynamicTyping, applyDynamicTypingToRow } from "./utils/dynamic-typing";

// CSV Generator utilities
export {
  csvGenerate,
  csvGenerateRows,
  csvGenerateAsync,
  csvGenerateData,
  createCsvGenerator,
  type CsvGenerateOptions,
  type CsvGenerateResult,
  type ColumnDef,
  type GeneratorColumnConfig,
  type BuiltinColumnType,
  type GeneratorFn,
  type GeneratorContext,
  type StopCondition,
  type StopContext
} from "./utils/generate";

// Number formatting utilities
export { formatNumberForCsv, parseNumberFromCsv, type DecimalSeparator } from "./utils/number";

// =============================================================================
// Errors
// =============================================================================

export { CsvError, CsvWorkerError } from "./errors";
