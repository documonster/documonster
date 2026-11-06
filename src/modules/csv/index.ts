/**
 * CSV Module - Public API
 *
 * Pure CSV parsing/formatting functionality with no Excel dependencies.
 * For CSV-Worksheet integration, use `ExcelCsv` from `documonster/excel/csv`
 * instead — they take a workbook handle rather than hanging off the `Workbook`
 * namespace, which has no CSV members.
 *
 * Design principles:
 * - Only export types and functions that are part of the PUBLIC API
 * - Internal utilities (like format helpers) are used internally but not exported
 * - This reduces bundle size and simplifies the public interface
 */

import type * as CsvAsync from "@csv/parse/async";
import type * as CsvTypes from "@csv/types";
import type * as CsvGenerate from "@csv/utils/generate";
import type * as CsvNumber from "@csv/utils/number";

// =============================================================================
// Core Types (from types.ts)
// =============================================================================

// Row types — `Csv`-prefixed so they do not collide with `Row` from documonster/excel.
export type {
  HeaderArray as CsvHeaderArray,
  RowHashArray as CsvRowHashArray,
  RowArray as CsvRowArray,
  RowMap as CsvRowMap,
  Row as CsvRow
} from "@csv/types";

// Every flat type carries the `Csv` prefix: its vocabulary (`Row`, `ColumnConfig`,
// `TransformResult`, `RecordInfo`) is generic enough to collide with the other entries'.
export type {
  // Transform types
  HeaderTransformFunction as CsvHeaderTransformFunction,
  RowTransformCallback as CsvRowTransformCallback,
  RowTransformFunction as CsvRowTransformFunction,
  RowValidateCallback as CsvRowValidateCallback,
  RowValidateFunction as CsvRowValidateFunction,
  TransformContext as CsvTransformContext,
  FormattedValue as CsvFormattedValue,
  TransformResult as CsvTransformResult,
  TypeTransformMap as CsvTypeTransformMap,

  // Dynamic typing
  DynamicTypingConfig as CsvDynamicTypingConfig,
  CastDateConfig as CsvCastDateConfig,

  // Column config
  ColumnConfig as CsvColumnConfig,

  // Error types (unified)
  CsvErrorCode,
  CsvRecordError,
  OnSkipCallback as CsvOnSkipCallback,

  // Options (general)
  CsvParseOptions,
  CsvFormatOptions,

  // Options (type-safe variants for better inference)
  CsvParseArrayOptions,
  CsvParseObjectOptions,

  // Parse results
  ChunkMeta as CsvChunkMeta,
  CsvParseMeta,
  RecordInfo as CsvRecordInfo,
  RecordWithInfo as CsvRecordWithInfo,
  CsvParseResult,
  CsvParseResultWithObjname,

  // Column mismatch config
  ColumnMismatchConfig as CsvColumnMismatchConfig,
  ColumnMismatchLess as CsvColumnMismatchLess,
  ColumnMismatchMore as CsvColumnMismatchMore
} from "@csv/types";

// =============================================================================
// Core API — the `Csv` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Csv from "@csv/surface/csv";

/** What `Csv.format` takes. */
export type { CsvFormatData } from "@csv/format/formatter";

// Type-only re-exports for the generator + number utilities (the value
// functions live on the `Csv` namespace; these are the option/result types).
export type {
  CsvGenerateOptions,
  CsvGenerateResult,
  CsvGenerator,
  ColumnDef as CsvColumnDef,
  GeneratorColumnConfig as CsvGeneratorColumnConfig,
  BuiltinColumnType as CsvBuiltinColumnType,
  GeneratorFn as CsvGeneratorFn,
  GeneratorContext as CsvGeneratorContext,
  StopCondition as CsvStopCondition,
  StopContext as CsvStopContext
} from "@csv/utils/generate";
export type { DecimalSeparator as CsvDecimalSeparator, CsvNumberOptions } from "@csv/utils/number";
export type { CsvDynamicTypingOptions } from "@csv/utils/dynamic-typing";

// Streaming/async parser input + row shapes.
export type {
  CsvAsyncInput,
  ParsedRow as CsvParsedRow,
  CsvParseAsyncOptions,
  CsvParseProgress
} from "@csv/parse/async";
export type { CsvDetectDelimiterOptions, CsvDetectLinebreakOptions } from "@csv/utils/detect";

/** @deprecated Use {@link CsvHeaderArray}. Removed in 1.0. */
export type HeaderArray = CsvTypes.HeaderArray;
/** @deprecated Use {@link CsvRowHashArray}. Removed in 1.0. */
export type RowHashArray<V = unknown> = CsvTypes.RowHashArray<V>;
/** @deprecated Use {@link CsvRowArray}. Removed in 1.0. */
export type RowArray = CsvTypes.RowArray;
/** @deprecated Use {@link CsvRowMap}. Removed in 1.0. */
export type RowMap = CsvTypes.RowMap;
/** @deprecated Use {@link CsvRow}. Removed in 1.0. */
export type Row = CsvTypes.Row;
/** @deprecated Use {@link CsvParsedRow}. Removed in 1.0. */
export type ParsedRow = CsvAsync.ParsedRow;
/** @deprecated Use {@link CsvHeaderTransformFunction}. Removed in 1.0. */
export type HeaderTransformFunction = CsvTypes.HeaderTransformFunction;
/** @deprecated Use {@link CsvRowTransformCallback}. Removed in 1.0. */
export type RowTransformCallback<T> = CsvTypes.RowTransformCallback<T>;
/** @deprecated Use {@link CsvRowTransformFunction}. Removed in 1.0. */
export type RowTransformFunction<
  I = CsvTypes.Row,
  O = CsvTypes.Row
> = CsvTypes.RowTransformFunction<I, O>;
/** @deprecated Use {@link CsvRowValidateCallback}. Removed in 1.0. */
export type RowValidateCallback = CsvTypes.RowValidateCallback;
/** @deprecated Use {@link CsvRowValidateFunction}. Removed in 1.0. */
export type RowValidateFunction<T = CsvTypes.Row> = CsvTypes.RowValidateFunction<T>;
/** @deprecated Use {@link CsvTransformContext}. Removed in 1.0. */
export type TransformContext = CsvTypes.TransformContext;
/** @deprecated Use {@link CsvFormattedValue}. Removed in 1.0. */
export type FormattedValue = CsvTypes.FormattedValue;
/** @deprecated Use {@link CsvTransformResult}. Removed in 1.0. */
export type TransformResult = CsvTypes.TransformResult;
/** @deprecated Use {@link CsvTypeTransformMap}. Removed in 1.0. */
export type TypeTransformMap = CsvTypes.TypeTransformMap;
/** @deprecated Use {@link CsvDynamicTypingConfig}. Removed in 1.0. */
export type DynamicTypingConfig = CsvTypes.DynamicTypingConfig;
/** @deprecated Use {@link CsvCastDateConfig}. Removed in 1.0. */
export type CastDateConfig = CsvTypes.CastDateConfig;
/** @deprecated Use {@link CsvColumnConfig}. Removed in 1.0. */
export type ColumnConfig = CsvTypes.ColumnConfig;
/** @deprecated Use {@link CsvOnSkipCallback}. Removed in 1.0. */
export type OnSkipCallback = CsvTypes.OnSkipCallback;
/** @deprecated Use {@link CsvChunkMeta}. Removed in 1.0. */
export type ChunkMeta = CsvTypes.ChunkMeta;
/** @deprecated Use {@link CsvRecordInfo}. Removed in 1.0. */
export type RecordInfo = CsvTypes.RecordInfo;
/** @deprecated Use {@link CsvRecordWithInfo}. Removed in 1.0. */
export type RecordWithInfo<T = Record<string, unknown>> = CsvTypes.RecordWithInfo<T>;
/** @deprecated Use {@link CsvColumnMismatchConfig}. Removed in 1.0. */
export type ColumnMismatchConfig = CsvTypes.ColumnMismatchConfig;
/** @deprecated Use {@link CsvColumnMismatchLess}. Removed in 1.0. */
export type ColumnMismatchLess = CsvTypes.ColumnMismatchLess;
/** @deprecated Use {@link CsvColumnMismatchMore}. Removed in 1.0. */
export type ColumnMismatchMore = CsvTypes.ColumnMismatchMore;
/** @deprecated Use {@link CsvColumnDef}. Removed in 1.0. */
export type ColumnDef = CsvGenerate.ColumnDef;
/** @deprecated Use {@link CsvGeneratorColumnConfig}. Removed in 1.0. */
export type GeneratorColumnConfig = CsvGenerate.GeneratorColumnConfig;
/** @deprecated Use {@link CsvBuiltinColumnType}. Removed in 1.0. */
export type BuiltinColumnType = CsvGenerate.BuiltinColumnType;
/** @deprecated Use {@link CsvGeneratorFn}. Removed in 1.0. */
export type GeneratorFn = CsvGenerate.GeneratorFn;
/** @deprecated Use {@link CsvGeneratorContext}. Removed in 1.0. */
export type GeneratorContext = CsvGenerate.GeneratorContext;
/** @deprecated Use {@link CsvStopCondition}. Removed in 1.0. */
export type StopCondition = CsvGenerate.StopCondition;
/** @deprecated Use {@link CsvStopContext}. Removed in 1.0. */
export type StopContext = CsvGenerate.StopContext;
/** @deprecated Use {@link CsvDecimalSeparator}. Removed in 1.0. */
export type DecimalSeparator = CsvNumber.DecimalSeparator;

// =============================================================================
// Errors
// =============================================================================

export { CsvError, CsvArgumentError, CsvLimitExceededError, isCsvError } from "@csv/errors";
export type { CsvLimit } from "@csv/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
