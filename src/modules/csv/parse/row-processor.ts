/**
 * CSV Row Processor
 *
 * Core row processing logic shared between sync and streaming parsers.
 * Handles header processing, column validation, and row completion.
 */

import type { ParseConfig } from "@csv/parse/config";
import { processHeaders, validateAndAdjustColumns, convertRowToObject } from "@csv/parse/helpers";
import type { ParseState } from "@csv/parse/state";
import type { CsvRecordError, HeaderArray, RecordInfo } from "@csv/types";
import { applyDynamicTypingToRow } from "@csv/utils/dynamic-typing";
import { isEmptyRow, hasAllEmptyValues } from "@csv/utils/row";

// =============================================================================
// Types
// =============================================================================

/**
 * Result of processing a single row
 */
export interface RowProcessResult {
  /** Whether to stop parsing (maxRows reached) */
  stop: boolean;
  /** Whether row was skipped (invalid, filtered, etc.) */
  skipped: boolean;
  /** Processed row data (if not skipped) */
  row?: string[];
  /** Record info (if info option enabled) */
  info?: RecordInfo;
  /** Error that occurred (if any) */
  error?: CsvRecordError;
  /** Reason for skipping/invalidating the row */
  reason?: string;
  /** Extra columns when columnMismatch.more is 'keep' */
  extras?: string[];
}

// =============================================================================
// Header Processing
// =============================================================================

/**
 * Process headers from a row (first data row or configured headers)
 * Returns true if the row should be skipped (was used as headers)
 */
export function processHeaderRow(
  row: string[],
  state: ParseState,
  config: Pick<ParseConfig, "headers" | "groupColumnsByName">
): boolean {
  const result = processHeaders(
    row,
    {
      headers: config.headers as boolean | string[] | ((h: string[]) => HeaderArray),
      groupColumnsByName: config.groupColumnsByName
    },
    state.headerRow
  );

  if (result) {
    state.headerRow = result.headers;
    state.originalHeaders = result.originalHeaders;
    state.renamedHeadersForMeta = result.renamedHeaders;
    state.headerRowProcessed = true;
    return result.skipCurrentRow;
  }

  state.headerRowProcessed = true;
  return false;
}

// =============================================================================
// Column Validation
// =============================================================================

/**
 * Validate row column count against headers
 * Returns error info if validation fails, null otherwise
 */
export function validateRowColumns(
  row: string[],
  state: ParseState,
  config: Pick<ParseConfig, "columnLess" | "columnMore">
): {
  errorCode: "TooManyFields" | "TooFewFields";
  message: string;
  isValid: boolean;
  reason?: string;
  extras?: string[];
} | null {
  if (!state.headerRow || state.headerRow.length === 0) {
    return null;
  }

  const expectedCols = state.headerRow.length;
  const actualCols = row.length;

  if (actualCols === expectedCols) {
    return null;
  }

  const validation = validateAndAdjustColumns(row, expectedCols, {
    columnLess: config.columnLess,
    columnMore: config.columnMore
  });

  if (validation.errorCode) {
    return {
      errorCode: validation.errorCode,
      message:
        validation.errorCode === "TooManyFields"
          ? `Too many fields: expected ${expectedCols}, found ${actualCols}`
          : `Too few fields: expected ${expectedCols}, found ${actualCols}`,
      isValid: validation.isValid,
      reason: validation.reason,
      extras: validation.extras
    };
  }

  return null;
}

// =============================================================================
// Record Info Building
// =============================================================================

/**
 * Build record info for a completed row
 */
export function buildRecordInfo(
  state: ParseState,
  dataRowIndex: number,
  includeRaw: boolean,
  fieldCount: number
): RecordInfo {
  const info: RecordInfo = {
    index: dataRowIndex,
    line: state.currentRowStartLine,
    offset: state.currentRowStartOffset,
    quoted: state.currentRowQuoted.slice(0, fieldCount) as boolean[]
  };
  if (includeRaw) {
    info.raw = state.currentRawRow;
  }
  return info;
}

// =============================================================================
// Row to Record Conversion
// =============================================================================

/**
 * Convert a raw row to an object record with optional dynamic typing
 */
export function rowToRecord(
  row: string[],
  state: ParseState,
  config: Pick<ParseConfig, "groupColumnsByName" | "dynamicTyping" | "castDate">
): Record<string, unknown> {
  if (state.headerRow) {
    let record: Record<string, unknown> = convertRowToObject(
      row,
      state.headerRow,
      state.originalHeaders,
      config.groupColumnsByName
    );
    if (config.dynamicTyping || config.castDate) {
      record = applyDynamicTypingToRow(
        record as Record<string, string>,
        config.dynamicTyping || false,
        config.castDate
      );
    }
    return record;
  }
  // No headers: use numeric indices as keys (O(n) instead of O(n²) reduce)
  const result: Record<number, string> = {};
  for (let i = 0; i < row.length; i++) {
    result[i] = row[i];
  }
  return result;
}

// =============================================================================
// Row Skip Logic
// =============================================================================

/**
 * Check if a row should be skipped (comment or empty)
 */
export function shouldSkipRow(
  row: string[],
  comment: string | undefined,
  shouldSkipEmpty: boolean | "greedy",
  skipRecordsWithEmptyValues: boolean
): boolean {
  // Comment line check - trim first field before checking for comment prefix
  // to handle lines like " # comment" with leading whitespace
  if (comment && row[0]?.trimStart().startsWith(comment)) {
    return true;
  }
  // Empty row check
  if (isEmptyRow(row, shouldSkipEmpty)) {
    return true;
  }
  // All empty values check
  if (skipRecordsWithEmptyValues && hasAllEmptyValues(row)) {
    return true;
  }
  return false;
}

// =============================================================================
// Main Row Processing
// =============================================================================

/**
 * Process a completed row through headers, validation, etc.
 * This is the core row processing logic shared between sync and streaming parsers.
 */
export function processCompletedRow(
  row: string[],
  state: ParseState,
  config: ParseConfig,
  errors: CsvRecordError[],
  lineNumber: number
): RowProcessResult {
  // Header handling
  if (state.useHeaders && !state.headerRowProcessed) {
    const shouldSkip = processHeaderRow(row, state, config);
    if (shouldSkip) {
      return { stop: false, skipped: true };
    }
  }

  // Skip data rows
  if (state.skippedDataRows < config.skipRows) {
    state.skippedDataRows++;
    return { stop: false, skipped: true };
  }

  // Column validation
  const validationError = validateRowColumns(row, state, config);
  let extras: string[] | undefined;

  if (validationError) {
    const errorObj: CsvRecordError = {
      code: validationError.errorCode,
      message: validationError.message,
      line: lineNumber
    };
    errors.push(errorObj);

    if (!validationError.isValid) {
      if (config.skipRecordsWithError) {
        config.invokeOnSkip?.(
          {
            code: validationError.errorCode,
            message: validationError.reason || "Column mismatch",
            line: lineNumber
          },
          row
        );
        return {
          stop: false,
          skipped: true,
          row,
          error: {
            code: validationError.errorCode,
            message: validationError.reason || "Column mismatch",
            line: lineNumber
          },
          reason: validationError.reason || "Column mismatch"
        };
      }
      // Column mismatch with error strategy - return as invalid
      return {
        stop: false,
        skipped: true,
        row,
        error: errorObj,
        reason: validationError.reason || "Column mismatch"
      };
    }

    // Valid but had extras (columnMore: 'keep')
    extras = validationError.extras;
  }

  // Skip records with all empty values
  if (config.skipRecordsWithEmptyValues && hasAllEmptyValues(row)) {
    return { stop: false, skipped: true };
  }

  // Check maxRows BEFORE incrementing count
  if (config.maxRows !== undefined && state.dataRowCount >= config.maxRows) {
    state.truncated = true;
    return { stop: true, skipped: false };
  }

  state.dataRowCount++;

  // Build info if needed
  let info: RecordInfo | undefined;
  if (config.infoOption) {
    info = buildRecordInfo(state, state.dataRowCount - 1, config.rawOption, row.length);
  }

  return { stop: false, skipped: false, row, info, extras };
}
