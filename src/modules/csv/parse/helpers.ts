/**
 * CSV Parse Utilities
 *
 * Shared parsing helpers used by both sync (parseCsv) and streaming (CsvParserStream)
 * parsers to ensure consistent behavior:
 *
 * - Header processing: Handle headers option (true/array/transform)
 * - Column validation: Check row length against expected column count
 * - Row-to-object conversion: Transform string[] to Record<string, any>
 * - Dynamic typing: Apply type coercion based on configuration
 *
 * These utilities are extracted to avoid code duplication between
 * the batch parser (parse.ts) and the streaming parser (csv-stream.ts).
 */

import { CsvError } from "../errors";
import type { CsvRecordError, OnSkipCallback } from "../types";
import { deduplicateHeadersWithRenames, type HeaderArray } from "../utils/row";

// =============================================================================
// Types
// =============================================================================

/**
 * Options for header processing
 */
interface HeaderProcessOptions {
  /** Headers configuration: true, array, or function */
  headers: boolean | string[] | ((row: string[]) => (string | null | undefined)[]);
  /** Whether to group columns by name (affects originalHeaders computation) */
  groupColumnsByName?: boolean;
}

/**
 * Options for column validation
 */
interface ColumnValidationOptions {
  /** Strategy for rows with fewer columns than expected */
  columnLess: "error" | "pad";
  /** Strategy for rows with more columns than expected */
  columnMore: "error" | "truncate" | "keep";
}

/**
 * Column validation result
 */
interface ColumnValidationResult {
  /** Whether the row is valid */
  isValid: boolean;
  /** Error code if invalid: 'TooManyFields' or 'TooFewFields' */
  errorCode?: "TooManyFields" | "TooFewFields";
  /** Error message if invalid */
  reason?: string;
  /** Whether the row was modified (padded or trimmed) */
  modified: boolean;
  /** Extra columns when columnMismatch.more is 'keep' */
  extras?: string[];
}

/**
 * Result of processing headers
 */
interface HeaderProcessResult {
  /** The processed (deduplicated) headers */
  headers: HeaderArray;
  /** The original (non-deduplicated) headers, for groupColumnsByName support. Null when groupColumnsByName is false. */
  originalHeaders: HeaderArray | null;
  /** Map of renamed headers (new name -> original name) */
  renamedHeaders: Record<string, string> | null;
  /** Whether the current row should be skipped (was used as headers) */
  skipCurrentRow: boolean;
}

// =============================================================================
// Header Processing
// =============================================================================

/**
 * Process headers from first row or configuration.
 * Shared logic between parseCsv and CsvParserStream.
 *
 * @param row - The current row being processed
 * @param options - Header processing options
 * @param existingHeaders - Already configured headers (for array case)
 * @returns Processing result or null if headers not applicable
 */
export function processHeaders(
  row: string[],
  options: HeaderProcessOptions,
  existingHeaders: HeaderArray | null
): HeaderProcessResult | null {
  const { headers, groupColumnsByName = false } = options;

  // If we already have headers from array config, no processing needed
  if (existingHeaders !== null && Array.isArray(headers)) {
    return null;
  }

  let rawHeaders: (string | null | undefined)[];
  let skipCurrentRow: boolean;

  if (typeof headers === "function") {
    // Function: call with row, skip current row
    rawHeaders = headers(row);
    // Validate returned array length matches the row
    if (rawHeaders.length !== row.length) {
      throw new CsvError(
        `Header function returned ${rawHeaders.length} headers but row has ${row.length} columns. ` +
          `The header function must return an array with the same length as the input row.`
      );
    }
    skipCurrentRow = true;
  } else if (Array.isArray(headers)) {
    // Array: use provided headers, don't skip current row (it's data)
    rawHeaders = headers;
    skipCurrentRow = false;
  } else if (headers) {
    // true: use first row as headers, skip it
    rawHeaders = row;
    skipCurrentRow = true;
  } else {
    // false/undefined: no headers
    return null;
  }

  // Deduplicate headers
  const { headers: dedupedHeaders, renamedHeaders } = deduplicateHeadersWithRenames(rawHeaders);

  // Only compute originalHeaders when groupColumnsByName is true (performance optimization)
  const originalHeaders: HeaderArray | null = groupColumnsByName
    ? rawHeaders.map(h => (h === null || h === undefined ? null : String(h)))
    : null;

  return {
    headers: dedupedHeaders,
    originalHeaders,
    renamedHeaders,
    skipCurrentRow
  };
}

/**
 * Validate and adjust row column count against expected headers.
 * Shared logic between parseCsv and CsvParserStream.
 *
 * @param row - The row to validate (will be modified in place if needed)
 * @param expectedCols - Expected number of columns (from headers)
 * @param options - Validation options
 * @returns Validation result
 */
export function validateAndAdjustColumns(
  row: string[],
  expectedCols: number,
  options: ColumnValidationOptions
): ColumnValidationResult {
  const { columnLess, columnMore } = options;
  const actualCols = row.length;

  if (actualCols === expectedCols) {
    return { isValid: true, modified: false };
  }

  // Too many columns
  if (actualCols > expectedCols) {
    switch (columnMore) {
      case "error":
        return {
          isValid: false,
          errorCode: "TooManyFields",
          reason: `expected ${expectedCols} columns, got ${actualCols}`,
          modified: false
        };
      case "truncate":
        row.length = expectedCols;
        return { isValid: true, errorCode: "TooManyFields", modified: true };
      case "keep": {
        const extras = row.splice(expectedCols);
        return { isValid: true, errorCode: "TooManyFields", modified: true, extras };
      }
      default: {
        const _never: never = columnMore;
        throw new Error(`Unknown columnMore strategy: ${_never}`);
      }
    }
  }

  // Too few columns
  switch (columnLess) {
    case "error":
      return {
        isValid: false,
        errorCode: "TooFewFields",
        reason: `expected ${expectedCols} columns, got ${actualCols}`,
        modified: false
      };
    case "pad":
      while (row.length < expectedCols) {
        row.push("");
      }
      return { isValid: true, errorCode: "TooFewFields", modified: true };
    default: {
      const _never: never = columnLess;
      throw new Error(`Unknown columnLess strategy: ${_never}`);
    }
  }
}

/**
 * Create a safe onSkip handler that catches errors from user callback.
 *
 * The onSkip callback is user-provided and may throw errors. We wrap it
 * to prevent callback errors from interrupting parsing. Errors in the
 * callback are silently ignored since there's no good way to surface them
 * in the sync parsing context.
 *
 * For better error visibility in async/streaming contexts, consider
 * emitting a warning event on the stream instead.
 */
export function createOnSkipHandler(
  onSkip: OnSkipCallback | undefined
): ((error: CsvRecordError, record: string[] | null) => void) | null {
  if (!onSkip) {
    return null;
  }
  return (error: CsvRecordError, record: string[] | null) => {
    try {
      onSkip(error, record);
    } catch (callbackError) {
      // Silently ignore errors in onSkip callback to prevent
      // callback bugs from interrupting CSV parsing.
      // In production, consider logging: console.warn('onSkip callback error:', callbackError);
      void callbackError;
    }
  };
}

/**
 * Convert a row array to an object using headers.
 * Internal helper for convertRowToObject.
 */
function rowToObject(row: string[], headers: HeaderArray): Record<string, string> {
  const obj: Record<string, string> = Object.create(null) as Record<string, string>;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header !== null && header !== undefined && header !== "__proto__") {
      obj[header] = row[i] ?? "";
    }
  }
  return obj;
}

/**
 * Convert a row array to an object, optionally grouping duplicate column names.
 * Unified function that handles both normal and grouped modes.
 *
 * @param row - The row values as an array
 * @param headers - The deduplicated header names
 * @param originalHeaders - The original (non-deduplicated) headers for grouping
 * @param groupColumnsByName - Whether to group duplicate column names
 * @returns Object with header keys and row values
 */
export function convertRowToObject(
  row: string[],
  headers: HeaderArray,
  originalHeaders: HeaderArray | null,
  groupColumnsByName: boolean
): Record<string, string | string[]> {
  if (groupColumnsByName && originalHeaders) {
    return rowToObjectGrouped(row, originalHeaders);
  }
  return rowToObject(row, headers);
}

/**
 * Convert a row array to an object, grouping duplicate column names.
 * Internal helper for convertRowToObject.
 */
function rowToObjectGrouped(
  row: string[],
  headers: HeaderArray
): Record<string, string | string[]> {
  const obj: Record<string, string | string[]> = Object.create(null) as Record<
    string,
    string | string[]
  >;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header !== null && header !== undefined && header !== "__proto__") {
      const value = row[i] ?? "";
      if (header in obj) {
        // Column name already exists - convert to array or push to existing array
        const existing = obj[header];
        if (Array.isArray(existing)) {
          existing.push(value);
        } else {
          obj[header] = [existing, value];
        }
      } else {
        obj[header] = value;
      }
    }
  }
  return obj;
}

/**
 * Filter out null/undefined values from a header array.
 * Returns only the valid string headers.
 *
 * @param headers - Header array that may contain null/undefined values
 * @returns Array of valid string headers (null/undefined removed)
 */
export function filterValidHeaders(headers: HeaderArray): string[] {
  return headers.filter((h): h is string => h !== null && h !== undefined);
}
