/**
 * CSV Parser - Synchronous
 *
 * RFC 4180 compliant CSV parser.
 * Provides parseCsv function and low-level parsing generators.
 */

import { getUtf8ByteLength } from "@csv/constants";
import type { ParseConfig } from "@csv/parse/config";
import { resolveParseConfig, toScannerConfig } from "@csv/parse/config";
import { filterValidHeaders } from "@csv/parse/helpers";
import { splitLinesWithEndings } from "@csv/parse/lines";
import type { RowProcessResult } from "@csv/parse/row-processor";
import { processCompletedRow, rowToRecord } from "@csv/parse/row-processor";
import { scanRow as scanRowImpl } from "@csv/parse/scanner";
import type { ParseState } from "@csv/parse/state";
import { createParseState, resetInfoState, getUnquotedArray } from "@csv/parse/state";
import type {
  CsvParseOptions,
  CsvParseArrayOptions,
  CsvParseObjectOptions,
  CsvParseResult,
  CsvParseResultWithObjname,
  CsvParseMeta,
  CsvRecordError,
  RecordWithInfo,
  DynamicTypingConfig,
  CastDateConfig
} from "@csv/types";
import { applyDynamicTypingToArrayRow } from "@csv/utils/dynamic-typing";
import { isEmptyRow } from "@csv/utils/row";

// =============================================================================
// Helper Functions
// =============================================================================

/**
 * Normalize validate result to { isValid, reason } form
 */
function normalizeValidateResult(result: boolean | { isValid: boolean; reason?: string }): {
  isValid: boolean;
  reason: string;
} {
  if (typeof result === "boolean") {
    return { isValid: result, reason: "Validation failed" };
  }
  return { isValid: result.isValid, reason: result.reason || "Validation failed" };
}

/**
 * Apply dynamic typing to an array row (wrapper to reduce code duplication)
 */
function applyArrayTyping(
  row: string[],
  dynamicTyping: DynamicTypingConfig | undefined,
  castDate: CastDateConfig | undefined
): unknown[] {
  return applyDynamicTypingToArrayRow(row, null, dynamicTyping || false, castDate);
}

/**
 * Return array only if non-empty, otherwise undefined
 */
function optionalArray<T>(arr: T[]): T[] | undefined {
  return arr.length > 0 ? arr : undefined;
}

/**
 * Build CsvParseMeta from config and state (avoids duplication between array and object mode)
 */
function buildMeta(config: ParseConfig, state: ParseState): CsvParseMeta {
  return {
    delimiter: config.delimiter,
    linebreak: config.linebreak,
    aborted: false,
    truncated: state.truncated,
    cursor: state.dataRowCount,
    fields: state.headerRow ? filterValidHeaders(state.headerRow) : undefined,
    renamedHeaders: state.renamedHeadersForMeta
  };
}

/**
 * Apply trim function to all fields in a row.
 * Uses cached trimFieldIsIdentity from config to avoid per-row checking.
 */
function trimFields(fields: string[], config: ParseConfig): string[] {
  // Fast path: if trim is identity function, return fields as-is
  if (config.trimFieldIsIdentity) {
    return fields;
  }
  return fields.map(config.trimField);
}

// =============================================================================
// Fast Mode Parser (No Quote Detection)
// =============================================================================

/**
 * Parse input using fast mode (no quote detection)
 */
export function* parseFastMode(
  input: string,
  config: ParseConfig,
  state: ParseState,
  errors: CsvRecordError[]
): Generator<RowProcessResult, void, undefined> {
  // Handle empty input - no rows to produce
  if (input === "") {
    return;
  }

  // Track character offset for info.offset
  let currentCharOffset = 0;

  for (const { line, lineLengthWithEnding: lineCharLength } of splitLinesWithEndings(
    input,
    config.linebreakRegex
  )) {
    state.lineNumber++;

    if (config.toLine !== undefined && state.lineNumber > config.toLine) {
      state.truncated = true;
      break;
    }
    if (state.lineNumber <= config.skipLines) {
      currentCharOffset += lineCharLength;
      continue;
    }
    // Only skip empty lines if skipEmptyLines option is enabled
    if (line === "" && config.shouldSkipEmpty) {
      currentCharOffset += lineCharLength;
      continue;
    }

    // Check maxRowBytes in fastMode using optimized byte length calculation
    if (config.maxRowBytes !== undefined) {
      const lineBytes = getUtf8ByteLength(line);
      if (lineBytes > config.maxRowBytes) {
        throw new Error(`Row exceeds the maximum size of ${config.maxRowBytes} bytes`);
      }
    }

    if (config.infoOption) {
      state.currentRowStartLine = state.lineNumber;
      state.currentRowStartOffset = currentCharOffset;
    }
    if (config.rawOption) {
      state.currentRawRow = line;
    }

    const row = line.split(config.delimiter);
    const trimmedRow = trimFields(row, config);

    if (config.infoOption) {
      state.currentRowQuoted = getUnquotedArray(trimmedRow.length);
    }

    if (config.comment && trimmedRow[0]?.trimStart().startsWith(config.comment)) {
      currentCharOffset += lineCharLength;
      continue;
    }
    if (config.shouldSkipEmpty && isEmptyRow(trimmedRow, config.shouldSkipEmpty)) {
      currentCharOffset += lineCharLength;
      continue;
    }

    const result = processCompletedRow(trimmedRow, state, config, errors, state.lineNumber);
    currentCharOffset += lineCharLength;

    if (result.stop) {
      yield result;
      return;
    }
    // Yield if not skipped, OR if skipped with an error (for invalidRows collection)
    if (!result.skipped || result.error) {
      yield result;
    }
    resetInfoState(
      state,
      config.infoOption,
      config.rawOption,
      state.lineNumber + 1,
      currentCharOffset
    );
  }
}

// =============================================================================
// Scanner-based Parser (High-Performance)
// =============================================================================

/**
 * Parse input using Scanner-based batch scanning.
 * This is a high-performance alternative that uses indexOf-based field scanning
 * instead of character-by-character parsing.
 *
 * Key optimizations:
 * 1. Uses indexOf to find delimiters/quotes/newlines in bulk
 * 2. Uses slice for field extraction (avoids string concatenation)
 * 3. Processes entire rows at once instead of character-by-character
 */
export function* parseWithScanner(
  input: string,
  config: ParseConfig,
  state: ParseState,
  errors: CsvRecordError[]
): Generator<RowProcessResult, void, undefined> {
  const scannerConfig = toScannerConfig(config);
  const len = input.length;
  let pos = 0;

  if (config.infoOption) {
    state.currentRowStartOffset = 0;
  }

  while (pos < len) {
    // Scan one row at a time
    const scanResult = scanRowImpl(input, pos, scannerConfig, true);

    // No fields and no progress - should not happen with isEof=true
    if (scanResult.fields.length === 0 && scanResult.endPos === pos) {
      break;
    }

    // Apply trim to fields
    const row = trimFields(scanResult.fields, config);

    // Save the start line BEFORE counting newlines (for accurate info.line on multi-line rows).
    // This must happen before any skip checks, so that skipped rows don't leave
    // currentRowStartLine stale from a previous iteration.
    const rowStartLine = state.lineNumber + 1;

    // Update line number (count newlines in raw content for multi-line quoted fields)
    {
      const rawStart = scanResult.rawStart;
      const rawEnd = scanResult.rawEnd;
      let newlines = 1; // At least one line per row
      for (let i = rawStart; i < rawEnd; i++) {
        const ch = input.charCodeAt(i);
        if (ch === 10) {
          // \n
          newlines++;
        } else if (ch === 13) {
          // \r — skip \r\n as single newline
          if (i + 1 < rawEnd && input.charCodeAt(i + 1) === 10) {
            i++;
          }
          newlines++;
        }
      }
      state.lineNumber += newlines;
    }

    // Check toLine limit
    if (config.toLine !== undefined && state.lineNumber > config.toLine) {
      state.truncated = true;
      break;
    }

    // Calculate positions for raw/info tracking
    // Use rawEnd directly from scan result (position before newline)
    const rawEndPos = scanResult.rawEnd;

    // Skip lines at beginning (must be before maxRowBytes to avoid errors on skipped rows)
    if (state.lineNumber <= config.skipLines) {
      pos = scanResult.endPos;
      continue;
    }

    // Check maxRowBytes limit
    if (config.maxRowBytes !== undefined) {
      const rawRow = input.slice(scanResult.rawStart, rawEndPos);
      const rowBytes = getUtf8ByteLength(rawRow);
      if (rowBytes > config.maxRowBytes) {
        throw new Error(`Row exceeds the maximum size of ${config.maxRowBytes} bytes`);
      }
    }

    // Skip comment lines
    if (config.comment && row[0]?.trimStart().startsWith(config.comment)) {
      pos = scanResult.endPos;
      continue;
    }

    // Skip empty lines
    if (config.shouldSkipEmpty && isEmptyRow(row, config.shouldSkipEmpty)) {
      pos = scanResult.endPos;
      continue;
    }

    // Set up info tracking
    // Use rowStartLine computed BEFORE newline counting — this gives the correct
    // 1-based line number where the row starts, even after skipped rows.
    if (config.infoOption) {
      state.currentRowStartLine = rowStartLine;
      state.currentRowStartOffset = scanResult.rawStart;
      state.currentRowQuoted = scanResult.quoted;
    }

    // Extract raw row using zero-copy from scan result
    if (config.rawOption) {
      state.currentRawRow = input.slice(scanResult.rawStart, rawEndPos);
    }

    // Populate state for processCompletedRow
    // (state.currentRow was removed as dead code - row is passed directly)

    // Check for unterminated quotes and report error
    if (scanResult.unterminatedQuote) {
      // Line number for error is 1-based
      errors.push({
        code: "MissingQuotes",
        message: "Quoted field unterminated",
        line: state.lineNumber
      });
    }

    const result = processCompletedRow(row, state, config, errors, state.lineNumber);

    if (result.stop) {
      yield result;
      return;
    }

    if (!result.skipped || result.error) {
      yield result;
    }

    // Reset for next row
    pos = scanResult.endPos;

    if (config.infoOption) {
      state.currentRowStartOffset = scanResult.endPos;
    }
  }
}

// =============================================================================
// Function Overloads for Better Type Inference
// =============================================================================

/**
 * Parse CSV string - returns string[][] when no options provided.
 */
export function parseCsv(input: string): string[][];

/**
 * Parse CSV string - returns string[][] when headers is false/undefined and no info option.
 *
 * Note: When `info: true` is set, returns CsvParseResult instead.
 */
export function parseCsv(
  input: string,
  options: CsvParseArrayOptions & { info?: false }
): string[][];

/**
 * Parse CSV string - returns CsvParseResult with RecordWithInfo when info: true (array mode).
 */
export function parseCsv(
  input: string,
  options: CsvParseArrayOptions & { info: true }
): CsvParseResult<RecordWithInfo<string[]>>;

/**
 * Parse CSV string - returns CsvParseResult when headers are enabled.
 */
export function parseCsv(
  input: string,
  options: CsvParseObjectOptions & { info?: false }
): CsvParseResult<Record<string, unknown>>;

/**
 * Parse CSV string - returns CsvParseResult with RecordWithInfo when info: true (object mode).
 */
export function parseCsv(
  input: string,
  options: CsvParseObjectOptions & { info: true }
): CsvParseResult<RecordWithInfo<Record<string, unknown>>>;

/**
 * Parse CSV string - general overload for backward compatibility.
 */
export function parseCsv(
  input: string,
  options: CsvParseOptions
):
  | string[][]
  | CsvParseResult<Record<string, string>>
  | CsvParseResult<Record<string, unknown>>
  | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
  | CsvParseResult<RecordWithInfo<string[]>>;

/**
 * Parse CSV string synchronously.
 *
 * @example
 * ```ts
 * // Simple array output (no headers)
 * const rows = parseCsv("a,b,c\n1,2,3");
 * // rows: string[][] = [["a","b","c"], ["1","2","3"]]
 *
 * // Object output with headers
 * const result = parseCsv("name,age\nAlice,30", { headers: true });
 * // result.rows: Record<string, unknown>[] = [{ name: "Alice", age: "30" }]
 *
 * // With info option
 * const result = parseCsv("a,b\n1,2", { info: true });
 * // result.rows: RecordWithInfo<string[]>[] = [{ record: ["a","b"], info: {...} }, ...]
 * ```
 */
export function parseCsv(
  input: string,
  options: CsvParseOptions = {}
):
  | string[][]
  | CsvParseResult<Record<string, string>>
  | CsvParseResult<Record<string, unknown>>
  | CsvParseResult<RecordWithInfo<Record<string, unknown>>>
  | CsvParseResult<RecordWithInfo<string[]>>
  | CsvParseResult<string[] | unknown[]>
  | CsvParseResultWithObjname<Record<string, unknown>> {
  // Resolve config and preprocess input
  const { config, processedInput } = resolveParseConfig(input, options);

  // Initialize state
  const state = createParseState(config);
  const errors: CsvRecordError[] = [];
  const invalidRows: { row: string[]; reason: string }[] = [];

  // Choose parser based on mode
  const parser = config.fastMode
    ? parseFastMode(processedInput, config, state, errors)
    : parseWithScanner(processedInput, config, state, errors);

  // ==========================================================================
  // Single-pass processing: parse + transform + validate + dynamicTyping
  // ==========================================================================

  // Simple array output (no headers) - True single pass processing
  if (!state.useHeaders) {
    // Use unified type for both info and non-info mode to avoid two-pass zipping
    const processedRows: (string[] | unknown[] | RecordWithInfo<string[] | unknown[]>)[] = [];

    for (const result of parser) {
      if (result.row && !result.skipped) {
        let row: string[] | unknown[] = result.row;

        // Apply rowTransform if provided
        if (options.rowTransform) {
          const transformed = options.rowTransform(row as string[]);
          if (transformed === null || transformed === undefined) {
            continue;
          }
          row = transformed as string[] | unknown[];
        }

        // Apply validate if provided
        if (options.validate) {
          const { isValid, reason } = normalizeValidateResult(options.validate(row as string[]));
          if (!isValid) {
            invalidRows.push({ row: row as string[], reason });
            continue;
          }
        }

        // Apply dynamicTyping/castDate if configured
        if (config.dynamicTyping || config.castDate) {
          row = applyArrayTyping(row as string[], config.dynamicTyping, config.castDate);
        }

        // Push with or without info in single pass
        if (config.infoOption && result.info) {
          processedRows.push({ record: row, info: result.info });
        } else {
          processedRows.push(row);
        }
      } else if (result.row && result.skipped && result.error) {
        // Handle invalid rows from columnMismatch errors
        invalidRows.push({ row: result.row, reason: result.reason || result.error.message });
      }
      if (result.stop) {
        break;
      }
    }

    // Build metadata
    const meta = buildMeta(config, state);

    // If info option is enabled, rows are already wrapped
    if (config.infoOption) {
      return {
        headers: undefined,
        rows: processedRows as RecordWithInfo<string[] | unknown[]>[],
        invalidRows: optionalArray(invalidRows),
        errors: optionalArray(errors),
        meta
      } as CsvParseResult<RecordWithInfo<string[]>>;
    }

    // If validate was used, always return result object for consistent API
    // This allows users to check invalidRows even when all rows pass validation
    if (options.validate) {
      return {
        headers: undefined,
        rows: processedRows,
        invalidRows: optionalArray(invalidRows),
        errors: optionalArray(errors),
        meta
      } as CsvParseResult<string[] | unknown[]>;
    }

    return processedRows as string[][];
  }

  // ==========================================================================
  // Object mode (with headers) - True single-pass processing
  // ==========================================================================

  // Process rows in single pass: parse + convert + transform + validate
  const objectRows: (Record<string, unknown> | RecordWithInfo<Record<string, unknown>>)[] = [];

  for (const result of parser) {
    if (result.row && !result.skipped) {
      // Convert to record immediately (single pass, no intermediate array)
      let record = rowToRecord(result.row, state, config);

      // Add extras if columnMismatch.more: 'keep' was used
      if (result.extras && result.extras.length > 0) {
        record._extra = result.extras;
      }

      // Apply rowTransform if provided
      if (options.rowTransform) {
        const transformed = options.rowTransform(record as Record<string, string>);
        if (transformed === null || transformed === undefined) {
          continue;
        }
        record = transformed as Record<string, unknown>;
      }

      // Apply validate if provided
      if (options.validate) {
        const { isValid, reason } = normalizeValidateResult(
          options.validate(record as Record<string, string>)
        );
        if (!isValid) {
          invalidRows.push({ row: result.row, reason });
          continue;
        }
      }

      if (config.infoOption && result.info) {
        objectRows.push({ record, info: result.info });
      } else {
        objectRows.push(record);
      }
    } else if (result.row && result.skipped && result.error) {
      invalidRows.push({ row: result.row, reason: result.reason || result.error.message });
    }
    if (result.stop) {
      break;
    }
  }

  // Build metadata
  const meta = buildMeta(config, state);

  // Handle objname option
  const { objname } = options;
  if (objname && state.headerRow) {
    const objResult: Record<
      string,
      Record<string, unknown> | RecordWithInfo<Record<string, unknown>>
    > = Object.create(null) as Record<
      string,
      Record<string, unknown> | RecordWithInfo<Record<string, unknown>>
    >;
    for (const item of objectRows) {
      const rec = config.infoOption
        ? (item as RecordWithInfo<Record<string, unknown>>).record
        : item;
      const key = (rec as Record<string, unknown>)[objname];
      // Convert undefined/null to empty string, otherwise convert to string
      const keyStr = key === undefined || key === null ? "" : String(key);
      // Skip __proto__ to prevent prototype pollution via JSON.
      // Note: constructor/prototype are safe on Object.create(null) objects.
      if (keyStr === "__proto__") {
        continue;
      }
      objResult[keyStr] = item;
    }
    return {
      headers: meta.fields,
      rows: objResult,
      invalidRows: optionalArray(invalidRows),
      errors: optionalArray(errors),
      meta
    } as CsvParseResultWithObjname<Record<string, unknown>>;
  }

  return {
    headers: meta.fields,
    rows: objectRows,
    invalidRows: optionalArray(invalidRows),
    errors: optionalArray(errors),
    meta
  } as CsvParseResult<Record<string, unknown>>;
}
