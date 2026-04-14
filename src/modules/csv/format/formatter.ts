/**
 * CSV Formatter
 *
 * Core formatting functions for converting data to CSV strings.
 *
 * Low-level exports:
 * - formatField(): Format a single field value
 * - formatRowWithLookup(): Format an entire row (used by CsvFormatterStream)
 * - applyTypeTransform(): Apply type-based transforms
 * - defaultToString(): Default value-to-string conversion
 *
 * High-level exports:
 * - formatCsv(): Batch format data to CSV string
 *
 * Features:
 * - Multiple input types (objects, arrays, RowHashArray)
 * - Flexible quoting (per-column, per-header, always, disabled)
 * - Type transforms with context
 * - Formula escaping (CSV injection protection)
 * - BOM support
 */

import type {
  CsvFormatOptions,
  Row,
  TypeTransformMap,
  TransformContext,
  TransformResult
} from "../types";
import { startsWithFormulaChar } from "../utils/detect";
import { formatNumberForCsv, type DecimalSeparator } from "../utils/number";
import {
  deduplicateHeaders,
  isRowHashArray,
  rowHashArrayToHeaders,
  rowHashArrayToValues,
  rowHashArrayMapByHeaders,
  processColumns
} from "../utils/row";
import type { CsvFormatRegex, FormatFieldContext, FormatRowOptions, FormatConfig } from "./config";
import { createFormatConfig } from "./config";
import { isFormattedValue } from "./formatted-value";

// =============================================================================
// Type Transform Functions
// =============================================================================

/**
 * Apply type-based transform to a single value.
 * Returns the transformed result, or undefined if no transform applies.
 */
export function applyTypeTransform(
  value: any,
  transform: TypeTransformMap,
  ctx: TransformContext
): TransformResult {
  if (value === null || value === undefined) {
    return undefined;
  }

  const type = typeof value;

  if (type === "boolean" && transform.boolean) {
    return transform.boolean(value, ctx);
  }
  if (value instanceof Date && transform.date) {
    return transform.date(value, ctx);
  }
  if (type === "number" && transform.number) {
    return transform.number(value, ctx);
  }
  if (type === "bigint" && transform.bigint) {
    return transform.bigint(value, ctx);
  }
  if (type === "string" && transform.string) {
    return transform.string(value, ctx);
  }
  // Handle plain objects (not Date, not Array, not null)
  if (type === "object" && !Array.isArray(value) && !(value instanceof Date)) {
    if (transform.object) {
      return transform.object(value, ctx);
    }
  }

  return undefined;
}

/**
 * Default type conversion to string.
 */
export function defaultToString(value: any, decimalSeparator: DecimalSeparator): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number") {
    return formatNumberForCsv(value, decimalSeparator);
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === "bigint") {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "true" : "false";
  }
  if (typeof value === "object") {
    try {
      return JSON.stringify(value);
    } catch {
      // Handle circular references or other JSON.stringify errors
      return "[object Object]";
    }
  }
  return String(value);
}

// =============================================================================
// Field Formatting
// =============================================================================

/**
 * Fast check if a string needs quoting (for single-char delimiter/quote/escape)
 * Uses indexOf for slightly better V8 optimization
 */
function needsQuoteFast(str: string, delimiter: string, quote: string, escape: string): boolean {
  return (
    str.indexOf(delimiter) !== -1 ||
    str.indexOf(quote) !== -1 ||
    (escape !== quote && str.indexOf(escape) !== -1) ||
    str.indexOf("\n") !== -1 ||
    str.indexOf("\r") !== -1
  );
}

/**
 * Format a single field value to CSV string
 */
export function formatField(
  value: unknown,
  regex: CsvFormatRegex,
  ctx: FormatFieldContext
): string {
  const {
    index,
    header,
    isHeader,
    outputRowIndex,
    forceQuote,
    quoteAll,
    escapeFormulae,
    decimalSeparator,
    transform
  } = ctx;

  // Apply type-based transform if provided (not for headers)
  let str: string;
  // Track if transform explicitly requested quoting control
  let transformQuoteHint: boolean | undefined;

  if (!isHeader && transform) {
    // Create fresh context for each call to ensure safety if user stores reference
    const transformCtx: TransformContext = { column: header ?? index, index: outputRowIndex };
    const transformed = applyTypeTransform(value, transform, transformCtx);

    if (transformed === undefined || transformed === null) {
      str = defaultToString(value, decimalSeparator);
    } else if (isFormattedValue(transformed)) {
      // FormattedValue contains explicit quoting hint
      str = transformed.value;
      transformQuoteHint = transformed.quote;
    } else {
      str = transformed as string;
    }
  } else {
    str = defaultToString(value, decimalSeparator);
  }

  // Escape formulae to prevent CSV injection (OWASP recommendation)
  // Prefix dangerous characters with single quote to neutralize them in spreadsheet apps
  // Using single quote (') as recommended by OWASP, which Excel interprets as a text prefix
  // Skip numeric types: negative numbers like -5.55 are not formula injection vectors
  if (
    escapeFormulae &&
    transformQuoteHint !== false &&
    typeof value !== "number" &&
    typeof value !== "bigint" &&
    startsWithFormulaChar(str)
  ) {
    str = "'" + str;
  }

  // If quoting is disabled, return raw string
  if (!regex.quoteEnabled) {
    return str;
  }

  // Check if quoting is needed
  // Transform quote hint takes precedence (explicit control via quoted()/unquoted())
  let needsQuote: boolean;
  if (transformQuoteHint !== undefined) {
    needsQuote = transformQuoteHint;
  } else {
    needsQuote =
      quoteAll ||
      forceQuote ||
      (regex.useFastCheck
        ? needsQuoteFast(str, regex.delimiter, regex.quote, regex.escape)
        : regex.needsQuoteRegex!.test(str));
  }

  if (needsQuote) {
    // Escape quotes (and escape chars if different from quote) using pre-compiled regex
    let escaped: string;
    if (regex.escape !== regex.quote) {
      // When escape !== quote, regex matches both quote and escape chars.
      // Use replacement function to escape each correctly.
      escaped = str.replace(regex.escapeQuoteRegex!, ch =>
        ch === regex.quote ? regex.escape + regex.quote : regex.escape + regex.escape
      );
    } else {
      escaped = str.replace(regex.escapeQuoteRegex!, regex.escapedQuote);
    }
    return regex.quote + escaped + regex.quote;
  }

  return str;
}

// =============================================================================
// Row Formatting
// =============================================================================

/**
 * Format an entire row to CSV string.
 *
 * Performance optimizations:
 * - Uses for loop with direct string building instead of map().join()
 * - Reuses a single mutable context object instead of creating one per field
 */
export function formatRowWithLookup(
  row: unknown[],
  regex: CsvFormatRegex,
  options: FormatRowOptions
): string {
  const {
    quoteLookup,
    delimiter,
    headers,
    isHeader,
    outputRowIndex,
    quoteAll,
    escapeFormulae,
    decimalSeparator,
    transform
  } = options;

  const len = row.length;
  if (len === 0) {
    return "";
  }

  // Reusable context object - mutate index/header/forceQuote per field
  // This avoids creating a new object for every field
  const ctx: FormatFieldContext = {
    index: 0,
    header: headers?.[0],
    isHeader,
    outputRowIndex,
    forceQuote: quoteLookup(0, headers?.[0]),
    quoteAll,
    escapeFormulae,
    decimalSeparator,
    transform
  };

  // Build string directly without intermediate array from map()
  let result = formatField(row[0], regex, ctx);

  for (let i = 1; i < len; i++) {
    ctx.index = i;
    ctx.header = headers?.[i];
    ctx.forceQuote = quoteLookup(i, ctx.header);
    result += delimiter + formatField(row[i], regex, ctx);
  }

  return result;
}

// =============================================================================
// Input Normalization
// =============================================================================

/**
 * Apply row transform if configured. Returns null to skip the row.
 */
function applyRowTransform(cfg: FormatConfig, row: Row, index: number): Row | null {
  if (!cfg.typeTransform?.row) {
    return row;
  }
  const t = cfg.typeTransform.row(row, index);
  return t === null ? null : t;
}

interface NormalizedInput {
  keys: string[] | null;
  displayHeaders: string[] | null;
  rows: unknown[][];
}

/**
 * Normalize all input types to a unified format.
 * Handles: objects, arrays, RowHashArray, and columns config.
 */
function normalizeInput(
  data: Row[] | Record<string, unknown>[],
  options: CsvFormatOptions,
  cfg: FormatConfig
): NormalizedInput {
  const { headers, columns } = options;

  // Empty data
  if (data.length === 0) {
    if (columns && columns.length > 0) {
      const displayHeaders = columns.map(c => (typeof c === "string" ? c : (c.header ?? c.key)));
      return { keys: null, displayHeaders, rows: [] };
    }
    if (Array.isArray(headers)) {
      return { keys: headers, displayHeaders: headers, rows: [] };
    }
    return { keys: null, displayHeaders: null, rows: [] };
  }

  const firstRow = data[0];

  // Columns config takes precedence
  if (columns && columns.length > 0) {
    const processed = processColumns(columns)!;
    const keys = processed.keys;
    const displayHeaders = processed.headers;

    const rows: unknown[][] = [];
    for (let i = 0; i < data.length; i++) {
      const row = applyRowTransform(cfg, data[i] as Row, i);
      if (row === null) {
        continue;
      }

      let values: unknown[];
      if (isRowHashArray(row)) {
        values = rowHashArrayMapByHeaders(row as [string, unknown][], keys);
      } else if (Array.isArray(row)) {
        values = row;
      } else {
        values = keys.map(k => (row as Record<string, unknown>)[k]);
      }

      rows.push(values);
    }

    return { keys, displayHeaders, rows };
  }

  // RowHashArray input
  if (isRowHashArray(firstRow)) {
    const hashArrays = data as [string, unknown][][];
    const keys =
      headers === true
        ? rowHashArrayToHeaders(hashArrays[0])
        : Array.isArray(headers)
          ? headers
          : null;

    const rows: unknown[][] = [];
    for (let i = 0; i < hashArrays.length; i++) {
      const row = applyRowTransform(cfg, hashArrays[i] as Row, i);
      if (row === null) {
        continue;
      }

      let values: unknown[];
      if (isRowHashArray(row)) {
        values = keys ? rowHashArrayMapByHeaders(row, keys) : rowHashArrayToValues(row);
      } else if (Array.isArray(row)) {
        values = row;
      } else {
        values = keys ? keys.map(k => (row as Record<string, unknown>)[k]) : Object.values(row);
      }

      rows.push(values);
    }

    return { keys, displayHeaders: keys, rows };
  }

  // Object input
  if (!Array.isArray(firstRow) && typeof firstRow === "object") {
    const objects = data as Record<string, unknown>[];
    const keys =
      headers === true ? Object.keys(objects[0]) : Array.isArray(headers) ? headers : null;

    const rows: unknown[][] = [];
    for (let i = 0; i < objects.length; i++) {
      const obj = applyRowTransform(cfg, objects[i] as Row, i);
      if (obj === null) {
        continue;
      }

      const values = keys
        ? keys.map(k => (obj as Record<string, unknown>)[k])
        : Object.values(obj as Record<string, unknown>);
      rows.push(values);
    }

    return { keys, displayHeaders: keys, rows };
  }

  // Array input
  const arrays = data as unknown[][];
  const keys = Array.isArray(headers) ? headers : null;

  const rows: unknown[][] = [];
  for (let i = 0; i < arrays.length; i++) {
    const row = applyRowTransform(cfg, arrays[i] as Row, i);
    if (row === null) {
      continue;
    }

    rows.push(row as unknown[]);
  }

  return { keys, displayHeaders: keys, rows };
}

// =============================================================================
// Main Formatter
// =============================================================================

/**
 * Format data as CSV string.
 *
 * Performance optimization: Builds result string directly without
 * intermediate arrays from map().join() operations.
 *
 * @example
 * ```ts
 * // Array of arrays
 * formatCsv([["a", "b"], ["1", "2"]])
 * // "a,b\n1,2"
 *
 * // Array of objects
 * formatCsv([{ name: "Alice", age: 30 }])
 * // "name,age\nAlice,30"
 *
 * // With options
 * formatCsv(data, {
 *   delimiter: ";",
 *   quoteColumns: { name: true },
 *   escapeFormulae: true,
 *   bom: true
 * })
 * ```
 */
export function formatCsv(
  data: Row[] | Record<string, unknown>[],
  options: CsvFormatOptions = {}
): string {
  const cfg = createFormatConfig(options);
  const { displayHeaders, rows } = normalizeInput(data, options, cfg);

  // Collect lines into array and join at the end.
  // V8 optimizes Array.join internally, outperforming repeated += for many rows.
  const lines: string[] = [];

  // Deduplicate headers once for both header and data rows (consistent with stream formatter)
  const effectiveHeaders = displayHeaders ? deduplicateHeaders(displayHeaders) : undefined;

  // Header row
  if (effectiveHeaders && cfg.writeHeaders) {
    lines.push(
      formatRowWithLookup(effectiveHeaders, cfg.regex, {
        quoteLookup: cfg.shouldQuoteHeader,
        delimiter: cfg.delimiter,
        headers: effectiveHeaders as string[],
        isHeader: true,
        outputRowIndex: 0,
        quoteAll: cfg.quoteAll,
        escapeFormulae: cfg.escapeFormulae,
        decimalSeparator: cfg.decimalSeparator,
        transform: undefined
      })
    );
  }

  // Data rows
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    lines.push(
      formatRowWithLookup(rows[rowIdx], cfg.regex, {
        quoteLookup: cfg.shouldQuoteColumn,
        delimiter: cfg.delimiter,
        headers: effectiveHeaders as string[] | undefined,
        isHeader: false,
        outputRowIndex: rowIdx,
        quoteAll: cfg.quoteAll,
        escapeFormulae: cfg.escapeFormulae,
        decimalSeparator: cfg.decimalSeparator,
        transform: cfg.typeTransform
      })
    );
  }

  let result = cfg.bom ? "\uFEFF" : "";
  result += lines.join(cfg.lineEnding);

  // Trailing newline
  if (lines.length > 0 && cfg.trailingNewline) {
    result += cfg.lineEnding;
  }

  return result;
}
