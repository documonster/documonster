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

import { CsvArgumentError } from "@csv/errors";
import type {
  CsvFormatRegex,
  FormatFieldContext,
  FormatRowOptions,
  FormatConfig
} from "@csv/format/config";
import { createFormatConfig } from "@csv/format/config";
import { isFormattedValue } from "@csv/format/formatted-value";
import type {
  CsvFormatOptions,
  Row,
  TypeTransformMap,
  TransformContext,
  TransformResult
} from "@csv/types";
import { checkFormatRow, checkFormatRows, describeValue } from "@csv/utils/check-options";
import { startsWithFormulaChar } from "@csv/utils/detect";
import type { DecimalSeparator } from "@csv/utils/number";
import { formatNumberForCsv } from "@csv/utils/number";
import {
  deduplicateHeaders,
  detectRowKeys,
  extractRowValues,
  processColumns
} from "@csv/utils/row";

// =============================================================================
// Type Transform Functions
// =============================================================================

/**
 * Apply type-based transform to a single value.
 * Returns the transformed result, or undefined if no transform applies.
 */
export function applyTypeTransform(
  value: unknown,
  transform: TypeTransformMap,
  ctx: TransformContext
): TransformResult {
  if (value === null || value === undefined) {
    return undefined;
  }

  if (typeof value === "boolean" && transform.boolean) {
    return transform.boolean(value, ctx);
  }
  if (value instanceof Date && transform.date) {
    return transform.date(value, ctx);
  }
  if (typeof value === "number" && transform.number) {
    return transform.number(value, ctx);
  }
  if (typeof value === "bigint" && transform.bigint) {
    return transform.bigint(value, ctx);
  }
  if (typeof value === "string" && transform.string) {
    return transform.string(value, ctx);
  }
  // Handle plain objects (not Date, not Array, not null)
  if (typeof value === "object" && !Array.isArray(value) && !(value instanceof Date)) {
    if (transform.object) {
      return transform.object(value as Record<string, unknown>, ctx);
    }
  }

  return undefined;
}

/**
 * Default type conversion to string.
 */
export function defaultToString(
  value: unknown,
  decimalSeparator: DecimalSeparator,
  where = "Csv.format"
): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number") {
    return formatNumberForCsv(value, decimalSeparator);
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      throw new CsvArgumentError(
        "data",
        `${where}: a cell holds an Invalid Date, which has no text form`
      );
    }
    return value.toISOString();
  }
  if (typeof value === "bigint") {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "true" : "false";
  }
  if (typeof value === "object") {
    // An object or array is written as JSON (`typeTransform.object` overrides it). One JSON
    // cannot write — a circular structure — used to come out as "[object Object]".
    let json: string | undefined;
    try {
      json = JSON.stringify(value);
    } catch (cause) {
      throw new CsvArgumentError(
        "data",
        `${where}: a cell holds an object JSON cannot write (a circular structure?); format it with typeTransform.object`,
        { cause }
      );
    }
    return json ?? "";
  }
  if (typeof value === "function" || typeof value === "symbol") {
    // A function used to be written as its source text and a symbol as "Symbol(x)".
    throw new CsvArgumentError(
      "data",
      `${where}: a cell holds a ${typeof value}, which has no CSV text`
    );
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
    transform,
    where
  } = ctx;

  // Apply type-based transform if provided (not for headers)
  let str: string;
  // Track if transform explicitly requested quoting control
  let transformQuoteHint: boolean | undefined;

  if (isFormattedValue(value)) {
    // `Csv.forceQuote(v)` / `Csv.preventQuote(v)` written straight into a row.
    str = value.value;
    transformQuoteHint = value.quote;
  } else if (!isHeader && transform) {
    // Create fresh context for each call to ensure safety if user stores reference
    const transformCtx: TransformContext = { column: header ?? index, index: outputRowIndex };
    const transformed = applyTypeTransform(value, transform, transformCtx);

    if (transformed === undefined || transformed === null) {
      str = defaultToString(value, decimalSeparator, where);
    } else if (isFormattedValue(transformed)) {
      // FormattedValue contains explicit quoting hint
      str = transformed.value;
      transformQuoteHint = transformed.quote;
    } else if (typeof transformed === "string") {
      str = transformed;
    } else {
      throw new CsvArgumentError(
        "options.typeTransform",
        `${where}: typeTransform must return a string, Csv.forceQuote(…), Csv.preventQuote(…), null or undefined, got ${describeValue(transformed)} for column ${header ?? index}`
      );
    }
  } else {
    str = defaultToString(value, decimalSeparator, where);
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

/** The names a header row's fields give their columns: each field's text, as it is written before any quoting. */
export function headerNames(row: unknown, cfg: FormatConfig): string[] {
  return (Array.isArray(row) ? row : []).map(value =>
    isFormattedValue(value) ? value.value : defaultToString(value, cfg.decimalSeparator, cfg.where)
  );
}

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
    transform,
    where
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
    transform,
    where
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
  if (t === null) {
    return null;
  }
  checkFormatRow(t, index, "typeTransform.row", cfg.where);
  return t;
}

/**
 * What `Csv.format` takes: rows of values (any value — a `Date`, `Csv.forceQuote(…)` — not
 * only the {@link Row} primitives), objects, or `[key, value]` lists.
 */
export type CsvFormatData =
  | readonly Row[]
  | readonly Record<string, unknown>[]
  | readonly (readonly unknown[] | Readonly<Record<string, unknown>>)[];

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
  data: CsvFormatData,
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

  // Columns config takes precedence, then explicit headers. Keyed rows (objects and
  // RowHashArrays) are otherwise mapped by the first row's keys, so a row that lists its
  // keys in another order still lands in the right columns; those keys are the header row
  // unless `headers: false`.
  let keys: string[] | null;
  let displayHeaders: string[] | null;
  if (columns && columns.length > 0) {
    const processed = processColumns(columns)!;
    keys = processed.keys;
    displayHeaders = processed.headers;
  } else if (Array.isArray(headers)) {
    keys = headers;
    displayHeaders = headers;
  } else {
    const detected = detectRowKeys(data[0]);
    keys = detected.length > 0 ? detected : null;
    displayHeaders = headers === false ? null : keys;
  }

  const rows: unknown[][] = [];
  for (let i = 0; i < data.length; i++) {
    checkFormatRow(data[i], i, undefined, cfg.where);
    const row = applyRowTransform(cfg, data[i] as Row, i);
    if (row === null) {
      continue;
    }
    rows.push(extractRowValues(row, keys));
  }

  return { keys, displayHeaders, rows };
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
 * Csv.format([["a", "b"], ["1", "2"]])
 * // "a,b\n1,2"
 *
 * // Array of objects
 * Csv.format([{ name: "Alice", age: 30 }])
 * // "name,age\nAlice,30"
 *
 * // With options
 * Csv.format(data, {
 *   delimiter: ";",
 *   quoteColumns: { name: true },
 *   escapeFormulae: true,
 *   bom: true
 * })
 * ```
 */
export function formatCsv(data: CsvFormatData, options: CsvFormatOptions = {}): string {
  const cfg = createFormatConfig(options);
  checkFormatRows(data, cfg.where);
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
        transform: undefined,
        where: cfg.where
      })
    );
  }

  // The first array row as the header row, when the rows carry their own (`firstRowIsHeader`).
  const headerFirst = options.firstRowIsHeader === true && !effectiveHeaders && rows.length > 0;
  // Its texts name the columns, as `headers` would: what `quoteHeaders` and `quoteColumns` match by name.
  const firstRowNames = headerFirst ? headerNames(rows[0], cfg) : undefined;
  if (headerFirst) {
    lines.push(
      formatRowWithLookup(rows[0], cfg.regex, {
        quoteLookup: cfg.shouldQuoteHeader,
        delimiter: cfg.delimiter,
        headers: firstRowNames,
        isHeader: true,
        outputRowIndex: 0,
        quoteAll: cfg.quoteAll,
        escapeFormulae: cfg.escapeFormulae,
        decimalSeparator: cfg.decimalSeparator,
        transform: undefined,
        where: cfg.where
      })
    );
  }

  // Data rows
  for (let rowIdx = headerFirst ? 1 : 0; rowIdx < rows.length; rowIdx++) {
    lines.push(
      formatRowWithLookup(rows[rowIdx], cfg.regex, {
        quoteLookup: cfg.shouldQuoteColumn,
        delimiter: cfg.delimiter,
        headers: firstRowNames ?? (effectiveHeaders as string[] | undefined),
        isHeader: false,
        outputRowIndex: rowIdx,
        quoteAll: cfg.quoteAll,
        escapeFormulae: cfg.escapeFormulae,
        decimalSeparator: cfg.decimalSeparator,
        transform: cfg.typeTransform,
        where: cfg.where
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
