/**
 * CSV Row Utilities
 *
 * Helper functions for working with different row formats:
 * - RowHashArray: Array of [key, value] tuples, e.g., [["name", "John"], ["age", "30"]]
 * - RowMap: Record<string, any> object format
 * - RowArray: Simple string[] array format
 *
 * Key functions:
 * - isRowHashArray(): Type guard for RowHashArray detection
 * - extractRowValues(): Extract values from any row format
 * - detectRowKeys(): Detect header keys from a row
 * - deduplicateHeaders(): Handle duplicate column names
 * - processColumns(): Process column configuration
 */

// Re-export types from central types.ts to avoid duplication
import type { ColumnConfig, HeaderArray, RowHashArray } from "@csv/types";
import { isSafeDynamicKey } from "@utils/object";
export type { HeaderArray, RowHashArray } from "@csv/types";

// =============================================================================
// RowHashArray Utilities
// =============================================================================

/**
 * Check if a row is a RowHashArray (array of [key, value] tuples)
 */
export function isRowHashArray(row: unknown): row is RowHashArray {
  if (!Array.isArray(row) || row.length === 0) {
    return false;
  }
  // Check if first element is a 2-element array with string key
  const first = row[0];
  return Array.isArray(first) && first.length === 2 && typeof first[0] === "string";
}

/**
 * Convert RowHashArray to RowMap
 * Note: Manual loop is ~4x faster than Object.fromEntries
 */
export function rowHashArrayToMap<V = any>(row: RowHashArray<V>): Record<string, V> {
  const obj: Record<string, V> = Object.create(null) as Record<string, V>;
  for (const [key, value] of row) {
    if (isSafeDynamicKey(key)) {
      obj[key] = value;
    }
  }
  return obj;
}

/**
 * Convert RowHashArray to values array (preserving order)
 */
export function rowHashArrayToValues<V = any>(row: RowHashArray<V>): V[] {
  return row.map(([, value]) => value);
}

/**
 * Get headers from RowHashArray
 */
export function rowHashArrayToHeaders(row: RowHashArray): string[] {
  return row.map(([key]) => key);
}

/**
 * Get value by key from RowHashArray (returns undefined if not found)
 * More efficient than creating a full map when you need only specific values
 */
export function rowHashArrayGet<V = any>(row: RowHashArray<V>, key: string): V | undefined {
  for (const [k, v] of row) {
    if (k === key) {
      return v;
    }
  }
  return undefined;
}

/**
 * Map RowHashArray values according to header order
 * Optimized: builds values array in single pass without intermediate object
 */
export function rowHashArrayMapByHeaders<V = any>(
  row: RowHashArray<V>,
  headers: string[]
): (V | undefined)[] {
  // For small headers array, linear search per header is faster than building a map
  // For larger headers (>10), build a map once
  if (headers.length <= 10) {
    return headers.map(h => rowHashArrayGet(row, h));
  }
  const map = rowHashArrayToMap(row);
  return headers.map(h => map[h]);
}

/**
 * Extract values from a row (array, object, or RowHashArray) in consistent order.
 * This is the unified function for row value extraction used by both
 * batch (formatCsv) and streaming (CsvFormatterStream) code paths.
 *
 * @param row - The row data (array, object, or RowHashArray)
 * @param keys - Optional key order for object/RowHashArray rows
 * @returns Array of values in the specified key order (or natural order if no keys)
 */
export function extractRowValues(row: unknown, keys: string[] | null | undefined): unknown[] {
  if (isRowHashArray(row)) {
    return keys ? rowHashArrayMapByHeaders(row, keys) : rowHashArrayToValues(row);
  }
  if (Array.isArray(row)) {
    return row;
  }
  if (row !== null && typeof row === "object") {
    return keys ? keys.map(key => (row as Record<string, unknown>)[key]) : Object.values(row);
  }
  // Primitive value: wrap in array
  return [row];
}

/**
 * Auto-detect keys (headers) from a row based on its type.
 * Returns empty array for arrays or primitive values.
 *
 * @param row - The row data to detect keys from
 * @returns Array of string keys, or empty array if not applicable
 */
export function detectRowKeys(row: unknown): string[] {
  if (isRowHashArray(row)) {
    return rowHashArrayToHeaders(row);
  }
  if (!Array.isArray(row) && row !== null && typeof row === "object") {
    return Object.keys(row);
  }
  return [];
}

// =============================================================================
// Header Utilities
// =============================================================================

/**
 * Deduplicate headers by appending suffix to duplicates.
 * Example: ["A", "B", "A", "A"] → ["A", "B", "A_1", "A_2"]
 *
 * @param headers - Original header array
 * @returns New array with unique header names
 */
export function deduplicateHeaders(headers: HeaderArray): HeaderArray {
  return deduplicateHeadersWithRenames(headers).headers;
}

export function deduplicateHeadersWithRenames(headers: HeaderArray): {
  headers: HeaderArray;
  renamedHeaders: Record<string, string> | null;
} {
  const headerCount = new Map<string, number>();
  const usedHeaders = new Set<string>();
  // Reserve all original header names so we don't generate a rename that
  // collides with a header that appears later in the row.
  const reservedHeaders = new Set<string>();
  const result: HeaderArray = [];
  const renamedHeaders: Record<string, string> = {};

  let hasRenames = false;
  let emptyHeaderCount = 0;

  for (const header of headers) {
    if (header !== null && header !== undefined && header !== "") {
      reservedHeaders.add(header);
    }
  }

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (header === null || header === undefined) {
      result.push(header);
      continue;
    }

    // Handle empty string headers by generating placeholder names
    if (header === "") {
      let placeholder = `_column_${i}`;
      // Ensure placeholder doesn't collide with existing headers
      while (usedHeaders.has(placeholder) || reservedHeaders.has(placeholder)) {
        placeholder = `_column_${i}_${emptyHeaderCount++}`;
      }
      usedHeaders.add(placeholder);
      result.push(placeholder);
      renamedHeaders[placeholder] = "";
      hasRenames = true;
      continue;
    }

    if (!usedHeaders.has(header)) {
      usedHeaders.add(header);
      headerCount.set(header, 1);
      result.push(header);
      continue;
    }

    // Duplicate: find a unique suffix, avoiding collisions with already-present headers
    let suffix = headerCount.get(header) ?? 1;
    let candidate = `${header}_${suffix}`;
    while (usedHeaders.has(candidate) || reservedHeaders.has(candidate)) {
      suffix++;
      candidate = `${header}_${suffix}`;
    }

    headerCount.set(header, suffix + 1);
    usedHeaders.add(candidate);
    result.push(candidate);
    renamedHeaders[candidate] = header;
    hasRenames = true;
  }

  return { headers: result, renamedHeaders: hasRenames ? renamedHeaders : null };
}

// =============================================================================
// Column Utilities
// =============================================================================

/**
 * Process columns configuration to extract keys and headers.
 * Returns null if columns is empty or undefined.
 *
 * This function is used by both formatCsv (batch) and CsvFormatterStream (streaming)
 * to normalize column configuration into separate key/header arrays.
 *
 * @param columns - Column configuration array (string names or ColumnConfig objects)
 * @returns Object with keys (data access) and headers (output names), or null if empty
 *
 * @example
 * ```ts
 * processColumns(['name', { key: 'age', header: 'Age (years)' }])
 * // { keys: ['name', 'age'], headers: ['name', 'Age (years)'] }
 * ```
 */
export function processColumns(
  columns: (string | ColumnConfig)[] | undefined
): { keys: string[]; headers: string[] } | null {
  if (!columns || columns.length === 0) {
    return null;
  }
  const keys = columns.map(c => (typeof c === "string" ? c : c.key));
  const headers = columns.map(c => (typeof c === "string" ? c : (c.header ?? c.key)));
  return { keys, headers };
}

// =============================================================================
// Row Validation Utilities
// =============================================================================

/** Pre-compiled regex for non-whitespace detection */
const NON_WHITESPACE_REGEX = /\S/;

/**
 * Check if a row should be skipped as empty.
 * When `shouldSkipEmpty` is "greedy", whitespace-only rows also count as empty.
 *
 * @param row - The row to check
 * @param shouldSkipEmpty - true, false, or "greedy"
 * @returns true if the row should be skipped
 */
export function isEmptyRow(row: string[], shouldSkipEmpty: boolean | "greedy"): boolean {
  if (!shouldSkipEmpty) {
    return false;
  }
  if (shouldSkipEmpty === "greedy") {
    // Greedy: whitespace-only fields count as empty
    for (const field of row) {
      if (NON_WHITESPACE_REGEX.test(field)) {
        return false;
      }
    }
  } else {
    // Non-greedy: only truly empty strings count as empty
    for (const field of row) {
      if (field !== "") {
        return false;
      }
    }
  }
  return true;
}

/**
 * Check if all values in a row are empty strings.
 * Used by skipRecordsWithEmptyValues option.
 *
 * @param row - The row to check
 * @returns true if all fields are empty strings
 */
export function hasAllEmptyValues(row: string[]): boolean {
  return isEmptyRow(row, true);
}
