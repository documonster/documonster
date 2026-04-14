/**
 * CSV Dynamic Typing - Automatic Type Conversion
 *
 * Functions for converting CSV string values to appropriate JavaScript types.
 * Supports boolean, number, null detection with customizable per-column config.
 */

import type { DynamicTypingConfig, CastDateConfig } from "@csv/types";
import { DateParser } from "@utils/datetime";

// =============================================================================
// Pre-compiled Regex Constants
// =============================================================================

/**
 * Pre-compiled regex for valid number format detection.
 * Matches integers, decimals, and scientific notation.
 * Pre-compiling avoids regex compilation overhead in the hot path.
 */
const NUMERIC_REGEX = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

// Singleton date parser for ISO formats (created lazily)
let isoDateParser: DateParser | null = null;

/**
 * Get or create the ISO date parser singleton
 */
function getIsoDateParser(): DateParser {
  if (!isoDateParser) {
    isoDateParser = DateParser.iso();
  }
  return isoDateParser;
}

/**
 * Try to parse a string as an ISO date.
 * Returns the Date if successful, or null if not a valid date.
 *
 * Supported formats:
 * - YYYY-MM-DD
 * - YYYY-MM-DDTHH:mm:ss
 * - YYYY-MM-DD HH:mm:ss
 * - YYYY-MM-DDTHH:mm:ssZ
 * - YYYY-MM-DDTHH:mm:ss.SSSZ
 * - YYYY-MM-DDTHH:mm:ss+HH:mm
 */
export function tryParseDate(value: string): Date | null {
  if (!value || value.length < 10) {
    return null;
  }
  return getIsoDateParser().parse(value);
}

/**
 * Check if castDate config enables date parsing for a column
 */
export function shouldCastDate(
  castDate: CastDateConfig | undefined,
  columnName: string | number | undefined
): boolean {
  if (!castDate) {
    return false;
  }
  if (castDate === true) {
    return true;
  }
  if (Array.isArray(castDate) && typeof columnName === "string") {
    return castDate.includes(columnName);
  }
  return false;
}

// =============================================================================
// Core Conversion
// =============================================================================

/**
 * Check if a charCode matches a lowercase letter (case-insensitive).
 * @param code - The charCode to check
 * @param lowercaseCode - The lowercase letter's charCode to match against
 * @returns true if code matches (case-insensitive)
 */
function isCharEqualIgnoreCase(code: number, lowercaseCode: number): boolean {
  // Lowercase letters are 32 higher than uppercase in ASCII
  // e.g., 'a' = 97, 'A' = 65, difference = 32
  return code === lowercaseCode || code === lowercaseCode - 32;
}

/**
 * Convert a string value to its appropriate JavaScript type.
 * Used internally by dynamicTyping feature.
 *
 * Conversion rules:
 * - Empty string → "" (unchanged)
 * - "true"/"TRUE"/"True" → true
 * - "false"/"FALSE"/"False" → false
 * - "null"/"NULL" → null
 * - Numeric strings → number (int or float)
 * - Everything else → original string
 *
 * Special cases:
 * - Leading zeros (e.g., "007") → preserved as string (for zip codes, IDs)
 * - "Infinity", "-Infinity", "NaN" → corresponding number values
 */
export function convertValue(value: string): string | number | boolean | null {
  const len = value.length;

  // Empty string stays empty (not converted to null)
  if (len === 0) {
    return "";
  }

  // Fast path: use charCodeAt for quick first-character checks
  const firstChar = value.charCodeAt(0);

  // Boolean/null detection using charCode comparison (avoids toLowerCase allocation)
  // 't' = 116, 'T' = 84, 'r' = 114, 'R' = 82, 'u' = 117, 'U' = 85, 'e' = 101, 'E' = 69
  // 'f' = 102, 'F' = 70, 'a' = 97, 'A' = 65, 'l' = 108, 'L' = 76, 's' = 115, 'S' = 83
  // 'n' = 110, 'N' = 78
  if (len === 4) {
    if (
      (firstChar === 116 || firstChar === 84) && // t/T
      isCharEqualIgnoreCase(value.charCodeAt(1), 114) && // r/R
      isCharEqualIgnoreCase(value.charCodeAt(2), 117) && // u/U
      isCharEqualIgnoreCase(value.charCodeAt(3), 101) // e/E
    ) {
      return true;
    }
    if (
      (firstChar === 110 || firstChar === 78) && // n/N
      isCharEqualIgnoreCase(value.charCodeAt(1), 117) && // u/U
      isCharEqualIgnoreCase(value.charCodeAt(2), 108) && // l/L
      isCharEqualIgnoreCase(value.charCodeAt(3), 108) // l/L
    ) {
      return null;
    }
  } else if (
    len === 5 &&
    (firstChar === 102 || firstChar === 70) && // f/F
    isCharEqualIgnoreCase(value.charCodeAt(1), 97) && // a/A
    isCharEqualIgnoreCase(value.charCodeAt(2), 108) && // l/L
    isCharEqualIgnoreCase(value.charCodeAt(3), 115) && // s/S
    isCharEqualIgnoreCase(value.charCodeAt(4), 101) // e/E
  ) {
    return false;
  }

  // Number detection - only if first char could start a number
  // '-' = 45, '.' = 46, '0'-'9' = 48-57, 'I' = 73, 'N' = 78
  if (
    (firstChar >= 48 && firstChar <= 57) || // 0-9
    firstChar === 45 || // -
    firstChar === 46 || // .
    firstChar === 73 || // I (Infinity)
    firstChar === 78 // N (NaN)
  ) {
    // Check for trailing whitespace - skip number conversion
    const lastChar = value.charCodeAt(len - 1);
    // Space = 32, Tab = 9, \n = 10, \r = 13
    if (lastChar <= 32) {
      return value;
    }

    // Special numeric values
    if (value === "Infinity") {
      return Infinity;
    }
    if (value === "-Infinity") {
      return -Infinity;
    }
    if (value === "NaN") {
      return NaN;
    }

    // Preserve leading zeros (important for zip codes, phone numbers, IDs)
    // Check for pattern like "007" but allow "0" and "0.xxx"
    if (firstChar === 48 && len > 1) {
      // starts with '0'
      const secondChar = value.charCodeAt(1);
      // If second char is a digit (not '.'), preserve as string
      if (secondChar >= 48 && secondChar <= 57) {
        return value;
      }
    }
    // Handle negative leading zeros like "-007"
    if (firstChar === 45 && len > 2 && value.charCodeAt(1) === 48) {
      // starts with '-0'
      const thirdChar = value.charCodeAt(2);
      if (thirdChar >= 48 && thirdChar <= 57) {
        return value;
      }
    }

    // Check for valid number format (avoid converting "123abc" or "1.2.3")
    if (NUMERIC_REGEX.test(value)) {
      const num = Number(value);
      if (!isNaN(num)) {
        return num;
      }
    }
  }

  // Default: keep as string
  return value;
}

// =============================================================================
// Type Guards
// =============================================================================

/**
 * Type guard to check if dynamicTyping config has custom converter function
 */
function isCustomConverter(
  config: boolean | ((value: string) => unknown)
): config is (value: string) => unknown {
  return typeof config === "function";
}

// =============================================================================
// Application Functions
// =============================================================================

/**
 * Apply dynamic typing to a single field value
 *
 * @param value - The string value to convert
 * @param columnConfig - Column-specific config (true, false, or custom function)
 * @returns Converted value
 */
export function applyDynamicTyping(
  value: string,
  columnConfig: boolean | ((value: string) => unknown)
): unknown {
  if (columnConfig === false) {
    return value;
  }

  if (isCustomConverter(columnConfig)) {
    return columnConfig(value);
  }

  // columnConfig === true → use default conversion
  return convertValue(value);
}

// =============================================================================
// Core Value Conversion Helper
// =============================================================================

/**
 * Apply dynamic typing and/or date casting to a single value.
 * Unified helper used by both object and array row processing.
 *
 * @param value - The string value to convert
 * @param columnName - Column identifier (string for objects, can be used for per-column config)
 * @param dynamicTyping - DynamicTyping configuration
 * @param castDate - CastDate configuration
 * @returns Converted value
 */
function convertSingleValue(
  value: string,
  columnName: string | undefined,
  dynamicTyping: DynamicTypingConfig,
  castDate: CastDateConfig | undefined
): unknown {
  // Try date parsing first if castDate is enabled for this column
  if (shouldCastDate(castDate, columnName)) {
    const dateValue = tryParseDate(value);
    if (dateValue !== null) {
      return dateValue;
    }
  }

  // Apply dynamic typing based on config type
  if (dynamicTyping === true) {
    return convertValue(value);
  }

  if (dynamicTyping === false) {
    return value;
  }

  // Per-column configuration
  if (columnName === undefined) {
    return value;
  }

  const config = (dynamicTyping as Record<string, boolean | ((value: string) => unknown)>)[
    columnName
  ];
  if (config === undefined) {
    return value;
  }

  return applyDynamicTyping(value, config);
}

// =============================================================================
// Row Conversion Functions
// =============================================================================

/**
 * Apply dynamic typing to an entire row (object form).
 *
 * Performance: Converts values IN PLACE to avoid allocating a new object.
 * The input object is mutated and returned with converted values.
 *
 * @param row - Row object with string values (will be mutated)
 * @param dynamicTyping - DynamicTyping configuration
 * @param castDate - CastDate configuration for date parsing
 * @returns The same row object with converted values
 */
export function applyDynamicTypingToRow(
  row: Record<string, string>,
  dynamicTyping: DynamicTypingConfig,
  castDate?: CastDateConfig
): Record<string, unknown> {
  if (dynamicTyping === false && !castDate) {
    // No conversion - return as-is (fast path)
    return row;
  }

  // Convert in place - mutate the input object directly
  for (const key in row) {
    if (Object.hasOwn(row, key)) {
      (row as Record<string, unknown>)[key] = convertSingleValue(
        row[key],
        key,
        dynamicTyping,
        castDate
      );
    }
  }
  return row as Record<string, unknown>;
}

/**
 * Apply dynamic typing to an array row
 *
 * @param row - Row array with string values
 * @param headers - Header names (for per-column config lookup)
 * @param dynamicTyping - DynamicTyping configuration
 * @param castDate - CastDate configuration for date parsing
 * @returns New row array with converted values
 */
export function applyDynamicTypingToArrayRow(
  row: string[],
  headers: string[] | null,
  dynamicTyping: DynamicTypingConfig,
  castDate?: CastDateConfig
): unknown[] {
  if (dynamicTyping === false && !castDate) {
    // No conversion - return as-is (fast path)
    return row;
  }

  // Per-column config requires headers
  if (dynamicTyping !== true && dynamicTyping !== false && !headers) {
    return row;
  }

  return row.map((value, index) => {
    const columnName = headers?.[index];
    return convertSingleValue(value, columnName, dynamicTyping, castDate);
  });
}
