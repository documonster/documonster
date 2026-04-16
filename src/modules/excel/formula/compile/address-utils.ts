/**
 * Formula Address Utilities — Shared helpers for address parsing and conversion.
 *
 * These pure functions are used across the compile, runtime, and materialize
 * layers. Centralizing them here eliminates duplication and ensures consistent
 * behavior.
 *
 * ## Design Constraint
 *
 * This module must NOT import from `@excel/utils/col-cache` or any live
 * workbook type. All functions are self-contained.
 */

// ============================================================================
// Column Letter ↔ Number
// ============================================================================

/**
 * Convert a column letter string (e.g. "A", "AA", "XFD") to a 1-based number.
 *
 * - `"A"` → 1
 * - `"Z"` → 26
 * - `"AA"` → 27
 * - `"XFD"` → 16384
 */
export function colLetterToNumber(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64); // 'A' = 65
  }
  return result;
}

// ============================================================================
// Defined Name Range Parsing
// ============================================================================

/**
 * A parsed defined-name range reference.
 */
export interface ParsedNameRange {
  readonly sheet: string;
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

const DEFINED_NAME_RANGE_RE =
  /^(?:'([^']*(?:''[^']*)*)'|([^!]+))!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/;

/**
 * Parse a defined-name range string like `"Sheet1!$A$1:$B$2"` or `"'Sheet Name'!$C$3"`
 * into a `ParsedNameRange`. Returns `null` if the format is unrecognized.
 */
export function parseDefinedNameRange(rangeStr: string): ParsedNameRange | null {
  const m = DEFINED_NAME_RANGE_RE.exec(rangeStr);
  if (!m) {
    return null;
  }
  // Sheet name: quoted (group 1, with '' unescaping) or unquoted (group 2)
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const startCol = colLetterToNumber(m[3]);
  const startRow = parseInt(m[4], 10);
  const endCol = m[5] ? colLetterToNumber(m[5]) : startCol;
  const endRow = m[6] ? parseInt(m[6], 10) : startRow;
  return { sheet, startRow, startCol, endRow, endCol };
}

// ============================================================================
// Simple Cell Address Parsing
// ============================================================================

/**
 * A parsed simple cell address (no sheet name).
 */
export interface ParsedAddress {
  readonly row: number;
  readonly col: number;
}

/**
 * Parse a simple cell address like `"A1"`, `"$B$2"`, or `"AA100"`.
 * Dollar signs are stripped. Returns `null` on parse failure.
 */
export function parseSimpleAddress(addr: string): ParsedAddress | null {
  const m = /^([A-Z]+)(\d+)$/.exec(addr.replace(/\$/g, ""));
  if (!m) {
    return null;
  }
  return {
    row: parseInt(m[2], 10),
    col: colLetterToNumber(m[1])
  };
}

// ============================================================================
// Ref Range Parsing (e.g. "A1:B2")
// ============================================================================

/**
 * A parsed rectangular range (no sheet name).
 */
export interface ParsedRefRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/**
 * Parse a cell range string like `"A1:B2"` into a `ParsedRefRange`.
 * Returns `null` if the format is unrecognized.
 */
export function parseRefRange(ref: string): ParsedRefRange | null {
  const parts = ref.split(":");
  if (parts.length !== 2) {
    return null;
  }
  const tl = parseSimpleAddress(parts[0]);
  const br = parseSimpleAddress(parts[1]);
  if (!tl || !br) {
    return null;
  }
  return {
    top: Math.min(tl.row, br.row),
    left: Math.min(tl.col, br.col),
    bottom: Math.max(tl.row, br.row),
    right: Math.max(tl.col, br.col)
  };
}
