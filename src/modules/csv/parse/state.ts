/**
 * CSV Parse State
 *
 * Defines the ParseState type and factory function for creating
 * mutable parsing state. Also includes field building operations.
 */

import type { ParseConfig } from "@csv/parse/config";
import { processHeaders } from "@csv/parse/helpers";
import type { HeaderArray } from "@csv/types";

// =============================================================================
// Types
// =============================================================================

/**
 * Mutable parsing state - shared between sync and streaming parsers
 */
export interface ParseState {
  // Position tracking
  lineNumber: number;

  // Data row tracking
  dataRowCount: number;
  skippedDataRows: number;
  truncated: boolean;

  // Header state
  headerRow: HeaderArray | null;
  originalHeaders: HeaderArray | null;
  useHeaders: boolean;
  headerRowProcessed: boolean;
  renamedHeadersForMeta: Record<string, string> | null;

  // Info tracking (for info/raw options)
  currentRowStartLine: number;
  currentRowStartOffset: number;
  /** Quoted status per field. May be a shared readonly array - copy before modifying. */
  currentRowQuoted: readonly boolean[];
  currentRawRow: string;
}

// =============================================================================
// State Factory
// =============================================================================

/**
 * Create initial parse state with optional header configuration
 */
export function createParseState(
  config: Pick<ParseConfig, "headers" | "groupColumnsByName" | "infoOption" | "rawOption">
): ParseState {
  const state: ParseState = {
    lineNumber: 0,
    dataRowCount: 0,
    skippedDataRows: 0,
    truncated: false,
    headerRow: null,
    originalHeaders: null,
    useHeaders: false,
    headerRowProcessed: false,
    renamedHeadersForMeta: null,
    currentRowStartLine: config.infoOption ? 1 : 0,
    currentRowStartOffset: 0,
    currentRowQuoted: [],
    currentRawRow: ""
  };

  // Determine header mode
  const { headers, groupColumnsByName } = config;
  if (headers === true) {
    state.useHeaders = true;
  } else if (Array.isArray(headers)) {
    const result = processHeaders([], { headers: headers as string[], groupColumnsByName }, null);
    if (result) {
      state.headerRow = result.headers;
      state.originalHeaders = result.originalHeaders;
      state.renamedHeadersForMeta = result.renamedHeaders;
    }
    state.useHeaders = true;
    state.headerRowProcessed = true;
  } else if (typeof headers === "function") {
    state.useHeaders = true;
  }

  return state;
}

// =============================================================================
// Info State Management
// =============================================================================

/**
 * Reset info state for next row
 */
export function resetInfoState(
  state: ParseState,
  trackInfo: boolean,
  trackRaw: boolean,
  nextLine: number,
  nextOffset: number
): void {
  if (trackInfo) {
    state.currentRowQuoted = [];
    state.currentRowStartLine = nextLine;
    state.currentRowStartOffset = nextOffset;
  }
  if (trackRaw) {
    state.currentRawRow = "";
  }
}

// =============================================================================
// Performance Optimization: Shared False Array for Fast Mode
// =============================================================================

/**
 * Pre-allocated frozen array of false values for fast mode quoted tracking.
 * In fast mode (no quote detection), all fields are unquoted, so we can
 * return a shared reference instead of allocating per row.
 *
 * IMPORTANT: This array is frozen and must NOT be modified.
 * Callers should copy if they need to store/modify the values.
 */
const SHARED_FALSE_ARRAY_SIZE = 256;
const SHARED_FALSE_ARRAY: readonly boolean[] = Object.freeze(
  new Array(SHARED_FALSE_ARRAY_SIZE).fill(false) as boolean[]
);

/**
 * Get a shared array of false values for unquoted field tracking.
 * Returns a frozen shared reference for common cases to avoid per-row allocation.
 *
 * IMPORTANT: The returned array must NOT be modified. If you need to store
 * the values, make a copy: `[...getUnquotedArray(n)]` or `.slice(0, n)`.
 *
 * @param length - Number of fields in the row
 * @returns Shared frozen array (for length <= 256) or new array (for larger rows)
 */
export function getUnquotedArray(length: number): readonly boolean[] {
  if (length <= SHARED_FALSE_ARRAY_SIZE) {
    // Return shared reference - caller must not modify
    return SHARED_FALSE_ARRAY;
  }
  // Fall back to creating new array for very wide rows
  return new Array(length).fill(false);
}
