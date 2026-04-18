/**
 * Shared helpers for native function implementations.
 *
 * Centralizes small utilities that were previously duplicated across multiple
 * function files (math, statistical, financial, text, date, engineering,
 * conditional, dynamic-array, database, lookup). None of these change runtime
 * semantics — they are the canonical extractions of the identical helpers that
 * appeared in several modules.
 */

import type {
  RuntimeValue,
  ScalarValue,
  NumberValue,
  ErrorValue,
  ArrayValue
} from "../runtime/values";
import { RVKind, BLANK, toNumberRV, topLeft } from "../runtime/values";

// ============================================================================
// Error propagation
// ============================================================================

/**
 * Return the error value if `v` (as a scalar, extracted via `topLeft`) is an
 * error; otherwise return `null`. Used by text / date / engineering functions
 * for the standard "propagate first-arg error" pattern.
 */
export function checkError(v: RuntimeValue): ErrorValue | null {
  const s = topLeft(v);
  return s.kind === RVKind.Error ? s : null;
}

// ============================================================================
// Argument coercion
// ============================================================================

/**
 * Coerce a single RuntimeValue argument to a number. Applies `topLeft` first
 * so that a 1×1 (or arbitrary) array yields its top-left scalar before
 * numeric coercion. Matches the semantics formerly found as `argToNumber` in
 * `math.ts` and `numArg(args, idx)` in `statistical.ts`.
 */
export function argToNumber(arg: RuntimeValue): NumberValue | ErrorValue {
  const s = topLeft(arg);
  if (s.kind === RVKind.Error) {
    return s;
  }
  return toNumberRV(s);
}

// ============================================================================
// Flattening
// ============================================================================

/**
 * Flatten a list of arguments into a sequence of numeric values (or errors).
 *
 * Array arguments contribute only their `Number` and `Error` cells — booleans,
 * strings, and blanks inside arrays are skipped (Excel range semantics).
 * Direct scalar arguments are coerced via `toNumberRV`, except for direct
 * `Blank` scalars which are skipped (Excel aggregate semantics).
 *
 * Returns a list of `NumberValue | ErrorValue`. Callers that need raw
 * `number[]` after an error check should map `.value` themselves.
 */
export function flattenNumbers(args: RuntimeValue[]): (NumberValue | ErrorValue)[] {
  const result: (NumberValue | ErrorValue)[] = [];
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            result.push(cell);
          } else if (cell.kind === RVKind.Number) {
            result.push(cell);
          }
          // Skip booleans, strings, blanks in array context (Excel behavior).
        }
      }
    } else if (arg.kind === RVKind.Error) {
      result.push(arg);
    } else if (arg.kind !== RVKind.Blank) {
      result.push(toNumberRV(arg));
    }
    // Skip blanks for direct scalar args.
  }
  return result;
}

/**
 * Flatten all cells from the arguments into a flat list of ScalarValue,
 * preserving every cell (including blanks, errors, booleans, strings).
 * Direct scalar arguments are projected via `topLeft`.
 */
export function flattenAll(args: RuntimeValue[]): ScalarValue[] {
  const result: ScalarValue[] = [];
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          result.push(cell);
        }
      }
    } else {
      result.push(topLeft(arg));
    }
  }
  return result;
}

/**
 * Return the first `ErrorValue` in a list of `NumberValue | ErrorValue`,
 * or `null` if none is present.
 */
export function firstError(values: readonly (NumberValue | ErrorValue)[]): ErrorValue | null {
  for (const v of values) {
    if (v.kind === RVKind.Error) {
      return v;
    }
  }
  return null;
}

// ============================================================================
// Array helpers
// ============================================================================

/**
 * Narrow a RuntimeValue to an ArrayValue, returning `null` if it is not an
 * array. Used by conditional / database / lookup / dynamic-array families.
 */
export function asArray(v: RuntimeValue): ArrayValue | null {
  return v.kind === RVKind.Array ? v : null;
}

/**
 * Safe cell accessor for ArrayValue — returns `BLANK` for out-of-bounds (r, c).
 */
export function getCell(arr: ArrayValue, r: number, c: number): ScalarValue {
  if (r < arr.height && c < arr.width) {
    return arr.rows[r][c];
  }
  return BLANK;
}

// ============================================================================
// Excel wildcard helpers
// ============================================================================

/**
 * Return `true` if `s` contains an unescaped `*` or `?`. Excel uses `~` as
 * the escape character, so `~*` and `~?` are literals while `*` and `?` at
 * any other position are wildcards. `~~` is an escaped tilde.
 *
 * Centralised here so every wildcard-consuming function (SEARCH, MATCH,
 * XLOOKUP, SUMIF/COUNTIF/…) agrees on whether a criterion should trigger
 * the wildcard code path.
 */
export function hasUnescapedWildcard(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "~" && i + 1 < s.length) {
      i++; // skip the escaped char
      continue;
    }
    if (ch === "*" || ch === "?") {
      return true;
    }
  }
  return false;
}

/** Escape a single character for use in a regex literal. */
function regexEscapeChar(ch: string): string {
  return /[.*+^${}()|[\]\\?]/.test(ch) ? "\\" + ch : ch;
}

/**
 * Convert an Excel wildcard pattern to a JavaScript regex source. Rules:
 *   `*`     → `.*`
 *   `?`     → `.`
 *   `~*`    → literal `*`
 *   `~?`    → literal `?`
 *   `~~`    → literal `~`
 *   `~x`    → literal `x` (any other character after `~` is treated literally,
 *             matching Excel's tolerant behaviour)
 *   everything else → regex-escaped literal
 *
 * Callers typically wrap the result in `^…$` and use the `i` flag for
 * case-insensitive matching.
 */
export function excelWildcardToRegex(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "~" && i + 1 < s.length) {
      out += regexEscapeChar(s[i + 1]);
      i++;
      continue;
    }
    if (ch === "*") {
      out += ".*";
      continue;
    }
    if (ch === "?") {
      out += ".";
      continue;
    }
    out += regexEscapeChar(ch);
  }
  return out;
}

/**
 * Strip `~` escape characters from an Excel criteria string so the remaining
 * text can be used for a literal comparison. `~*` → `*`, `~?` → `?`,
 * `~~` → `~`, etc. Only used when the caller has already determined that the
 * pattern contains no unescaped wildcards.
 */
export function unescapeExcelWildcard(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "~" && i + 1 < s.length) {
      out += s[i + 1];
      i++;
      continue;
    }
    out += ch;
  }
  return out;
}
