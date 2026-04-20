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
 * Streaming fold over numeric arguments.
 *
 * Same selection rules as `flattenNumbers` (array cells contribute only
 * Number/Error; direct scalar coercion via `toNumberRV`; blanks dropped),
 * but the caller's `onNumber` callback fires inline — no intermediate
 * array is allocated. On the first error encountered the scan short-
 * circuits and returns that error.
 *
 * Returns:
 *   - `null` when iteration finished without encountering an error, or
 *   - the `ErrorValue` that aborted the scan.
 *
 * Prefer this over `flattenNumbers` + `firstError` + manual loop in hot
 * aggregates (SUM / AVERAGE / MIN / MAX / …). The allocation saved is
 * one `NumberValue | ErrorValue` array per invocation — meaningful
 * when the engine sums tens of thousands of cells.
 */
export function forEachNumber(
  args: readonly RuntimeValue[],
  onNumber: (n: number) => void
): ErrorValue | null {
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            return cell;
          }
          if (cell.kind === RVKind.Number) {
            onNumber(cell.value);
          }
          // Booleans, strings, blanks inside arrays are skipped.
        }
      }
    } else if (arg.kind === RVKind.Error) {
      return arg;
    } else if (arg.kind !== RVKind.Blank) {
      const n = toNumberRV(arg);
      if (n.kind === RVKind.Error) {
        return n;
      }
      onNumber(n.value);
    }
    // Direct blank scalars are dropped.
  }
  return null;
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

/**
 * Replace every cell marked by the array's `subtotalMask` with BLANK so an
 * outer SUBTOTAL/AGGREGATE call does not double-count the inner aggregate's
 * result. Only ArrayValue args carrying a mask are rewritten; scalars and
 * arrays without masks pass through unchanged.
 *
 * Excel behavior: SUBTOTAL and AGGREGATE deliberately skip any cell whose
 * own formula is itself SUBTOTAL or AGGREGATE — this is how the classic
 * "totals row inside a filtered range" case works without double-counting.
 */
export function stripSubtotalMaskedCells(args: RuntimeValue[]): RuntimeValue[] {
  let needsCopy = false;
  for (const arg of args) {
    if (arg.kind === RVKind.Array && arg.subtotalMask) {
      needsCopy = true;
      break;
    }
  }
  if (!needsCopy) {
    return args;
  }
  const out: RuntimeValue[] = [];
  for (const arg of args) {
    if (arg.kind !== RVKind.Array || !arg.subtotalMask) {
      out.push(arg);
      continue;
    }
    const mask = arg.subtotalMask;
    const newRows: ScalarValue[][] = [];
    for (let r = 0; r < arg.height; r++) {
      const srcRow = arg.rows[r];
      const maskRow = mask[r];
      const newRow: ScalarValue[] = new Array<ScalarValue>(arg.width);
      for (let c = 0; c < arg.width; c++) {
        newRow[c] = maskRow?.[c] ? BLANK : srcRow[c];
      }
      newRows.push(newRow);
    }
    // Drop the mask on the rewritten array — masked cells are already BLANK.
    // Preserve hiddenRowMask so downstream SUBTOTAL 1xx / AGGREGATE opt
    // 5/7 handling still applies to rows whose visibility was recorded.
    out.push({
      kind: RVKind.Array,
      rows: newRows,
      height: arg.height,
      width: arg.width,
      ...(arg.originRow !== undefined
        ? { originRow: arg.originRow, originCol: arg.originCol }
        : {}),
      ...(arg.hiddenRowMask ? { hiddenRowMask: arg.hiddenRowMask } : {})
    } satisfies ArrayValue);
  }
  return out;
}

/**
 * Replace every cell in a hidden row with BLANK so aggregate functions
 * drop them during flattening. Callers: SUBTOTAL's 1xx-variant codes
 * (101-111) and AGGREGATE with option 5 or 7.
 *
 * Only ArrayValue args carrying a hiddenRowMask are rewritten.
 */
export function stripHiddenRowCells(args: RuntimeValue[]): RuntimeValue[] {
  let needsCopy = false;
  for (const arg of args) {
    if (arg.kind === RVKind.Array && arg.hiddenRowMask) {
      needsCopy = true;
      break;
    }
  }
  if (!needsCopy) {
    return args;
  }
  const out: RuntimeValue[] = [];
  for (const arg of args) {
    if (arg.kind !== RVKind.Array || !arg.hiddenRowMask) {
      out.push(arg);
      continue;
    }
    const mask = arg.hiddenRowMask;
    const newRows: ScalarValue[][] = [];
    for (let r = 0; r < arg.height; r++) {
      if (mask[r]) {
        const blankRow: ScalarValue[] = new Array<ScalarValue>(arg.width).fill(BLANK);
        newRows.push(blankRow);
      } else {
        newRows.push(arg.rows[r].slice());
      }
    }
    // Drop hiddenRowMask on rewritten output; preserve subtotalMask if any.
    out.push({
      kind: RVKind.Array,
      rows: newRows,
      height: arg.height,
      width: arg.width,
      ...(arg.originRow !== undefined
        ? { originRow: arg.originRow, originCol: arg.originCol }
        : {}),
      ...(arg.subtotalMask ? { subtotalMask: arg.subtotalMask } : {})
    } satisfies ArrayValue);
  }
  return out;
}

/**
 * Replace every error cell inside ArrayValue args with BLANK so
 * aggregate functions simply skip them during flattening. Callers:
 * AGGREGATE with option 2, 3, 6, or 7.
 *
 * Direct scalar error args are left alone — AGGREGATE's caller
 * already passes them as scalars; our aggregators will surface them
 * as-is, which matches Excel when a direct literal is an error.
 *
 * This is distinct from the standard error-propagation path: arrays
 * that contain errors normally propagate those errors out of the
 * aggregate. With option 2/3/6/7, errors inside the *arrays* are
 * deliberately ignored.
 */
export function stripErrorCells(args: RuntimeValue[]): RuntimeValue[] {
  let needsCopy = false;
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            needsCopy = true;
            break;
          }
        }
        if (needsCopy) {
          break;
        }
      }
      if (needsCopy) {
        break;
      }
    }
  }
  if (!needsCopy) {
    return args;
  }
  const out: RuntimeValue[] = [];
  for (const arg of args) {
    if (arg.kind !== RVKind.Array) {
      out.push(arg);
      continue;
    }
    const newRows: ScalarValue[][] = [];
    for (const srcRow of arg.rows) {
      const newRow: ScalarValue[] = new Array<ScalarValue>(srcRow.length);
      for (let c = 0; c < srcRow.length; c++) {
        newRow[c] = srcRow[c].kind === RVKind.Error ? BLANK : srcRow[c];
      }
      newRows.push(newRow);
    }
    out.push({
      kind: RVKind.Array,
      rows: newRows,
      height: arg.height,
      width: arg.width,
      ...(arg.originRow !== undefined
        ? { originRow: arg.originRow, originCol: arg.originCol }
        : {}),
      ...(arg.subtotalMask ? { subtotalMask: arg.subtotalMask } : {}),
      ...(arg.hiddenRowMask ? { hiddenRowMask: arg.hiddenRowMask } : {})
    } satisfies ArrayValue);
  }
  return out;
}
