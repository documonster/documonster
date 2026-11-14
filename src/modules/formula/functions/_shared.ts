/**
 * Shared helpers for native function implementations.
 *
 * Centralizes small utilities that were previously duplicated across multiple
 * function files (math, statistical, financial, text, date, engineering,
 * conditional, dynamic-array, database, lookup), and the rounding every
 * Excel-rounding function shares.
 */

import type {
  RuntimeValue,
  ScalarValue,
  NumberValue,
  ErrorValue,
  ArrayValue
} from "@formula/runtime/values";
import { RVKind, BLANK, ERRORS, rvArray, toNumberRV, topLeft } from "@formula/runtime/values";

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
// Rounding
// ============================================================================

/**
 * A number as Excel sees it: rounded to 15 significant digits.
 *
 * Excel keeps 15 significant decimal digits and rounds at that precision, so `1.005` is the
 * decimal 1.005 to it — not the binary 1.00499999999999989… a double holds — and
 * `0.1 + 0.2` is 0.3. Rounding the binary value instead is how `ROUND(1.005, 2)` came out 1.
 */
export function toExcelPrecision(value: number): number {
  return Number.isFinite(value) ? Number(value.toPrecision(15)) : value;
}

/**
 * Move the decimal point of `value` by `exponent` places without a binary multiplication.
 *
 * Done on the number's shortest decimal form, so `shiftDecimal(1.005, 2)` is exactly 100.5;
 * `1.005 * 100` is 100.49999999999999.
 */
function shiftDecimal(value: number, exponent: number): number {
  const text = String(value);
  const e = text.indexOf("e");
  return e < 0
    ? Number(`${text}e${exponent}`)
    : Number(`${text.slice(0, e)}e${Number(text.slice(e + 1)) + exponent}`);
}

/** 10^0 … 10^22: every one is exact in a double. */
const POWERS_OF_TEN = Array.from({ length: 23 }, (_, i) => Number(`1e${i}`));

/** Relative margin around a scaled value inside which a rounding boundary sends `applyRounding` the slow, decimal way. */
const ROUNDING_GUARD = 1e-12;

/**
 * The furthest `applyRounding` moves the decimal point. Doubles span about 10^-324 to 10^308, so
 * moving any of them by this much already overflows or underflows; a larger `digits` rounds the
 * same way, and clamping keeps it from writing an exponent like `e1e+300` into a number literal.
 */
const MAX_DECIMAL_SHIFT = 1000;

/**
 * Round `value` to `digits` decimal places (negative: to the left of the point) with `round`,
 * the way Excel's `ROUND` family does.
 *
 * The value is first read at Excel's precision ({@link toExcelPrecision}), then the decimal
 * point is moved in decimal rather than by multiplying by a power of ten, so the only rounding
 * that happens is the one asked for. `digits` is truncated. Where `digits` reaches past the
 * value's last digit there is nothing to round; where it reaches past its first, the value rounds
 * as the smallest amount of its sign would — to zero, or to one whole step away from it (`ROUNDUP`).
 * That step can be larger than any double, and so can a value read at 15 digits just below the
 * largest one: the result is then infinite, and the caller reports `#NUM!`.
 */
export function applyRounding(value: number, digits: number, round: (n: number) => number): number {
  if (!Number.isFinite(value)) {
    return value;
  }
  const d = Math.max(-MAX_DECIMAL_SHIFT, Math.min(MAX_DECIMAL_SHIFT, Math.trunc(digits)));
  // Fast path. Scaling by a power of ten in binary is off by far less than one part in 10^12, so
  // when nudging the scaled value by that much either way rounds to the same integer, no
  // rounding boundary is near and the decimal route below would give that integer too. Turning
  // it back is one correctly rounded operation by an exact power of ten, which is the same double
  // as parsing the decimal result. None of that holds once the scaling overflowed to infinity or
  // lost a non-zero value to zero — both ends then round alike for the wrong reason.
  if (d <= 22 && d >= -22) {
    const factor = POWERS_OF_TEN[Math.abs(d)];
    const fast = d >= 0 ? value * factor : value / factor;
    if (Number.isFinite(fast) && (fast !== 0 || value === 0)) {
      const lower = round(fast * (1 - ROUNDING_GUARD));
      if (lower === round(fast * (1 + ROUNDING_GUARD))) {
        // `+ 0` turns -0 into 0: a spreadsheet has no negative zero, and the decimal route has none.
        return (d >= 0 ? lower / factor : lower * factor) + 0;
      }
    }
  }
  const decimal = toExcelPrecision(value);
  const scaled = shiftDecimal(decimal, d);
  // Overflow means `digits` reaches past the value's last significant digit: nothing to round.
  if (!Number.isFinite(scaled)) {
    return decimal;
  }
  // Underflow means it reaches past the first: what is left is smaller than any double, and
  // rounds as the smallest one of its sign does.
  const rounded = round(
    scaled !== 0 || decimal === 0 ? scaled : Math.sign(decimal) * Number.MIN_VALUE
  );
  return shiftDecimal(rounded, -d);
}

/**
 * Round half away from zero, matching Excel's ROUND semantics.
 *
 * JavaScript's `Math.round` rounds half toward +∞ (so `Math.round(-0.5)`
 * returns `-0` instead of `-1`), whereas Excel rounds the magnitude and
 * preserves the sign: `ROUND(-0.5, 0) = -1`.
 */
export function roundHalfAwayFromZero(n: number): number {
  return n < 0 ? -Math.round(-n) : Math.round(n);
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
 * An argument Excel reads as an array — a range, an array constant, or a function's array
 * result — as an `ArrayValue`. A single value is a 1×1 array: a one-cell reference reaches a
 * function as the cell's value (see `dereferenceValue`), and Excel treats it, like any
 * scalar, as a one-element array (`SORT(A1)` is A1, `MATCH(5,A1,0)` is 1). A single error
 * value is returned as the error, which the caller propagates (`SORT(NA())` is #N/A); a
 * LAMBDA, or a reference left unresolved, is no array at all and is #VALUE!.
 */
export function asArrayArg(v: RuntimeValue): ArrayValue | ErrorValue {
  switch (v.kind) {
    case RVKind.Array:
    case RVKind.Error:
      return v;
    case RVKind.Lambda:
    case RVKind.Reference:
      return ERRORS.VALUE;
    default:
      return rvArray([[v]]);
  }
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
