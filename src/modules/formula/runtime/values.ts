/**
 * Runtime Value System — The engine's canonical value representation.
 *
 * Every value that flows through the formula evaluator is a `RuntimeValue`.
 * This is a properly tagged discriminated union that all function
 * implementations operate on directly — no adapter layer.
 *
 * ## Design Principles
 *
 * 1. **No Date objects** — dates are represented as their Excel serial
 *    number (a plain `number`). Date formatting is an output concern.
 * 2. **No monkey-patched arrays** — `ArrayValue` is a proper type with
 *    explicit `height`, `width`, and optional origin metadata.
 * 3. **Reference is a first-class value** — `ReferenceValue` can flow
 *    through the evaluator and be passed to functions.
 * 4. **Error codes are strict** — `ErrorValue` uses a typed code enum.
 * 5. **Lambda is a value** — `LambdaValue` is part of the value union.
 */

import type { BoundExpr } from "@formula/compile/bound-ast";
import type { ErrorCode } from "@formula/integration/workbook-snapshot";

// ============================================================================
// Value Tag Enum
// ============================================================================

export const enum RVKind {
  /** Empty cell / blank. Numeric value 0, string value "". */
  Blank = 0,
  /** A number (includes dates as serial numbers). */
  Number = 1,
  /** A string. */
  String = 2,
  /** A boolean (TRUE / FALSE). */
  Boolean = 3,
  /** An error value (#N/A, #VALUE!, etc.). */
  Error = 4,
  /** A 2D array of scalar values. */
  Array = 5,
  /** A cell or area reference (lazy — not yet resolved to values). */
  Reference = 6,
  /** A lambda (closure). */
  Lambda = 7
}

// ============================================================================
// Scalar Value Types
// ============================================================================

export interface BlankValue {
  readonly kind: RVKind.Blank;
}

export interface NumberValue {
  readonly kind: RVKind.Number;
  readonly value: number;
}

export interface StringValue {
  readonly kind: RVKind.String;
  readonly value: string;
}

export interface BooleanValue {
  readonly kind: RVKind.Boolean;
  readonly value: boolean;
}

export interface ErrorValue {
  readonly kind: RVKind.Error;
  readonly code: ErrorCode;
}

// ============================================================================
// Array Value
// ============================================================================

/**
 * A 2D array of scalar values.
 *
 * `rows[r][c]` is a `ScalarValue`. Arrays are always rectangular.
 * Origin metadata (for implicit intersection) is carried inline.
 */
export interface ArrayValue {
  readonly kind: RVKind.Array;
  /** Rows of scalar values. rows[0] is the first row. */
  readonly rows: readonly (readonly ScalarValue[])[];
  /** Number of rows. */
  readonly height: number;
  /** Number of columns. */
  readonly width: number;
  /** Origin row in the worksheet (1-based). Used for implicit intersection. */
  readonly originRow?: number;
  /** Origin column in the worksheet (1-based). Used for implicit intersection. */
  readonly originCol?: number;
  /**
   * Per-cell mask marking cells whose source formula is SUBTOTAL or
   * AGGREGATE. When a range is passed to an outer SUBTOTAL/AGGREGATE
   * call, those cells must be skipped so their results are not
   * double-counted (Excel semantics — standard totals-row behavior).
   *
   * Same shape as `rows`: `subtotalMask[r][c]` is true when the cell
   * should be excluded from outer SUBTOTAL/AGGREGATE aggregation.
   * Omitted when no cell in the array is a SUBTOTAL/AGGREGATE output.
   */
  readonly subtotalMask?: readonly (readonly boolean[])[];
  /**
   * Per-row mask marking rows whose source worksheet row is hidden.
   * Used by SUBTOTAL's 1xx-variant codes (101-111) and by AGGREGATE
   * options 5/7 to skip hidden rows. `hiddenRowMask[r]` is true when
   * row `r` of the array should be excluded under those semantics.
   * Omitted when no row in the array is hidden.
   */
  readonly hiddenRowMask?: readonly boolean[];
}

// ============================================================================
// Reference Value
// ============================================================================

/**
 * The shape of a reference.
 */
export type RefArea = {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
};

/**
 * A reference value that represents one or more areas in the workbook.
 *
 * References are first-class values in the new engine — they can be
 * passed to functions, returned from functions, and stored in variables.
 * They are resolved to actual values lazily when needed.
 */
export interface ReferenceValue {
  readonly kind: RVKind.Reference;
  /** The areas this reference covers. Usually one, but UNION produces multiple. */
  readonly areas: readonly RefArea[];
  /** Whether this reference originated from a single-cell ref (A1) vs an area ref (A1:A1). */
  readonly singleCell?: boolean;
}

// ============================================================================
// Lambda Value
// ============================================================================

/**
 * A lambda closure value.
 */
export interface LambdaValue {
  readonly kind: RVKind.Lambda;
  /** Parameter names (uppercase). */
  readonly params: readonly string[];
  /** The body expression to evaluate when called. */
  readonly body: BoundExpr;
  /** Captured variable bindings from the enclosing scope. */
  readonly closureBindings?: ReadonlyMap<string, RuntimeValue>;
}

// ============================================================================
// Discriminated Unions
// ============================================================================

/**
 * A scalar value (non-array, non-reference, non-lambda).
 */
export type ScalarValue = BlankValue | NumberValue | StringValue | BooleanValue | ErrorValue;

/**
 * Any value that can flow through the evaluator.
 */
export type RuntimeValue =
  | BlankValue
  | NumberValue
  | StringValue
  | BooleanValue
  | ErrorValue
  | ArrayValue
  | ReferenceValue
  | LambdaValue;

// ============================================================================
// Singleton Constants
// ============================================================================

/** The singleton blank value. */
export const BLANK: BlankValue = { kind: RVKind.Blank };

/** Common error values. */
export const ERRORS = {
  VALUE: { kind: RVKind.Error, code: "#VALUE!" } as ErrorValue,
  REF: { kind: RVKind.Error, code: "#REF!" } as ErrorValue,
  NAME: { kind: RVKind.Error, code: "#NAME?" } as ErrorValue,
  DIV0: { kind: RVKind.Error, code: "#DIV/0!" } as ErrorValue,
  NA: { kind: RVKind.Error, code: "#N/A" } as ErrorValue,
  NUM: { kind: RVKind.Error, code: "#NUM!" } as ErrorValue,
  NULL: { kind: RVKind.Error, code: "#NULL!" } as ErrorValue,
  SPILL: { kind: RVKind.Error, code: "#SPILL!" } as ErrorValue,
  CALC: { kind: RVKind.Error, code: "#CALC!" } as ErrorValue
} as const;

// ============================================================================
// Constructor Helpers
// ============================================================================

export function rvNumber(value: number): NumberValue {
  return { kind: RVKind.Number, value };
}

export function rvString(value: string): StringValue {
  return { kind: RVKind.String, value };
}

export function rvBoolean(value: boolean): BooleanValue {
  return { kind: RVKind.Boolean, value };
}

export function rvError(code: ErrorCode): ErrorValue {
  return { kind: RVKind.Error, code };
}

export function rvArray(
  rows: ScalarValue[][],
  originRow?: number,
  originCol?: number,
  subtotalMask?: readonly (readonly boolean[])[],
  hiddenRowMask?: readonly boolean[]
): ArrayValue {
  const height = rows.length;
  // Determine max width across all rows for rectangular normalisation.
  let width = 0;
  for (const row of rows) {
    if (row.length > width) {
      width = row.length;
    }
  }
  // Pad short rows with BLANK so the resulting ArrayValue is rectangular.
  // The old implementation did `row.push(BLANK)` directly — mutating the
  // caller's arrays. Callers that shared row references across multiple
  // `rvArray` calls could observe surprise modifications; we now copy any
  // row that needs padding and leave the caller's arrays untouched.
  let normalisedRows: ScalarValue[][] = rows;
  if (height > 0 && width > 0) {
    let anyNeedPadding = false;
    for (const row of rows) {
      if (row.length < width) {
        anyNeedPadding = true;
        break;
      }
    }
    if (anyNeedPadding) {
      normalisedRows = new Array<ScalarValue[]>(height);
      for (let r = 0; r < height; r++) {
        const row = rows[r];
        if (row.length === width) {
          normalisedRows[r] = row;
          continue;
        }
        const padded = new Array<ScalarValue>(width);
        for (let c = 0; c < row.length; c++) {
          padded[c] = row[c];
        }
        for (let c = row.length; c < width; c++) {
          padded[c] = BLANK;
        }
        normalisedRows[r] = padded;
      }
    }
  }
  return buildArrayValue(
    normalisedRows,
    height,
    width,
    originRow,
    originCol,
    subtotalMask,
    hiddenRowMask
  );
}

/**
 * Fast-path rectangular ArrayValue constructor.
 *
 * Callers that have already produced strictly-rectangular `rows` (every
 * row is the same length — the length they explicitly `new Array(width)`
 * allocated) can skip the two-pass width-scan + padding loop in
 * `rvArray`. Examples: `buildRangeArray`, `broadcastBinaryOp`,
 * `evaluateArrayLiteral`, `TRANSPOSE` — they all know `width` up front.
 *
 * Rows MUST be rectangular; passing ragged data will silently surface as
 * `undefined` cells downstream.
 */
export function rvArrayRect(
  rows: ScalarValue[][],
  height: number,
  width: number,
  originRow?: number,
  originCol?: number,
  subtotalMask?: readonly (readonly boolean[])[],
  hiddenRowMask?: readonly boolean[]
): ArrayValue {
  return buildArrayValue(rows, height, width, originRow, originCol, subtotalMask, hiddenRowMask);
}

function buildArrayValue(
  rows: ScalarValue[][],
  height: number,
  width: number,
  originRow: number | undefined,
  originCol: number | undefined,
  subtotalMask: readonly (readonly boolean[])[] | undefined,
  hiddenRowMask: readonly boolean[] | undefined
): ArrayValue {
  return originRow !== undefined
    ? {
        kind: RVKind.Array,
        rows,
        height,
        width,
        originRow,
        originCol,
        ...(subtotalMask ? { subtotalMask } : {}),
        ...(hiddenRowMask ? { hiddenRowMask } : {})
      }
    : {
        kind: RVKind.Array,
        rows,
        height,
        width,
        ...(subtotalMask ? { subtotalMask } : {}),
        ...(hiddenRowMask ? { hiddenRowMask } : {})
      };
}

export function rvRef(
  sheet: string,
  top: number,
  left: number,
  bottom: number,
  right: number
): ReferenceValue {
  return {
    kind: RVKind.Reference,
    areas: [{ sheet, top, left, bottom, right }]
  };
}

export function rvCellRef(sheet: string, row: number, col: number): ReferenceValue {
  return {
    kind: RVKind.Reference,
    areas: [{ sheet, top: row, left: col, bottom: row, right: col }],
    singleCell: true
  };
}

export function rvLambda(
  params: string[],
  body: BoundExpr,
  closureBindings?: ReadonlyMap<string, RuntimeValue>
): LambdaValue {
  return { kind: RVKind.Lambda, params, body, closureBindings };
}

// ============================================================================
// Type Guards
// ============================================================================

export function isError(v: RuntimeValue): v is ErrorValue {
  return v.kind === RVKind.Error;
}

export function isArray(v: RuntimeValue): v is ArrayValue {
  return v.kind === RVKind.Array;
}

export function isLambda(v: RuntimeValue): v is LambdaValue {
  return v.kind === RVKind.Lambda;
}

export function isScalar(v: RuntimeValue): v is ScalarValue {
  return (
    v.kind === RVKind.Blank ||
    v.kind === RVKind.Number ||
    v.kind === RVKind.String ||
    v.kind === RVKind.Boolean ||
    v.kind === RVKind.Error
  );
}

// ============================================================================
// Coercion Helpers
// ============================================================================

/**
 * Parse a user-facing numeric string the way Excel does.
 *
 * Accepts:
 *   - plain decimals: `"1"`, `"-1.5"`, `"+.25"`
 *   - scientific notation: `"1.2e3"`, `"2E-4"`
 *   - percentage suffix: `"50%"` → 0.5
 *   - leading/trailing whitespace around the above
 *
 * Rejects (unlike JavaScript's `Number()`):
 *   - empty strings and whitespace-only (`" "` would become 0)
 *   - `"Infinity"`, `"-Infinity"`, `"NaN"` (Excel treats as text)
 *   - hexadecimal (`"0x10"`), octal, binary literals
 *   - currency symbols, thousands separators, locale-specific formats
 *     (these are out of scope for the engine; callers should strip before
 *     calling)
 *
 * Returns `#VALUE!` on any rejection so the error bubbles naturally
 * through formula evaluation.
 */
function parseNumericString(raw: string): NumberValue | ErrorValue {
  const s = raw.trim();
  if (s === "") {
    return ERRORS.VALUE;
  }
  let body = s;
  let percentFactor = 1;
  if (body.endsWith("%")) {
    percentFactor = 0.01;
    body = body.slice(0, -1).trim();
    if (body === "") {
      return ERRORS.VALUE;
    }
  }
  // Require at least one digit somewhere; this shuts the door on
  // "Infinity", "NaN", "0x10", "1e" (Excel's own lexer refuses these).
  // The strict decimal grammar below also rejects "1_000" etc.
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(body)) {
    return ERRORS.VALUE;
  }
  const n = Number(body);
  if (!Number.isFinite(n)) {
    return ERRORS.VALUE;
  }
  return rvNumber(n * percentFactor);
}

/**
 * Coerce a runtime value to a number.
 * - Blank → 0
 * - Number → itself
 * - Boolean → 1 / 0
 * - String → parse or #VALUE!
 * - Error → propagate
 */
export function toNumberRV(v: RuntimeValue): NumberValue | ErrorValue {
  switch (v.kind) {
    case RVKind.Number:
      return v;
    case RVKind.Blank:
      return rvNumber(0);
    case RVKind.Boolean:
      return rvNumber(v.value ? 1 : 0);
    case RVKind.String: {
      return parseNumericString(v.value);
    }
    case RVKind.Error:
      return v;
    default:
      return ERRORS.VALUE;
  }
}

/**
 * Coerce a runtime value to a string.
 */
export function toStringRV(v: RuntimeValue): string {
  switch (v.kind) {
    case RVKind.String:
      return v.value;
    case RVKind.Number:
      return String(v.value);
    case RVKind.Boolean:
      return v.value ? "TRUE" : "FALSE";
    case RVKind.Blank:
      return "";
    case RVKind.Error:
      return v.code;
    default:
      return "";
  }
}

/**
 * Coerce a runtime value to a boolean.
 */
export function toBooleanRV(v: RuntimeValue): BooleanValue | ErrorValue {
  switch (v.kind) {
    case RVKind.Boolean:
      return v;
    case RVKind.Number:
      return rvBoolean(v.value !== 0);
    case RVKind.Blank:
      return rvBoolean(false);
    case RVKind.String: {
      const u = v.value.toUpperCase();
      if (u === "TRUE") {
        return rvBoolean(true);
      }
      if (u === "FALSE") {
        return rvBoolean(false);
      }
      return ERRORS.VALUE;
    }
    case RVKind.Error:
      return v;
    default:
      return ERRORS.VALUE;
  }
}

/**
 * Structural equality of scalar values.
 *
 * - Different kinds → false
 * - Number / Boolean / Blank → strict value equality (Blank always equal)
 * - String → case-insensitive comparison (Excel semantics)
 * - Error → not equal (errors do not compare equal to each other)
 */
/**
 * Three-way compare two scalars that share a kind.
 *
 * Returns a negative number if `a < b`, zero if equal, positive if `a > b`.
 * Returns `NaN` when the kinds differ or cannot be ordered (e.g. errors);
 * callers decide how to surface the incomparability — sort helpers usually
 * skip NaN pairs, while comparison operators route to a kind-priority
 * tiebreak. Strings are compared case-insensitively to match Excel.
 */
export function compareScalarsSameKind(a: ScalarValue, b: ScalarValue): number {
  if (a.kind !== b.kind) {
    return Number.NaN;
  }
  switch (a.kind) {
    case RVKind.Number:
      return a.value - (b as NumberValue).value;
    case RVKind.String: {
      const al = a.value.toLowerCase();
      const bl = (b as StringValue).value.toLowerCase();
      return al < bl ? -1 : al > bl ? 1 : 0;
    }
    case RVKind.Boolean: {
      const bv = (b as BooleanValue).value;
      return a.value === bv ? 0 : a.value ? 1 : -1;
    }
    case RVKind.Blank:
      return 0;
    default:
      return Number.NaN;
  }
}

export function scalarEquals(a: ScalarValue, b: ScalarValue): boolean {
  if (a.kind !== b.kind) {
    return false;
  }
  switch (a.kind) {
    case RVKind.Number:
      return a.value === (b as NumberValue).value;
    case RVKind.String:
      return a.value.toLowerCase() === (b as StringValue).value.toLowerCase();
    case RVKind.Boolean:
      return a.value === (b as BooleanValue).value;
    case RVKind.Blank:
      return true;
    default:
      return false;
  }
}

/**
 * Get the top-left scalar from any value (for implicit intersection fallback).
 */
export function topLeft(v: RuntimeValue): ScalarValue {
  if (isScalar(v)) {
    return v;
  }
  if (v.kind === RVKind.Array) {
    if (v.height === 0 || v.width === 0) {
      return BLANK;
    }
    return v.rows[0][0];
  }
  // Reference, Lambda → need context to resolve
  return ERRORS.VALUE;
}

// ============================================================================
// Snapshot Value Conversion
// ============================================================================

/**
 * Convert a SnapshotCellValue to a RuntimeValue.
 */
export function fromSnapshotValue(
  v: number | string | boolean | { error: string } | null
): ScalarValue {
  if (v === null) {
    return BLANK;
  }
  if (typeof v === "number") {
    return rvNumber(v);
  }
  if (typeof v === "string") {
    return rvString(v);
  }
  if (typeof v === "boolean") {
    return rvBoolean(v);
  }
  if (typeof v === "object" && "error" in v) {
    return rvError(v.error as ErrorCode);
  }
  return BLANK;
}
