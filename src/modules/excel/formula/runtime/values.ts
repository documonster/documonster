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
 * 6. **Blank and MissingArg are distinct** — blank means "empty cell",
 *    MissingArg means "omitted function argument".
 */

import type { BoundExpr } from "../compile/bound-ast";
import type { ErrorCode } from "../integration/workbook-snapshot";

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
  Lambda = 7,
  /** An omitted function argument (different from blank). */
  MissingArg = 8
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
// Missing Argument Value
// ============================================================================

export interface MissingArgValue {
  readonly kind: RVKind.MissingArg;
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
  | LambdaValue
  | MissingArgValue;

// ============================================================================
// Singleton Constants
// ============================================================================

/** The singleton blank value. */
export const BLANK: BlankValue = { kind: RVKind.Blank };

/** The singleton missing argument value. */
export const MISSING_ARG: MissingArgValue = { kind: RVKind.MissingArg };

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

export function rvArray(rows: ScalarValue[][], originRow?: number, originCol?: number): ArrayValue {
  const height = rows.length;
  // Determine max width across all rows for rectangular normalization
  let width = 0;
  for (const row of rows) {
    if (row.length > width) {
      width = row.length;
    }
  }
  // Normalize: pad short rows with BLANK to ensure rectangular shape
  if (height > 0 && width > 0) {
    for (let r = 0; r < height; r++) {
      const row = rows[r];
      if (row.length < width) {
        for (let c = row.length; c < width; c++) {
          row.push(BLANK);
        }
      }
    }
  }
  return originRow !== undefined
    ? { kind: RVKind.Array, rows, height, width, originRow, originCol }
    : { kind: RVKind.Array, rows, height, width };
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

export function isBlank(v: RuntimeValue): v is BlankValue {
  return v.kind === RVKind.Blank;
}

export function isArray(v: RuntimeValue): v is ArrayValue {
  return v.kind === RVKind.Array;
}

export function isRef(v: RuntimeValue): v is ReferenceValue {
  return v.kind === RVKind.Reference;
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
 * Coerce a runtime value to a number.
 * - Blank → 0
 * - Number → itself
 * - Boolean → 1 / 0
 * - String → parse or #VALUE!
 * - Error → propagate
 * - MissingArg → 0
 */
export function toNumberRV(v: RuntimeValue): NumberValue | ErrorValue {
  switch (v.kind) {
    case RVKind.Number:
      return v;
    case RVKind.Blank:
    case RVKind.MissingArg:
      return rvNumber(0);
    case RVKind.Boolean:
      return rvNumber(v.value ? 1 : 0);
    case RVKind.String: {
      if (v.value === "") {
        return rvNumber(0);
      }
      const n = Number(v.value);
      return isNaN(n) ? ERRORS.VALUE : rvNumber(n);
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
    case RVKind.MissingArg:
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
    case RVKind.MissingArg:
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
  if (v.kind === RVKind.MissingArg) {
    return BLANK;
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
