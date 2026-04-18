/**
 * Bound Expression AST — The output of the compilation (binding) phase.
 *
 * After parsing produces a raw AST, the binder resolves all static symbols:
 * - Defined names → resolved cell/range references or formula expressions
 * - Structured references → resolved cell/range references
 * - Sheet names → validated against the snapshot
 * - Function names → validated against the registry
 *
 * The result is a `BoundExpr` tree where every reference is resolved and
 * every function call is validated. The runtime evaluator only executes
 * `BoundExpr` nodes — it never needs to perform symbol resolution.
 *
 * ## Node Types (compile output)
 *
 * After binding, only these node types exist:
 *
 * - `BoundLiteral` — number, string, boolean, error, null (missing)
 * - `BoundCellRef` — resolved cell reference (sheet + row + col)
 * - `BoundAreaRef` — resolved rectangular range (sheet + bounds)
 * - `BoundColRangeRef` — resolved whole-column range
 * - `BoundRowRangeRef` — resolved whole-row range
 * - `BoundRef3D` — resolved 3D reference (multiple sheets)
 * - `BoundBinaryOp` — binary operation
 * - `BoundUnaryOp` — unary operation (including @)
 * - `BoundPercent` — postfix %
 * - `BoundCall` — function call (eagerly evaluated args)
 * - `BoundSpecialCall` — special form (IF, LET, LAMBDA, etc.) with lazy args
 * - `BoundArray` — array constant {1,2;3,4}
 * - `BoundNameExpr` — UNRESOLVED name (only for names that must be resolved at runtime,
 *   e.g. formula-based defined names that produce lambdas, or truly unknown names)
 * - `BoundLambda` — LAMBDA expression (params + body)
 */

// ============================================================================
// Bound Expression Tag Enum
// ============================================================================

export const enum BoundExprKind {
  Literal = 1,
  CellRef = 2,
  AreaRef = 3,
  ColRangeRef = 4,
  RowRangeRef = 5,
  Ref3D = 6,
  BinaryOp = 7,
  UnaryOp = 8,
  Percent = 9,
  Call = 10,
  SpecialCall = 11,
  Array = 12,
  NameExpr = 13,
  Lambda = 14,
  StructuredRef = 15
}

// ============================================================================
// Individual Bound Expression Types
// ============================================================================

/**
 * A resolved literal value.
 */
export interface BoundLiteral {
  readonly kind: BoundExprKind.Literal;
  /** number | string | boolean | null (for Missing/blank) */
  readonly value: number | string | boolean | null;
  /** If this is an error literal (e.g. #N/A), the error code. */
  readonly errorCode?: string;
}

/**
 * A fully resolved cell reference.
 * After binding, sheet is always resolved (never undefined).
 */
export interface BoundCellRef {
  readonly kind: BoundExprKind.CellRef;
  /** Resolved sheet name. */
  readonly sheet: string;
  /** 1-based row number. */
  readonly row: number;
  /** 1-based column number. */
  readonly col: number;
}

/**
 * A fully resolved rectangular area reference.
 */
export interface BoundAreaRef {
  readonly kind: BoundExprKind.AreaRef;
  /** Resolved sheet name. */
  readonly sheet: string;
  /** 1-based bounds. */
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/**
 * A resolved whole-column range (e.g. A:B).
 * Row bounds are determined at runtime from sheet dimensions.
 */
export interface BoundColRangeRef {
  readonly kind: BoundExprKind.ColRangeRef;
  readonly sheet: string;
  readonly leftCol: number;
  readonly rightCol: number;
}

/**
 * A resolved whole-row range (e.g. 1:5).
 * Column bounds are determined at runtime from sheet dimensions.
 */
export interface BoundRowRangeRef {
  readonly kind: BoundExprKind.RowRangeRef;
  readonly sheet: string;
  readonly topRow: number;
  readonly bottomRow: number;
}

/**
 * A resolved 3D reference spanning multiple sheets.
 * The inner reference is either a cell or an area.
 */
export interface BoundRef3D {
  readonly kind: BoundExprKind.Ref3D;
  /** Sheet names in workbook order, from start to end (inclusive). */
  readonly sheets: readonly string[];
  /** The cell/area that is replicated across sheets. */
  readonly inner: BoundCellRef | BoundAreaRef;
}

/**
 * Binary operation with two sub-expressions.
 */
export interface BoundBinaryOp {
  readonly kind: BoundExprKind.BinaryOp;
  readonly op: string;
  readonly left: BoundExpr;
  readonly right: BoundExpr;
}

/**
 * Unary operation (prefix +, -, or @ implicit intersection).
 */
export interface BoundUnaryOp {
  readonly kind: BoundExprKind.UnaryOp;
  readonly op: string;
  readonly operand: BoundExpr;
}

/**
 * Postfix percent operation.
 */
export interface BoundPercent {
  readonly kind: BoundExprKind.Percent;
  readonly operand: BoundExpr;
}

/**
 * A standard (eager) function call.
 * All arguments are evaluated before the function is called.
 */
export interface BoundCall {
  readonly kind: BoundExprKind.Call;
  /** Uppercase, canonical function name. */
  readonly name: string;
  /** Bound argument expressions. */
  readonly args: readonly BoundExpr[];
}

/**
 * Special form identifiers for lazy/short-circuit evaluation.
 */
export type SpecialFormName =
  | "IF"
  | "IFERROR"
  | "IFNA"
  | "IFS"
  | "SWITCH"
  | "CHOOSE"
  | "LET"
  | "LAMBDA"
  | "INDIRECT"
  | "OFFSET"
  | "MAP"
  | "REDUCE"
  | "SCAN"
  | "MAKEARRAY"
  | "BYROW"
  | "BYCOL";

/**
 * A special-form function call with lazy argument evaluation.
 * The evaluator handles these with custom logic (short-circuit, binding, etc.).
 */
export interface BoundSpecialCall {
  readonly kind: BoundExprKind.SpecialCall;
  readonly name: SpecialFormName;
  /** Argument expressions (evaluated lazily by the evaluator). */
  readonly args: readonly BoundExpr[];
}

/**
 * Array constant: {1,2;3,4}.
 */
export interface BoundArray {
  readonly kind: BoundExprKind.Array;
  /** rows[i][j] = element expression. */
  readonly rows: readonly (readonly BoundExpr[])[];
}

/**
 * An unresolved name reference.
 *
 * This exists for names that cannot be fully resolved at compile time:
 * - Formula-based defined names (e.g. =LAMBDA(x,y,x+y))
 * - Names that resolve to lambdas stored in cells
 * - LET-bound local variables (resolved at runtime from localBindings)
 * - Truly unknown names (#NAME? at runtime)
 *
 * For names that resolve to simple cell/range references, the binder
 * produces `BoundCellRef` or `BoundAreaRef` directly.
 */
export interface BoundNameExpr {
  readonly kind: BoundExprKind.NameExpr;
  /** The original name (case-preserved). */
  readonly name: string;
  /** Uppercase name for lookup. */
  readonly upperName: string;
}

/**
 * A LAMBDA expression with parameter names and a body.
 * Produced by the binder when it encounters a LAMBDA special form.
 */
export interface BoundLambda {
  readonly kind: BoundExprKind.Lambda;
  /** Parameter names (uppercase). */
  readonly params: readonly string[];
  /** The body expression. */
  readonly body: BoundExpr;
}

/**
 * An unresolved structured reference that requires runtime context.
 *
 * This is produced by the binder when:
 * - The table name is empty (implicit `@` syntax — needs current cell position)
 * - The reference includes `#This Row` (needs current cell's row)
 *
 * The evaluator resolves these at runtime using the current cell address.
 */
export interface BoundStructuredRef {
  readonly kind: BoundExprKind.StructuredRef;
  /** Table name (empty string for implicit table). */
  readonly tableName: string;
  /** Column names. */
  readonly columns: readonly string[];
  /** Special items (#Headers, #Data, #Totals, #All, #This Row). */
  readonly specials: readonly string[];
}

// ============================================================================
// Discriminated Union
// ============================================================================

export type BoundExpr =
  | BoundLiteral
  | BoundCellRef
  | BoundAreaRef
  | BoundColRangeRef
  | BoundRowRangeRef
  | BoundRef3D
  | BoundBinaryOp
  | BoundUnaryOp
  | BoundPercent
  | BoundCall
  | BoundSpecialCall
  | BoundArray
  | BoundNameExpr
  | BoundLambda
  | BoundStructuredRef;

// ============================================================================
// Constructor Helpers
// ============================================================================

export function boundLiteral(
  value: number | string | boolean | null,
  errorCode?: string
): BoundLiteral {
  return errorCode !== undefined
    ? { kind: BoundExprKind.Literal, value, errorCode }
    : { kind: BoundExprKind.Literal, value };
}

export function boundCellRef(sheet: string, row: number, col: number): BoundCellRef {
  return { kind: BoundExprKind.CellRef, sheet, row, col };
}

export function boundAreaRef(
  sheet: string,
  top: number,
  left: number,
  bottom: number,
  right: number
): BoundAreaRef {
  return { kind: BoundExprKind.AreaRef, sheet, top, left, bottom, right };
}

export function boundCall(name: string, args: BoundExpr[]): BoundCall {
  return { kind: BoundExprKind.Call, name, args };
}

export function boundSpecialCall(name: SpecialFormName, args: BoundExpr[]): BoundSpecialCall {
  return { kind: BoundExprKind.SpecialCall, name, args };
}

export function boundNameExpr(name: string): BoundNameExpr {
  return { kind: BoundExprKind.NameExpr, name, upperName: name.toUpperCase() };
}

export function boundErrorLiteral(errorCode: string): BoundLiteral {
  return { kind: BoundExprKind.Literal, value: null, errorCode };
}
