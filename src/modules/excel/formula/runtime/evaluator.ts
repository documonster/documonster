/**
 * Evaluator — Execute BoundExpr using the RuntimeValue system.
 *
 * The evaluator operates on BoundExpr (from the compile phase),
 * WorkbookSnapshot (from the snapshot phase), and RuntimeValue
 * (the value system).
 */

import { parseDefinedNameRange } from "../compile/address-utils";
import { bind, type BindingContext } from "../compile/binder";
import type {
  BoundExpr,
  BoundCellRef,
  BoundAreaRef,
  BoundCall,
  BoundSpecialCall,
  BoundNameExpr,
  BoundLambda
} from "../compile/bound-ast";
import { BoundExprKind } from "../compile/bound-ast";
import type { CompiledFormula } from "../compile/compiled-formula";
import {
  resolveStructuredRefRows,
  buildTableGeometry,
  resolveStructuredRefColumns
} from "../compile/structured-ref-utils";
import type { WorkbookSnapshot } from "../integration/workbook-snapshot";
import {
  snapshotCellKey,
  formulaCellKey,
  resolveDefinedName as resolveDefinedNameFromSnapshot
} from "../integration/workbook-snapshot";
import { parse } from "../syntax/parser";
import { tokenize } from "../syntax/tokenizer";
import { lookupFunction, ensureRegistryInitialized } from "./function-registry";
import type {
  RuntimeValue,
  ScalarValue,
  ArrayValue,
  LambdaValue,
  ErrorValue,
  RefArea,
  NumberValue
} from "./values";
import {
  RVKind,
  BLANK,
  ERRORS,
  rvNumber,
  rvString,
  rvBoolean,
  rvError,
  rvArray,
  rvRef,
  rvCellRef,
  rvLambda,
  isError,
  isScalar,
  isLambda,
  toNumberRV,
  toStringRV,
  toBooleanRV,
  topLeft,
  fromSnapshotValue
} from "./values";

// ============================================================================
// Eval Session
// ============================================================================

/**
 * Cached formula evaluation result with both scalar and raw forms.
 * The scalar form is used by dependent formulas (implicit intersection applied).
 * The raw form preserves the full array for dynamic array / CSE materialization.
 */
interface CachedResult {
  /** Scalar (implicit-intersected, dereferenced) result. */
  readonly scalar: RuntimeValue;
  /** Raw (full array) result. Only differs from scalar for array formulas. */
  readonly raw: RuntimeValue;
}

/**
 * Per-calculation mutable state.
 */
export class EvalSession {
  /** Cells currently on the evaluation call stack (circular-ref detection). */
  readonly evaluating = new Set<string>();
  /**
   * Unified formula result cache.
   * Each entry holds both the scalar form (for dependents) and the raw form
   * (for materialize). This replaces the previous separate cache/rawCache
   * pattern with a single, self-documenting structure.
   */
  readonly resultCache = new Map<string, CachedResult>();
  /** Cache for runtime name resolution (defined names that need parsing). */
  readonly nameCache = new Map<string, RuntimeValue>();
  /** Fallback values for circular references during iterative calculation. */
  readonly circularFallback = new Map<string, RuntimeValue>();

  /**
   * Runtime dependency recorder — tracks cell accesses made during evaluation.
   *
   * When a formula with `hasDynamicRefs` (INDIRECT/OFFSET) is being evaluated,
   * every `getCellValue` / `buildRangeArray` call records the accessed cell/range
   * key here. After evaluation, these dynamic edges can be merged with the
   * compiled static dependency set to produce a complete dependency graph.
   *
   * Key: formula cell key being evaluated → Set of accessed cell keys.
   * Only populated for formulas that have `hasDynamicRefs === true`.
   */
  readonly dynamicDeps = new Map<string, Set<string>>();

  /**
   * The formula cell key currently being recorded (null if recording is off).
   * Set before evaluating a formula with dynamic refs, cleared after.
   */
  recordingKey: string | null = null;

  makeKey(sheet: string, row: number, col: number): string {
    return formulaCellKey(sheet, row, col);
  }

  /**
   * Record a cell access for the currently-recording formula.
   */
  recordAccess(sheet: string, row: number, col: number): void {
    if (this.recordingKey === null) {
      return;
    }
    let deps = this.dynamicDeps.get(this.recordingKey);
    if (!deps) {
      deps = new Set();
      this.dynamicDeps.set(this.recordingKey, deps);
    }
    deps.add(formulaCellKey(sheet, row, col));
  }
}

// ============================================================================
// Eval Context
// ============================================================================

/**
 * The evaluation context. Carries the snapshot and compiled formula map
 * for the evaluator to access cell values and resolve names at runtime.
 */
export interface EvalContext {
  /** The workbook snapshot. */
  readonly snapshot: WorkbookSnapshot;
  /** Map from formula cell key to CompiledFormula. */
  readonly compiledFormulas: ReadonlyMap<string, CompiledFormula>;
  /** AST cache for INDIRECT re-parsing. */
  readonly astCache: Map<string, BoundExpr>;
  /** The current sheet name (for relative references). */
  currentSheet: string;
  /** Current cell address being evaluated. */
  currentAddress?: { sheet: string; row: number; col: number };
  /** Local variable bindings from LET expressions. */
  localBindings?: Map<string, RuntimeValue>;
}

// ============================================================================
// Main Evaluate Function
// ============================================================================

/**
 * Evaluate a BoundExpr to produce a RuntimeValue.
 */
export function evaluate(expr: BoundExpr, ctx: EvalContext, session: EvalSession): RuntimeValue {
  switch (expr.kind) {
    case BoundExprKind.Literal:
      return evaluateLiteral(expr);

    case BoundExprKind.CellRef:
      return evaluateCellRef(expr, ctx, session);

    case BoundExprKind.AreaRef:
      return evaluateAreaRef(expr, ctx, session);

    case BoundExprKind.ColRangeRef:
      return evaluateColRange(expr, ctx, session);

    case BoundExprKind.RowRangeRef:
      return evaluateRowRange(expr, ctx, session);

    case BoundExprKind.Ref3D:
      return evaluateRef3D(expr, ctx, session);

    case BoundExprKind.BinaryOp:
      return evaluateBinaryOp(expr.op, expr.left, expr.right, ctx, session);

    case BoundExprKind.UnaryOp:
      return evaluateUnaryOp(expr.op, expr.operand, ctx, session);

    case BoundExprKind.Percent:
      return evaluatePercent(expr.operand, ctx, session);

    case BoundExprKind.Call:
      return evaluateCall(expr, ctx, session);

    case BoundExprKind.SpecialCall:
      return evaluateSpecialCall(expr, ctx, session);

    case BoundExprKind.Array:
      return evaluateArrayLiteral(expr, ctx, session);

    case BoundExprKind.NameExpr:
      return evaluateNameExpr(expr, ctx, session);

    case BoundExprKind.Lambda:
      return evaluateLambdaExpr(expr, ctx);

    case BoundExprKind.StructuredRef:
      return evaluateStructuredRef(expr, ctx, session);

    default: {
      const _: never = expr;
      return ERRORS.VALUE;
    }
  }
}

// ============================================================================
// Literal
// ============================================================================

function evaluateLiteral(expr: {
  value: number | string | boolean | null;
  errorCode?: string;
}): ScalarValue {
  if (expr.errorCode) {
    return rvError(expr.errorCode as ErrorValue["code"]);
  }
  if (expr.value === null) {
    return BLANK;
  }
  if (typeof expr.value === "number") {
    return rvNumber(expr.value);
  }
  if (typeof expr.value === "string") {
    return rvString(expr.value);
  }
  if (typeof expr.value === "boolean") {
    return rvBoolean(expr.value);
  }
  return BLANK;
}

// ============================================================================
// Cell Reference
// ============================================================================

function evaluateCellRef(expr: BoundCellRef, ctx: EvalContext, session: EvalSession): RuntimeValue {
  return rvCellRef(expr.sheet, expr.row, expr.col);
}

// ============================================================================
// Area Reference → ReferenceValue
// ============================================================================

function evaluateAreaRef(expr: BoundAreaRef, ctx: EvalContext, session: EvalSession): RuntimeValue {
  return rvRef(expr.sheet, expr.top, expr.left, expr.bottom, expr.right);
}

function evaluateColRange(
  expr: { sheet: string; leftCol: number; rightCol: number },
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const ws = ctx.snapshot.worksheetsByName.get(expr.sheet.toLowerCase());
  if (!ws || !ws.dimensions) {
    return rvArray([]);
  }
  return rvRef(expr.sheet, ws.dimensions.top, expr.leftCol, ws.dimensions.bottom, expr.rightCol);
}

function evaluateRowRange(
  expr: { sheet: string; topRow: number; bottomRow: number },
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const ws = ctx.snapshot.worksheetsByName.get(expr.sheet.toLowerCase());
  if (!ws || !ws.dimensions) {
    return rvArray([]);
  }
  return rvRef(expr.sheet, expr.topRow, ws.dimensions.left, expr.bottomRow, ws.dimensions.right);
}

function evaluateRef3D(
  expr: { sheets: readonly string[]; inner: BoundCellRef | BoundAreaRef },
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const areas: RefArea[] = [];
  for (const sheet of expr.sheets) {
    if (expr.inner.kind === BoundExprKind.CellRef) {
      areas.push({
        sheet,
        top: expr.inner.row,
        left: expr.inner.col,
        bottom: expr.inner.row,
        right: expr.inner.col
      });
    } else {
      areas.push({
        sheet,
        top: expr.inner.top,
        left: expr.inner.left,
        bottom: expr.inner.bottom,
        right: expr.inner.right
      });
    }
  }
  return { kind: RVKind.Reference, areas };
}

// ============================================================================
// Build Range Array from Snapshot
// ============================================================================

function buildRangeArray(
  ctx: EvalContext,
  session: EvalSession,
  sheet: string,
  top: number,
  left: number,
  bottom: number,
  right: number
): ArrayValue {
  const rows: ScalarValue[][] = [];
  for (let r = top; r <= bottom; r++) {
    const row: ScalarValue[] = [];
    for (let c = left; c <= right; c++) {
      const val = getCellValue(sheet, r, c, ctx, session);
      row.push(toScalar(val));
    }
    rows.push(row);
  }
  return rvArray(rows, top, left);
}

// ============================================================================
// Dereference: Reference → concrete value
// ============================================================================

/**
 * Resolve a `ReferenceValue` to its concrete value (scalar or array).
 * - Single-cell references (from CellRef nodes) resolve to a scalar.
 * - Area references (even 1x1 from A1:A1) produce an array.
 * - Multi-area references (from 3D refs) flatten all areas into one array.
 * Non-reference values are returned unchanged.
 */
function dereferenceValue(v: RuntimeValue, ctx: EvalContext, session: EvalSession): RuntimeValue {
  if (v.kind !== RVKind.Reference) {
    return v;
  }
  if (v.areas.length === 0) {
    return BLANK;
  }
  // Single-cell references (from CellRef nodes) resolve to scalar
  if (v.singleCell) {
    const area = v.areas[0];
    return getCellValue(area.sheet, area.top, area.left, ctx, session);
  }
  // Single area — build range array
  if (v.areas.length === 1) {
    const area = v.areas[0];
    return buildRangeArray(ctx, session, area.sheet, area.top, area.left, area.bottom, area.right);
  }
  // Multi-area (3D reference) — flatten all areas into one array
  const allRows: ScalarValue[][] = [];
  for (const area of v.areas) {
    const arr = buildRangeArray(
      ctx,
      session,
      area.sheet,
      area.top,
      area.left,
      area.bottom,
      area.right
    );
    for (const row of arr.rows) {
      allRows.push([...row]);
    }
  }
  return rvArray(allRows);
}

// ============================================================================
// Get Cell Value from Snapshot
// ============================================================================

function getCellValue(
  sheetName: string,
  row: number,
  col: number,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  // Record this access for runtime dependency tracking
  session.recordAccess(sheetName, row, col);

  const ws = ctx.snapshot.worksheetsByName.get(sheetName.toLowerCase());
  if (!ws) {
    return BLANK;
  }

  const cellKey = snapshotCellKey(row, col);
  const cell = ws.cells.get(cellKey);
  if (!cell) {
    return BLANK;
  }

  // If this cell has a formula, evaluate it
  if (cell.formulaKind !== "none" && cell.formula) {
    const fKey = formulaCellKey(sheetName, row, col);

    // Check cache — return scalar form for dependency resolution
    const cached = session.resultCache.get(fKey);
    if (cached !== undefined) {
      return cached.scalar;
    }

    // Get compiled formula
    const compiled = ctx.compiledFormulas.get(fKey);
    if (compiled) {
      return evaluateFormula(compiled, ctx, session);
    }
  }

  // Non-formula cell — return snapshot value
  return fromSnapshotValue(cell.value);
}

// ============================================================================
// Evaluate a Compiled Formula
// ============================================================================

/**
 * Shared implementation for evaluateFormula and evaluateFormulaRaw.
 *
 * Handles key computation, cache lookup, circular reference detection,
 * expression evaluation, and result caching.
 */
function evaluateFormulaInner(
  compiled: CompiledFormula,
  ctx: EvalContext,
  session: EvalSession
): CachedResult {
  const inst = compiled.instance;
  const key = session.makeKey(inst.sheetName, inst.row, inst.col);

  // Check cache
  const cached = session.resultCache.get(key);
  if (cached !== undefined) {
    return cached;
  }

  // Circular reference detection
  if (session.evaluating.has(key)) {
    const fallback = session.circularFallback.get(key);
    const val = fallback !== undefined ? fallback : rvNumber(0);
    return { scalar: val, raw: val };
  }

  session.evaluating.add(key);
  const prevAddress = ctx.currentAddress;
  const prevSheet = ctx.currentSheet;
  const prevRecording = session.recordingKey;
  ctx.currentAddress = { sheet: inst.sheetName, row: inst.row, col: inst.col };
  ctx.currentSheet = inst.sheetName;
  // Enable runtime dependency recording for formulas with dynamic refs
  if (compiled.hasDynamicRefs) {
    session.recordingKey = key;
  }

  try {
    const result = evaluate(compiled.bound, ctx, session);
    const intersected = implicitIntersect(result, ctx);
    const scalar = dereferenceValue(intersected, ctx, session);
    const raw = dereferenceValue(result, ctx, session);
    const entry: CachedResult = { scalar, raw };
    session.resultCache.set(key, entry);
    return entry;
  } finally {
    session.evaluating.delete(key);
    ctx.currentAddress = prevAddress;
    ctx.currentSheet = prevSheet;
    session.recordingKey = prevRecording;
  }
}

/**
 * Evaluate a compiled formula and return its **scalar** result.
 *
 * This is the standard evaluation path for regular (non-array) formulas.
 * The result is:
 * 1. Evaluated from the bound expression tree
 * 2. Implicit-intersected to a single value/reference
 * 3. Dereferenced if it's a reference
 * 4. Cached for subsequent lookups by dependent formulas
 *
 * Use `evaluateFormulaRaw` instead when the full array result is needed
 * (dynamic array formulas, CSE formulas).
 */
export function evaluateFormula(
  compiled: CompiledFormula,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  return evaluateFormulaInner(compiled, ctx, session).scalar;
}

/**
 * Evaluate a compiled formula and return the **raw** (possibly array) result.
 *
 * This is the evaluation path for dynamic array and CSE formulas where
 * the full array shape must be preserved for spill/distribution.
 *
 * Semantics:
 * - Both scalar and raw forms are stored in `session.resultCache` as a
 *   `CachedResult{scalar, raw}`. Dependent scalar formulas see the scalar
 *   form; the materialize layer retrieves the raw form.
 * - The return value is the full dereferenced result — may be an ArrayValue
 *   with height > 1 or width > 1.
 */
export function evaluateFormulaRaw(
  compiled: CompiledFormula,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  return evaluateFormulaInner(compiled, ctx, session).raw;
}

// ============================================================================
// Binary Operations
// ============================================================================

function evaluateBinaryOp(
  op: string,
  leftExpr: BoundExpr,
  rightExpr: BoundExpr,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  // Intersection operator (whitespace between two refs). Must be handled
  // BEFORE dereferencing so we can inspect the reference areas.
  if (op === " ") {
    return evaluateIntersection(leftExpr, rightExpr, ctx, session);
  }

  const left = dereferenceValue(evaluate(leftExpr, ctx, session), ctx, session);
  const right = dereferenceValue(evaluate(rightExpr, ctx, session), ctx, session);

  const lIsArr = left.kind === RVKind.Array;
  const rIsArr = right.kind === RVKind.Array;

  if (lIsArr || rIsArr) {
    return broadcastBinaryOp(op, left, right);
  }

  return applyScalarBinaryOp(op, toScalar(left), toScalar(right));
}

/**
 * Excel's intersection operator — a whitespace character separating two
 * references (e.g. `A1:A10 B1:B10`).
 *
 * Semantics:
 * - Both operands must evaluate to single-area references. Otherwise the
 *   result is `#VALUE!` (matches Excel's behaviour when non-refs or
 *   multi-area refs are intersected).
 * - Intersection is the rectangle overlap of the two areas on the same
 *   sheet.
 * - If the areas do not overlap (or are on different sheets) the result
 *   is `#NULL!`, Excel's canonical "empty intersection" error.
 */
function evaluateIntersection(
  leftExpr: BoundExpr,
  rightExpr: BoundExpr,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const left = evaluate(leftExpr, ctx, session);
  const right = evaluate(rightExpr, ctx, session);

  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }
  if (left.kind !== RVKind.Reference || right.kind !== RVKind.Reference) {
    return ERRORS.VALUE;
  }
  if (left.areas.length !== 1 || right.areas.length !== 1) {
    return ERRORS.VALUE;
  }

  const la = left.areas[0];
  const ra = right.areas[0];

  if (la.sheet.toLowerCase() !== ra.sheet.toLowerCase()) {
    return ERRORS.NULL;
  }

  const top = Math.max(la.top, ra.top);
  const left_ = Math.max(la.left, ra.left);
  const bottom = Math.min(la.bottom, ra.bottom);
  const right_ = Math.min(la.right, ra.right);

  if (top > bottom || left_ > right_) {
    return ERRORS.NULL;
  }

  return rvRef(la.sheet, top, left_, bottom, right_);
}

function applyScalarBinaryOp(op: string, left: ScalarValue, right: ScalarValue): ScalarValue {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }

  // Concatenation
  if (op === "&") {
    const lStr = toStringRV(left);
    const rStr = toStringRV(right);
    return rvString(lStr + rStr);
  }

  // Comparison
  if (op === "=" || op === "<>" || op === "<" || op === ">" || op === "<=" || op === ">=") {
    return rvBoolean(compareScalars(left, right, op));
  }

  // Arithmetic
  const lNum = toNumberRV(left);
  if (isError(lNum)) {
    return lNum;
  }
  const rNum = toNumberRV(right);
  if (isError(rNum)) {
    return rNum;
  }

  let result: number;
  switch (op) {
    case "+":
      result = lNum.value + rNum.value;
      break;
    case "-":
      result = lNum.value - rNum.value;
      break;
    case "*":
      result = lNum.value * rNum.value;
      break;
    case "/":
      if (rNum.value === 0) {
        return ERRORS.DIV0;
      }
      result = lNum.value / rNum.value;
      break;
    case "^":
      result = Math.pow(lNum.value, rNum.value);
      break;
    default:
      return ERRORS.VALUE;
  }

  return !isFinite(result) ? ERRORS.NUM : rvNumber(result);
}

function compareScalars(left: ScalarValue, right: ScalarValue, op: string): boolean {
  // Normalize blanks
  const l =
    left.kind === RVKind.Blank
      ? right.kind === RVKind.String
        ? rvString("")
        : right.kind === RVKind.Boolean
          ? rvBoolean(false)
          : rvNumber(0)
      : left;
  const r =
    right.kind === RVKind.Blank
      ? left.kind === RVKind.String
        ? rvString("")
        : left.kind === RVKind.Boolean
          ? rvBoolean(false)
          : rvNumber(0)
      : right;

  let cmp: number;
  if (l.kind === r.kind) {
    if (l.kind === RVKind.String && r.kind === RVKind.String) {
      cmp = l.value.toLowerCase().localeCompare(r.value.toLowerCase());
    } else if (l.kind === RVKind.Number && r.kind === RVKind.Number) {
      cmp = l.value - r.value;
    } else if (l.kind === RVKind.Boolean && r.kind === RVKind.Boolean) {
      cmp = l.value === r.value ? 0 : l.value ? 1 : -1;
    } else {
      cmp = 0;
    }
  } else {
    const order = (v: ScalarValue): number => {
      if (v.kind === RVKind.Number) {
        return 0;
      }
      if (v.kind === RVKind.String) {
        return 1;
      }
      if (v.kind === RVKind.Boolean) {
        return 2;
      }
      return 3;
    };
    cmp = order(l) - order(r);
  }

  switch (op) {
    case "=":
      return cmp === 0;
    case "<>":
      return cmp !== 0;
    case "<":
      return cmp < 0;
    case ">":
      return cmp > 0;
    case "<=":
      return cmp <= 0;
    case ">=":
      return cmp >= 0;
    default:
      return false;
  }
}

function broadcastBinaryOp(op: string, left: RuntimeValue, right: RuntimeValue): RuntimeValue {
  const lArr = left.kind === RVKind.Array ? left : null;
  const rArr = right.kind === RVKind.Array ? right : null;

  const lRows = lArr ? lArr.height : 1;
  const lCols = lArr ? lArr.width : 1;
  const rRows = rArr ? rArr.height : 1;
  const rCols = rArr ? rArr.width : 1;

  const outRows = Math.max(lRows, rRows);
  const outCols = Math.max(lCols, rCols);

  if (
    (lRows !== 1 && rRows !== 1 && lRows !== rRows) ||
    (lCols !== 1 && rCols !== 1 && lCols !== rCols)
  ) {
    return ERRORS.VALUE;
  }

  const rows: ScalarValue[][] = [];
  for (let r = 0; r < outRows; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < outCols; c++) {
      const lR = lRows === 1 ? 0 : r;
      const lC = lCols === 1 ? 0 : c;
      const rR = rRows === 1 ? 0 : r;
      const rC = rCols === 1 ? 0 : c;

      const lVal: ScalarValue = lArr ? (lArr.rows[lR]?.[lC] ?? BLANK) : toScalar(left);
      const rVal: ScalarValue = rArr ? (rArr.rows[rR]?.[rC] ?? BLANK) : toScalar(right);

      row.push(applyScalarBinaryOp(op, lVal, rVal));
    }
    rows.push(row);
  }

  // Propagate origin metadata
  const originRow = lArr?.originRow ?? rArr?.originRow;
  const originCol = lArr?.originCol ?? rArr?.originCol;
  return rvArray(rows, originRow, originCol);
}

// ============================================================================
// Unary Operations
// ============================================================================

function evaluateUnaryOp(
  op: string,
  operandExpr: BoundExpr,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const rawVal = evaluate(operandExpr, ctx, session);

  // @ implicit intersection
  if (op === "@") {
    const intersected = implicitIntersect(rawVal, ctx);
    return dereferenceValue(intersected, ctx, session);
  }

  const val = dereferenceValue(rawVal, ctx, session);

  if (val.kind === RVKind.Array) {
    const rows: ScalarValue[][] = [];
    for (const row of val.rows) {
      rows.push(row.map(cell => applyScalarUnary(op, cell)));
    }
    return rvArray(rows, val.originRow, val.originCol);
  }

  return applyScalarUnary(op, toScalar(val));
}

function applyScalarUnary(op: string, val: ScalarValue): ScalarValue {
  if (isError(val)) {
    return val;
  }
  const n = toNumberRV(val);
  if (isError(n)) {
    return n;
  }
  switch (op) {
    case "-":
      return rvNumber(-n.value);
    case "+":
      return n;
    default:
      return ERRORS.VALUE;
  }
}

// ============================================================================
// Percent
// ============================================================================

function evaluatePercent(
  operandExpr: BoundExpr,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const val = dereferenceValue(evaluate(operandExpr, ctx, session), ctx, session);
  if (val.kind === RVKind.Array) {
    const rows: ScalarValue[][] = [];
    for (const row of val.rows) {
      rows.push(
        row.map(cell => {
          if (isError(cell)) {
            return cell;
          }
          const n = toNumberRV(cell);
          if (isError(n)) {
            return n;
          }
          return rvNumber(n.value / 100);
        })
      );
    }
    return rvArray(rows);
  }
  const scalar = toScalar(val);
  if (isError(scalar)) {
    return scalar;
  }
  const n = toNumberRV(scalar);
  if (isError(n)) {
    return n;
  }
  return rvNumber(n.value / 100);
}

// ============================================================================
// Function Call (Eager)
// ============================================================================

function evaluateCall(expr: BoundCall, ctx: EvalContext, session: EvalSession): RuntimeValue {
  ensureRegistryInitialized();

  // Reference functions: ROW, COLUMN, ROWS, COLUMNS
  const refResult = tryEvaluateRefFunction(expr.name, expr.args, ctx);
  if (refResult !== undefined) {
    return refResult;
  }

  // Evaluate all arguments eagerly and dereference references
  const args: RuntimeValue[] = expr.args.map(arg =>
    dereferenceValue(evaluate(arg, ctx, session), ctx, session)
  );

  // Look up function
  const desc = lookupFunction(expr.name);
  if (desc) {
    // Validate arity — produce #VALUE! for wrong argument count
    if (args.length < desc.minArity || args.length > desc.maxArity) {
      return ERRORS.VALUE;
    }
    // Context-aware overrides for functions that need evaluator state
    switch (expr.name) {
      case "SHEET": {
        // SHEET() → current sheet number; SHEET(ref) → sheet number of ref
        if (args.length === 0) {
          const idx = ctx.snapshot.worksheets.findIndex(
            ws => ws.name.toLowerCase() === ctx.currentSheet.toLowerCase()
          );
          return rvNumber(idx >= 0 ? idx + 1 : 1);
        }
        return desc.invoke(args);
      }
      case "SHEETS": {
        // SHEETS() → total sheet count
        if (args.length === 0) {
          return rvNumber(ctx.snapshot.worksheets.length);
        }
        return desc.invoke(args);
      }
      case "ISFORMULA": {
        // ISFORMULA requires a reference argument. When the raw argument is
        // a CellRef/AreaRef we can look up the underlying cell's formulaKind.
        // Any other shape (literal, computed value, etc.) yields #N/A per
        // Excel's behavior for non-reference arguments.
        const raw = expr.args[0];
        if (raw && raw.kind === BoundExprKind.CellRef) {
          const ws = ctx.snapshot.worksheetsByName.get(raw.sheet.toLowerCase());
          if (!ws) {
            return ERRORS.REF;
          }
          const cell = ws.cells.get(snapshotCellKey(raw.row, raw.col));
          return rvBoolean(cell !== undefined && cell.formulaKind !== "none");
        }
        if (raw && raw.kind === BoundExprKind.AreaRef) {
          const ws = ctx.snapshot.worksheetsByName.get(raw.sheet.toLowerCase());
          if (!ws) {
            return ERRORS.REF;
          }
          // ISFORMULA on an area ref inspects the top-left cell.
          const cell = ws.cells.get(snapshotCellKey(raw.top, raw.left));
          return rvBoolean(cell !== undefined && cell.formulaKind !== "none");
        }
        return ERRORS.NA;
      }
      case "FORMULATEXT": {
        // FORMULATEXT returns the formula source text at the referenced cell,
        // or #N/A if the cell has no formula. Non-reference arguments also
        // yield #N/A.
        const raw = expr.args[0];
        if (raw && raw.kind === BoundExprKind.CellRef) {
          const ws = ctx.snapshot.worksheetsByName.get(raw.sheet.toLowerCase());
          if (!ws) {
            return ERRORS.REF;
          }
          const cell = ws.cells.get(snapshotCellKey(raw.row, raw.col));
          if (cell && cell.formulaKind !== "none" && cell.formula !== undefined) {
            return rvString(`=${cell.formula}`);
          }
          return ERRORS.NA;
        }
        if (raw && raw.kind === BoundExprKind.AreaRef) {
          const ws = ctx.snapshot.worksheetsByName.get(raw.sheet.toLowerCase());
          if (!ws) {
            return ERRORS.REF;
          }
          const cell = ws.cells.get(snapshotCellKey(raw.top, raw.left));
          if (cell && cell.formulaKind !== "none" && cell.formula !== undefined) {
            return rvString(`=${cell.formula}`);
          }
          return ERRORS.NA;
        }
        return ERRORS.NA;
      }
      default:
        return desc.invoke(args);
    }
  }

  // Check if name resolves to a lambda (defined name or local binding)
  const lambda = resolveLambdaName(expr.name, args, ctx, session);
  if (lambda !== undefined) {
    return lambda;
  }

  return ERRORS.NAME;
}

// ============================================================================
// Special Form Call (Lazy)
// ============================================================================

function evaluateSpecialCall(
  expr: BoundSpecialCall,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  switch (expr.name) {
    case "IF":
      return evaluateIF(expr.args, ctx, session);
    case "IFERROR":
      return evaluateIFERROR(expr.args, ctx, session);
    case "IFNA":
      return evaluateIFNA(expr.args, ctx, session);
    case "IFS":
      return evaluateIFS(expr.args, ctx, session);
    case "SWITCH":
      return evaluateSWITCH(expr.args, ctx, session);
    case "CHOOSE":
      return evaluateCHOOSE(expr.args, ctx, session);
    case "LET":
      return evaluateLET(expr.args, ctx, session);
    case "LAMBDA":
      return evaluateLAMBDA(expr.args, ctx);
    case "INDIRECT":
      return evaluateINDIRECT(expr.args, ctx, session);
    case "OFFSET":
      return evaluateOFFSET(expr.args, ctx, session);
    case "MAP":
    case "REDUCE":
    case "SCAN":
    case "MAKEARRAY":
    case "BYROW":
    case "BYCOL":
      return evaluateHigherOrder(expr.name, expr.args, ctx, session);
    default:
      return ERRORS.VALUE;
  }
}

// ============================================================================
// Special Forms Implementation
// ============================================================================

function evaluateIF(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const cond = toScalar(evalDeref(args[0], ctx, session));
  if (isError(cond)) {
    return cond;
  }
  const bool = toBooleanRV(cond);
  if (isError(bool)) {
    return bool;
  }
  if (bool.value) {
    return evaluate(args[1], ctx, session);
  }
  return args.length > 2 ? evaluate(args[2], ctx, session) : rvBoolean(false);
}

function evaluateIFERROR(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const val = evalDeref(args[0], ctx, session);
  // Array: element-wise replacement of errors with the value-if-error
  if (val.kind === RVKind.Array) {
    const replacement = evalDeref(args[1], ctx, session);
    const replaceScalar = topLeft(replacement);
    const rows: ScalarValue[][] = [];
    for (const row of val.rows) {
      const newRow: ScalarValue[] = [];
      for (const cell of row) {
        newRow.push(cell.kind === RVKind.Error ? replaceScalar : cell);
      }
      rows.push(newRow);
    }
    return rvArray(rows);
  }
  return isError(val) ? evaluate(args[1], ctx, session) : val;
}

function evaluateIFNA(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const val = evalDeref(args[0], ctx, session);
  // Array: element-wise replacement of #N/A with the value-if-na
  if (val.kind === RVKind.Array) {
    const replacement = evalDeref(args[1], ctx, session);
    const replaceScalar = topLeft(replacement);
    const rows: ScalarValue[][] = [];
    for (const row of val.rows) {
      const newRow: ScalarValue[] = [];
      for (const cell of row) {
        newRow.push(cell.kind === RVKind.Error && cell.code === "#N/A" ? replaceScalar : cell);
      }
      rows.push(newRow);
    }
    return rvArray(rows);
  }
  return isError(val) && val.code === "#N/A" ? evaluate(args[1], ctx, session) : val;
}

function evaluateIFS(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  for (let i = 0; i < args.length - 1; i += 2) {
    const cond = toScalar(evalDeref(args[i], ctx, session));
    if (isError(cond)) {
      return cond;
    }
    const bool = toBooleanRV(cond);
    if (isError(bool)) {
      return bool;
    }
    if (bool.value) {
      return evaluate(args[i + 1], ctx, session);
    }
  }
  return ERRORS.NA;
}

function evaluateSWITCH(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  const expr = toScalar(evalDeref(args[0], ctx, session));
  for (let i = 1; i < args.length - 1; i += 2) {
    const caseVal = toScalar(evalDeref(args[i], ctx, session));
    if (scalarEquals(expr, caseVal)) {
      return evaluate(args[i + 1], ctx, session);
    }
  }
  if (args.length % 2 === 0) {
    return evaluate(args[args.length - 1], ctx, session);
  }
  return ERRORS.NA;
}

function evaluateCHOOSE(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const idxVal = toScalar(evalDeref(args[0], ctx, session));
  if (isError(idxVal)) {
    return idxVal;
  }
  const num = toNumberRV(idxVal);
  if (isError(num)) {
    return num;
  }
  const idx = Math.floor(num.value);
  if (idx < 1 || idx >= args.length) {
    return ERRORS.VALUE;
  }
  return evaluate(args[idx], ctx, session);
}

function evaluateLET(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3 || args.length % 2 !== 1) {
    return ERRORS.VALUE;
  }
  const prevBindings = ctx.localBindings;
  const newBindings = new Map<string, RuntimeValue>(prevBindings);

  try {
    const pairCount = (args.length - 1) / 2;
    for (let i = 0; i < pairCount; i++) {
      const nameExpr = args[i * 2];
      const valueExpr = args[i * 2 + 1];
      if (nameExpr.kind !== BoundExprKind.NameExpr) {
        return ERRORS.VALUE;
      }
      ctx.localBindings = newBindings;
      const val = evaluate(valueExpr, ctx, session);
      newBindings.set(nameExpr.upperName, val);
    }

    ctx.localBindings = newBindings;
    return evaluate(args[args.length - 1], ctx, session);
  } finally {
    ctx.localBindings = prevBindings;
  }
}

function evaluateLAMBDA(args: readonly BoundExpr[], ctx: EvalContext): RuntimeValue {
  if (args.length < 1) {
    return ERRORS.VALUE;
  }
  const paramExprs = args.slice(0, -1);
  const bodyExpr = args[args.length - 1];
  const params: string[] = [];
  for (const p of paramExprs) {
    if (p.kind !== BoundExprKind.NameExpr) {
      return ERRORS.VALUE;
    }
    params.push(p.upperName);
  }
  return rvLambda(params, bodyExpr, ctx.localBindings ? new Map(ctx.localBindings) : undefined);
}

function evaluateINDIRECT(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 1) {
    return ERRORS.VALUE;
  }
  const refArg = evalDeref(args[0], ctx, session);
  const refText = toStringRV(topLeft(refArg));
  if (!refText) {
    return ERRORS.REF;
  }

  let a1 = true;
  if (args.length >= 2) {
    const a1Val = topLeft(evalDeref(args[1], ctx, session));
    a1 =
      !(a1Val.kind === RVKind.Boolean && !a1Val.value) &&
      !(a1Val.kind === RVKind.Number && a1Val.value === 0);
  }

  if (!a1) {
    // R1C1 — delegate to runtime R1C1 parser
    return resolveR1C1(refText, ctx, session);
  }

  // A1 style — parse and bind at runtime
  try {
    const cacheKey = `__INDIRECT__${ctx.currentSheet}__${refText}`;
    let bound = ctx.astCache.get(cacheKey);
    if (!bound) {
      const tokens = tokenize(refText);
      const ast = parse(tokens);
      const bindCtx: BindingContext = { snapshot: ctx.snapshot, currentSheet: ctx.currentSheet };
      bound = bind(ast, bindCtx);
      ctx.astCache.set(cacheKey, bound);
    }
    return evaluate(bound, ctx, session);
  } catch {
    return ERRORS.REF;
  }
}

function evaluateOFFSET(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  const refExpr = args[0];
  let baseRow: number;
  let baseCol: number;
  let baseSheet: string;

  if (refExpr.kind === BoundExprKind.CellRef) {
    baseRow = refExpr.row;
    baseCol = refExpr.col;
    baseSheet = refExpr.sheet;
  } else if (refExpr.kind === BoundExprKind.AreaRef) {
    baseRow = refExpr.top;
    baseCol = refExpr.left;
    baseSheet = refExpr.sheet;
  } else {
    return ERRORS.VALUE;
  }

  const rowsVal = topLeft(evalDeref(args[1], ctx, session));
  const rowsNum = toNumberRV(rowsVal);
  if (isError(rowsNum)) {
    return rowsNum;
  }
  const colsVal = topLeft(evalDeref(args[2], ctx, session));
  const colsNum = toNumberRV(colsVal);
  if (isError(colsNum)) {
    return colsNum;
  }

  const newRow = baseRow + rowsNum.value;
  const newCol = baseCol + colsNum.value;
  if (newRow < 1 || newCol < 1) {
    return ERRORS.REF;
  }

  let height = 1;
  let width = 1;
  if (args.length > 3) {
    const h = toNumberRV(topLeft(evalDeref(args[3], ctx, session)));
    if (isError(h)) {
      return h;
    }
    height = Math.trunc(h.value);
    if (height === 0) {
      return ERRORS.REF;
    }
  }
  if (args.length > 4) {
    const w = toNumberRV(topLeft(evalDeref(args[4], ctx, session)));
    if (isError(w)) {
      return w;
    }
    width = Math.trunc(w.value);
    if (width === 0) {
      return ERRORS.REF;
    }
  }

  // Resolve range coordinates — negative height/width extend upward/leftward
  let top = newRow;
  let bottom = newRow + height - 1;
  if (height < 0) {
    top = newRow + height + 1;
    bottom = newRow;
  }
  let left = newCol;
  let right = newCol + width - 1;
  if (width < 0) {
    left = newCol + width + 1;
    right = newCol;
  }

  if (top < 1 || left < 1) {
    return ERRORS.REF;
  }

  if (top === bottom && left === right) {
    return getCellValue(baseSheet, top, left, ctx, session);
  }

  return buildRangeArray(ctx, session, baseSheet, top, left, bottom, right);
}

// ============================================================================
// Higher-Order Functions
// ============================================================================

function evaluateHigherOrder(
  name: string,
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  switch (name) {
    case "MAP":
      return evaluateMAP(args, ctx, session);
    case "REDUCE":
      return evaluateREDUCE(args, ctx, session);
    case "SCAN":
      return evaluateSCAN(args, ctx, session);
    case "MAKEARRAY":
      return evaluateMAKEARRAY(args, ctx, session);
    case "BYROW":
      return evaluateBYROW(args, ctx, session);
    case "BYCOL":
      return evaluateBYCOL(args, ctx, session);
    default:
      return ERRORS.VALUE;
  }
}

function invokeLambda(
  lambda: LambdaValue,
  lambdaArgs: RuntimeValue[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (lambdaArgs.length !== lambda.params.length) {
    return ERRORS.VALUE;
  }
  const prevBindings = ctx.localBindings;
  const newBindings = new Map<string, RuntimeValue>(lambda.closureBindings);
  for (let i = 0; i < lambda.params.length; i++) {
    newBindings.set(lambda.params[i], lambdaArgs[i]);
  }
  ctx.localBindings = newBindings;
  try {
    return dereferenceValue(evaluate(lambda.body, ctx, session), ctx, session);
  } finally {
    ctx.localBindings = prevBindings;
  }
}

function evaluateMAP(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const arrVal = evalDeref(args[0], ctx, session);
  const lambdaVal = evalDeref(args[args.length - 1], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  if (arrVal.kind !== RVKind.Array) {
    return invokeLambda(lambdaVal, [arrVal], ctx, session);
  }
  const rows: ScalarValue[][] = [];
  for (const row of arrVal.rows) {
    const outRow: ScalarValue[] = [];
    for (const cell of row) {
      outRow.push(toScalar(invokeLambda(lambdaVal, [cell], ctx, session)));
    }
    rows.push(outRow);
  }
  return rvArray(rows);
}

function evaluateREDUCE(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  let acc = evalDeref(args[0], ctx, session);
  const arrVal = evalDeref(args[1], ctx, session);
  const lambdaVal = evalDeref(args[2], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  if (arrVal.kind === RVKind.Array) {
    for (const row of arrVal.rows) {
      for (const cell of row) {
        acc = invokeLambda(lambdaVal, [topLeft(acc), cell], ctx, session);
      }
    }
  }
  return acc;
}

function evaluateSCAN(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  let acc = evalDeref(args[0], ctx, session);
  const arrVal = evalDeref(args[1], ctx, session);
  const lambdaVal = evalDeref(args[2], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  const rows: ScalarValue[][] = [];
  if (arrVal.kind === RVKind.Array) {
    for (const row of arrVal.rows) {
      const outRow: ScalarValue[] = [];
      for (const cell of row) {
        acc = invokeLambda(lambdaVal, [topLeft(acc), cell], ctx, session);
        outRow.push(toScalar(acc));
      }
      rows.push(outRow);
    }
  }
  return rows.length > 0 ? rvArray(rows) : ERRORS.CALC;
}

function evaluateMAKEARRAY(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 3) {
    return ERRORS.VALUE;
  }
  const rowsNum = toNumberRV(topLeft(evalDeref(args[0], ctx, session)));
  if (isError(rowsNum)) {
    return rowsNum;
  }
  const colsNum = toNumberRV(topLeft(evalDeref(args[1], ctx, session)));
  if (isError(colsNum)) {
    return colsNum;
  }
  const lambdaVal = evalDeref(args[2], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  const rows: ScalarValue[][] = [];
  for (let r = 1; r <= rowsNum.value; r++) {
    const outRow: ScalarValue[] = [];
    for (let c = 1; c <= colsNum.value; c++) {
      outRow.push(toScalar(invokeLambda(lambdaVal, [rvNumber(r), rvNumber(c)], ctx, session)));
    }
    rows.push(outRow);
  }
  return rvArray(rows);
}

function evaluateBYROW(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const arrVal = evalDeref(args[0], ctx, session);
  const lambdaVal = evalDeref(args[1], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  if (arrVal.kind !== RVKind.Array) {
    return invokeLambda(lambdaVal, [rvArray([[toScalar(arrVal)]])], ctx, session);
  }
  const rows: ScalarValue[][] = [];
  for (const row of arrVal.rows) {
    const rowArr = rvArray([row.map(c => c)]);
    rows.push([toScalar(invokeLambda(lambdaVal, [rowArr], ctx, session))]);
  }
  return rvArray(rows);
}

function evaluateBYCOL(
  args: readonly BoundExpr[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  if (args.length < 2) {
    return ERRORS.VALUE;
  }
  const arrVal = evalDeref(args[0], ctx, session);
  const lambdaVal = evalDeref(args[1], ctx, session);
  if (!isLambda(lambdaVal)) {
    return ERRORS.VALUE;
  }
  if (arrVal.kind !== RVKind.Array) {
    return invokeLambda(lambdaVal, [rvArray([[toScalar(arrVal)]])], ctx, session);
  }
  const numCols = arrVal.width;
  const outRow: ScalarValue[] = [];
  for (let c = 0; c < numCols; c++) {
    const colArr = rvArray(arrVal.rows.map(row => [row[c]]));
    outRow.push(toScalar(invokeLambda(lambdaVal, [colArr], ctx, session)));
  }
  return rvArray([outRow]);
}

// ============================================================================
// Reference Functions (ROW, COLUMN, ROWS, COLUMNS)
// ============================================================================

function tryEvaluateRefFunction(
  name: string,
  args: readonly BoundExpr[],
  ctx: EvalContext
): RuntimeValue | undefined {
  switch (name) {
    case "ROW":
      if (args.length === 0) {
        return ctx.currentAddress ? rvNumber(ctx.currentAddress.row) : ERRORS.VALUE;
      }
      if (args[0].kind === BoundExprKind.CellRef) {
        return rvNumber(args[0].row);
      }
      if (args[0].kind === BoundExprKind.AreaRef) {
        return rvNumber(args[0].top);
      }
      return undefined;
    case "COLUMN":
      if (args.length === 0) {
        return ctx.currentAddress ? rvNumber(ctx.currentAddress.col) : ERRORS.VALUE;
      }
      if (args[0].kind === BoundExprKind.CellRef) {
        return rvNumber(args[0].col);
      }
      if (args[0].kind === BoundExprKind.AreaRef) {
        return rvNumber(args[0].left);
      }
      return undefined;
    case "ROWS":
      if (args.length > 0 && args[0].kind === BoundExprKind.AreaRef) {
        return rvNumber(args[0].bottom - args[0].top + 1);
      }
      if (args.length > 0 && args[0].kind === BoundExprKind.CellRef) {
        return rvNumber(1);
      }
      return undefined;
    case "COLUMNS":
      if (args.length > 0 && args[0].kind === BoundExprKind.AreaRef) {
        return rvNumber(args[0].right - args[0].left + 1);
      }
      if (args.length > 0 && args[0].kind === BoundExprKind.CellRef) {
        return rvNumber(1);
      }
      return undefined;
    default:
      return undefined;
  }
}

// ============================================================================
// Name Expression
// ============================================================================

function evaluateNameExpr(
  expr: BoundNameExpr,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  // Check local bindings first (LET variables)
  if (ctx.localBindings?.has(expr.upperName)) {
    return ctx.localBindings.get(expr.upperName)!;
  }

  // Try snapshot defined name resolution (for formula-based names).
  // Respects scope precedence: sheet-local > workbook-global.
  const dn = resolveDefinedNameFromSnapshot(ctx.snapshot.definedNames, expr.name, ctx.currentSheet);
  if (dn && dn.ranges.length > 0) {
    if (dn.ranges.length > 1) {
      return ERRORS.VALUE;
    }
    const rangeStr = dn.ranges[0];
    const parsed = parseDefinedNameRange(rangeStr);
    if (parsed) {
      if (parsed.startRow === parsed.endRow && parsed.startCol === parsed.endCol) {
        return getCellValue(parsed.sheet, parsed.startRow, parsed.startCol, ctx, session);
      }
      return buildRangeArray(
        ctx,
        session,
        parsed.sheet,
        Math.min(parsed.startRow, parsed.endRow),
        Math.min(parsed.startCol, parsed.endCol),
        Math.max(parsed.startRow, parsed.endRow),
        Math.max(parsed.startCol, parsed.endCol)
      );
    }
    // Formula expression — parse and evaluate via shared helper
    const nameResult = evaluateFormulaName(expr.upperName, rangeStr, ctx, session);
    return nameResult ?? ERRORS.NAME;
  }

  return ERRORS.NAME;
}

// ============================================================================
// Formula-based Defined Name Evaluation (shared helper)
// ============================================================================

/**
 * Evaluate a formula-based defined name expression, with caching.
 * Cache key includes name + sheet + cell address to handle position-dependent
 * formulas like ROW()/COLUMN().
 *
 * Returns the evaluated result, or undefined if parsing/evaluation fails.
 */
function evaluateFormulaName(
  upperName: string,
  formulaExpr: string,
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue | undefined {
  const addr = ctx.currentAddress;
  const cacheKey = addr
    ? `__NAME__${upperName}__${ctx.currentSheet}__${addr.row}:${addr.col}`
    : `__NAME__${upperName}__${ctx.currentSheet}`;
  const cached = session.nameCache.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }
  try {
    const tokens = tokenize(formulaExpr);
    const ast = parse(tokens);
    const bindCtx: BindingContext = { snapshot: ctx.snapshot, currentSheet: ctx.currentSheet };
    const bound = bind(ast, bindCtx);
    const result = evaluate(bound, ctx, session);
    session.nameCache.set(cacheKey, result);
    return result;
  } catch {
    return undefined;
  }
}

// ============================================================================
// Lambda Expression
// ============================================================================

function evaluateLambdaExpr(expr: BoundLambda, ctx: EvalContext): RuntimeValue {
  return rvLambda(
    [...expr.params],
    expr.body,
    ctx.localBindings ? new Map(ctx.localBindings) : undefined
  );
}

// ============================================================================
// Structured Reference (runtime resolution)
// ============================================================================

function evaluateStructuredRef(
  expr: { tableName: string; columns: readonly string[]; specials: readonly string[] },
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const snapshot = ctx.snapshot;
  const addr = ctx.currentAddress;

  // Find the table
  let tableName = expr.tableName;
  let tableSheet: string | null = null;
  let tableInfo: {
    name: string;
    columns: readonly { name: string }[];
    topLeft: { row: number; col: number };
    dataRowCount: number;
    hasHeaderRow: boolean;
    hasTotalsRow: boolean;
  } | null = null;

  if (tableName === "") {
    // Implicit table — find the table containing the current cell
    if (!addr) {
      return ERRORS.REF;
    }
    for (const ws of snapshot.worksheets) {
      if (ws.name !== addr.sheet) {
        continue;
      }
      for (const t of ws.tables) {
        const g = buildTableGeometry(t);
        const width = t.columns.length;
        if (
          addr.row >= g.dataRowStart &&
          addr.row <= g.dataRowEnd &&
          addr.col >= t.topLeft.col &&
          addr.col < t.topLeft.col + width
        ) {
          tableInfo = t;
          tableSheet = ws.name;
          tableName = t.name;
          break;
        }
      }
      if (tableInfo) {
        break;
      }
    }
  } else {
    // Named table — use the pre-built index for O(1) lookup
    const resolved = snapshot.tablesByName.get(tableName.toLowerCase());
    if (resolved) {
      tableInfo = resolved.table;
      tableSheet = resolved.sheetName;
    }
  }

  if (!tableInfo || !tableSheet) {
    return ERRORS.REF;
  }

  const geo = buildTableGeometry(tableInfo);

  // Strict column resolution — unknown column names surface as #REF!
  const colRange = resolveStructuredRefColumns(expr.columns, tableInfo, "strict");
  if (colRange === "error") {
    return ERRORS.REF;
  }

  const rowRange = resolveStructuredRefRows(expr.specials, geo);

  let rowTop: number;
  let rowBottom: number;
  if (rowRange === "error") {
    return ERRORS.REF;
  } else if (rowRange === "thisRow") {
    if (addr) {
      rowTop = addr.row;
      rowBottom = addr.row;
    } else {
      return ERRORS.VALUE;
    }
  } else {
    rowTop = rowRange.rowTop;
    rowBottom = rowRange.rowBottom;
  }

  // Single cell — return as single-cell ReferenceValue
  if (rowTop === rowBottom && colRange.colLeft === colRange.colRight) {
    return rvCellRef(tableSheet, rowTop, colRange.colLeft);
  }

  // Range — return as area ReferenceValue
  return rvRef(tableSheet, rowTop, colRange.colLeft, rowBottom, colRange.colRight);
}

// ============================================================================
// Array Literal
// ============================================================================

function evaluateArrayLiteral(
  expr: { rows: readonly (readonly BoundExpr[])[] },
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue {
  const rows: ScalarValue[][] = [];
  for (const row of expr.rows) {
    const evalRow: ScalarValue[] = [];
    for (const elem of row) {
      evalRow.push(toScalar(evalDeref(elem, ctx, session)));
    }
    rows.push(evalRow);
  }
  return rvArray(rows);
}

// ============================================================================
// Implicit Intersection
// ============================================================================

export function implicitIntersect(val: RuntimeValue, ctx: EvalContext): RuntimeValue {
  if (isScalar(val)) {
    return val;
  }
  if (val.kind === RVKind.MissingArg) {
    return BLANK;
  }
  if (val.kind === RVKind.Lambda) {
    return val;
  }
  if (val.kind === RVKind.Reference) {
    // Implicit intersection on a reference: resolve using formula cell's row/col
    if (val.areas.length === 0) {
      return BLANK;
    }
    const area = val.areas[0];
    const isSingleCell = area.top === area.bottom && area.left === area.right;
    if (isSingleCell) {
      return val; // Single cell ref — keep as-is, will be dereferenced at use site
    }
    // Multi-cell reference — apply implicit intersection
    const addr = ctx.currentAddress;
    if (!addr) {
      return val;
    }
    // Single column — pick row
    if (area.left === area.right && addr.row >= area.top && addr.row <= area.bottom) {
      return rvCellRef(area.sheet, addr.row, area.left);
    }
    // Single row — pick column
    if (area.top === area.bottom && addr.col >= area.left && addr.col <= area.right) {
      return rvCellRef(area.sheet, area.top, addr.col);
    }
    // Both row and col
    if (
      addr.row >= area.top &&
      addr.row <= area.bottom &&
      addr.col >= area.left &&
      addr.col <= area.right
    ) {
      return rvCellRef(area.sheet, addr.row, addr.col);
    }
    return val;
  }
  if (val.kind !== RVKind.Array) {
    return val;
  }

  const arr = val;
  if (arr.height === 0 || arr.width === 0) {
    return BLANK;
  }
  if (arr.height === 1 && arr.width === 1) {
    return arr.rows[0][0];
  }

  const addr = ctx.currentAddress;
  if (!addr) {
    return arr.rows[0][0];
  }

  // Single row — pick column by offset
  if (arr.height === 1) {
    if (arr.originCol !== undefined) {
      const colIdx = addr.col - arr.originCol;
      if (colIdx >= 0 && colIdx < arr.width) {
        return arr.rows[0][colIdx];
      }
    }
    return arr.rows[0][0];
  }

  // Single column — pick row by offset
  if (arr.width === 1) {
    if (arr.originRow !== undefined) {
      const rowIdx = addr.row - arr.originRow;
      if (rowIdx >= 0 && rowIdx < arr.height) {
        return arr.rows[rowIdx][0];
      }
    }
    return arr.rows[0][0];
  }

  // Multi-row, multi-column
  if (arr.originRow !== undefined && arr.originCol !== undefined) {
    const rowIdx = addr.row - arr.originRow;
    const colIdx = addr.col - arr.originCol;
    if (rowIdx >= 0 && rowIdx < arr.height && colIdx >= 0 && colIdx < arr.width) {
      return arr.rows[rowIdx][colIdx];
    }
  }

  return arr.rows[0][0];
}

// ============================================================================
// R1C1 Reference Resolution
// ============================================================================

function resolveR1C1(refText: string, ctx: EvalContext, session: EvalSession): RuntimeValue {
  const upper = refText.toUpperCase().trim();
  // Check for range separator
  let depth = 0;
  let sepIdx = -1;
  for (let i = 0; i < upper.length; i++) {
    if (upper[i] === "[") {
      depth++;
    } else if (upper[i] === "]") {
      depth--;
    } else if (upper[i] === ":" && depth === 0) {
      sepIdx = i;
      break;
    }
  }

  if (sepIdx !== -1) {
    const s = parseR1C1Single(upper.slice(0, sepIdx), ctx);
    const e = parseR1C1Single(upper.slice(sepIdx + 1), ctx);
    if (!s || !e) {
      return ERRORS.REF;
    }
    const top = Math.min(s.row, e.row);
    const bottom = Math.max(s.row, e.row);
    const left = Math.min(s.col, e.col);
    const right = Math.max(s.col, e.col);
    return rvRef(ctx.currentSheet, top, left, bottom, right);
  }

  const ref = parseR1C1Single(upper, ctx);
  if (!ref || ref.row < 1 || ref.col < 1) {
    return ERRORS.REF;
  }
  return rvCellRef(ctx.currentSheet, ref.row, ref.col);
}

function parseR1C1Single(text: string, ctx: EvalContext): { row: number; col: number } | null {
  const re = /^R(\[(-?\d+)\]|(\d+))C(\[(-?\d+)\]|(\d+))$/;
  const m = re.exec(text);
  if (!m) {
    return null;
  }
  const addr = ctx.currentAddress;
  const row = m[2] !== undefined ? (addr?.row ?? 1) + parseInt(m[2], 10) : parseInt(m[3], 10);
  const col = m[5] !== undefined ? (addr?.col ?? 1) + parseInt(m[5], 10) : parseInt(m[6], 10);
  return { row, col };
}

// ============================================================================
// Helpers
// ============================================================================

function toScalar(v: RuntimeValue): ScalarValue {
  return topLeft(v);
}

/**
 * Evaluate a BoundExpr and dereference any resulting ReferenceValue.
 * Use this whenever a concrete (non-reference) value is needed.
 */
function evalDeref(expr: BoundExpr, ctx: EvalContext, session: EvalSession): RuntimeValue {
  return dereferenceValue(evaluate(expr, ctx, session), ctx, session);
}

function scalarEquals(a: ScalarValue, b: ScalarValue): boolean {
  if (a.kind !== b.kind) {
    return false;
  }
  switch (a.kind) {
    case RVKind.Number:
      return a.value === (b as NumberValue).value;
    case RVKind.String:
      return a.value.toLowerCase() === (b as { value: string }).value.toLowerCase();
    case RVKind.Boolean:
      return a.value === (b as { value: boolean }).value;
    case RVKind.Blank:
      return true;
    default:
      return false;
  }
}

function resolveLambdaName(
  name: string,
  args: RuntimeValue[],
  ctx: EvalContext,
  session: EvalSession
): RuntimeValue | undefined {
  // Check local bindings
  if (ctx.localBindings?.has(name)) {
    const val = ctx.localBindings.get(name)!;
    if (isLambda(val)) {
      return invokeLambda(val, args, ctx, session);
    }
  }

  // Check defined names that resolve to lambdas (via snapshot, scope-aware)
  const dn = resolveDefinedNameFromSnapshot(ctx.snapshot.definedNames, name, ctx.currentSheet);
  if (dn && dn.ranges.length === 1) {
    const rangeStr = dn.ranges[0];
    const parsed = parseDefinedNameRange(rangeStr);
    if (parsed && parsed.startRow === parsed.endRow && parsed.startCol === parsed.endCol) {
      const cellVal = getCellValue(parsed.sheet, parsed.startRow, parsed.startCol, ctx, session);
      if (isLambda(cellVal)) {
        return invokeLambda(cellVal, args, ctx, session);
      }
    }
    // Formula-based name — evaluate via shared helper
    if (!parsed) {
      const nameVal = evaluateFormulaName(name.toUpperCase(), rangeStr, ctx, session);
      if (nameVal !== undefined && isLambda(nameVal)) {
        return invokeLambda(nameVal, args, ctx, session);
      }
    }
  }

  return undefined;
}
