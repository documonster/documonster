/**
 * Formula Evaluator
 *
 * Evaluates an AST node against a workbook context to produce a result value.
 * Handles cell references, range references, cross-sheet references,
 * operators, and function calls.
 *
 * All evaluation state (circular-reference stack, memoization cache) is held
 * in an `EvalSession` instance to avoid global state pollution.
 */

import { colCache } from "@excel/utils/col-cache";

import {
  FUNCTIONS,
  isError,
  toNumber,
  makeError,
  type CalcArray,
  type CalcValue
} from "./formula-functions";
import {
  NodeType,
  parse,
  type AstNode,
  type CellRefNode,
  type RangeRefNode,
  type ColRangeRefNode,
  type RowRangeRefNode,
  type NameNode
} from "./formula-parser";
import { tokenize } from "./formula-tokenizer";

// ============================================================================
// Evaluator Context
// ============================================================================

/**
 * Resolved named range — either a single cell or a rectangular range.
 */
export interface ResolvedName {
  sheet?: string;
  /** Single cell (when startRow === endRow && startCol === endCol) or range */
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/**
 * Resolved structured reference — a rectangular region in the worksheet.
 */
export interface ResolvedStructuredRef {
  sheetName: string;
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * Interface for accessing cell values during evaluation.
 * This decouples the evaluator from the Workbook/Worksheet classes.
 */
export interface EvalContext {
  /**
   * Get the value of a cell. For formula cells this must recursively evaluate
   * the formula (using the session) and persist the result.
   * @param sheetName - Sheet name (undefined = current sheet)
   * @param row - 1-based row number
   * @param col - 1-based column number
   */
  getCellValue(sheetName: string | undefined, row: number, col: number): CalcValue;

  /**
   * Resolve a defined name to its cell/range coordinates.
   * Returns null if the name is not defined in the workbook.
   */
  resolveName?(name: string): ResolvedName | null;

  /**
   * Resolve a defined name that is a formula expression (e.g. =LAMBDA(x,y,x+y)).
   * Returns the evaluated CalcValue/CalcArray if the name maps to a formula,
   * or undefined if it's not a formula-based name.
   */
  resolveNameValue?(name: string): CalcValue | CalcArray | undefined;

  /**
   * Resolve a structured reference (Table[Column]) to a cell range.
   * @param tableName - Table name (empty string = infer from current cell)
   * @param columns - Column name(s)
   * @param specials - Special items (#Headers, #Data, #Totals, #All, #This Row)
   */
  resolveStructuredRef?(
    tableName: string,
    columns: string[],
    specials: string[]
  ): ResolvedStructuredRef | null;

  /**
   * Get the used range dimensions of a worksheet.
   * Returns { top, bottom, left, right } (1-based) or null if sheet is empty/not found.
   */
  getSheetDimensions?(
    sheetName: string | undefined
  ): { top: number; bottom: number; left: number; right: number } | null;

  /**
   * Get all worksheet names in workbook order.
   * Required for 3D references (e.g. Sheet1:Sheet3!A1).
   */
  getSheetNames?(): string[];

  /** The current sheet name (for relative references) */
  currentSheet: string;

  /** Current cell address being evaluated (for circular reference detection) */
  currentAddress?: { sheet: string; row: number; col: number };

  /**
   * Local variable bindings from LET expressions.
   * Maps variable name (uppercase) → value.
   */
  localBindings?: Map<string, CalcValue | CalcArray>;
}

// ============================================================================
// Lambda Value — represents an unevaluated LAMBDA function
// ============================================================================

/**
 * A lambda value returned by evaluating a LAMBDA() expression.
 * This is a tagged object that carries the parameter names and the body AST
 * so it can be invoked later with actual arguments.
 *
 * Note: This is NOT a CalcValue — it exists outside the normal value domain
 * and is only used internally by the evaluator for LAMBDA/LET patterns.
 */
export interface LambdaValue {
  readonly __lambda: true;
  /** Parameter names (uppercase) */
  readonly params: string[];
  /** The body AST node to evaluate when the lambda is called */
  readonly body: AstNode;
  /** Captured local bindings from the enclosing scope (closure) */
  readonly closureBindings?: Map<string, CalcValue | CalcArray>;
}

function isLambdaValue(v: unknown): v is LambdaValue {
  return v !== null && typeof v === "object" && "__lambda" in v && (v as LambdaValue).__lambda;
}

// ============================================================================
// Eval Session — per-calculation instance state
// ============================================================================

function cellKey(sheet: string, row: number, col: number): string {
  return `${sheet}!${row}:${col}`;
}

/**
 * Holds all mutable state for a single `calculateFormulas()` invocation.
 * This avoids module-level globals and prevents cross-contamination between
 * concurrent calculations of different workbooks.
 */
export class EvalSession {
  /** Cells currently on the evaluation call stack (circular-ref detection). */
  readonly evaluating = new Set<string>();
  /** Memoized results — avoids re-evaluating the same formula cell twice. */
  readonly cache = new Map<string, CalcValue | CalcValue[][]>();
  /**
   * Fallback values for circular references during iterative calculation.
   * When set, encountering a circular ref returns the stored value (previous
   * iteration result) instead of the default 0.
   */
  readonly circularFallback = new Map<string, CalcValue>();

  makeKey(sheet: string, row: number, col: number): string {
    return cellKey(sheet, row, col);
  }
}

// ============================================================================
// Cell Reference Resolution
// ============================================================================

function resolveCol(colStr: string): number {
  return colCache.l2n(colStr);
}

function resolveRow(rowStr: string): number {
  return parseInt(rowStr, 10);
}

function resolveCellRef(
  node: CellRefNode,
  _ctx: EvalContext
): { sheet: string | undefined; row: number; col: number } {
  return {
    sheet: node.sheet,
    row: resolveRow(node.row),
    col: resolveCol(node.col)
  };
}

/**
 * Build a 2D CalcArray from a rectangular region of cells, tagged with
 * origin metadata for implicit intersection.
 */
function buildRangeArray(
  ctx: EvalContext,
  sheet: string | undefined,
  top: number,
  bottom: number,
  left: number,
  right: number
): CalcArray {
  const result: CalcArray = [];
  for (let r = top; r <= bottom; r++) {
    const row: CalcValue[] = [];
    for (let c = left; c <= right; c++) {
      row.push(ctx.getCellValue(sheet, r, c));
    }
    result.push(row);
  }
  // Tag with origin for implicit intersection
  result._originRow = top;
  result._originCol = left;
  return result;
}

/**
 * Resolve a range reference into a 2D array of values.
 */
function resolveRange(node: RangeRefNode, ctx: EvalContext): CalcArray {
  const startRow = resolveRow(node.start.row);
  const startCol = resolveCol(node.start.col);
  const endRow = resolveRow(node.end.row);
  const endCol = resolveCol(node.end.col);

  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);

  return buildRangeArray(ctx, node.sheet, top, bottom, left, right);
}

/**
 * Resolve a whole-column range (e.g. A:B) into a 2D array using the sheet's used dimensions.
 */
function resolveColRange(node: ColRangeRefNode, ctx: EvalContext): CalcArray {
  const startCol = resolveCol(node.startCol);
  const endCol = resolveCol(node.endCol);
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);

  // Use sheet dimensions to determine row extent
  const dims = ctx.getSheetDimensions?.(node.sheet);
  if (!dims) {
    return [[]]; // Empty sheet
  }

  return buildRangeArray(ctx, node.sheet, dims.top, dims.bottom, left, right);
}

/**
 * Resolve a whole-row range (e.g. 1:5) into a 2D array using the sheet's used dimensions.
 */
function resolveRowRange(node: RowRangeRefNode, ctx: EvalContext): CalcArray {
  const top = Math.min(node.startRow, node.endRow);
  const bottom = Math.max(node.startRow, node.endRow);

  // Use sheet dimensions to determine column extent
  const dims = ctx.getSheetDimensions?.(node.sheet);
  if (!dims) {
    return [[]]; // Empty sheet
  }

  return buildRangeArray(ctx, node.sheet, top, bottom, dims.left, dims.right);
}

// ============================================================================
// 3D Reference Helpers
// ============================================================================

/**
 * Get the list of sheet names between startSheet and endSheet (inclusive)
 * based on workbook sheet order. Returns null if getSheetNames is not available
 * or if either sheet name is not found.
 */
function getSheetsInRange(ctx: EvalContext, startSheet: string, endSheet: string): string[] | null {
  if (!ctx.getSheetNames) {
    return null;
  }
  const allSheets = ctx.getSheetNames();
  const startIdx = allSheets.findIndex(s => s.toLowerCase() === startSheet.toLowerCase());
  const endIdx = allSheets.findIndex(s => s.toLowerCase() === endSheet.toLowerCase());
  if (startIdx === -1 || endIdx === -1) {
    return null;
  }
  const lo = Math.min(startIdx, endIdx);
  const hi = Math.max(startIdx, endIdx);
  return allSheets.slice(lo, hi + 1);
}

/**
 * Resolve a 3D cell reference (e.g. Sheet1:Sheet3!A1) into a flat column array.
 * Returns one value per sheet in the range, arranged as rows in a CalcArray.
 */
function resolve3DCellRef(row: number, col: number, sheets: string[], ctx: EvalContext): CalcArray {
  const result: CalcArray = [];
  for (const sheet of sheets) {
    result.push([ctx.getCellValue(sheet, row, col)]);
  }
  return result;
}

/**
 * Resolve a 3D range reference (e.g. Sheet1:Sheet3!A1:B2).
 * Flattens all sheets' ranges into a single CalcArray by appending rows.
 */
function resolve3DRange(node: RangeRefNode, sheets: string[], ctx: EvalContext): CalcArray {
  const startRow = resolveRow(node.start.row);
  const startCol = resolveCol(node.start.col);
  const endRow = resolveRow(node.end.row);
  const endCol = resolveCol(node.end.col);

  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);

  const result: CalcArray = [];
  for (const sheet of sheets) {
    for (let r = top; r <= bottom; r++) {
      const row: CalcValue[] = [];
      for (let c = left; c <= right; c++) {
        row.push(ctx.getCellValue(sheet, r, c));
      }
      result.push(row);
    }
  }
  return result;
}

// ============================================================================
// R1C1 Reference Resolution
// ============================================================================

/**
 * Parse an R1C1-style reference string and resolve it to a cell value or range.
 *
 * Supported formats:
 * - Absolute: R5C3 → row 5, col 3
 * - Relative: R[2]C[-1] → current row + 2, current col - 1
 * - Mixed: R5C[-1], R[2]C3
 * - Range: R1C1:R5C3
 */
function resolveR1C1Reference(
  refText: string,
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  void session; // not needed for R1C1 resolution, but kept for API consistency
  const upper = refText.toUpperCase().trim();

  // Check for range: R...C...:R...C...
  const rangeIdx = findR1C1RangeSeparator(upper);
  if (rangeIdx !== -1) {
    const startText = upper.slice(0, rangeIdx);
    const endText = upper.slice(rangeIdx + 1);
    const startRef = parseR1C1Single(startText, ctx);
    const endRef = parseR1C1Single(endText, ctx);
    if (!startRef || !endRef) {
      return makeError("#REF!");
    }
    const top = Math.min(startRef.row, endRef.row);
    const bottom = Math.max(startRef.row, endRef.row);
    const left = Math.min(startRef.col, endRef.col);
    const right = Math.max(startRef.col, endRef.col);
    const result: CalcArray = [];
    for (let r = top; r <= bottom; r++) {
      const row: CalcValue[] = [];
      for (let c = left; c <= right; c++) {
        row.push(ctx.getCellValue(undefined, r, c));
      }
      result.push(row);
    }
    result._originRow = top;
    result._originCol = left;
    return result;
  }

  // Single cell
  const ref = parseR1C1Single(upper, ctx);
  if (!ref) {
    return makeError("#REF!");
  }
  if (ref.row < 1 || ref.col < 1) {
    return makeError("#REF!");
  }
  return ctx.getCellValue(undefined, ref.row, ref.col);
}

/**
 * Find the colon that separates two R1C1 references in a range expression.
 * We need to be careful not to match colons inside brackets (e.g. R[-1]C[-1]).
 */
function findR1C1RangeSeparator(text: string): number {
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "[") {
      depth++;
    } else if (text[i] === "]") {
      depth--;
    } else if (text[i] === ":" && depth === 0) {
      return i;
    }
  }
  return -1;
}

/**
 * Parse a single R1C1 cell reference like R5C3 or R[2]C[-1].
 * Returns { row, col } or null on parse failure.
 */
function parseR1C1Single(text: string, ctx: EvalContext): { row: number; col: number } | null {
  // Match: R followed by optional [n] or n, then C followed by optional [n] or n
  const re = /^R(\[(-?\d+)\]|(\d+))C(\[(-?\d+)\]|(\d+))$/;
  const m = re.exec(text);
  if (!m) {
    return null;
  }

  const addr = ctx.currentAddress;
  let row: number;
  let col: number;

  if (m[2] !== undefined) {
    // Relative row: R[offset]
    row = (addr?.row ?? 1) + parseInt(m[2], 10);
  } else {
    // Absolute row: Rn
    row = parseInt(m[3], 10);
  }

  if (m[5] !== undefined) {
    // Relative col: C[offset]
    col = (addr?.col ?? 1) + parseInt(m[5], 10);
  } else {
    // Absolute col: Cn
    col = parseInt(m[6], 10);
  }

  return { row, col };
}

// ============================================================================
// Evaluator
// ============================================================================

export function evaluate(
  node: AstNode,
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  switch (node.type) {
    case NodeType.Number:
      return node.value;

    case NodeType.String:
      return node.value;

    case NodeType.Boolean:
      return node.value;

    case NodeType.Error:
      return makeError(node.value);

    case NodeType.Missing:
      return null;

    case NodeType.Name: {
      // Check local bindings first (from LET expressions)
      const upperName = node.name.toUpperCase();
      if (ctx.localBindings?.has(upperName)) {
        return ctx.localBindings.get(upperName)!;
      }
      if (!ctx.resolveName) {
        return makeError("#NAME?");
      }
      const resolved = ctx.resolveName(node.name);
      if (!resolved) {
        // resolveName returned null — try resolveNameValue as fallback.
        // This handles formula-based names and returns error values for
        // unsupported cases (e.g. multi-area defined names).
        if (ctx.resolveNameValue) {
          const nameVal = ctx.resolveNameValue(node.name);
          if (nameVal !== undefined) {
            return nameVal;
          }
        }
        return makeError("#NAME?");
      }
      // Single cell
      if (resolved.startRow === resolved.endRow && resolved.startCol === resolved.endCol) {
        const sheet = resolved.sheet;
        const key = cellKey(sheet ?? ctx.currentSheet, resolved.startRow, resolved.startCol);
        if (session.evaluating.has(key)) {
          // During iterative calculation, return previous iteration's result
          const fallback = session.circularFallback.get(key);
          if (fallback !== undefined) {
            return fallback;
          }
          return 0; // Circular reference — Excel returns 0
        }
        return ctx.getCellValue(sheet, resolved.startRow, resolved.startCol);
      }
      // Range — return 2D array with origin metadata for implicit intersection
      {
        const top = Math.min(resolved.startRow, resolved.endRow);
        const bottom = Math.max(resolved.startRow, resolved.endRow);
        const left = Math.min(resolved.startCol, resolved.endCol);
        const right = Math.max(resolved.startCol, resolved.endCol);
        return buildRangeArray(ctx, resolved.sheet, top, bottom, left, right);
      }
    }

    case NodeType.CellRef: {
      // 3D cell reference: Sheet1:Sheet3!A1
      if (node.endSheet) {
        const sheets = getSheetsInRange(ctx, node.sheet ?? ctx.currentSheet, node.endSheet);
        if (!sheets) {
          return makeError("#REF!");
        }
        const row = resolveRow(node.row);
        const col = resolveCol(node.col);
        return resolve3DCellRef(row, col, sheets, ctx);
      }
      const ref = resolveCellRef(node, ctx);
      const sheet = ref.sheet ?? ctx.currentSheet;
      const key = cellKey(sheet, ref.row, ref.col);
      if (session.evaluating.has(key)) {
        // During iterative calculation, return previous iteration's result
        const fallback = session.circularFallback.get(key);
        if (fallback !== undefined) {
          return fallback;
        }
        return 0; // Circular reference — Excel returns 0
      }
      return ctx.getCellValue(ref.sheet, ref.row, ref.col);
    }

    case NodeType.RangeRef:
      // 3D range reference: Sheet1:Sheet3!A1:B2
      if (node.endSheet) {
        const sheets = getSheetsInRange(ctx, node.sheet ?? ctx.currentSheet, node.endSheet);
        if (!sheets) {
          return makeError("#REF!");
        }
        return resolve3DRange(node, sheets, ctx);
      }
      return resolveRange(node, ctx);

    case NodeType.ColRangeRef:
      return resolveColRange(node, ctx);

    case NodeType.RowRangeRef:
      return resolveRowRange(node, ctx);

    case NodeType.BinaryOp:
      return evaluateBinaryOp(node.op, node.left, node.right, ctx, session);

    case NodeType.UnaryOp:
      return evaluateUnaryOp(node.op, node.operand, ctx, session);

    case NodeType.Percent: {
      const operand = evaluate(node.operand, ctx, session);
      if (Array.isArray(operand)) {
        // Apply % element-wise
        const result: CalcArray = [];
        for (const row of operand) {
          const outRow: CalcValue[] = [];
          for (const cell of row) {
            if (isError(cell)) {
              outRow.push(cell);
            } else {
              const n = typeof cell === "number" ? cell : Number(cell);
              outRow.push(isNaN(n) ? makeError("#VALUE!") : n / 100);
            }
          }
          result.push(outRow);
        }
        return result;
      }
      if (isError(operand)) {
        return operand;
      }
      const n = typeof operand === "number" ? operand : Number(operand);
      return isNaN(n) ? makeError("#VALUE!") : n / 100;
    }

    case NodeType.FunctionCall:
      return evaluateFunction(node.name, node.args, ctx, session);

    case NodeType.Array: {
      const rows: CalcValue[][] = [];
      for (const row of node.rows) {
        const evalRow: CalcValue[] = [];
        for (const elem of row) {
          const v = evaluate(elem, ctx, session);
          if (Array.isArray(v)) {
            evalRow.push(v[0]?.[0] ?? null);
          } else {
            evalRow.push(v);
          }
        }
        rows.push(evalRow);
      }
      return rows;
    }

    case NodeType.StructuredRef: {
      if (!ctx.resolveStructuredRef) {
        return makeError("#REF!");
      }
      const resolved = ctx.resolveStructuredRef(node.tableName, node.columns, node.specials);
      if (!resolved) {
        return makeError("#REF!");
      }
      // Single cell
      if (resolved.top === resolved.bottom && resolved.left === resolved.right) {
        return ctx.getCellValue(resolved.sheetName, resolved.top, resolved.left);
      }
      // Range — return 2D array
      const sResult: CalcArray = [];
      for (let r = resolved.top; r <= resolved.bottom; r++) {
        const row: CalcValue[] = [];
        for (let c = resolved.left; c <= resolved.right; c++) {
          row.push(ctx.getCellValue(resolved.sheetName, r, c));
        }
        sResult.push(row);
      }
      return sResult;
    }

    default: {
      // Exhaustive check
      const _: never = node;
      return makeError("#VALUE!");
    }
  }
}

// ============================================================================
// Implicit Intersection Helper
// ============================================================================

/**
 * Perform implicit intersection: given a CalcArray and the current cell's
 * position, extract the value from the array that aligns with the formula
 * cell's row or column.
 *
 * - If the array is a single row, pick the element at the formula cell's column offset.
/**
 * Implicit intersection: when an array appears in a scalar context, pick a
 * single value based on the formula cell's position relative to the array's
 * origin in the worksheet.
 *
 * If the array carries _originRow/_originCol metadata (set by resolveRange),
 * we use the formula cell's row/col to compute the correct offset into the
 * array. For single-row arrays, the column offset is used; for single-column
 * arrays, the row offset is used. This matches Excel's implicit intersection
 * behavior where a range reference in a scalar context picks the value from
 * the same row or column as the formula cell.
 *
 * If no origin metadata is available (e.g., computed arrays from functions),
 * falls back to the top-left element.
 */
function implicitIntersect(arr: CalcArray, ctx: EvalContext): CalcValue {
  if (arr.length === 0 || (arr.length === 1 && arr[0].length === 0)) {
    return null;
  }
  // 1x1 — trivial
  if (arr.length === 1 && arr[0].length === 1) {
    return arr[0][0];
  }

  const addr = ctx.currentAddress;
  if (!addr) {
    return arr[0]?.[0] ?? null;
  }

  const originRow = arr._originRow;
  const originCol = arr._originCol;

  // Single row — pick column by offset from origin
  if (arr.length === 1) {
    if (originCol !== undefined) {
      const colIdx = addr.col - originCol;
      if (colIdx >= 0 && colIdx < arr[0].length) {
        return arr[0][colIdx];
      }
    }
    return arr[0][0];
  }

  // Single column — pick row by offset from origin
  if (arr[0].length === 1) {
    if (originRow !== undefined) {
      const rowIdx = addr.row - originRow;
      if (rowIdx >= 0 && rowIdx < arr.length) {
        return arr[rowIdx][0];
      }
    }
    return arr[0][0];
  }

  // Multi-row, multi-column: try both dimensions
  if (originRow !== undefined && originCol !== undefined) {
    const rowIdx = addr.row - originRow;
    const colIdx = addr.col - originCol;
    if (rowIdx >= 0 && rowIdx < arr.length && colIdx >= 0 && colIdx < arr[0].length) {
      return arr[rowIdx][colIdx];
    }
  }

  return arr[0][0] ?? null;
}

// ============================================================================
// Array Arithmetic Broadcasting
// ============================================================================

/**
 * Apply a binary operation element-wise over two arrays with broadcasting.
 * Broadcasting rules (Excel-style):
 * - scalar op array: apply scalar to each element
 * - single-row op multi-row: broadcast the single row across all rows
 * - single-col op multi-col: broadcast the single col across all cols
 * - same dimensions: element-wise
 * - incompatible dimensions: #VALUE! error
 */
function broadcastBinaryOp(
  op: string,
  leftVal: CalcValue | CalcArray,
  rightVal: CalcValue | CalcArray
): CalcValue | CalcArray {
  const lIsArr = Array.isArray(leftVal);
  const rIsArr = Array.isArray(rightVal);

  // Neither is array — scalar op
  if (!lIsArr && !rIsArr) {
    return applyScalarBinaryOp(op, leftVal, rightVal);
  }

  // Determine dimensions
  const lRows = lIsArr ? leftVal.length : 1;
  const lCols = lIsArr ? (leftVal[0]?.length ?? 0) : 1;
  const rRows = rIsArr ? rightVal.length : 1;
  const rCols = rIsArr ? (rightVal[0]?.length ?? 0) : 1;

  // Result dimensions: max of each, with broadcasting validation
  const outRows = Math.max(lRows, rRows);
  const outCols = Math.max(lCols, rCols);

  // Validate broadcasting compatibility
  if (
    (lRows !== 1 && rRows !== 1 && lRows !== rRows) ||
    (lCols !== 1 && rCols !== 1 && lCols !== rCols)
  ) {
    return makeError("#VALUE!");
  }

  const result: CalcArray = [];
  for (let r = 0; r < outRows; r++) {
    const row: CalcValue[] = [];
    for (let c = 0; c < outCols; c++) {
      const lR = lRows === 1 ? 0 : r;
      const lC = lCols === 1 ? 0 : c;
      const rR = rRows === 1 ? 0 : r;
      const rC = rCols === 1 ? 0 : c;

      const lVal: CalcValue = lIsArr ? (leftVal[lR]?.[lC] ?? null) : leftVal;
      const rVal: CalcValue = rIsArr ? (rightVal[rR]?.[rC] ?? null) : rightVal;

      row.push(applyScalarBinaryOp(op, lVal, rVal));
    }
    result.push(row);
  }

  // Propagate origin metadata from input arrays for implicit intersection
  const lArr = lIsArr ? (leftVal as CalcArray) : null;
  const rArr = rIsArr ? (rightVal as CalcArray) : null;
  if (lArr?._originRow !== undefined) {
    result._originRow = lArr._originRow;
  } else if (rArr?._originRow !== undefined) {
    result._originRow = rArr._originRow;
  }
  if (lArr?._originCol !== undefined) {
    result._originCol = lArr._originCol;
  } else if (rArr?._originCol !== undefined) {
    result._originCol = rArr._originCol;
  }

  return result;
}

/**
 * Apply a binary operation to two scalar values.
 */
function applyScalarBinaryOp(op: string, left: CalcValue, right: CalcValue): CalcValue {
  if (isError(left)) {
    return left;
  }
  if (isError(right)) {
    return right;
  }

  // Concatenation — Excel uses uppercase TRUE/FALSE
  if (op === "&") {
    const lStr = left === true ? "TRUE" : left === false ? "FALSE" : String(left ?? "");
    const rStr = right === true ? "TRUE" : right === false ? "FALSE" : String(right ?? "");
    return lStr + rStr;
  }

  // Comparison operators
  if (op === "=" || op === "<>" || op === "<" || op === ">" || op === "<=" || op === ">=") {
    return compareValues(left, right, op);
  }

  // Arithmetic operators
  const lNum = toNumber(left);
  if (isError(lNum)) {
    return lNum;
  }
  const rNum = toNumber(right);
  if (isError(rNum)) {
    return rNum;
  }

  switch (op) {
    case "+": {
      const r = lNum + rNum;
      return !isFinite(r) ? makeError("#NUM!") : r;
    }
    case "-": {
      const r = lNum - rNum;
      return !isFinite(r) ? makeError("#NUM!") : r;
    }
    case "*": {
      const r = lNum * rNum;
      return !isFinite(r) ? makeError("#NUM!") : r;
    }
    case "/":
      if (rNum === 0) {
        return makeError("#DIV/0!");
      }
      {
        const r = lNum / rNum;
        return !isFinite(r) ? makeError("#NUM!") : r;
      }
    case "^": {
      const r = Math.pow(lNum, rNum);
      return !isFinite(r) ? makeError("#NUM!") : r;
    }
    default:
      return makeError("#VALUE!");
  }
}

// ============================================================================
// Binary Operations
// ============================================================================

function evaluateBinaryOp(
  op: string,
  leftNode: AstNode,
  rightNode: AstNode,
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  const leftVal = evaluate(leftNode, ctx, session);
  const rightVal = evaluate(rightNode, ctx, session);

  const lIsArr = Array.isArray(leftVal);
  const rIsArr = Array.isArray(rightVal);

  // If either operand is an array, use array broadcasting
  if (lIsArr || rIsArr) {
    return broadcastBinaryOp(op, leftVal, rightVal);
  }

  // Both scalars — fast path
  return applyScalarBinaryOp(op, leftVal as CalcValue, rightVal as CalcValue);
}

function compareValues(left: CalcValue, right: CalcValue, op: string): boolean {
  // Excel comparison rules:
  // - Different types: numbers < strings < booleans
  // - Same type: natural comparison
  // - Null is 0 for numbers, "" for strings
  const l = left ?? (typeof right === "string" ? "" : 0);
  const r = right ?? (typeof left === "string" ? "" : 0);

  let cmp: number;
  if (typeof l === typeof r) {
    if (typeof l === "string") {
      cmp = (l as string).toLowerCase().localeCompare((r as string).toLowerCase());
    } else if (typeof l === "number") {
      cmp = (l as number) - (r as number);
    } else if (typeof l === "boolean") {
      cmp = l === r ? 0 : l ? 1 : -1;
    } else {
      cmp = 0;
    }
  } else {
    // Cross-type comparison
    const typeOrder = (v: CalcValue): number => {
      if (typeof v === "number") {
        return 0;
      }
      if (typeof v === "string") {
        return 1;
      }
      if (typeof v === "boolean") {
        return 2;
      }
      return 3;
    };
    cmp = typeOrder(l) - typeOrder(r);
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

// ============================================================================
// Unary Operations
// ============================================================================

function applyScalarUnaryOp(op: string, scalar: CalcValue): CalcValue {
  if (isError(scalar)) {
    return scalar;
  }
  const n = toNumber(scalar);
  if (isError(n)) {
    return n;
  }
  switch (op) {
    case "-":
      return -n;
    case "+":
      return n;
    default:
      return makeError("#VALUE!");
  }
}

function evaluateUnaryOp(
  op: string,
  operandNode: AstNode,
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  const val = evaluate(operandNode, ctx, session);

  // @ implicit intersection operator: reduce array to scalar
  if (op === "@") {
    if (Array.isArray(val)) {
      return implicitIntersect(val, ctx);
    }
    return val;
  }

  if (Array.isArray(val)) {
    // Apply unary op element-wise
    const result: CalcArray = [];
    for (const row of val) {
      const outRow: CalcValue[] = [];
      for (const cell of row) {
        outRow.push(applyScalarUnaryOp(op, cell));
      }
      result.push(outRow);
    }
    return result;
  }

  return applyScalarUnaryOp(op, val);
}

// ============================================================================
// Function Evaluation
// ============================================================================

function evaluateFunction(
  name: string,
  argNodes: AstNode[],
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  const upperName = name.toUpperCase();

  // Special handling for ROW() and COLUMN() with cell ref argument
  if (upperName === "ROW" || upperName === "COLUMN") {
    if (argNodes.length === 0) {
      // ROW() / COLUMN() with no args = current cell
      if (ctx.currentAddress) {
        return upperName === "ROW" ? ctx.currentAddress.row : ctx.currentAddress.col;
      }
      return makeError("#VALUE!");
    }
    const argNode = argNodes[0];
    if (argNode.type === NodeType.CellRef) {
      return upperName === "ROW" ? resolveRow(argNode.row) : resolveCol(argNode.col);
    }
    if (argNode.type === NodeType.RangeRef) {
      return upperName === "ROW" ? resolveRow(argNode.start.row) : resolveCol(argNode.start.col);
    }
  }

  // Special handling for ROWS() and COLUMNS() with range argument
  if (upperName === "ROWS" || upperName === "COLUMNS") {
    if (argNodes.length > 0) {
      const argNode = argNodes[0];
      if (argNode.type === NodeType.RangeRef) {
        const startRow = resolveRow(argNode.start.row);
        const endRow = resolveRow(argNode.end.row);
        const startCol = resolveCol(argNode.start.col);
        const endCol = resolveCol(argNode.end.col);
        return upperName === "ROWS"
          ? Math.abs(endRow - startRow) + 1
          : Math.abs(endCol - startCol) + 1;
      }
    }
  }

  // ── Short-circuit functions: evaluate args lazily ──
  // IF(cond, trueVal, falseVal) — only evaluate the taken branch
  if (upperName === "IF" && argNodes.length >= 2) {
    const cond = evaluate(argNodes[0], ctx, session);
    const condScalar = Array.isArray(cond) ? (cond[0]?.[0] ?? null) : cond;
    if (isError(condScalar)) {
      return condScalar;
    }
    const truthy = condScalar === true || (typeof condScalar === "number" && condScalar !== 0);
    if (truthy) {
      return argNodes.length > 1 ? evaluate(argNodes[1], ctx, session) : true;
    }
    return argNodes.length > 2 ? evaluate(argNodes[2], ctx, session) : false;
  }

  // IFERROR(value, valueIfError) — only evaluate fallback if first arg is error
  if (upperName === "IFERROR" && argNodes.length >= 2) {
    const val = evaluate(argNodes[0], ctx, session);
    const scalar = Array.isArray(val) ? (val[0]?.[0] ?? null) : val;
    return isError(scalar) ? evaluate(argNodes[1], ctx, session) : val;
  }

  // IFNA(value, valueIfNA) — only evaluate fallback if first arg is #N/A
  if (upperName === "IFNA" && argNodes.length >= 2) {
    const val = evaluate(argNodes[0], ctx, session);
    const scalar = Array.isArray(val) ? (val[0]?.[0] ?? null) : val;
    return isError(scalar) && scalar.error === "#N/A" ? evaluate(argNodes[1], ctx, session) : val;
  }

  // IFS(cond1, val1, cond2, val2, ...) — stop at first true condition
  if (upperName === "IFS" && argNodes.length >= 2) {
    for (let i = 0; i < argNodes.length - 1; i += 2) {
      const cond = evaluate(argNodes[i], ctx, session);
      const condScalar = Array.isArray(cond) ? (cond[0]?.[0] ?? null) : cond;
      if (isError(condScalar)) {
        return condScalar;
      }
      if (condScalar === true || (typeof condScalar === "number" && condScalar !== 0)) {
        return evaluate(argNodes[i + 1], ctx, session);
      }
    }
    return makeError("#N/A");
  }

  // SWITCH(expr, val1, result1, ..., [default]) — stop at first match
  if (upperName === "SWITCH" && argNodes.length >= 3) {
    const expr = evaluate(argNodes[0], ctx, session);
    const exprScalar = Array.isArray(expr) ? (expr[0]?.[0] ?? null) : expr;
    for (let i = 1; i < argNodes.length - 1; i += 2) {
      const caseVal = evaluate(argNodes[i], ctx, session);
      const caseScalar = Array.isArray(caseVal) ? (caseVal[0]?.[0] ?? null) : caseVal;
      if (exprScalar === caseScalar) {
        return evaluate(argNodes[i + 1], ctx, session);
      }
    }
    // Default value when arg count is even
    if (argNodes.length % 2 === 0) {
      return evaluate(argNodes[argNodes.length - 1], ctx, session);
    }
    return makeError("#N/A");
  }

  // CHOOSE(index, val1, val2, ...) — only evaluate chosen branch
  if (upperName === "CHOOSE" && argNodes.length >= 2) {
    const idxVal = evaluate(argNodes[0], ctx, session);
    const idxScalar = Array.isArray(idxVal) ? (idxVal[0]?.[0] ?? null) : idxVal;
    if (isError(idxScalar)) {
      return idxScalar;
    }
    const idx = typeof idxScalar === "number" ? Math.floor(idxScalar) : 0;
    if (idx < 1 || idx >= argNodes.length) {
      return makeError("#VALUE!");
    }
    return evaluate(argNodes[idx], ctx, session);
  }

  // ── LET(name1, value1, name2, value2, ..., calculation) ──
  // Introduces local variable bindings. Args must be odd count (≥ 3).
  // Names are Name nodes that become local bindings; values are lazily evaluated.
  if (upperName === "LET" && argNodes.length >= 3 && argNodes.length % 2 === 1) {
    // Save current local bindings so we can restore them after
    const prevBindings = ctx.localBindings;
    const newBindings = new Map<string, CalcValue | CalcArray>(prevBindings);

    // Process name/value pairs
    const pairCount = (argNodes.length - 1) / 2;
    for (let i = 0; i < pairCount; i++) {
      const nameNode = argNodes[i * 2];
      const valueNode = argNodes[i * 2 + 1];

      // The name argument must be a Name node
      if (nameNode.type !== NodeType.Name) {
        return makeError("#VALUE!");
      }
      const varName = (nameNode as NameNode).name.toUpperCase();

      // Evaluate the value with current bindings (so later bindings can reference earlier ones)
      ctx.localBindings = newBindings;
      const val = evaluate(valueNode, ctx, session);
      newBindings.set(varName, val);
    }

    // Evaluate the final calculation expression with all bindings active
    ctx.localBindings = newBindings;
    try {
      return evaluate(argNodes[argNodes.length - 1], ctx, session);
    } finally {
      ctx.localBindings = prevBindings;
    }
  }

  // ── LAMBDA(param1, param2, ..., body) ──
  // Creates a function value. When directly invoked (not stored), this returns
  // a LambdaValue. When a LAMBDA result is called (via a defined name),
  // the parameters are bound and the body is evaluated.
  if (upperName === "LAMBDA" && argNodes.length >= 1) {
    // All args except the last are parameter names (Name nodes)
    const paramNodes = argNodes.slice(0, -1);
    const bodyNode = argNodes[argNodes.length - 1];

    const params: string[] = [];
    for (const pNode of paramNodes) {
      if (pNode.type !== NodeType.Name) {
        return makeError("#VALUE!");
      }
      params.push((pNode as NameNode).name.toUpperCase());
    }

    // Return a LambdaValue (tagged object).
    // This is not a standard CalcValue, but it can be stored in localBindings
    // and invoked when called as a function.
    const closureBindings = ctx.localBindings ? new Map(ctx.localBindings) : undefined;
    return {
      __lambda: true,
      params,
      body: bodyNode,
      closureBindings
    } as unknown as CalcValue;
  }

  // ── MAP(array, lambda) ──
  // Applies lambda to each element of the array.
  if (upperName === "MAP" || upperName === "_XLFN.MAP") {
    if (argNodes.length < 2) {
      return makeError("#VALUE!");
    }
    const arrVal = evaluate(argNodes[0], ctx, session);
    const lambdaVal = evaluate(argNodes[argNodes.length - 1], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    if (!Array.isArray(arrVal)) {
      return invokeLambda(lambdaVal, [arrVal], ctx, session);
    }
    const result: CalcArray = [];
    for (const row of arrVal) {
      const outRow: CalcValue[] = [];
      for (const cell of row) {
        const v = invokeLambda(lambdaVal, [cell], ctx, session);
        outRow.push(Array.isArray(v) ? (v[0]?.[0] ?? null) : v);
      }
      result.push(outRow);
    }
    return result;
  }

  // ── REDUCE(initial_value, array, lambda) ──
  // Reduces array to a single value by applying lambda(accumulator, value).
  if (upperName === "REDUCE" || upperName === "_XLFN.REDUCE") {
    if (argNodes.length < 3) {
      return makeError("#VALUE!");
    }
    let acc: CalcValue | CalcArray = evaluate(argNodes[0], ctx, session);
    const arrVal = evaluate(argNodes[1], ctx, session);
    const lambdaVal = evaluate(argNodes[2], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    if (Array.isArray(arrVal)) {
      for (const row of arrVal) {
        for (const cell of row) {
          acc = invokeLambda(
            lambdaVal,
            [Array.isArray(acc) ? (acc[0]?.[0] ?? null) : acc, cell],
            ctx,
            session
          );
        }
      }
    }
    return acc;
  }

  // ── SCAN(initial_value, array, lambda) ──
  // Like REDUCE but returns intermediate results as an array.
  if (upperName === "SCAN" || upperName === "_XLFN.SCAN") {
    if (argNodes.length < 3) {
      return makeError("#VALUE!");
    }
    let acc: CalcValue | CalcArray = evaluate(argNodes[0], ctx, session);
    const arrVal = evaluate(argNodes[1], ctx, session);
    const lambdaVal = evaluate(argNodes[2], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    const result: CalcArray = [];
    if (Array.isArray(arrVal)) {
      for (const row of arrVal) {
        const outRow: CalcValue[] = [];
        for (const cell of row) {
          acc = invokeLambda(
            lambdaVal,
            [Array.isArray(acc) ? (acc[0]?.[0] ?? null) : acc, cell],
            ctx,
            session
          );
          outRow.push(Array.isArray(acc) ? (acc[0]?.[0] ?? null) : acc);
        }
        result.push(outRow);
      }
    }
    return result.length > 0 ? result : makeError("#CALC!");
  }

  // ── MAKEARRAY(rows, cols, lambda) ──
  // Creates an array by calling lambda(row_index, col_index) for each cell.
  if (upperName === "MAKEARRAY" || upperName === "_XLFN.MAKEARRAY") {
    if (argNodes.length < 3) {
      return makeError("#VALUE!");
    }
    const rowsVal = evaluate(argNodes[0], ctx, session);
    const colsVal = evaluate(argNodes[1], ctx, session);
    const lambdaVal = evaluate(argNodes[2], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    const rows = toNumber(Array.isArray(rowsVal) ? (rowsVal[0]?.[0] ?? null) : rowsVal);
    if (isError(rows)) {
      return rows;
    }
    const cols = toNumber(Array.isArray(colsVal) ? (colsVal[0]?.[0] ?? null) : colsVal);
    if (isError(cols)) {
      return cols;
    }
    const result: CalcArray = [];
    for (let r = 1; r <= (rows as number); r++) {
      const outRow: CalcValue[] = [];
      for (let c = 1; c <= (cols as number); c++) {
        const v = invokeLambda(lambdaVal, [r, c], ctx, session);
        outRow.push(Array.isArray(v) ? (v[0]?.[0] ?? null) : v);
      }
      result.push(outRow);
    }
    return result;
  }

  // ── BYROW(array, lambda) ──
  // Applies lambda to each row (as a 1×N array) and returns a column of results.
  if (upperName === "BYROW" || upperName === "_XLFN.BYROW") {
    if (argNodes.length < 2) {
      return makeError("#VALUE!");
    }
    const arrVal = evaluate(argNodes[0], ctx, session);
    const lambdaVal = evaluate(argNodes[1], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    if (!Array.isArray(arrVal)) {
      return invokeLambda(lambdaVal, [[[arrVal]]], ctx, session);
    }
    const result: CalcArray = [];
    for (const row of arrVal) {
      const rowArr: CalcArray = [row];
      const v = invokeLambda(lambdaVal, [rowArr], ctx, session);
      result.push([Array.isArray(v) ? (v[0]?.[0] ?? null) : v]);
    }
    return result;
  }

  // ── BYCOL(array, lambda) ──
  // Applies lambda to each column (as an N×1 array) and returns a row of results.
  if (upperName === "BYCOL" || upperName === "_XLFN.BYCOL") {
    if (argNodes.length < 2) {
      return makeError("#VALUE!");
    }
    const arrVal = evaluate(argNodes[0], ctx, session);
    const lambdaVal = evaluate(argNodes[1], ctx, session);
    if (!isLambdaValue(lambdaVal)) {
      return makeError("#VALUE!");
    }
    if (!Array.isArray(arrVal)) {
      return invokeLambda(lambdaVal, [[[arrVal]]], ctx, session);
    }
    const numCols = arrVal[0]?.length ?? 0;
    const outRow: CalcValue[] = [];
    for (let c = 0; c < numCols; c++) {
      const colArr: CalcArray = arrVal.map(row => [row[c]]);
      const v = invokeLambda(lambdaVal, [colArr], ctx, session);
      outRow.push(Array.isArray(v) ? (v[0]?.[0] ?? null) : v);
    }
    return [outRow];
  }

  // ── INDIRECT(ref_text, [a1]) ──
  // Parses a string as a cell/range reference at runtime and returns the value(s).
  if (upperName === "INDIRECT" && argNodes.length >= 1) {
    const refArg = evaluate(argNodes[0], ctx, session);
    const refText = Array.isArray(refArg) ? String(refArg[0]?.[0] ?? "") : String(refArg ?? "");
    if (!refText) {
      return makeError("#REF!");
    }

    // Check for a1 argument (default true). If false, parse as R1C1 style.
    let a1 = true;
    if (argNodes.length >= 2) {
      const a1Arg = evaluate(argNodes[1], ctx, session);
      const a1Scalar = Array.isArray(a1Arg) ? (a1Arg[0]?.[0] ?? null) : a1Arg;
      a1 = a1Scalar !== false && a1Scalar !== 0;
    }

    if (!a1) {
      // R1C1 style: R5C3 (absolute) or R[2]C[-1] (relative)
      return resolveR1C1Reference(refText, ctx, session);
    }

    // A1 style: try parsing as a reference using the tokenizer+parser
    try {
      const tokens = tokenize(refText);
      const ast = parse(tokens);
      return evaluate(ast, ctx, session);
    } catch {
      return makeError("#REF!");
    }
  }

  // ── OFFSET(reference, rows, cols, [height], [width]) ──
  // Returns a reference offset from a starting cell.
  if (upperName === "OFFSET" && argNodes.length >= 3) {
    // Evaluate the reference argument — we need to extract the cell address
    const refNode = argNodes[0];
    let baseRow: number;
    let baseCol: number;
    let baseSheet: string | undefined;

    if (refNode.type === NodeType.CellRef) {
      baseRow = resolveRow(refNode.row);
      baseCol = resolveCol(refNode.col);
      baseSheet = refNode.sheet;
    } else if (refNode.type === NodeType.RangeRef) {
      baseRow = resolveRow(refNode.start.row);
      baseCol = resolveCol(refNode.start.col);
      baseSheet = refNode.sheet;
    } else {
      // Fallback: evaluate and try to use as value
      return makeError("#VALUE!");
    }

    const rowsVal = evaluate(argNodes[1], ctx, session);
    const rowsNum = toNumber(Array.isArray(rowsVal) ? (rowsVal[0]?.[0] ?? null) : rowsVal);
    if (isError(rowsNum)) {
      return rowsNum;
    }
    const colsVal = evaluate(argNodes[2], ctx, session);
    const colsNum = toNumber(Array.isArray(colsVal) ? (colsVal[0]?.[0] ?? null) : colsVal);
    if (isError(colsNum)) {
      return colsNum;
    }

    const newRow = baseRow + (rowsNum as number);
    const newCol = baseCol + (colsNum as number);

    if (newRow < 1 || newCol < 1) {
      return makeError("#REF!");
    }

    // Height and width
    let height = 1;
    let width = 1;
    if (argNodes.length > 3) {
      const hVal = evaluate(argNodes[3], ctx, session);
      const h = toNumber(Array.isArray(hVal) ? (hVal[0]?.[0] ?? null) : hVal);
      if (!isError(h)) {
        height = h as number;
      }
    }
    if (argNodes.length > 4) {
      const wVal = evaluate(argNodes[4], ctx, session);
      const w = toNumber(Array.isArray(wVal) ? (wVal[0]?.[0] ?? null) : wVal);
      if (!isError(w)) {
        width = w as number;
      }
    }

    if (height === 1 && width === 1) {
      // Single cell
      return ctx.getCellValue(baseSheet, newRow, newCol);
    }

    // Range
    const result: CalcArray = [];
    for (let r = 0; r < height; r++) {
      const row: CalcValue[] = [];
      for (let c = 0; c < width; c++) {
        row.push(ctx.getCellValue(baseSheet, newRow + r, newCol + c));
      }
      result.push(row);
    }
    result._originRow = newRow;
    result._originCol = newCol;
    return result;
  }

  // ── Default: eager argument evaluation ──
  const args: (CalcValue | CalcArray)[] = argNodes.map(arg => evaluate(arg, ctx, session));

  // Look up function
  const fn = FUNCTIONS[upperName];
  if (fn) {
    return fn(args);
  }

  // ── Unknown function name: check if it's a defined name resolving to a lambda ──
  // This handles patterns like: MyFunc(1, 2) where MyFunc is defined as =LAMBDA(x, y, x+y)
  if (ctx.localBindings?.has(upperName)) {
    const boundVal = ctx.localBindings.get(upperName)!;
    if (isLambdaValue(boundVal)) {
      return invokeLambda(boundVal, args, ctx, session);
    }
  }
  // Check workbook defined names that resolve to a lambda
  if (ctx.resolveName) {
    const resolved = ctx.resolveName(upperName);
    if (resolved) {
      // Evaluate the defined name — if it's a cell containing a LAMBDA formula,
      // the cell's value would have been computed as a LambdaValue
      const sheet = resolved.sheet;
      if (resolved.startRow === resolved.endRow && resolved.startCol === resolved.endCol) {
        const cellVal = ctx.getCellValue(sheet, resolved.startRow, resolved.startCol);
        if (isLambdaValue(cellVal)) {
          return invokeLambda(cellVal, args, ctx, session);
        }
      }
    }
  }
  // Check formula-based defined names that evaluate to a lambda
  if (ctx.resolveNameValue) {
    const nameVal = ctx.resolveNameValue(upperName);
    if (nameVal !== undefined && isLambdaValue(nameVal)) {
      return invokeLambda(nameVal, args, ctx, session);
    }
  }

  return makeError("#NAME?");
}

/**
 * Invoke a LambdaValue with the given arguments.
 * Binds each parameter to the corresponding argument value and evaluates the body.
 */
function invokeLambda(
  lambda: LambdaValue,
  args: (CalcValue | CalcArray)[],
  ctx: EvalContext,
  session: EvalSession
): CalcValue | CalcArray {
  if (args.length !== lambda.params.length) {
    return makeError("#VALUE!");
  }

  const prevBindings = ctx.localBindings;
  const newBindings = new Map<string, CalcValue | CalcArray>(lambda.closureBindings);

  // Bind parameters to argument values
  for (let i = 0; i < lambda.params.length; i++) {
    newBindings.set(lambda.params[i], args[i]);
  }

  ctx.localBindings = newBindings;
  try {
    return evaluate(lambda.body, ctx, session);
  } finally {
    ctx.localBindings = prevBindings;
  }
}

// ============================================================================
// Public: Evaluate with Circular Reference Protection
// ============================================================================

/**
 * Internal helper for formula evaluation with caching, circular reference
 * detection, and address tracking. Both `evaluateFormula` and
 * `evaluateFormulaRaw` delegate to this.
 *
 * @returns `{ scalar, raw }` where `scalar` is the implicit-intersected value
 *          (always cached) and `raw` is the unmodified evaluation result.
 */
function evaluateFormulaCore(
  node: AstNode,
  ctx: EvalContext,
  session: EvalSession,
  cellSheet: string,
  cellRow: number,
  cellCol: number
): { scalar: CalcValue; raw: CalcValue | CalcArray } {
  const key = session.makeKey(cellSheet, cellRow, cellCol);

  // Check memo cache first — avoids re-evaluating the same cell
  const cached = session.cache.get(key);
  if (cached !== undefined && !Array.isArray(cached)) {
    return { scalar: cached, raw: cached };
  }

  // Circular reference detection
  if (session.evaluating.has(key)) {
    // During iterative calculation, return the previous iteration's result
    const fallback = session.circularFallback.get(key);
    if (fallback !== undefined) {
      return { scalar: fallback, raw: fallback };
    }
    return { scalar: 0, raw: 0 }; // Excel returns 0 for circular references
  }

  session.evaluating.add(key);
  const prevAddress = ctx.currentAddress;
  ctx.currentAddress = { sheet: cellSheet, row: cellRow, col: cellCol };

  try {
    const result = evaluate(node, ctx, session);
    // Unwrap array via implicit intersection for the scalar form
    const scalar = Array.isArray(result) ? implicitIntersect(result, ctx) : result;
    // Memoize the scalar value for dependency resolution
    session.cache.set(key, scalar);
    return { scalar, raw: result };
  } finally {
    session.evaluating.delete(key);
    ctx.currentAddress = prevAddress;
  }
}

/**
 * Evaluate a formula and return the scalar result.
 * Array results are unwrapped via implicit intersection (using the formula
 * cell's position) to produce a single value. This is the function used
 * for dependency resolution (getCellValue) and backward compatibility.
 */
export function evaluateFormula(
  node: AstNode,
  ctx: EvalContext,
  session: EvalSession,
  cellSheet: string,
  cellRow: number,
  cellCol: number
): CalcValue {
  return evaluateFormulaCore(node, ctx, session, cellSheet, cellRow, cellCol).scalar;
}

/**
 * Evaluate a formula and return the raw result, preserving arrays.
 * Used by the spill engine in `calculateFormulas` to detect array results
 * and write them to adjacent cells.
 *
 * The scalar result is still memoized in the session cache for dependency
 * resolution, but the full array is returned to the caller.
 */
export function evaluateFormulaRaw(
  node: AstNode,
  ctx: EvalContext,
  session: EvalSession,
  cellSheet: string,
  cellRow: number,
  cellCol: number
): CalcValue | CalcArray {
  return evaluateFormulaCore(node, ctx, session, cellSheet, cellRow, cellCol).raw;
}
