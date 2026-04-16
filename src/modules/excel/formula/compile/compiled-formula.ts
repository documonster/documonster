/**
 * Compiled Formula — The output of the full compilation pipeline.
 *
 * A `CompiledFormula` packages together:
 * - The original `FormulaInstance` metadata
 * - The raw AST (from the parser)
 * - The bound expression tree (from the binder)
 * - Static dependency information
 * - Metadata flags (volatile, dynamic refs, etc.)
 *
 * ## Static Dependency Extraction
 *
 * Dependencies are extracted from the `BoundExpr` tree. Since names and
 * structured refs are already resolved by the binder, the dependency
 * extraction is a simple tree walk that collects `BoundCellRef` and
 * `BoundAreaRef` nodes.
 *
 * Runtime-dependent references (INDIRECT, OFFSET) cannot be captured
 * statically — the `hasDynamicRefs` flag marks formulas that may have
 * additional runtime dependencies.
 */

import type { FormulaInstance } from "../integration/formula-instance";
import type { ResolvedTable, WorkbookSnapshot } from "../integration/workbook-snapshot";
import type { AstNode } from "../syntax/ast";
import { NodeType } from "../syntax/ast";
import type { BoundExpr } from "./bound-ast";
import { BoundExprKind } from "./bound-ast";

// ============================================================================
// CompiledFormula
// ============================================================================

/**
 * The complete compiled representation of a formula.
 */
export interface CompiledFormula {
  /** The original formula instance metadata. */
  readonly instance: FormulaInstance;
  /** The raw AST (kept for INDIRECT/OFFSET that re-parse at runtime). */
  readonly ast: AstNode;
  /** The bound expression tree (the evaluator executes this). */
  readonly bound: BoundExpr;
  /** Statically extractable dependencies. */
  readonly staticDeps: StaticDependencySet;
  /** Whether this formula uses volatile functions (RAND, NOW, etc.). */
  readonly isVolatile: boolean;
  /** Whether this formula contains INDIRECT/OFFSET (runtime-dependent refs). */
  readonly hasDynamicRefs: boolean;
  /** Whether this formula contains any lambda expressions. */
  readonly containsLambda: boolean;
  /**
   * Whether the top-level function is a known dynamic array function
   * (FILTER, SORT, UNIQUE, SEQUENCE, etc.). Determined once at compile
   * time — consumers should use this flag instead of re-checking the AST
   * or bound expression.
   */
  readonly isDynamicArrayFunction: boolean;
}

// ============================================================================
// Static Dependency Set
// ============================================================================

/**
 * A single cell dependency.
 */
export interface CellDep {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
}

/**
 * A rectangular area dependency.
 */
export interface AreaDep {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/**
 * All statically extractable dependencies of a formula.
 */
export interface StaticDependencySet {
  /** Individual cell references. */
  readonly cells: readonly CellDep[];
  /** Range references. */
  readonly areas: readonly AreaDep[];
}

// ============================================================================
// Dependency Extraction
// ============================================================================

/**
 * Callback to resolve a defined name into its parsed/bound expression
 * for dependency extraction. Returns the bound expression if the name
 * resolves to a formula, or undefined if it can't be resolved.
 */
export type NameDepResolver = (name: string) =>
  | {
      deps: StaticDependencySet;
      hasDynamicRefs: boolean;
    }
  | undefined;

/**
 * Extract static dependencies from a bound expression tree.
 *
 * @param expr - The bound expression tree to analyze
 * @param snapshot - Optional snapshot for resolving structured references
 * @param nameResolver - Optional resolver for formula-based defined name dependencies
 */
export function extractStaticDeps(
  expr: BoundExpr,
  snapshot?: WorkbookSnapshot,
  nameResolver?: NameDepResolver
): StaticDependencySet {
  const cells: CellDep[] = [];
  const areas: AreaDep[] = [];
  walkDeps(expr, cells, areas, snapshot?.tablesByName, nameResolver, new Set());
  return { cells, areas };
}

function walkDeps(
  expr: BoundExpr,
  cells: CellDep[],
  areas: AreaDep[],
  tablesByName?: ReadonlyMap<string, ResolvedTable>,
  nameResolver?: NameDepResolver,
  visitedNames?: Set<string>
): void {
  switch (expr.kind) {
    case BoundExprKind.CellRef:
      cells.push({ sheet: expr.sheet, row: expr.row, col: expr.col });
      break;

    case BoundExprKind.AreaRef:
      areas.push({
        sheet: expr.sheet,
        top: expr.top,
        left: expr.left,
        bottom: expr.bottom,
        right: expr.right
      });
      break;

    case BoundExprKind.ColRangeRef:
      areas.push({
        sheet: expr.sheet,
        top: 1,
        left: expr.leftCol,
        bottom: 1_048_576,
        right: expr.rightCol
      });
      break;

    case BoundExprKind.RowRangeRef:
      areas.push({
        sheet: expr.sheet,
        top: expr.topRow,
        left: 1,
        bottom: expr.bottomRow,
        right: 16_384
      });
      break;

    case BoundExprKind.Ref3D:
      for (const sheet of expr.sheets) {
        if (expr.inner.kind === BoundExprKind.CellRef) {
          cells.push({ sheet, row: expr.inner.row, col: expr.inner.col });
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
      break;

    case BoundExprKind.BinaryOp:
      walkDeps(expr.left, cells, areas, tablesByName, nameResolver, visitedNames);
      walkDeps(expr.right, cells, areas, tablesByName, nameResolver, visitedNames);
      break;

    case BoundExprKind.UnaryOp:
      walkDeps(expr.operand, cells, areas, tablesByName, nameResolver, visitedNames);
      break;

    case BoundExprKind.Percent:
      walkDeps(expr.operand, cells, areas, tablesByName, nameResolver, visitedNames);
      break;

    case BoundExprKind.Call:
      for (const arg of expr.args) {
        walkDeps(arg, cells, areas, tablesByName, nameResolver, visitedNames);
      }
      break;

    case BoundExprKind.SpecialCall:
      for (const arg of expr.args) {
        walkDeps(arg, cells, areas, tablesByName, nameResolver, visitedNames);
      }
      break;

    case BoundExprKind.Array:
      for (const row of expr.rows) {
        for (const elem of row) {
          walkDeps(elem, cells, areas, tablesByName, nameResolver, visitedNames);
        }
      }
      break;

    case BoundExprKind.Lambda:
      walkDeps(expr.body, cells, areas, tablesByName, nameResolver, visitedNames);
      break;

    case BoundExprKind.StructuredRef:
      // Try to resolve structured reference to static deps using the table index.
      // Only possible when we have the table name and the table exists in the snapshot.
      if (tablesByName && expr.tableName) {
        const resolved = tablesByName.get(expr.tableName.toLowerCase());
        if (resolved) {
          const t = resolved.table;
          const tl = t.topLeft;
          const width = t.columns.length;
          const dataRowStart = tl.row + (t.hasHeaderRow ? 1 : 0);
          const dataRowEnd = dataRowStart + t.dataRowCount - 1;

          // Determine column range
          let colLeft = tl.col;
          let colRight = tl.col + width - 1;
          if (expr.columns.length > 0) {
            const indices: number[] = [];
            for (const colName of expr.columns) {
              const idx = t.columns.findIndex(c => c.name.toLowerCase() === colName.toLowerCase());
              if (idx !== -1) {
                indices.push(idx);
              }
            }
            if (indices.length > 0) {
              colLeft = tl.col + Math.min(...indices);
              colRight = tl.col + Math.max(...indices);
            }
          }

          // Determine row range based on specials
          let rowTop = dataRowStart;
          let rowBottom = dataRowEnd;
          const hasAll = expr.specials.includes("#All");
          const hasHeaders = expr.specials.includes("#Headers");
          const hasTotals = expr.specials.includes("#Totals");
          if (hasAll) {
            rowTop = tl.row;
            rowBottom = t.hasTotalsRow ? dataRowEnd + 1 : dataRowEnd;
          } else if (hasHeaders && hasTotals) {
            rowTop = tl.row;
            rowBottom = t.hasTotalsRow ? dataRowEnd + 1 : dataRowEnd;
          } else if (hasHeaders) {
            rowTop = tl.row;
            rowBottom = tl.row;
          } else if (hasTotals && t.hasTotalsRow) {
            rowTop = dataRowEnd + 1;
            rowBottom = dataRowEnd + 1;
          }
          // #This Row can't be resolved statically, but we add the data range
          // as a conservative dependency bound.

          areas.push({
            sheet: resolved.sheetName,
            top: rowTop,
            left: colLeft,
            bottom: rowBottom,
            right: colRight
          });
          break;
        }
      }
      // Cannot resolve — no static deps (will be handled at runtime)
      break;

    case BoundExprKind.Literal:
      // No static dependencies
      break;

    case BoundExprKind.NameExpr:
      // Try to resolve the defined name and extract its dependencies.
      // Guard against infinite recursion with visitedNames.
      if (nameResolver && visitedNames && !visitedNames.has(expr.upperName)) {
        visitedNames.add(expr.upperName);
        const resolved = nameResolver(expr.upperName);
        if (resolved) {
          for (const c of resolved.deps.cells) {
            cells.push(c);
          }
          for (const a of resolved.deps.areas) {
            areas.push(a);
          }
        }
      }
      break;
  }
}

// ============================================================================
// Volatility / Dynamic Ref / Dynamic Array Detection
// ============================================================================

const VOLATILE_FUNCTIONS = new Set(["RAND", "RANDBETWEEN", "NOW", "TODAY", "RANDARRAY"]);

const DYNAMIC_REF_FUNCTIONS = new Set(["INDIRECT", "OFFSET"]);

/**
 * Function names that naturally produce arrays and should trigger spill
 * behavior even when `isDynamicArray` is not explicitly set in the XLSX model.
 */
const DYNAMIC_ARRAY_FUNCTION_NAMES = new Set([
  "FILTER",
  "SORT",
  "UNIQUE",
  "SORTBY",
  "SEQUENCE",
  "RANDARRAY",
  "TOCOL",
  "TOROW",
  "CHOOSEROWS",
  "CHOOSECOLS",
  "VSTACK",
  "HSTACK",
  "WRAPROWS",
  "WRAPCOLS",
  "EXPAND",
  "TAKE",
  "DROP"
]);

/**
 * Check if the formula's top-level expression is a known dynamic array
 * function. Checks both the raw AST (for prefixed names like `_XLFN.FILTER`)
 * and the bound expression (for the canonical uppercase name).
 */
export function detectDynamicArrayFunction(ast: AstNode, bound: BoundExpr): boolean {
  // Check AST level (handles _XLFN. prefixed names)
  if (ast.type === NodeType.FunctionCall) {
    const upper = ast.name.toUpperCase();
    if (DYNAMIC_ARRAY_FUNCTION_NAMES.has(upper)) {
      return true;
    }
    // Strip prefixes
    let canonical = upper;
    if (canonical.startsWith("_XLFN._XLWS.")) {
      canonical = canonical.slice(13);
    } else if (canonical.startsWith("_XLFN.")) {
      canonical = canonical.slice(6);
    }
    if (DYNAMIC_ARRAY_FUNCTION_NAMES.has(canonical)) {
      return true;
    }
  }
  // Check bound expression level
  if (bound.kind === BoundExprKind.Call && DYNAMIC_ARRAY_FUNCTION_NAMES.has(bound.name)) {
    return true;
  }
  return false;
}

/**
 * Analyze a bound expression for volatile functions and dynamic references.
 *
 * @param nameResolver - Optional; if a NameExpr resolves to a formula-based
 *   defined name containing INDIRECT/OFFSET, the outer formula inherits
 *   `hasDynamicRefs = true`.
 */
export function analyzeExpr(
  expr: BoundExpr,
  nameResolver?: NameDepResolver
): {
  isVolatile: boolean;
  hasDynamicRefs: boolean;
  containsLambda: boolean;
} {
  let isVolatile = false;
  let hasDynamicRefs = false;
  let containsLambda = false;

  walkAnalyze(expr);

  return { isVolatile, hasDynamicRefs, containsLambda };

  function walkAnalyze(e: BoundExpr): void {
    switch (e.kind) {
      case BoundExprKind.Call:
        if (VOLATILE_FUNCTIONS.has(e.name)) {
          isVolatile = true;
        }
        for (const arg of e.args) {
          walkAnalyze(arg);
        }
        break;

      case BoundExprKind.SpecialCall:
        if (DYNAMIC_REF_FUNCTIONS.has(e.name)) {
          hasDynamicRefs = true;
        }
        if (VOLATILE_FUNCTIONS.has(e.name)) {
          isVolatile = true;
        }
        for (const arg of e.args) {
          walkAnalyze(arg);
        }
        break;

      case BoundExprKind.Lambda:
        containsLambda = true;
        walkAnalyze(e.body);
        break;

      case BoundExprKind.NameExpr:
        // If the name resolves to a formula that contains dynamic refs,
        // the outer formula inherits hasDynamicRefs.
        if (nameResolver) {
          const resolved = nameResolver(e.upperName);
          if (resolved?.hasDynamicRefs) {
            hasDynamicRefs = true;
          }
        }
        break;

      case BoundExprKind.BinaryOp:
        walkAnalyze(e.left);
        walkAnalyze(e.right);
        break;

      case BoundExprKind.UnaryOp:
        walkAnalyze(e.operand);
        break;

      case BoundExprKind.Percent:
        walkAnalyze(e.operand);
        break;

      case BoundExprKind.Array:
        for (const row of e.rows) {
          for (const elem of row) {
            walkAnalyze(elem);
          }
        }
        break;

      default:
        // Literal, CellRef, AreaRef, etc. — no children to analyze
        break;
    }
  }
}
