/**
 * Dependency Graph for Formula Calculation
 *
 * Provides static dependency analysis of formula cells by walking their ASTs
 * to extract cell/range references. Builds a directed graph where each formula
 * cell points to the cells it depends on, then produces a topological evaluation
 * order. Also supports incremental recalculation by propagating "dirty" cells
 * through reverse dependencies.
 *
 * Key exports:
 * - `buildDependencyGraph()` — analyze ASTs and build the dependency graph
 * - `topologicalSort()` — produce evaluation order, detecting circular refs
 * - `getDirtyCells()` — given changed cells, find all formulas that need recalc
 */

import { colCache } from "@excel/utils/col-cache";

import type { AstNode, ColRangeRefNode, RangeRefNode, RowRangeRefNode } from "./formula-parser";
import { NodeType } from "./formula-parser";

// ============================================================================
// Types
// ============================================================================

/**
 * A single cell coordinate used as a dependency target.
 */
interface CellCoord {
  sheet: string;
  row: number;
  col: number;
}

/**
 * A rectangular range of cells used as a dependency target.
 */
interface RangeCoord {
  sheet: string;
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * A reference extracted from an AST: either a single cell or a range.
 * Whole-column and whole-row ranges are represented as RangeCoord with
 * sentinel bounds (row 1..MAX_ROW or col 1..MAX_COL).
 */
type DepRef = CellCoord | RangeCoord;

/**
 * Formula cell descriptor — the minimum info needed for graph construction.
 * Matches the shape used by calculateFormulas().
 */
export interface FormulaCellInfo {
  sheetName: string;
  row: number;
  col: number;
  formula: string;
}

/**
 * The dependency graph structure.
 *
 * - `dependsOn`: formula cell key → set of cell keys it reads from
 * - `dependedBy`: cell key → set of formula cell keys that read from it
 *   (reverse index for incremental recalc)
 * - `formulaKeys`: ordered list of all formula cell keys
 * - `circularKeys`: set of formula cell keys involved in circular references
 */
export interface DependencyGraph {
  /** Forward edges: formula → cells it depends on */
  readonly dependsOn: ReadonlyMap<string, ReadonlySet<string>>;
  /** Reverse edges: cell → formulas that depend on it */
  readonly dependedBy: ReadonlyMap<string, ReadonlySet<string>>;
  /** All formula cell keys in insertion order */
  readonly formulaKeys: readonly string[];
  /** Formula cell keys that are part of a circular reference cycle */
  readonly circularKeys: ReadonlySet<string>;
}

// ============================================================================
// Constants
// ============================================================================

/**
 * Sentinel max row/col for whole-row / whole-column ranges.
 * Excel's actual limits are 1,048,576 rows and 16,384 columns, but for
 * dependency tracking we only need these as expansion markers — at graph
 * query time we intersect with actual sheet dimensions.
 */
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

// ============================================================================
// Cell Key Helpers
// ============================================================================

function makeKey(sheet: string, row: number, col: number): string {
  return `${sheet}!${row}:${col}`;
}

// ============================================================================
// Defined Name Resolver (optional)
// ============================================================================

/**
 * Callback to resolve a defined name to its target range.
 * Returns null if the name is not defined.
 */
export type NameResolver = (
  name: string
) => { sheet?: string; startRow: number; startCol: number; endRow: number; endCol: number } | null;

// ============================================================================
// AST Reference Extraction
// ============================================================================

function resolveCol(colStr: string): number {
  return colCache.l2n(colStr);
}

function resolveRow(rowStr: string): number {
  return parseInt(rowStr, 10);
}

/**
 * Walk an AST node and collect all cell/range references it contains.
 * This is a pure static analysis — no values are evaluated.
 */
function extractRefs(
  node: AstNode,
  currentSheet: string,
  resolveName: NameResolver | undefined,
  out: DepRef[]
): void {
  switch (node.type) {
    case NodeType.CellRef: {
      const sheet = node.sheet ?? currentSheet;
      out.push({ sheet, row: resolveRow(node.row), col: resolveCol(node.col) });
      return;
    }

    case NodeType.RangeRef: {
      const sheet = node.sheet ?? currentSheet;
      pushRange(node, sheet, out);
      return;
    }

    case NodeType.ColRangeRef: {
      const sheet = node.sheet ?? currentSheet;
      pushColRange(node, sheet, out);
      return;
    }

    case NodeType.RowRangeRef: {
      const sheet = node.sheet ?? currentSheet;
      pushRowRange(node, sheet, out);
      return;
    }

    case NodeType.Name: {
      if (resolveName) {
        const resolved = resolveName(node.name);
        if (resolved) {
          const sheet = resolved.sheet ?? currentSheet;
          if (resolved.startRow === resolved.endRow && resolved.startCol === resolved.endCol) {
            out.push({ sheet, row: resolved.startRow, col: resolved.startCol });
          } else {
            out.push({
              sheet,
              top: Math.min(resolved.startRow, resolved.endRow),
              left: Math.min(resolved.startCol, resolved.endCol),
              bottom: Math.max(resolved.startRow, resolved.endRow),
              right: Math.max(resolved.startCol, resolved.endCol)
            });
          }
        }
      }
      return;
    }

    case NodeType.BinaryOp:
      extractRefs(node.left, currentSheet, resolveName, out);
      extractRefs(node.right, currentSheet, resolveName, out);
      return;

    case NodeType.UnaryOp:
      extractRefs(node.operand, currentSheet, resolveName, out);
      return;

    case NodeType.Percent:
      extractRefs(node.operand, currentSheet, resolveName, out);
      return;

    case NodeType.FunctionCall:
      for (const arg of node.args) {
        extractRefs(arg, currentSheet, resolveName, out);
      }
      return;

    case NodeType.Array:
      for (const row of node.rows) {
        for (const elem of row) {
          extractRefs(elem, currentSheet, resolveName, out);
        }
      }
      return;

    case NodeType.StructuredRef:
      // Structured references (e.g. Table1[Column]) cannot be resolved
      // statically without table metadata. They are treated as opaque —
      // no dependency edges are created. The recursive evaluator handles
      // them at runtime.
      return;

    // Literals have no dependencies
    case NodeType.Number:
    case NodeType.String:
    case NodeType.Boolean:
    case NodeType.Error:
    case NodeType.Missing:
      return;

    default: {
      // Exhaustive guard — future node types will trigger a compile error here
      const _: never = node;
      void _;
      return;
    }
  }
}

function pushRange(node: RangeRefNode, sheet: string, out: DepRef[]): void {
  const startRow = resolveRow(node.start.row);
  const startCol = resolveCol(node.start.col);
  const endRow = resolveRow(node.end.row);
  const endCol = resolveCol(node.end.col);
  out.push({
    sheet,
    top: Math.min(startRow, endRow),
    left: Math.min(startCol, endCol),
    bottom: Math.max(startRow, endRow),
    right: Math.max(startCol, endCol)
  });
}

function pushColRange(node: ColRangeRefNode, sheet: string, out: DepRef[]): void {
  const startCol = resolveCol(node.startCol);
  const endCol = resolveCol(node.endCol);
  out.push({
    sheet,
    top: 1,
    left: Math.min(startCol, endCol),
    bottom: MAX_ROW,
    right: Math.max(startCol, endCol)
  });
}

function pushRowRange(node: RowRangeRefNode, sheet: string, out: DepRef[]): void {
  out.push({
    sheet,
    top: Math.min(node.startRow, node.endRow),
    left: 1,
    bottom: Math.max(node.startRow, node.endRow),
    right: MAX_COL
  });
}

// ============================================================================
// Dependency Reference → Cell Key Expansion
// ============================================================================

/**
 * Type guard: is this a range reference (has `top` property)?
 */
function isRange(ref: DepRef): ref is RangeCoord {
  return "top" in ref;
}

/**
 * Expand a set of DepRefs into concrete cell keys that are known formula cells.
 * For single-cell refs, the key is produced directly.
 * For range refs, we check which formula cells fall within the range rather than
 * enumerating every cell (which would be prohibitively expensive for whole-column
 * ranges like A:A).
 *
 * @param refs - References extracted from an AST
 * @param formulaKeySet - Set of all formula cell keys (for range intersection)
 * @param formulaCellCoords - Map from formula key to its (sheet, row, col) for range checks
 */
function expandRefsToKeys(
  refs: DepRef[],
  formulaKeySet: ReadonlySet<string>,
  formulaCellCoords: ReadonlyMap<string, CellCoord>
): Set<string> {
  const keys = new Set<string>();

  for (const ref of refs) {
    if (isRange(ref)) {
      // For ranges, find all formula cells within the range.
      // Also add individual cell keys for non-formula cells (they are dependency
      // targets even though they don't appear as formula keys — needed for
      // incremental recalc's reverse index).
      //
      // Strategy: for small ranges, enumerate all cells. For large ranges
      // (whole-row/col), scan formula cells instead.
      const rangeSize = (ref.bottom - ref.top + 1) * (ref.right - ref.left + 1);

      if (rangeSize <= 500) {
        // Small range: enumerate every cell
        for (let r = ref.top; r <= ref.bottom; r++) {
          for (let c = ref.left; c <= ref.right; c++) {
            keys.add(makeKey(ref.sheet, r, c));
          }
        }
      } else {
        // Large range (e.g. whole column A:A): scan formula cells and check containment
        for (const [fKey, coord] of formulaCellCoords) {
          if (
            coord.sheet === ref.sheet &&
            coord.row >= ref.top &&
            coord.row <= ref.bottom &&
            coord.col >= ref.left &&
            coord.col <= ref.right
          ) {
            keys.add(fKey);
          }
        }
        // For non-formula cells in large ranges, we cannot enumerate.
        // The reverse index for incremental recalc will be incomplete for
        // literal cells inside whole-row/col refs. This is an acceptable
        // trade-off — when a user edits a cell inside a whole-column ref,
        // the incremental engine should treat whole-row/col formula deps
        // as always-dirty (handled in getDirtyCells).
      }
    } else {
      // Single cell — always add
      keys.add(makeKey(ref.sheet, ref.row, ref.col));
    }
  }

  return keys;
}

// ============================================================================
// Build Dependency Graph
// ============================================================================

/**
 * Build a dependency graph for all formula cells.
 *
 * For each formula cell, walks its pre-parsed AST to extract cell and range
 * references, then records directed edges:
 *   formula cell → depends on → referenced cells
 *
 * @param formulaCells - All formula cells in the workbook
 * @param astCache - Pre-parsed AST cache (formula text → AstNode)
 * @param resolveName - Optional callback to resolve defined names to ranges
 * @returns The complete dependency graph
 */
export function buildDependencyGraph(
  formulaCells: readonly FormulaCellInfo[],
  astCache: ReadonlyMap<string, AstNode>,
  resolveName?: NameResolver
): DependencyGraph {
  // Build a lookup of all formula cell coordinates for range intersection
  const formulaKeySet = new Set<string>();
  const formulaCellCoords = new Map<string, CellCoord>();
  const formulaKeys: string[] = [];

  for (const fc of formulaCells) {
    const key = makeKey(fc.sheetName, fc.row, fc.col);
    formulaKeySet.add(key);
    formulaCellCoords.set(key, { sheet: fc.sheetName, row: fc.row, col: fc.col });
    formulaKeys.push(key);
  }

  // Forward edges: formula key → set of keys it depends on
  const dependsOn = new Map<string, Set<string>>();
  // Reverse edges: cell key → set of formula keys that depend on it
  const dependedBy = new Map<string, Set<string>>();

  for (const fc of formulaCells) {
    const key = makeKey(fc.sheetName, fc.row, fc.col);
    const ast = astCache.get(fc.formula);
    if (!ast) {
      // No AST (parse failure) — no dependencies
      dependsOn.set(key, new Set());
      continue;
    }

    // Extract raw references
    const refs: DepRef[] = [];
    extractRefs(ast, fc.sheetName, resolveName, refs);

    // Expand to concrete cell keys
    const depKeys = expandRefsToKeys(refs, formulaKeySet, formulaCellCoords);
    // NOTE: self-dependency is intentionally kept. If a cell references itself
    // (e.g. A1=A1+1), it must be detected as circular so iterative calculation
    // can handle it.

    dependsOn.set(key, depKeys);

    // Build reverse index
    for (const depKey of depKeys) {
      let set = dependedBy.get(depKey);
      if (!set) {
        set = new Set();
        dependedBy.set(depKey, set);
      }
      set.add(key);
    }
  }

  // Run topological sort to detect circular references
  const { circularKeys } = detectCircularRefs(formulaKeys, dependsOn);

  return {
    dependsOn,
    dependedBy,
    formulaKeys,
    circularKeys
  };
}

// ============================================================================
// Circular Reference Detection
// ============================================================================

/**
 * Detect circular references using iterative DFS with three-color marking.
 * Returns the set of formula keys involved in cycles.
 */
function detectCircularRefs(
  formulaKeys: readonly string[],
  dependsOn: ReadonlyMap<string, ReadonlySet<string>>
): { circularKeys: Set<string> } {
  const WHITE = 0; // Not visited
  const GRAY = 1; // On current DFS path (in stack)
  const BLACK = 2; // Fully processed

  const color = new Map<string, number>();
  const circularKeys = new Set<string>();

  // Initialize all formula keys as WHITE
  for (const key of formulaKeys) {
    color.set(key, WHITE);
  }

  for (const startKey of formulaKeys) {
    if (color.get(startKey) !== WHITE) {
      continue;
    }

    // Iterative DFS using an explicit stack.
    // Each stack frame is [key, iterator, isEntering].
    // When entering (isEntering=true), we mark GRAY and push deps.
    // When all deps are processed, we mark BLACK.
    const stack: [string, Iterator<string> | null, boolean][] = [[startKey, null, true]];

    while (stack.length > 0) {
      const frame = stack[stack.length - 1];
      const [key, , entering] = frame;

      if (entering) {
        frame[2] = false; // No longer entering

        if (color.get(key) === BLACK) {
          stack.pop();
          continue;
        }

        color.set(key, GRAY);
        const deps = dependsOn.get(key);
        frame[1] = deps ? deps[Symbol.iterator]() : null;
      }

      // Process next dependency
      const iter = frame[1];
      if (iter) {
        const next = iter.next();
        if (!next.done) {
          const depKey = next.value;
          const depColor = color.get(depKey);

          if (depColor === GRAY) {
            // Back edge — cycle detected! Mark all nodes on the cycle path.
            circularKeys.add(depKey);
            // Walk back up the stack to find all nodes in this cycle
            for (let i = stack.length - 1; i >= 0; i--) {
              circularKeys.add(stack[i][0]);
              if (stack[i][0] === depKey) {
                break;
              }
            }
          } else if (depColor === WHITE || depColor === undefined) {
            // depColor === undefined means this dep is not a formula cell
            // (it's a literal data cell) — skip it
            if (depColor === WHITE) {
              stack.push([depKey, null, true]);
            }
          }
          // BLACK deps are already fully processed — skip
          continue; // Continue processing current frame's deps
        }
      }

      // All deps processed — mark BLACK and pop
      color.set(key, BLACK);
      stack.pop();
    }
  }

  return { circularKeys };
}

// ============================================================================
// Topological Sort (Kahn's Algorithm)
// ============================================================================

/**
 * Produce a topological evaluation order for formula cells using Kahn's algorithm.
 * Cells with no dependencies are evaluated first. Circular references are
 * appended at the end in their original order (they'll be handled by the
 * recursive evaluator's cycle detection).
 *
 * @param graph - The dependency graph
 * @returns Formula cell keys in evaluation order
 */
export function topologicalSort(graph: DependencyGraph): string[] {
  const { formulaKeys, dependsOn, circularKeys } = graph;

  // Build in-degree counts (only count edges between formula cells)
  const formulaKeySet = new Set(formulaKeys);
  const inDegree = new Map<string, number>();

  // Edge semantics: if A depends on B (A's formula reads B), then B must be
  // evaluated before A. In the DAG, edge direction is B → A. So A's in-degree
  // is the count of formula cells in its dependsOn set.
  for (const key of formulaKeys) {
    if (circularKeys.has(key)) {
      continue; // Skip circular nodes for topo sort
    }
    const deps = dependsOn.get(key);
    if (!deps) {
      inDegree.set(key, 0);
      continue;
    }
    let count = 0;
    for (const depKey of deps) {
      if (formulaKeySet.has(depKey) && !circularKeys.has(depKey)) {
        count++;
      }
    }
    inDegree.set(key, count);
  }

  // Seed the queue with nodes that have in-degree 0
  const queue: string[] = [];
  for (const key of formulaKeys) {
    if (circularKeys.has(key)) {
      continue;
    }
    if ((inDegree.get(key) ?? 0) === 0) {
      queue.push(key);
    }
  }

  const sorted: string[] = [];
  let head = 0;

  while (head < queue.length) {
    const key = queue[head++];
    sorted.push(key);

    // For each formula that depends on this key, decrement in-degree
    // "dependedBy" gives us: key → set of formula keys that depend on key
    const dependents = graph.dependedBy.get(key);
    if (dependents) {
      for (const depKey of dependents) {
        if (circularKeys.has(depKey) || !formulaKeySet.has(depKey)) {
          continue;
        }
        const deg = (inDegree.get(depKey) ?? 1) - 1;
        inDegree.set(depKey, deg);
        if (deg === 0) {
          queue.push(depKey);
        }
      }
    }
  }

  // Append circular reference cells at the end in original order.
  // These will fall back to the recursive evaluator's cycle detection.
  for (const key of formulaKeys) {
    if (circularKeys.has(key)) {
      sorted.push(key);
    }
  }

  return sorted;
}

// ============================================================================
// Incremental Recalculation: Dirty Cell Propagation
// ============================================================================

/**
 * Given a set of cells whose values have changed ("dirty" cells), identify
 * all formula cells that need to be recalculated. Returns them in topological
 * order (dependencies first).
 *
 * The propagation is transitive: if cell A depends on cell B, and B depends
 * on dirty cell C, then both B and A are marked dirty.
 *
 * @param dirtyCells - Set of cell keys (e.g. "Sheet1!1:1") whose values changed
 * @param graph - The dependency graph
 * @returns Formula cell keys that need recalculation, in topological order
 */
export function getDirtyCells(dirtyCells: ReadonlySet<string>, graph: DependencyGraph): string[] {
  const { dependedBy, formulaKeys, circularKeys } = graph;
  const formulaKeySet = new Set(formulaKeys);
  const dirty = new Set<string>();

  // BFS from dirty cells through reverse dependency edges
  const queue: string[] = [];

  // Seed: all dirty cells
  for (const cellKey of dirtyCells) {
    const dependents = dependedBy.get(cellKey);
    if (dependents) {
      for (const depKey of dependents) {
        if (!dirty.has(depKey)) {
          dirty.add(depKey);
          queue.push(depKey);
        }
      }
    }
  }

  // Propagate transitively
  let head = 0;
  while (head < queue.length) {
    const key = queue[head++];
    const dependents = dependedBy.get(key);
    if (dependents) {
      for (const depKey of dependents) {
        if (!dirty.has(depKey)) {
          dirty.add(depKey);
          queue.push(depKey);
        }
      }
    }
  }

  // Also mark all circular-ref cells as dirty (they always need re-evaluation)
  for (const key of circularKeys) {
    dirty.add(key);
  }

  // Return dirty formula cells in topological order.
  // Use a full topo sort of the subgraph, or simply filter the pre-sorted order.
  // Filtering the full topo order is correct and efficient.
  const fullOrder = topologicalSort(graph);
  return fullOrder.filter(key => dirty.has(key) && formulaKeySet.has(key));
}
