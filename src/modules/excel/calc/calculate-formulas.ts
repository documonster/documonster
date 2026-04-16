/**
 * Formula Calculation Engine
 *
 * Provides `calculateFormulas()` to recalculate all formula cells in a workbook.
 * Evaluates each formula using the tokenizer → parser → evaluator pipeline with
 * memoization and circular reference detection via an `EvalSession`.
 *
 * This module is the sole public entry point for formula calculation.
 */

import type { Cell, FormulaResult } from "@excel/cell";
import { Enums } from "@excel/enums";
import type { CellErrorValue } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import type { Worksheet } from "@excel/worksheet";

import { buildDependencyGraph, topologicalSort, type NameResolver } from "./dependency-graph";
import {
  evaluate,
  evaluateFormula,
  evaluateFormulaRaw,
  EvalSession,
  type EvalContext,
  type ResolvedName,
  type ResolvedStructuredRef
} from "./formula-evaluator";
import { isError, type CalcArray, type CalcValue } from "./formula-functions";
import { parse, NodeType, type AstNode } from "./formula-parser";
import { tokenize } from "./formula-tokenizer";

// ============================================================================
// Workbook Interface (to avoid circular imports)
// ============================================================================

interface DefinedNamesLike {
  getRanges(name: string): { name: string; ranges: string[] };
}

interface WorkbookLike {
  worksheets: Worksheet[];
  getWorksheet(id?: number | string): Worksheet | undefined;
  definedNames?: DefinedNamesLike;
  /** Calculation properties — used for iterative calculation settings. */
  calcProperties?: {
    fullCalcOnLoad?: boolean;
    /** Enable iterative calculation for circular references. */
    iterate?: boolean;
    /** Maximum number of iterations (default 100). */
    iterateCount?: number;
    /** Maximum change threshold for convergence (default 0.001). */
    iterateDelta?: number;
  };
  /** Workbook properties including date system. */
  properties?: {
    date1904?: boolean;
  };
}

// ============================================================================
// Formula Cell Info
// ============================================================================

interface FormulaCell {
  sheet: Worksheet;
  sheetName: string;
  row: number;
  col: number;
  cell: Cell;
  formula: string;
}

// ============================================================================
// Collect All Formula Cells
// ============================================================================

/**
 * Parse a defined name range string like "Sheet1!$A$1:$B$2" or "'Sheet Name'!$C$3"
 * into a ResolvedName. Returns null if the format is unrecognized.
 */
const definedNameRangeRegex =
  /^(?:'([^']*(?:''[^']*)*)'|([^!]+))!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/;

function parseDefinedNameRange(rangeStr: string): ResolvedName | null {
  const m = definedNameRangeRegex.exec(rangeStr);
  if (!m) {
    return null;
  }
  // Sheet name: quoted (group 1, with '' unescaping) or unquoted (group 2)
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const startCol = colCache.l2n(m[3]);
  const startRow = parseInt(m[4], 10);
  const endCol = m[5] ? colCache.l2n(m[5]) : startCol;
  const endRow = m[6] ? parseInt(m[6], 10) : startRow;
  return { sheet, startRow, startCol, endRow, endCol };
}

/**
 * Resolve a defined name to a single-area cell/range reference.
 * Multi-area names (e.g. "A1:B2,C3:D4") are rejected — returns null.
 * Used by both the runtime evaluator (resolveName) and the dependency graph
 * builder (nameResolver) to ensure consistent semantics.
 */
function resolveDefinedNameToRange(dn: DefinedNamesLike, name: string): ResolvedName | null {
  const { ranges } = dn.getRanges(name);
  if (!ranges || ranges.length === 0) {
    return null;
  }
  // Multi-area defined names are not supported — reject rather than
  // silently using only the first range.
  if (ranges.length > 1) {
    return null;
  }
  return parseDefinedNameRange(ranges[0]);
}

// ============================================================================
// Spill Engine Types
// ============================================================================

/**
 * Tracks a spill region: the source formula cell and the range of cells it
 * has spilled into. Used for cleanup when a formula is re-evaluated with
 * different-sized results.
 */
interface SpillRegion {
  /** The worksheet id — stable across renames */
  worksheetId: number;
  /** The source formula cell's row */
  sourceRow: number;
  /** The source formula cell's col */
  sourceCol: number;
  /** Number of rows in the spill (including source) */
  rows: number;
  /** Number of cols in the spill (including source) */
  cols: number;
}

/**
 * Persistent spill metadata: survives across calculateFormulas invocations.
 * Key: "SheetName!row:col" of the source cell → SpillRegion.
 * Stored in a WeakMap keyed by the workbook object to allow GC.
 */
const persistentSpillRegions = new WeakMap<WeakKey, Map<string, SpillRegion>>();

function getPersistentSpillMap(workbook: WorkbookLike): Map<string, SpillRegion> {
  let map = persistentSpillRegions.get(workbook);
  if (!map) {
    map = new Map();
    persistentSpillRegions.set(workbook, map);
  }
  return map;
}

/**
 * Persistent snapshot of values written to ghost (spill target) cells.
 * Key: ghost cell key "SheetName!row:col" → raw value written by spill.
 * Used to detect whether a ghost cell has been modified by the user since
 * the last calculation.
 */
const persistentGhostSnapshots = new WeakMap<WeakKey, Map<string, unknown>>();

function getGhostSnapshots(workbook: WorkbookLike): Map<string, unknown> {
  let map = persistentGhostSnapshots.get(workbook);
  if (!map) {
    map = new Map();
    persistentGhostSnapshots.set(workbook, map);
  }
  return map;
}

/**
 * Check whether a ghost cell's current value matches the snapshot recorded
 * when we last wrote it via spill. Returns true if the cell is still
 * unmodified (i.e. value matches what we wrote), false if the user changed it.
 * A cell that doesn't exist or is Null-type is always considered unmodified
 * (it was cleared, which is fine — we treat it as still "ours").
 */
function isGhostCellUnmodified(
  cell: Cell | undefined,
  ghostKey: string,
  ghostSnapshots: Map<string, unknown>
): boolean {
  if (!cell) {
    return true; // Cell doesn't exist — empty, so still "ours"
  }
  if (cell.type === Enums.ValueType.Null) {
    return true; // Empty cell — still "ours"
  }
  // If the cell now has a formula, user definitely changed it
  if (cell.formula) {
    return false;
  }
  // Compare current value with what we wrote
  const snapshot = ghostSnapshots.get(ghostKey);
  if (snapshot === undefined) {
    // No snapshot recorded — this ghost predates snapshot tracking.
    // Conservatively treat it as unmodified to avoid breaking existing
    // spills. On the next write cycle, a snapshot will be recorded.
    return true;
  }
  return valuesEqual(cell.value, snapshot);
}

/**
 * Shallow equality check for cell values, handling error objects.
 */
function valuesEqual(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  // NaN === NaN is false, but we consider two NaN values equal
  if (typeof a === "number" && typeof b === "number" && Number.isNaN(a) && Number.isNaN(b)) {
    return true;
  }
  // Handle error objects like { error: "#REF!" }
  if (
    a !== null &&
    b !== null &&
    typeof a === "object" &&
    typeof b === "object" &&
    "error" in (a as Record<string, unknown>) &&
    "error" in (b as Record<string, unknown>)
  ) {
    return (a as Record<string, unknown>).error === (b as Record<string, unknown>).error;
  }
  // Handle Date comparison
  if (a instanceof Date && b instanceof Date) {
    return a.getTime() === b.getTime();
  }
  return false;
}

/**
 * Key for tracking which cells are "ghost" (spill target) cells.
 * Maps from "SheetName!row:col" → source cell key "SheetName!row:col"
 */
type SpillGhostMap = Map<string, string>;

/**
 * Build a cell key for spill tracking, keyed by worksheet id (stable across
 * renames) rather than sheet name.  The `ws:` prefix ensures this key space
 * never collides with the `SheetName!row:col` keys used by the eval cache
 * and dependency graph.
 */
function spillCellKey(worksheetId: number, row: number, col: number): string {
  return `ws:${worksheetId}!${row}:${col}`;
}

/**
 * Clean up stale spill regions whose source formula cells no longer exist.
 * This must run even when there are no formula cells remaining, so that
 * ghost cells from deleted formulas are properly cleaned up.
 *
 * @param workbook - The workbook to clean up
 * @param formulaKeys - Set of cell keys for all current formula cells
 */
function cleanupStaleSpillRegions(workbook: WorkbookLike, formulaKeys: Set<string>): void {
  const persistentSpills = getPersistentSpillMap(workbook);
  const snapshots = getGhostSnapshots(workbook);

  for (const [srcKey, region] of [...persistentSpills]) {
    if (!formulaKeys.has(srcKey)) {
      // Source formula cell no longer exists — clean up ghosts
      const ws = workbook.getWorksheet(region.worksheetId);
      if (ws) {
        for (let r = 0; r < region.rows; r++) {
          for (let c = 0; c < region.cols; c++) {
            if (r === 0 && c === 0) {
              continue;
            }
            const targetKey = spillCellKey(
              region.worksheetId,
              region.sourceRow + r,
              region.sourceCol + c
            );
            const ghostCell = ws.findCell(region.sourceRow + r, region.sourceCol + c);
            // Only clear the cell if it hasn't been modified by the user
            if (ghostCell && isGhostCellUnmodified(ghostCell, targetKey, snapshots)) {
              ghostCell.value = null;
            }
            snapshots.delete(targetKey);
          }
        }
      }
      persistentSpills.delete(srcKey);
    }
  }
}

/**
 * Parse a cell ref string like "A1:B2" into { top, left, bottom, right }.
 */
function parseRefRange(
  ref: string
): { top: number; left: number; bottom: number; right: number } | null {
  const parts = ref.split(":");
  if (parts.length !== 2) {
    return null;
  }
  try {
    const tl = colCache.decodeAddress(parts[0]);
    const br = colCache.decodeAddress(parts[1]);
    return {
      top: Math.min(tl.row, br.row),
      left: Math.min(tl.col, br.col),
      bottom: Math.max(tl.row, br.row),
      right: Math.max(tl.col, br.col)
    };
  } catch {
    return null;
  }
}

// ============================================================================
// CalcValue ↔ Cell Value Conversion Helpers
// ============================================================================

/**
 * Convert a CalcValue to a FormulaResult suitable for `cell.result`.
 */
function calcValueToResult(val: CalcValue): FormulaResult {
  if (val === null) {
    return 0;
  }
  if (typeof val === "number" || typeof val === "string" || typeof val === "boolean") {
    return val;
  }
  if (val instanceof Date) {
    return val;
  }
  if (isError(val)) {
    return val as FormulaResult;
  }
  return 0;
}

/**
 * Convert a CalcValue to a raw cell value suitable for `cell.value`.
 * Used for writing spill ghost cells.
 */
function calcValueToRawValue(
  val: CalcValue
): number | string | boolean | Date | CellErrorValue | null {
  if (val === null) {
    return null;
  }
  if (typeof val === "number" || typeof val === "string" || typeof val === "boolean") {
    return val;
  }
  if (val instanceof Date) {
    return val;
  }
  if (isError(val)) {
    return val;
  }
  return null;
}

function collectFormulaCells(workbook: WorkbookLike): FormulaCell[] {
  const cells: FormulaCell[] = [];

  for (const sheet of workbook.worksheets) {
    sheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (cell.type === Enums.ValueType.Formula) {
          // `cell.formula` triggers shared-formula translation internally
          // via FormulaValue._getTranslatedFormula() for slave cells
          const formula = cell.formula;
          if (formula) {
            cells.push({
              sheet,
              sheetName: sheet.name,
              row: rowNumber,
              col: colNumber,
              cell,
              formula
            });
          }
        }
      });
    });
  }

  return cells;
}

/**
 * Set of function names that naturally produce arrays and should trigger
 * spill behavior even when isDynamicArray is not explicitly set.
 */
const DYNAMIC_ARRAY_FUNCTIONS = new Set([
  "FILTER",
  "_XLFN._XLWS.FILTER",
  "SORT",
  "_XLFN._XLWS.SORT",
  "UNIQUE",
  "_XLFN._XLWS.UNIQUE",
  "SORTBY",
  "_XLFN._XLWS.SORTBY",
  "SEQUENCE",
  "_XLFN.SEQUENCE",
  "RANDARRAY",
  "_XLFN.RANDARRAY",
  "TOCOL",
  "_XLFN.TOCOL",
  "TOROW",
  "_XLFN.TOROW",
  "CHOOSEROWS",
  "_XLFN.CHOOSEROWS",
  "CHOOSECOLS",
  "_XLFN.CHOOSECOLS",
  "VSTACK",
  "_XLFN.VSTACK",
  "HSTACK",
  "_XLFN.HSTACK",
  "WRAPROWS",
  "_XLFN.WRAPROWS",
  "WRAPCOLS",
  "_XLFN.WRAPCOLS",
  "EXPAND",
  "_XLFN.EXPAND",
  "TAKE",
  "_XLFN.TAKE",
  "DROP",
  "_XLFN.DROP"
]);

/**
 * Check if an AST's top-level function is a known dynamic array function.
 */
function isDynamicArrayFormula(ast: AstNode): boolean {
  if (ast.type === NodeType.FunctionCall) {
    return DYNAMIC_ARRAY_FUNCTIONS.has(ast.name.toUpperCase());
  }
  return false;
}

// ============================================================================
// AST Cache (avoids re-parsing the same formula text)
// ============================================================================

function parseFormula(formula: string, astCache: Map<string, AstNode>): AstNode | null {
  const cached = astCache.get(formula);
  if (cached) {
    return cached;
  }
  try {
    const tokens = tokenize(formula);
    const ast = parse(tokens);
    astCache.set(formula, ast);
    return ast;
  } catch {
    return null;
  }
}

// ============================================================================
// Build Evaluation Context
// ============================================================================

function buildContext(
  workbook: WorkbookLike,
  currentSheet: Worksheet,
  session: EvalSession,
  astCache: Map<string, AstNode>
): EvalContext {
  const ctx: EvalContext = {
    currentSheet: currentSheet.name,

    resolveName(name: string): ResolvedName | null {
      const dn = workbook.definedNames;
      if (!dn) {
        return null;
      }
      return resolveDefinedNameToRange(dn, name);
    },

    resolveNameValue(name: string): CalcValue | CalcArray | undefined {
      const dn = workbook.definedNames;
      if (!dn) {
        return undefined;
      }
      const { ranges } = dn.getRanges(name);
      if (!ranges || ranges.length === 0) {
        return undefined;
      }
      // Multi-area defined names are not supported — return #VALUE! error
      if (ranges.length > 1) {
        return { error: "#VALUE!" };
      }
      const rangeStr = ranges[0];
      // If the range string looks like a cell reference, skip — resolveName handles it
      if (parseDefinedNameRange(rangeStr)) {
        return undefined;
      }
      // The range string is a formula expression (e.g. "LAMBDA(x,y,x+y)")
      // Check the cache to avoid re-evaluating
      const cacheKey = `__NAME__${name.toUpperCase()}`;
      const cached = session.cache.get(cacheKey);
      if (cached !== undefined) {
        return cached;
      }
      // Parse and evaluate the formula
      const ast = parseFormula(rangeStr, astCache);
      if (!ast) {
        return undefined;
      }
      const savedSheet = ctx.currentSheet;
      try {
        const result = evaluate(ast, ctx, session);
        // Cache the full result so subsequent references to the same name
        // in the same calculation round get the complete value (including
        // arrays). Previously only the scalar (top-left) was cached, which
        // caused array-returning formula names to degrade on second use.
        session.cache.set(cacheKey, result);
        return result;
      } catch {
        return undefined;
      } finally {
        ctx.currentSheet = savedSheet;
      }
    },

    getSheetDimensions(
      sheetName: string | undefined
    ): { top: number; bottom: number; left: number; right: number } | null {
      const ws = sheetName ? workbook.getWorksheet(sheetName) : currentSheet;
      if (!ws) {
        return null;
      }
      const dims = ws.dimensions;
      if (!dims) {
        return null;
      }
      return { top: dims.top, bottom: dims.bottom, left: dims.left, right: dims.right };
    },

    getSheetNames(): string[] {
      return workbook.worksheets.map(ws => ws.name);
    },

    resolveStructuredRef(
      tableName: string,
      columns: string[],
      specials: string[]
    ): ResolvedStructuredRef | null {
      // Find the table in any worksheet
      let table: {
        columns: { name: string }[];
        tl?: { row: number; col: number };
        rows: unknown[][];
        headerRow?: boolean;
        totalsRow?: boolean;
        tableRef?: string;
      } | null = null;
      let tableSheet: Worksheet | null = null;

      for (const ws of workbook.worksheets) {
        if (!ws.getTables) {
          continue;
        }
        for (const t of ws.getTables()) {
          const tModel = t.table;
          const tName = tModel.name || tModel.displayName || "";
          if (tableName === "" || tName.toLowerCase() === tableName.toLowerCase()) {
            // For empty tableName (implicit @), match if current cell is within this table
            if (tableName === "") {
              const addr = ctx.currentAddress;
              if (!addr || addr.sheet !== ws.name) {
                continue;
              }
              if (!tModel.tl || !tModel.tableRef) {
                continue;
              }
              const tl = tModel.tl;
              const width = tModel.columns.length;
              const hasHeader = tModel.headerRow !== false;
              const dataStart = tl.row + (hasHeader ? 1 : 0);
              const dataEnd = dataStart + tModel.rows.length - 1;
              if (
                addr.row < dataStart ||
                addr.row > dataEnd ||
                addr.col < tl.col ||
                addr.col >= tl.col + width
              ) {
                continue;
              }
            }
            table = tModel;
            tableSheet = ws;
            break;
          }
        }
        if (table) {
          break;
        }
      }

      if (!table || !tableSheet || !table.tl) {
        return null;
      }

      const tl = table.tl;
      const width = table.columns.length;
      const hasHeader = table.headerRow !== false;
      const hasTotals = table.totalsRow === true;
      const dataRowStart = tl.row + (hasHeader ? 1 : 0);
      const dataRowEnd = dataRowStart + table.rows.length - 1;

      // Determine column range
      let colLeft = tl.col;
      let colRight = tl.col + width - 1;

      if (columns.length > 0) {
        const indices: number[] = [];
        for (const colName of columns) {
          const idx = table.columns.findIndex(c => c.name.toLowerCase() === colName.toLowerCase());
          if (idx === -1) {
            return null; // Column not found
          }
          indices.push(idx);
        }
        colLeft = tl.col + Math.min(...indices);
        colRight = tl.col + Math.max(...indices);
      }

      // Determine row range based on specials
      let rowTop = dataRowStart;
      let rowBottom = dataRowEnd;

      const hasThisRow = specials.includes("#This Row");
      const hasHeaders = specials.includes("#Headers");
      const hasTotalsSpec = specials.includes("#Totals");
      const hasAll = specials.includes("#All");
      const hasData = specials.includes("#Data");

      if (hasAll) {
        rowTop = tl.row;
        rowBottom = hasTotals ? dataRowEnd + 1 : dataRowEnd;
      } else if (hasThisRow) {
        const addr = ctx.currentAddress;
        if (addr) {
          rowTop = addr.row;
          rowBottom = addr.row;
        }
      } else if (hasHeaders && hasTotalsSpec) {
        rowTop = tl.row;
        rowBottom = hasTotals ? dataRowEnd + 1 : dataRowEnd;
      } else if (hasHeaders) {
        if (hasHeader) {
          rowTop = tl.row;
          rowBottom = tl.row;
        }
      } else if (hasTotalsSpec) {
        if (hasTotals) {
          rowTop = dataRowEnd + 1;
          rowBottom = dataRowEnd + 1;
        } else {
          return null; // No totals row
        }
      } else if (hasData || specials.length === 0) {
        // Default: data rows
        rowTop = dataRowStart;
        rowBottom = dataRowEnd;
      }

      return {
        sheetName: tableSheet.name,
        top: rowTop,
        left: colLeft,
        bottom: rowBottom,
        right: colRight
      };
    },

    getCellValue(sheetName: string | undefined, row: number, col: number): CalcValue {
      const ws = sheetName ? workbook.getWorksheet(sheetName) : currentSheet;
      if (!ws) {
        return null;
      }

      const cell = ws.findCell(row, col);
      if (!cell) {
        return null;
      }

      const cellType = cell.type;

      // If this cell is a formula, recursively evaluate it
      if (cellType === Enums.ValueType.Formula) {
        const key = session.makeKey(ws.name, row, col);

        // Check memo cache first — avoids re-evaluating the same cell
        const cachedResult = session.cache.get(key);
        if (cachedResult !== undefined && !Array.isArray(cachedResult)) {
          return cachedResult;
        }

        // Get the formula text (triggers shared-formula translation)
        const formula = cell.formula;
        if (formula) {
          const ast = parseFormula(formula, astCache);
          if (ast) {
            const savedSheet = ctx.currentSheet;
            ctx.currentSheet = ws.name;
            try {
              const result = evaluateFormula(ast, ctx, session, ws.name, row, col);
              // Only overwrite cached result if engine produced a real value
              if (!isError(result) || result.error !== "#NAME?" || cell.result === undefined) {
                cell.result = result as FormulaResult;
              }
              return result;
            } catch {
              return (cell.result as CalcValue) ?? null;
            } finally {
              ctx.currentSheet = savedSheet;
            }
          }
        }

        // Formula cell with no parseable formula — return cached result
        return (cell.result as CalcValue) ?? null;
      }

      // Non-formula cell — extract value
      const value = cell.value;
      if (value === null || value === undefined) {
        return null;
      }
      if (typeof value === "number") {
        return value;
      }
      if (typeof value === "string") {
        return value;
      }
      if (typeof value === "boolean") {
        return value;
      }
      if (value instanceof Date) {
        return value;
      }
      if (typeof value === "object" && "error" in value) {
        return value as CellErrorValue;
      }

      // Rich text — extract text content
      if (typeof value === "object" && "richText" in value) {
        return (value as { richText: { text: string }[] }).richText.map(r => r.text).join("");
      }
      if (typeof value === "object" && "text" in value) {
        return (value as { text: string }).text;
      }

      return cell.text || null;
    }
  };

  return ctx;
}

// ============================================================================
// Main: Calculate All Formulas
// ============================================================================

/**
 * Recalculate all formula cells in a workbook.
 *
 * Evaluates every formula cell using the built-in calculation engine
 * and updates each cell's `result` value. Formulas are evaluated lazily
 * with recursive dependency resolution, memoization, and circular
 * reference detection.
 *
 * All evaluation state is scoped to this invocation — concurrent calls
 * for different workbooks are safe.
 *
 * **Supported formula features:**
 * - Cell references: `A1`, `$B$2`, `Sheet1!A1`, `'Sheet Name'!A1:B10`
 * - Operators: `+ - * / ^`, `& (concat)`, `= <> < > <= >=`, `%`
 * - 120+ built-in functions (SUM, IF, VLOOKUP, SUMIF, FILTER, etc.)
 * - Shared formulas, array constants, nested expressions
 * - Dynamic array spill: FILTER, SORT, UNIQUE, SORTBY results are
 *   written to adjacent cells. #SPILL! error if target cells are occupied.
 * - CSE array formulas: `{=formula}` with a ref range distribute results
 *   across the designated range.
 * - Array arithmetic broadcasting: `{1,2,3} + {4;5;6}` produces a 3x3 matrix.
 * - Implicit intersection: range references in scalar context pick the
 *   value from the formula cell's row or column.
 *
 * **Unsupported formula behavior:**
 * - If a formula uses a function the engine does not implement, the engine
 *   returns `#NAME?`. However, if the cell already has a cached result
 *   (e.g., pre-computed by Excel when the XLSX was saved), that cached
 *   result is **preserved** — the engine will not overwrite usable data.
 * - If no cached result exists, the cell's result becomes `#NAME?`.
 *
 * **Volatile functions:**
 * - `RAND`, `RANDBETWEEN`, `NOW`, `TODAY` are re-evaluated on every call.
 *   This is intentional — these functions are expected to produce fresh values.
 *
 * **Side effects:**
 * - This function **mutates** the workbook by updating formula cells' `result`
 *   property in-place. For dynamic array formulas, adjacent cells are also
 *   written with spill results. If you need the original cached results
 *   preserved, clone the workbook before calling this function.
 *
 * @param workbook - The workbook whose formulas should be recalculated
 */
export function calculateFormulas(workbook: WorkbookLike): void {
  const formulaCells = collectFormulaCells(workbook);

  if (formulaCells.length === 0) {
    // Even when there are no formula cells, we must clean up stale spill
    // regions from previous calculations (e.g. when a spill formula was deleted).
    cleanupStaleSpillRegions(workbook, new Set());
    return;
  }

  // Per-invocation state — no global pollution
  const session = new EvalSession();
  const astCache = new Map<string, AstNode>();

  // Spill tracking: maps ghost cell keys to their source formula cell key
  const spillGhosts: SpillGhostMap = new Map();
  // Tracks current spill regions for cleanup (within this invocation)
  const spillRegions = new Map<string, SpillRegion>();
  // Persistent spill map: survives across invocations for the same workbook
  const persistentSpills = getPersistentSpillMap(workbook);
  // Persistent ghost value snapshots — used to detect user modifications
  const ghostSnapshots = getGhostSnapshots(workbook);

  // Pre-populate spillGhosts from persistent spill data from previous invocations.
  // Validate each ghost cell: if it has been overwritten by the user (e.g., has
  // a formula or a different value than what we wrote), remove it from the ghost
  // map so it will be detected as a conflict during spill availability checks.
  for (const [srcKey, region] of persistentSpills) {
    const ws = workbook.getWorksheet(region.worksheetId);
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        if (r === 0 && c === 0) {
          continue;
        }
        const targetRow = region.sourceRow + r;
        const targetCol = region.sourceCol + c;
        const targetKey = spillCellKey(region.worksheetId, targetRow, targetCol);
        // Verify the ghost cell hasn't been overwritten by the user
        if (ws) {
          const ghostCell = ws.findCell(targetRow, targetCol);
          if (!isGhostCellUnmodified(ghostCell, targetKey, ghostSnapshots)) {
            // User modified this ghost cell — no longer a ghost
            continue;
          }
        }
        spillGhosts.set(targetKey, srcKey);
      }
    }
    spillRegions.set(srcKey, region);
  }

  // Clean up stale spill regions: if the source formula cell no longer exists,
  // clean up its ghost cells and remove the persistent entry.
  // formulaKeys uses spillCellKey (worksheet id based) to match persistentSpills keys.
  const formulaKeys = new Set<string>();
  for (const fc of formulaCells) {
    formulaKeys.add(spillCellKey(fc.sheet.id, fc.row, fc.col));
  }
  // Remove stale ghost entries from the per-invocation maps before cleanup
  for (const [srcKey] of [...persistentSpills]) {
    if (!formulaKeys.has(srcKey)) {
      spillRegions.delete(srcKey);
      // Also remove any ghost entries belonging to this stale source
      for (const [ghostKey, owner] of [...spillGhosts]) {
        if (owner === srcKey) {
          spillGhosts.delete(ghostKey);
        }
      }
    }
  }
  cleanupStaleSpillRegions(workbook, formulaKeys);

  // Create a context for each sheet (lazy)
  const contexts = new Map<string, EvalContext>();

  function getContext(sheet: Worksheet): EvalContext {
    let ctx = contexts.get(sheet.name);
    if (!ctx) {
      ctx = buildContext(workbook, sheet, session, astCache);
      contexts.set(sheet.name, ctx);
    }
    return ctx;
  }

  /**
   * Check if a target cell is occupied by non-spill data (i.e., would cause
   * a #SPILL! conflict). A cell is "available" if:
   * - It does not exist (empty)
   * - It is a Null-type cell (empty)
   * - It is an unmodified ghost cell belonging to the same spill source
   */
  function isCellAvailable(ws: Worksheet, row: number, col: number, sourceKey: string): boolean {
    const targetKey = spillCellKey(ws.id, row, col);

    const existing = ws.findCell(row, col);
    if (!existing) {
      return true;
    }

    // Null type = empty cell
    if (existing.type === Enums.ValueType.Null) {
      return true;
    }

    // Check if this cell is a ghost of the same source AND hasn't been modified
    const ghostSource = spillGhosts.get(targetKey);
    if (ghostSource === sourceKey) {
      return isGhostCellUnmodified(existing, targetKey, ghostSnapshots);
    }

    return false;
  }

  /**
   * Clean up a previous spill region. Clears ghost cells that were written
   * by a prior evaluation of the same formula. Only clears cells that haven't
   * been modified by the user since the last spill write.
   */
  function cleanupSpillRegion(region: SpillRegion, ws: Worksheet): void {
    const srcKey = spillCellKey(region.worksheetId, region.sourceRow, region.sourceCol);
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        if (r === 0 && c === 0) {
          continue; // Skip source cell
        }
        const targetRow = region.sourceRow + r;
        const targetCol = region.sourceCol + c;
        const targetKey = spillCellKey(region.worksheetId, targetRow, targetCol);
        if (spillGhosts.get(targetKey) === srcKey) {
          spillGhosts.delete(targetKey);
          // Only clear if the ghost cell hasn't been modified by the user
          const ghostCell = ws.findCell(targetRow, targetCol);
          if (ghostCell && isGhostCellUnmodified(ghostCell, targetKey, ghostSnapshots)) {
            ghostCell.value = null;
          }
          ghostSnapshots.delete(targetKey);
        }
      }
    }
  }

  /**
   * Write a CalcArray result to the worksheet as a spill region.
   * Returns true if successful, false if there was a #SPILL! conflict.
   */
  function writeSpillRegion(
    ws: Worksheet,
    sourceRow: number,
    sourceCol: number,
    arr: CalcArray,
    sourceCell: Cell
  ): boolean {
    const wsId = ws.id;
    const srcKey = spillCellKey(wsId, sourceRow, sourceCol);
    const numRows = arr.length;
    const numCols = arr[0]?.length ?? 0;

    if (numRows === 0 || numCols === 0) {
      return true; // Empty array — nothing to spill
    }

    // 1x1 array — just set the source cell result, no spilling needed
    if (numRows === 1 && numCols === 1) {
      // Clean up any previous spill region first
      const prevRegion = spillRegions.get(srcKey);
      if (prevRegion) {
        cleanupSpillRegion(prevRegion, ws);
      }
      sourceCell.result = calcValueToResult(arr[0][0]);
      spillRegions.delete(srcKey);
      persistentSpills.delete(srcKey);
      return true;
    }

    // Check for spill conflicts on ALL target cells BEFORE cleaning up the old region.
    // This prevents data loss: if a user wrote into a former ghost cell, we must
    // detect the conflict and return #SPILL! without destroying the user's data.
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        if (r === 0 && c === 0) {
          continue; // Source cell is always available
        }
        if (!isCellAvailable(ws, sourceRow + r, sourceCol + c, srcKey)) {
          // #SPILL! conflict — do NOT clean up old spill region, preserve user data
          sourceCell.result = { error: "#SPILL!" } as FormulaResult;
          // Don't delete spillRegions/persistentSpills: keep old ghost tracking
          // so the next recalc can still detect these ghosts
          return false;
        }
      }
    }

    // No conflicts found — now safe to clean up the old spill region
    const prevRegion = spillRegions.get(srcKey);
    if (prevRegion) {
      cleanupSpillRegion(prevRegion, ws);
    }

    // Write all cells
    sourceCell.result = calcValueToResult(arr[0][0]);

    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        if (r === 0 && c === 0) {
          continue; // Source cell already set
        }
        const targetRow = sourceRow + r;
        const targetCol = sourceCol + c;
        const targetKey = spillCellKey(wsId, targetRow, targetCol);
        const val = arr[r]?.[c] ?? null;
        const rawVal = calcValueToRawValue(val);
        const targetCell = ws.getCell(targetRow, targetCol);
        targetCell.value = rawVal;
        spillGhosts.set(targetKey, srcKey);
        // Record the written value so we can detect user modifications later
        ghostSnapshots.set(targetKey, rawVal);
      }
    }

    // Record the spill region for future cleanup (within and across invocations)
    const region: SpillRegion = {
      worksheetId: wsId,
      sourceRow,
      sourceCol,
      rows: numRows,
      cols: numCols
    };
    spillRegions.set(srcKey, region);
    persistentSpills.set(srcKey, region);

    return true;
  }

  /**
   * Handle CSE (Ctrl+Shift+Enter) array formulas.
   * These have shareType === "array" and a ref property defining the output range.
   * The formula result is distributed across the ref range.
   */
  function handleCSEArrayFormula(
    ws: Worksheet,
    fc: FormulaCell,
    result: CalcValue | CalcArray,
    evalSession: EvalSession
  ): void {
    const cell = fc.cell;
    const model = cell.model;
    const ref = model.ref;

    if (!ref) {
      return;
    }

    const range = parseRefRange(ref);
    if (!range) {
      return;
    }

    const numRows = range.bottom - range.top + 1;
    const numCols = range.right - range.left + 1;

    if (Array.isArray(result)) {
      // Distribute array elements across the ref range
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          const targetRow = range.top + r;
          const targetCol = range.left + c;
          const val = result[r]?.[c] ?? null;
          const targetCell = ws.getCell(targetRow, targetCol);
          if (targetCell.type === Enums.ValueType.Formula) {
            targetCell.result = calcValueToResult(val);
            // Mark this cell as already evaluated in the session cache
            // so it won't be re-evaluated in the main loop
            const scalarVal = val ?? null;
            evalSession.cache.set(
              evalSession.makeKey(fc.sheetName, targetRow, targetCol),
              scalarVal
            );
          }
        }
      }
    } else {
      // Scalar result: fill entire ref range with the same value
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          const targetRow = range.top + r;
          const targetCol = range.left + c;
          const targetCell = ws.getCell(targetRow, targetCol);
          if (targetCell.type === Enums.ValueType.Formula) {
            targetCell.result = calcValueToResult(result);
            evalSession.cache.set(evalSession.makeKey(fc.sheetName, targetRow, targetCol), result);
          }
        }
      }
    }
  }

  // Pre-parse all formulas into the AST cache so the dependency graph can use them.
  for (const fc of formulaCells) {
    parseFormula(fc.formula, astCache);
  }

  // Build dependency graph and compute topological evaluation order.
  // This ensures cells are evaluated after their dependencies, reducing
  // recursive evaluation and catching circular references statically.
  const nameResolver: NameResolver | undefined = workbook.definedNames
    ? (name: string) => resolveDefinedNameToRange(workbook.definedNames!, name)
    : undefined;

  const graph = buildDependencyGraph(formulaCells, astCache, nameResolver);
  const evalOrder = topologicalSort(graph);

  // Build a lookup from cell key → FormulaCell for quick access during iteration
  const formulaCellMap = new Map<string, FormulaCell>();
  for (const fc of formulaCells) {
    const key = session.makeKey(fc.sheetName, fc.row, fc.col);
    formulaCellMap.set(key, fc);
  }

  // Evaluate formula cells in topological order.
  // Cells with no dependencies are evaluated first, so by the time we reach
  // a cell, all its dependencies have already been computed and memoized.
  // The recursive getCellValue still handles any missed cases (e.g. cross-sheet
  // refs to formula cells not in the graph) as a fallback.
  for (const key of evalOrder) {
    const fc = formulaCellMap.get(key);
    if (!fc) {
      continue; // Should not happen, but guard defensively
    }

    // Skip if already evaluated via a dependency chain
    if (session.cache.has(key)) {
      // Ensure result is persisted (it was set during recursive eval)
      continue;
    }

    const ctx = getContext(fc.sheet);
    ctx.currentSheet = fc.sheetName;

    try {
      const ast = parseFormula(fc.formula, astCache);
      if (ast) {
        const model = fc.cell.model;
        const isCSE = model.shareType === "array" && model.ref && !model.isDynamicArray;
        const isDynamic = !!model.isDynamicArray || isDynamicArrayFormula(ast);

        if (isCSE || isDynamic) {
          // Use raw evaluation to preserve array results
          const result = evaluateFormulaRaw(ast, ctx, session, fc.sheetName, fc.row, fc.col);

          if (
            !shouldPreserveCachedResult(
              Array.isArray(result) ? (result[0]?.[0] ?? null) : result,
              fc.cell
            )
          ) {
            if (isCSE) {
              // CSE array formula: distribute across ref range
              handleCSEArrayFormula(fc.sheet, fc, result, session);
            } else if (Array.isArray(result) && result.length > 0) {
              // Dynamic array formula: spill to adjacent cells
              writeSpillRegion(fc.sheet, fc.row, fc.col, result, fc.cell);
            } else {
              // Scalar result or empty array
              const scalar = Array.isArray(result) ? (result[0]?.[0] ?? null) : result;
              fc.cell.result = calcValueToResult(scalar);
            }
          }
        } else {
          // Standard scalar formula evaluation
          const result = evaluateFormula(ast, ctx, session, fc.sheetName, fc.row, fc.col);
          if (!shouldPreserveCachedResult(result, fc.cell)) {
            fc.cell.result = result as FormulaResult;
          }
        }
      }
    } catch {
      // If evaluation fails, keep the existing cached result
    }
  }

  // ── Iterative Calculation for Circular References ──
  // When enabled, re-evaluate circular formula cells until convergence.
  const iterateEnabled = workbook.calcProperties?.iterate === true;
  if (iterateEnabled && graph.circularKeys.size > 0) {
    const maxIter = workbook.calcProperties?.iterateCount ?? 100;
    const delta = workbook.calcProperties?.iterateDelta ?? 0.001;

    // Collect circular formula cells in topo order
    const circularCells: FormulaCell[] = [];
    for (const key of evalOrder) {
      if (graph.circularKeys.has(key)) {
        const fc = formulaCellMap.get(key);
        if (fc) {
          circularCells.push(fc);
        }
      }
    }

    for (let iter = 0; iter < maxIter; iter++) {
      let maxChange = 0;

      // Seed circularFallback with previous iteration results so that
      // self-references resolve to the last computed value (not 0).
      for (const fc of circularCells) {
        const key = session.makeKey(fc.sheetName, fc.row, fc.col);
        const prev = fc.cell.result;
        if (prev !== undefined && prev !== null && !isError(prev)) {
          session.circularFallback.set(key, prev as CalcValue);
        } else {
          session.circularFallback.set(key, 0);
        }
        // Clear cache so the formula gets re-evaluated
        session.cache.delete(key);
      }

      for (const fc of circularCells) {
        const ctx = getContext(fc.sheet);
        ctx.currentSheet = fc.sheetName;
        try {
          const ast = parseFormula(fc.formula, astCache);
          if (ast) {
            const oldResult = fc.cell.result;
            const result = evaluateFormula(ast, ctx, session, fc.sheetName, fc.row, fc.col);
            if (!shouldPreserveCachedResult(result, fc.cell)) {
              fc.cell.result = result as FormulaResult;
            }
            // Update fallback for cells evaluated later in this iteration
            const key = session.makeKey(fc.sheetName, fc.row, fc.col);
            session.circularFallback.set(key, result);
            // Compute change
            if (typeof result === "number" && typeof oldResult === "number") {
              maxChange = Math.max(maxChange, Math.abs(result - oldResult));
            }
          }
        } catch {
          // Keep existing result
        }
      }

      if (maxChange <= delta) {
        break; // Converged
      }
    }

    // Clean up circular fallback values
    session.circularFallback.clear();
  }
}

/**
 * Determine whether we should keep the cell's existing cached result
 * instead of overwriting it with the newly computed value.
 *
 * Returns true when the engine produced a #NAME? error (unsupported function)
 * and the cell already has a usable cached result from the XLSX file.
 */
function shouldPreserveCachedResult(computed: CalcValue, cell: Cell): boolean {
  if (!isError(computed) || computed.error !== "#NAME?") {
    return false; // Engine produced a real result — use it
  }
  // #NAME? means the engine couldn't evaluate the formula.
  // If the cell already has a cached result, preserve it.
  const existing = cell.result;
  return existing !== undefined && existing !== null;
}
