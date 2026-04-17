/**
 * Binder — Transform raw AST into BoundExpr.
 *
 * The binder performs static symbol resolution:
 * - Cell/range references → BoundCellRef / BoundAreaRef with resolved sheet
 * - Defined names → BoundCellRef / BoundAreaRef (if resolvable) or BoundNameExpr
 * - Structured references → BoundCellRef / BoundAreaRef
 * - Function calls → BoundCall or BoundSpecialCall
 * - 3D references → BoundRef3D
 *
 * The binder operates on snapshot data only — no live workbook objects.
 */

import type { DefinedNameSnapshot, WorkbookSnapshot } from "../integration/workbook-snapshot";
import { resolveDefinedName as resolveDefinedNameFromSnapshot } from "../integration/workbook-snapshot";
import type { AstNode, CellRefNode, RangeRefNode } from "../syntax/ast";
import { NodeType } from "../syntax/ast";
import { stripFunctionPrefix } from "../syntax/token-types";
import { colLetterToNumber, parseDefinedNameRange } from "./address-utils";
import type { BoundExpr, SpecialFormName } from "./bound-ast";
import {
  BoundExprKind,
  boundAreaRef,
  boundCall,
  boundCellRef,
  boundErrorLiteral,
  boundLiteral,
  boundNameExpr,
  boundSpecialCall
} from "./bound-ast";
import {
  resolveStructuredRefRows,
  buildTableGeometry,
  resolveStructuredRefColumns
} from "./structured-ref-utils";

// ============================================================================
// Binding Context
// ============================================================================

/**
 * Static context for the binder. Contains all information needed to
 * resolve symbols at compile time.
 */
export interface BindingContext {
  /** The workbook snapshot. */
  readonly snapshot: WorkbookSnapshot;
  /** The current sheet name (for relative references). */
  readonly currentSheet: string;
}

// ============================================================================
// Special Form Registry
// ============================================================================

const SPECIAL_FORMS = new Set<string>([
  "IF",
  "IFERROR",
  "IFNA",
  "IFS",
  "SWITCH",
  "CHOOSE",
  "LET",
  "LAMBDA",
  "INDIRECT",
  "OFFSET",
  "MAP",
  "REDUCE",
  "SCAN",
  "MAKEARRAY",
  "BYROW",
  "BYCOL"
]);

/**
 * Map prefixed names to their canonical special form name.
 */
function canonicalSpecialForm(name: string): SpecialFormName | undefined {
  const canonical = stripFunctionPrefix(name);
  if (SPECIAL_FORMS.has(canonical)) {
    return canonical as SpecialFormName;
  }
  return undefined;
}

// ============================================================================
// Main Bind Function
// ============================================================================

/**
 * Bind a raw AST node to produce a BoundExpr.
 *
 * This is the main entry point for the compilation phase. It recursively
 * walks the AST and resolves all symbols against the snapshot.
 */
export function bind(node: AstNode, ctx: BindingContext): BoundExpr {
  switch (node.type) {
    case NodeType.Number:
      return boundLiteral(node.value);

    case NodeType.String:
      return boundLiteral(node.value);

    case NodeType.Boolean:
      return boundLiteral(node.value);

    case NodeType.Error:
      return boundErrorLiteral(node.value);

    case NodeType.Missing:
      return boundLiteral(null);

    case NodeType.CellRef:
      return bindCellRef(node, ctx);

    case NodeType.RangeRef:
      return bindRangeRef(node, ctx);

    case NodeType.ColRangeRef:
      return bindColRangeRef(node, ctx);

    case NodeType.RowRangeRef:
      return bindRowRangeRef(node, ctx);

    case NodeType.BinaryOp:
      return {
        kind: BoundExprKind.BinaryOp,
        op: node.op,
        left: bind(node.left, ctx),
        right: bind(node.right, ctx)
      };

    case NodeType.UnaryOp:
      return {
        kind: BoundExprKind.UnaryOp,
        op: node.op,
        operand: bind(node.operand, ctx)
      };

    case NodeType.Percent:
      return {
        kind: BoundExprKind.Percent,
        operand: bind(node.operand, ctx)
      };

    case NodeType.FunctionCall:
      return bindFunctionCall(node.name, node.args, ctx);

    case NodeType.Array:
      return {
        kind: BoundExprKind.Array,
        rows: node.rows.map(row => row.map(elem => bind(elem, ctx)))
      };

    case NodeType.Name:
      return bindName(node.name, ctx);

    case NodeType.StructuredRef:
      return bindStructuredRef(node.tableName, node.columns, node.specials, ctx);

    case NodeType.ExternalRef:
      // External workbook references (e.g. [Book1]Sheet1!A1) are recognised
      // but unsupported — lower to `#REF!` at compile time so the runtime
      // simply surfaces the error.
      return boundErrorLiteral("#REF!");

    default: {
      const _: never = node;
      return boundErrorLiteral("#VALUE!");
    }
  }
}

// ============================================================================
// Cell Reference Binding
// ============================================================================

function bindCellRef(node: CellRefNode, ctx: BindingContext): BoundExpr {
  const sheet = node.sheet ?? ctx.currentSheet;
  const row = parseInt(node.row, 10);
  const col = colLetterToNumber(node.col);

  // 3D cell reference: Sheet1:Sheet3!A1
  if (node.endSheet) {
    const sheets = getSheetsInRange(ctx.snapshot, sheet, node.endSheet);
    if (!sheets) {
      return boundErrorLiteral("#REF!");
    }
    const inner = boundCellRef(sheets[0], row, col);
    return {
      kind: BoundExprKind.Ref3D,
      sheets,
      inner
    };
  }

  // Validate sheet exists
  if (!ctx.snapshot.worksheetsByName.has(sheet.toLowerCase())) {
    return boundErrorLiteral("#REF!");
  }

  return boundCellRef(sheet, row, col);
}

function bindRangeRef(node: RangeRefNode, ctx: BindingContext): BoundExpr {
  const sheet = node.sheet ?? ctx.currentSheet;
  const startRow = parseInt(node.start.row, 10);
  const startCol = colLetterToNumber(node.start.col);
  const endRow = parseInt(node.end.row, 10);
  const endCol = colLetterToNumber(node.end.col);

  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);

  // 3D range reference: Sheet1:Sheet3!A1:B2
  if (node.endSheet) {
    const sheets = getSheetsInRange(ctx.snapshot, sheet, node.endSheet);
    if (!sheets) {
      return boundErrorLiteral("#REF!");
    }
    const inner = boundAreaRef(sheets[0], top, left, bottom, right);
    return {
      kind: BoundExprKind.Ref3D,
      sheets,
      inner
    };
  }

  return boundAreaRef(sheet, top, left, bottom, right);
}

function bindColRangeRef(
  node: { sheet?: string; startCol: string; endCol: string; endSheet?: string },
  ctx: BindingContext
): BoundExpr {
  const sheet = node.sheet ?? ctx.currentSheet;
  const startCol = colLetterToNumber(node.startCol);
  const endCol = colLetterToNumber(node.endCol);
  const leftCol = Math.min(startCol, endCol);
  const rightCol = Math.max(startCol, endCol);

  // Validate sheet exists
  if (!ctx.snapshot.worksheetsByName.has(sheet.toLowerCase())) {
    return boundErrorLiteral("#REF!");
  }

  // 3D col range: Sheet1:Sheet3!A:B
  if (node.endSheet) {
    const sheets = getSheetsInRange(ctx.snapshot, sheet, node.endSheet);
    if (!sheets) {
      return boundErrorLiteral("#REF!");
    }
    // Return Ref3D wrapping an area for the full column range
    const inner = boundAreaRef(sheets[0], 1, leftCol, 1_048_576, rightCol);
    return {
      kind: BoundExprKind.Ref3D,
      sheets,
      inner
    };
  }

  return {
    kind: BoundExprKind.ColRangeRef,
    sheet,
    leftCol,
    rightCol
  };
}

function bindRowRangeRef(
  node: { sheet?: string; startRow: number; endRow: number; endSheet?: string },
  ctx: BindingContext
): BoundExpr {
  const sheet = node.sheet ?? ctx.currentSheet;
  const topRow = Math.min(node.startRow, node.endRow);
  const bottomRow = Math.max(node.startRow, node.endRow);

  // Validate sheet exists
  if (!ctx.snapshot.worksheetsByName.has(sheet.toLowerCase())) {
    return boundErrorLiteral("#REF!");
  }

  // 3D row range: Sheet1:Sheet3!1:5
  if (node.endSheet) {
    const sheets = getSheetsInRange(ctx.snapshot, sheet, node.endSheet);
    if (!sheets) {
      return boundErrorLiteral("#REF!");
    }
    const inner = boundAreaRef(sheets[0], topRow, 1, bottomRow, 16_384);
    return {
      kind: BoundExprKind.Ref3D,
      sheets,
      inner
    };
  }

  return {
    kind: BoundExprKind.RowRangeRef,
    sheet,
    topRow,
    bottomRow
  };
}

// ============================================================================
// Name Binding
// ============================================================================

function bindName(name: string, ctx: BindingContext): BoundExpr {
  // Try to resolve as a defined name that maps to a cell/range.
  // Respects scope precedence: sheet-local > workbook-global.
  const dn = resolveDefinedNameFromSnapshot(ctx.snapshot.definedNames, name, ctx.currentSheet);
  if (dn) {
    return resolveDefinedName(dn, ctx);
  }

  // Unresolved — could be a LET-bound local variable, a formula-based name,
  // or a truly unknown name. Return BoundNameExpr for runtime resolution.
  return boundNameExpr(name);
}

function resolveDefinedName(dn: DefinedNameSnapshot, ctx: BindingContext): BoundExpr {
  if (dn.ranges.length === 0) {
    return boundErrorLiteral("#NAME?");
  }

  // Multi-area names are not supported
  if (dn.ranges.length > 1) {
    return boundErrorLiteral("#VALUE!");
  }

  const rangeStr = dn.ranges[0];

  // Try to parse as a cell/range reference
  const parsed = parseDefinedNameRange(rangeStr);
  if (parsed) {
    // Single cell
    if (parsed.startRow === parsed.endRow && parsed.startCol === parsed.endCol) {
      return boundCellRef(parsed.sheet, parsed.startRow, parsed.startCol);
    }
    // Range
    return boundAreaRef(
      parsed.sheet,
      Math.min(parsed.startRow, parsed.endRow),
      Math.min(parsed.startCol, parsed.endCol),
      Math.max(parsed.startRow, parsed.endRow),
      Math.max(parsed.startCol, parsed.endCol)
    );
  }

  // The range string is a formula expression (e.g. "LAMBDA(x,y,x+y)").
  // This must be resolved at runtime since it may produce a lambda or
  // depend on the evaluation context. Return a BoundNameExpr.
  return boundNameExpr(dn.name);
}

// ============================================================================
// Structured Reference Binding
// ============================================================================

function bindStructuredRef(
  tableName: string,
  columns: string[],
  specials: string[],
  ctx: BindingContext
): BoundExpr {
  // If the table name is empty (implicit @) or specials include #This Row,
  // we need runtime context (current cell position) to resolve.
  // Defer to the evaluator.
  const needsRuntime = tableName === "" || specials.includes("#This Row");

  // Find the table in the snapshot
  const table = findTable(ctx.snapshot, tableName);
  if (!table) {
    if (needsRuntime) {
      // Defer to runtime — return BoundStructuredRef
      return {
        kind: BoundExprKind.StructuredRef,
        tableName,
        columns: [...columns],
        specials: [...specials]
      };
    }
    return boundErrorLiteral("#REF!");
  }

  // If #This Row is present, defer to runtime even if table is found
  if (needsRuntime) {
    return {
      kind: BoundExprKind.StructuredRef,
      tableName: table.table.name,
      columns: [...columns],
      specials: [...specials]
    };
  }

  const resolved = resolveStructuredRefBounds(table, columns, specials);
  if (!resolved) {
    return boundErrorLiteral("#REF!");
  }

  if (resolved.top === resolved.bottom && resolved.left === resolved.right) {
    return boundCellRef(resolved.sheet, resolved.top, resolved.left);
  }
  return boundAreaRef(resolved.sheet, resolved.top, resolved.left, resolved.bottom, resolved.right);
}

interface TableWithSheet {
  readonly table: {
    readonly name: string;
    readonly columns: readonly { readonly name: string }[];
    readonly topLeft: { readonly row: number; readonly col: number };
    readonly dataRowCount: number;
    readonly hasHeaderRow: boolean;
    readonly hasTotalsRow: boolean;
  };
  readonly sheetName: string;
}

function findTable(snapshot: WorkbookSnapshot, tableName: string): TableWithSheet | null {
  if (!tableName) {
    return null;
  }
  // Use the pre-built tablesByName index for O(1) lookup
  const resolved = snapshot.tablesByName.get(tableName.toLowerCase());
  if (resolved) {
    return { table: resolved.table, sheetName: resolved.sheetName };
  }
  return null;
}

function resolveStructuredRefBounds(
  tw: TableWithSheet,
  columns: string[],
  specials: string[]
): { sheet: string; top: number; left: number; bottom: number; right: number } | null {
  const t = tw.table;
  const geo = buildTableGeometry(t);

  const colRange = resolveStructuredRefColumns(columns, t, "strict");
  if (colRange === "error") {
    return null;
  }

  const rowRange = resolveStructuredRefRows(specials, geo);
  let rowTop: number;
  let rowBottom: number;
  if (rowRange === "error") {
    return null;
  } else if (rowRange === "thisRow") {
    // #This Row requires runtime context — use data range as static fallback
    rowTop = geo.dataRowStart;
    rowBottom = geo.dataRowEnd;
  } else {
    rowTop = rowRange.rowTop;
    rowBottom = rowRange.rowBottom;
  }

  return {
    sheet: tw.sheetName,
    top: rowTop,
    left: colRange.colLeft,
    bottom: rowBottom,
    right: colRange.colRight
  };
}

// ============================================================================
// Function Call Binding
// ============================================================================

function bindFunctionCall(name: string, args: AstNode[], ctx: BindingContext): BoundExpr {
  const upperName = name.toUpperCase();

  // Check for special forms (lazy evaluation)
  const specialName = canonicalSpecialForm(upperName);
  if (specialName) {
    // LAMBDA is special: bind it into a BoundLambda if possible
    if (specialName === "LAMBDA" && args.length >= 1) {
      return bindLambda(args, ctx);
    }
    // All other special forms: bind args recursively but wrap as BoundSpecialCall
    const boundArgs = args.map(arg => bind(arg, ctx));
    return boundSpecialCall(specialName, boundArgs);
  }

  // Reference functions that need AST-level access (ROW, COLUMN, ROWS, COLUMNS)
  // are bound as regular calls — the evaluator handles the AST → BoundExpr
  // mapping for these since they work the same way with BoundExpr.

  // Standard eager function call
  const boundArgs = args.map(arg => bind(arg, ctx));
  return boundCall(upperName, boundArgs);
}

function bindLambda(args: AstNode[], ctx: BindingContext): BoundExpr {
  if (args.length < 1) {
    return boundErrorLiteral("#VALUE!");
  }

  const paramNodes = args.slice(0, -1);
  const bodyNode = args[args.length - 1];
  const params: string[] = [];

  for (const pNode of paramNodes) {
    if (pNode.type !== NodeType.Name) {
      return boundErrorLiteral("#VALUE!");
    }
    params.push(pNode.name.toUpperCase());
  }

  return {
    kind: BoundExprKind.Lambda,
    params,
    body: bind(bodyNode, ctx)
  };
}

// ============================================================================
// 3D Reference Helpers
// ============================================================================

function getSheetsInRange(
  snapshot: WorkbookSnapshot,
  startSheet: string,
  endSheet: string
): string[] | null {
  const allSheets = snapshot.worksheets;
  const startIdx = allSheets.findIndex(s => s.name.toLowerCase() === startSheet.toLowerCase());
  const endIdx = allSheets.findIndex(s => s.name.toLowerCase() === endSheet.toLowerCase());
  if (startIdx === -1 || endIdx === -1) {
    return null;
  }
  const lo = Math.min(startIdx, endIdx);
  const hi = Math.max(startIdx, endIdx);
  return allSheets.slice(lo, hi + 1).map(s => s.name);
}
