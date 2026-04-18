/**
 * Unit tests for the formula binder.
 *
 * Covers the transform from raw AST (parser output) to BoundExpr:
 * - Literal coverage (numbers, strings, booleans, errors, #REF! fallback for
 *   external refs).
 * - Static reference resolution (cells, areas, column / row ranges, 3D refs).
 * - Sheet-existence checks lowering to `#REF!`.
 * - Defined-name resolution with scoping, and fallback to `BoundNameExpr`.
 * - Structured-reference resolution against table geometry.
 * - Function calls: eager (`BoundCall`) vs special forms (`BoundSpecialCall`)
 *   and LAMBDA bound to a `BoundLambda`.
 * - The `_XLFN.` / `_XLFN._XLWS.` prefix being stripped.
 *
 * The binder operates on snapshot data, so each test builds a minimal
 * `WorkbookSnapshot` by hand (no live workbook objects are involved).
 */

import { describe, it, expect } from "vitest";

import type {
  DefinedNameSnapshot,
  ResolvedTable,
  TableSnapshot,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "../../integration/workbook-snapshot";
import { scopedNameKey } from "../../integration/workbook-snapshot";
import { NodeType, type AstNode } from "../../syntax/ast";
import { parse } from "../../syntax/parser";
import { tokenize } from "../../syntax/tokenizer";
import { bind, type BindingContext } from "../binder";
import {
  BoundExprKind,
  type BoundAreaRef,
  type BoundCall,
  type BoundCellRef,
  type BoundColRangeRef,
  type BoundExpr,
  type BoundLambda,
  type BoundLiteral,
  type BoundNameExpr,
  type BoundRef3D,
  type BoundRowRangeRef,
  type BoundSpecialCall,
  type BoundStructuredRef
} from "../bound-ast";

// ---------------------------------------------------------------------------
// Test helpers — minimal WorkbookSnapshot builders.
// ---------------------------------------------------------------------------

interface SheetSpec {
  readonly name: string;
  readonly tables?: readonly TableSnapshot[];
}

function makeWorksheet(
  name: string,
  id: number,
  tables: readonly TableSnapshot[] = []
): WorksheetSnapshot {
  return {
    id,
    name,
    dimensions: null,
    cells: new Map(),
    hiddenRows: new Set(),
    tables
  };
}

function makeSnapshot(
  sheets: readonly SheetSpec[],
  definedNames: ReadonlyMap<string, DefinedNameSnapshot> = new Map()
): WorkbookSnapshot {
  const worksheets: WorksheetSnapshot[] = [];
  const worksheetsByName = new Map<string, WorksheetSnapshot>();
  const worksheetsById = new Map<number, WorksheetSnapshot>();
  const tablesByName = new Map<string, ResolvedTable>();

  sheets.forEach((spec, idx) => {
    const ws = makeWorksheet(spec.name, idx + 1, spec.tables ?? []);
    worksheets.push(ws);
    worksheetsByName.set(spec.name.toLowerCase(), ws);
    worksheetsById.set(ws.id, ws);
    for (const table of ws.tables) {
      tablesByName.set(table.name.toLowerCase(), { table, sheetName: ws.name });
    }
  });

  return {
    worksheets,
    worksheetsByName,
    worksheetsById,
    definedNames,
    tablesByName,
    calcProperties: {},
    properties: {}
  };
}

function compile(formula: string): AstNode {
  return parse(tokenize(formula));
}

function bindFormula(formula: string, snapshot: WorkbookSnapshot, currentSheet: string): BoundExpr {
  const ctx: BindingContext = { snapshot, currentSheet };
  return bind(compile(formula), ctx);
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("binder — literals", () => {
  const snapshot = makeSnapshot([{ name: "Sheet1" }]);

  it("binds a number AST node to a BoundLiteral", () => {
    const out = bindFormula("42", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.value).toBe(42);
    expect(out.errorCode).toBeUndefined();
  });

  it("binds a string AST node to a BoundLiteral", () => {
    const out = bindFormula('"hi"', snapshot, "Sheet1") as BoundLiteral;
    expect(out.value).toBe("hi");
  });

  it("binds a boolean AST node to a BoundLiteral", () => {
    const out = bindFormula("TRUE", snapshot, "Sheet1") as BoundLiteral;
    expect(out.value).toBe(true);
  });

  it("binds an error AST node to a BoundLiteral carrying the errorCode", () => {
    const out = bindFormula("#N/A", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#N/A");
  });

  it("lowers an external workbook reference to #REF! via the tokenizer's Error token", () => {
    // External workbook refs are rejected at the tokenizer level (Error token
    // with value "#REF!"), which the binder then surfaces as an error literal.
    const out = bindFormula("[Book1]Sheet1!A1", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#REF!");
  });
});

describe("binder — cell and range references", () => {
  const snapshot = makeSnapshot([{ name: "Sheet1" }, { name: "Sheet2" }, { name: "Sheet3" }]);

  it("binds A1 on the current sheet to BoundCellRef", () => {
    const out = bindFormula("A1", snapshot, "Sheet1") as BoundCellRef;
    expect(out.kind).toBe(BoundExprKind.CellRef);
    expect(out.sheet).toBe("Sheet1");
    expect(out.row).toBe(1);
    expect(out.col).toBe(1);
  });

  it("binds Sheet2!B3 with an explicit sheet qualifier", () => {
    const out = bindFormula("Sheet2!B3", snapshot, "Sheet1") as BoundCellRef;
    expect(out.sheet).toBe("Sheet2");
    expect(out.row).toBe(3);
    expect(out.col).toBe(2);
  });

  it("binds a range A1:B2 to BoundAreaRef with ordered bounds", () => {
    const out = bindFormula("A1:B2", snapshot, "Sheet1") as BoundAreaRef;
    expect(out.kind).toBe(BoundExprKind.AreaRef);
    expect(out.top).toBe(1);
    expect(out.left).toBe(1);
    expect(out.bottom).toBe(2);
    expect(out.right).toBe(2);
  });

  it("binds a reversed range B2:A1 and normalises to top-left / bottom-right", () => {
    const out = bindFormula("B2:A1", snapshot, "Sheet1") as BoundAreaRef;
    expect(out.top).toBe(1);
    expect(out.left).toBe(1);
    expect(out.bottom).toBe(2);
    expect(out.right).toBe(2);
  });

  it("binds a whole-column range A:C to BoundColRangeRef", () => {
    const out = bindFormula("A:C", snapshot, "Sheet1") as BoundColRangeRef;
    expect(out.kind).toBe(BoundExprKind.ColRangeRef);
    expect(out.leftCol).toBe(1);
    expect(out.rightCol).toBe(3);
  });

  it("binds a whole-row range 2:5 to BoundRowRangeRef", () => {
    const out = bindFormula("2:5", snapshot, "Sheet1") as BoundRowRangeRef;
    expect(out.kind).toBe(BoundExprKind.RowRangeRef);
    expect(out.topRow).toBe(2);
    expect(out.bottomRow).toBe(5);
  });

  it("binds a 3D cell ref to BoundRef3D with sheets in workbook order", () => {
    const out = bindFormula("Sheet1:Sheet3!A1", snapshot, "Sheet1") as BoundRef3D;
    expect(out.kind).toBe(BoundExprKind.Ref3D);
    expect(out.sheets).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    expect(out.inner.kind).toBe(BoundExprKind.CellRef);
  });

  it("binds a 3D range ref Sheet1:Sheet3!A1:B2", () => {
    const out = bindFormula("Sheet1:Sheet3!A1:B2", snapshot, "Sheet1") as BoundRef3D;
    expect(out.sheets).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    expect(out.inner.kind).toBe(BoundExprKind.AreaRef);
  });
});

describe("binder — sheet existence checks", () => {
  const snapshot = makeSnapshot([{ name: "Sheet1" }]);

  it("lowers a cell ref on a non-existent sheet to #REF!", () => {
    const out = bindFormula("NoSuch!A1", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#REF!");
  });

  it("lowers a whole-column range on a non-existent sheet to #REF!", () => {
    const out = bindFormula("NoSuch!A:B", snapshot, "Sheet1") as BoundLiteral;
    expect(out.errorCode).toBe("#REF!");
  });

  it("lowers a whole-row range on a non-existent sheet to #REF!", () => {
    const out = bindFormula("NoSuch!1:5", snapshot, "Sheet1") as BoundLiteral;
    expect(out.errorCode).toBe("#REF!");
  });

  it("lowers a 3D ref whose end sheet does not exist to #REF!", () => {
    const out = bindFormula("Sheet1:Nope!A1", snapshot, "Sheet1") as BoundLiteral;
    expect(out.errorCode).toBe("#REF!");
  });
});

describe("binder — defined names", () => {
  it("resolves a workbook-scoped name pointing at a single cell", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("TOTAL", { name: "Total", ranges: ["Sheet1!$A$1"] });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Total", snapshot, "Sheet1") as BoundCellRef;
    expect(out.kind).toBe(BoundExprKind.CellRef);
    expect(out.sheet).toBe("Sheet1");
    expect(out.row).toBe(1);
    expect(out.col).toBe(1);
  });

  it("resolves a defined name pointing at a range", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("DATA", { name: "Data", ranges: ["Sheet1!$A$1:$B$3"] });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Data", snapshot, "Sheet1") as BoundAreaRef;
    expect(out.kind).toBe(BoundExprKind.AreaRef);
    expect(out.top).toBe(1);
    expect(out.bottom).toBe(3);
    expect(out.right).toBe(2);
  });

  it("prefers a sheet-scoped name over a global with the same name", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("TOTAL", { name: "Total", ranges: ["Sheet1!$A$1"] });
    defined.set(scopedNameKey("Sheet1", "Total"), {
      name: "Total",
      ranges: ["Sheet1!$Z$9"],
      scope: "Sheet1"
    });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Total", snapshot, "Sheet1") as BoundCellRef;
    expect(out.row).toBe(9);
    expect(out.col).toBe(26); // Z
  });

  it("returns BoundNameExpr for an unknown identifier (deferred to runtime)", () => {
    const snapshot = makeSnapshot([{ name: "Sheet1" }]);
    const out = bindFormula("Unknown", snapshot, "Sheet1") as BoundNameExpr;
    expect(out.kind).toBe(BoundExprKind.NameExpr);
    expect(out.name).toBe("Unknown");
    expect(out.upperName).toBe("UNKNOWN");
  });

  it("returns BoundNameExpr for a formula-valued defined name (e.g. LAMBDA body)", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("ADD", { name: "Add", ranges: ["LAMBDA(x,y,x+y)"] });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Add", snapshot, "Sheet1") as BoundNameExpr;
    expect(out.kind).toBe(BoundExprKind.NameExpr);
    expect(out.name).toBe("Add");
  });

  it("returns #NAME? for a defined name with no ranges", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("EMPTY", { name: "Empty", ranges: [] });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Empty", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#NAME?");
  });

  it("returns #VALUE! for a multi-area defined name", () => {
    const defined = new Map<string, DefinedNameSnapshot>();
    defined.set("MULTI", { name: "Multi", ranges: ["Sheet1!$A$1", "Sheet1!$B$2"] });
    const snapshot = makeSnapshot([{ name: "Sheet1" }], defined);

    const out = bindFormula("Multi", snapshot, "Sheet1") as BoundLiteral;
    expect(out.errorCode).toBe("#VALUE!");
  });
});

describe("binder — structured references", () => {
  function tableSnapshot(): TableSnapshot {
    // Table "T1" on Sheet1 with:
    //   - columns: ["A", "B"]
    //   - top-left at row 2, col 3 (=C2)
    //   - 4 data rows, with header
    // Geometry: header row 2, data rows 3..6, totals row 7 if hasTotalsRow
    return {
      name: "T1",
      columns: [{ name: "A" }, { name: "B" }],
      topLeft: { row: 2, col: 3 },
      dataRowCount: 4,
      hasHeaderRow: true,
      hasTotalsRow: false
    };
  }

  it("resolves Table1[Col] to a BoundAreaRef over the data body of that column", () => {
    const snapshot = makeSnapshot([{ name: "Sheet1", tables: [tableSnapshot()] }]);
    const out = bindFormula("T1[A]", snapshot, "Sheet1") as BoundAreaRef;
    expect(out.kind).toBe(BoundExprKind.AreaRef);
    expect(out.sheet).toBe("Sheet1");
    // Column A is the first table column → col 3. Data rows: 3..6.
    expect(out.left).toBe(3);
    expect(out.right).toBe(3);
    expect(out.top).toBe(3);
    expect(out.bottom).toBe(6);
  });

  it("returns #REF! for an unknown table when no runtime resolution is needed", () => {
    const snapshot = makeSnapshot([{ name: "Sheet1" }]);
    const out = bindFormula("NoTable[Col]", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#REF!");
  });

  it("defers [@Col] (implicit #This Row) to runtime as BoundStructuredRef", () => {
    const snapshot = makeSnapshot([{ name: "Sheet1", tables: [tableSnapshot()] }]);
    const out = bindFormula("[@A]", snapshot, "Sheet1") as BoundStructuredRef;
    expect(out.kind).toBe(BoundExprKind.StructuredRef);
    expect(out.tableName).toBe("");
    expect(out.specials).toContain("#This Row");
    expect(out.columns).toEqual(["A"]);
  });
});

describe("binder — function calls", () => {
  const snapshot = makeSnapshot([{ name: "Sheet1" }]);

  it("binds an ordinary function as BoundCall with uppercased name", () => {
    const out = bindFormula("sum(A1,B1)", snapshot, "Sheet1") as BoundCall;
    expect(out.kind).toBe(BoundExprKind.Call);
    expect(out.name).toBe("SUM");
    expect(out.args).toHaveLength(2);
  });

  it("strips the _XLFN. prefix when the function is a special form", () => {
    // FILTER isn't a special form, but CHOOSE is. Verify the prefix stripping
    // path by using a prefixed special form.
    const out = bindFormula("_XLFN.CHOOSE(1,A1,B1)", snapshot, "Sheet1") as BoundSpecialCall;
    expect(out.kind).toBe(BoundExprKind.SpecialCall);
    expect(out.name).toBe("CHOOSE");
  });

  it("preserves the _XLFN. prefix for a regular function call (not a special form)", () => {
    // Regular eager calls keep the upper-cased name (binder doesn't strip).
    const out = bindFormula("_XLFN.FILTER(A1,B1)", snapshot, "Sheet1") as BoundCall;
    expect(out.kind).toBe(BoundExprKind.Call);
    expect(out.name).toBe("_XLFN.FILTER");
  });

  it("binds IF(...) as a BoundSpecialCall (lazy arguments)", () => {
    const out = bindFormula("IF(A1,1,0)", snapshot, "Sheet1") as BoundSpecialCall;
    expect(out.kind).toBe(BoundExprKind.SpecialCall);
    expect(out.name).toBe("IF");
    expect(out.args).toHaveLength(3);
  });

  it("binds LAMBDA(x,y,x+y) as a BoundLambda with parameter list and body", () => {
    const out = bindFormula("LAMBDA(x,y,x+y)", snapshot, "Sheet1") as BoundLambda;
    expect(out.kind).toBe(BoundExprKind.Lambda);
    expect(out.params).toEqual(["X", "Y"]);
    expect(out.body.kind).toBe(BoundExprKind.BinaryOp);
  });

  it("rejects a LAMBDA whose parameter slot isn't a Name and returns #VALUE!", () => {
    // LAMBDA(1, x+y) — first "param" is a number literal, not a name.
    const out = bindFormula("LAMBDA(1, 1+1)", snapshot, "Sheet1") as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.errorCode).toBe("#VALUE!");
  });

  it("recurses into arguments of a BinaryOp expression", () => {
    const out = bindFormula("A1+B1", snapshot, "Sheet1");
    expect(out.kind).toBe(BoundExprKind.BinaryOp);
  });

  it("binds an array constant to a BoundArray with bound children", () => {
    const out = bindFormula("{1,2;3,4}", snapshot, "Sheet1");
    expect(out.kind).toBe(BoundExprKind.Array);
  });

  it("treats a Missing AST arg (empty argument slot) as a null literal", () => {
    // IF(A1,,0) — middle slot is missing.
    const out = bindFormula("IF(A1,,0)", snapshot, "Sheet1") as BoundSpecialCall;
    expect(out.args).toHaveLength(3);
    const mid = out.args[1] as BoundLiteral;
    expect(mid.kind).toBe(BoundExprKind.Literal);
    expect(mid.value).toBeNull();
    expect(mid.errorCode).toBeUndefined();
  });
});

describe("binder — raw AST type routing", () => {
  // Exercises bind() at the AST level without going through the parser, to
  // guarantee all top-level switch branches are reachable even for paths
  // that the parser never produces directly.
  const snapshot = makeSnapshot([{ name: "Sheet1" }]);
  const ctx: BindingContext = { snapshot, currentSheet: "Sheet1" };

  it("binds NodeType.Missing to a null BoundLiteral", () => {
    const ast: AstNode = { type: NodeType.Missing };
    const out = bind(ast, ctx) as BoundLiteral;
    expect(out.kind).toBe(BoundExprKind.Literal);
    expect(out.value).toBeNull();
  });

  it("binds NodeType.Percent as a BoundPercent wrapping its operand", () => {
    const ast: AstNode = {
      type: NodeType.Percent,
      operand: { type: NodeType.Number, value: 50 }
    };
    const out = bind(ast, ctx);
    expect(out.kind).toBe(BoundExprKind.Percent);
  });
});
