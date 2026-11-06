/**
 * `@` through the containers, the engine and structural edits: written as `_xlfn.SINGLE(…)` in every place a formula
 * is stored and read back as `@`; `@x` and `_xlfn.SINGLE(x)` compute alike; a row inserted above moves the references
 * inside it.
 */
import { Archive } from "@archive";
import { calculateFormulas } from "@excel/core/formula-adapter";
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Row,
  Table,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.addMany(ws, [[1], [2], [3]]);
  return { wb, ws };
}

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode((await Archive.unpackZip(bytes)).get(path)!.data);
}

/**
 * Ordinary formulas. XLSB writes `@x` as Excel does (`single-tokens.node.test.ts`): `x` alone where `x` is intersected
 * there anyway — a whole formula, an operator's operand — so the two are one formula and read back bare, as Excel's
 * own XLSB does; the `_xlfn.SINGLE` call everywhere else, which reads back as `@`.
 */
const FORMULAS: [string, string][] = [
  ["@A1:A3", "A1:A3"],
  ["SUM(@A1:A3)*2", "SUM(@A1:A3)*2"],
  ["@S!$A$1:$A$3+1", "S!$A$1:$A$3+1"],
  ["ROWS(@A1:A3)", "ROWS(@A1:A3)"],
  ["@A1*2", "@A1*2"]
];

describe.each(["xlsx", "xlsb"] as const)("@ through %s", format => {
  it("round-trips cell formulas", async () => {
    const { wb, ws } = book();
    FORMULAS.forEach(([formula], i) => Cell.setValue(ws, { row: 2, col: 3 + i * 2 }, { formula }));
    calculateFormulas(wb);
    const before = FORMULAS.map((_, i) => Cell.getResult(ws, { row: 2, col: 3 + i * 2 }));
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    const back = Workbook.getWorksheet(read, "S")!;
    expect(FORMULAS.map((_, i) => Cell.getFormula(back, { row: 2, col: 3 + i * 2 }))).toEqual(
      FORMULAS.map(([formula, xlsb]) => (format === "xlsx" ? formula : xlsb))
    );
    calculateFormulas(read);
    expect(FORMULAS.map((_, i) => Cell.getResult(back, { row: 2, col: 3 + i * 2 }))).toEqual(
      before
    );
  });

  it("round-trips a defined name, a conditional format and a validation", async () => {
    const { wb, ws } = book();
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "One", "SUM(@S!$A$1:$A$3)");
    ConditionalFormatting.add(ws, {
      ref: "B1:B3",
      rules: [
        { type: "expression", formulae: ["SUM(@$A$1:$A$3)>1"], style: { font: { bold: true } } }
      ]
    });
    DataValidation.add(ws, "D2", { type: "custom", formulae: ["SUM(@$A$1:$A$3)>0"] });
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    expect(DefinedNames.get(Workbook.getDefinedNames(read), "One")?.formulaExpression).toBe(
      "SUM(@S!$A$1:$A$3)"
    );
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(ConditionalFormatting.list(sheet)[0]!.rules[0]).toMatchObject({
      formulae: ["SUM(@$A$1:$A$3)>1"]
    });
    expect(DataValidation.find(sheet, "D2")).toMatchObject({ formulae: ["SUM(@$A$1:$A$3)>0"] });
  });

  it("keeps a dynamic array with @ inside, and what it computes", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C2", { formula: "_xlfn.SEQUENCE(3)*@A1:A3" });
    calculateFormulas(wb);
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(Cell.getFormula(sheet, "C2")).toBe("_xlfn.SEQUENCE(3)*@A1:A3");
    calculateFormulas(read);
    expect(
      [2, 3, 4].map(row => Cell.getValue(sheet, `C${row}`) as { result?: number } | number)
    ).toEqual([expect.objectContaining({ result: 2 }), 4, 6]);
  });
});

describe("the XLSX file holds _xlfn.SINGLE", () => {
  it("in cells, shared formulas, defined names, rules and table columns", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "SUM(@A1:A3)" });
    Cell.setValue(ws, "D1", { formula: "@A1:A3*2", shareType: "shared", ref: "D1:D2" } as never);
    Cell.setValue(ws, "D2", { sharedFormula: "D1" });
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "One", "@S!$A$1:$A$3");
    ConditionalFormatting.add(ws, {
      ref: "B1",
      rules: [{ type: "expression", formulae: ["@$A$1:$A$3>1"], style: { font: { bold: true } } }]
    });
    DataValidation.add(ws, "E1", { type: "custom", formulae: ["@$A$1:$A$3>0"] });
    Table.add(ws, {
      name: "T",
      ref: "H1",
      headerRow: true,
      columns: [{ name: "a" }, { name: "b", calculatedColumnFormula: "@$A$1:$A$3+T[@a]" }],
      rows: [[1, 2]]
    } as never);
    const bytes = await Workbook.toBuffer(wb);
    const xml =
      (await part(bytes, "xl/worksheets/sheet1.xml")) + (await part(bytes, "xl/workbook.xml"));
    const table = await part(bytes, "xl/tables/table1.xml");
    expect(xml).toContain("SUM(_xlfn.SINGLE(A1:A3))</f>");
    expect(xml).toContain(">_xlfn.SINGLE(A1:A3)*2</f>");
    expect(xml).toContain(">_xlfn.SINGLE(S!$A$1:$A$3)</definedName>");
    expect(xml).toContain("<formula>_xlfn.SINGLE($A$1:$A$3)&gt;1</formula>");
    expect(xml).toContain("<formula1>_xlfn.SINGLE($A$1:$A$3)&gt;0</formula1>");
    expect(table).toContain("_xlfn.SINGLE($A$1:$A$3)+T[[#This Row],[a]]");
    expect(xml + table).not.toMatch(/[>(,=]@/);
  });

  it("reads Excel's spelling as @, and a call @ cannot stand for as written", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "_xlfn.SINGLE(A1:A3)" });
    Cell.setValue(ws, "C2", { formula: "_xlfn.SINGLE(A1:A3*2)" });
    const back = Workbook.listWorksheets(await Workbook.read(await Workbook.toBuffer(wb)))[0]!;
    expect(Cell.getFormula(back, "C1")).toBe("@A1:A3");
    expect(Cell.getFormula(back, "C2")).toBe("_xlfn.SINGLE(A1:A3*2)");
  });
});

describe("the engine computes @x and _xlfn.SINGLE(x) alike", () => {
  it.each([
    ["@A1:A3", "_xlfn.SINGLE(A1:A3)"],
    ["@A1:A3", "SINGLE(A1:A3)"],
    ["SUM(@A1:A3)", "SUM(_xlfn.SINGLE(A1:A3))"],
    ["@SEQUENCE(3)", "_xlfn.SINGLE(SEQUENCE(3))"],
    ["@A1:A3*10", "_xlfn.SINGLE(A1:A3)*10"]
  ])("%s ≡ %s", (operator, call) => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C2", { formula: operator });
    Cell.setValue(ws, "D2", { formula: call });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D2")).toEqual(Cell.getResult(ws, "C2"));
    expect(Cell.getValue(ws, "D3")).toEqual(Cell.getValue(ws, "C3"));
  });
});

describe("structural edits move the references inside @", () => {
  it("a row inserted above", () => {
    const { ws } = book();
    Cell.setValue(ws, "C5", { formula: "@A1:A3+SUM(@$A$1:$A$3)" });
    Row.insert(ws, 1, []);
    expect(Cell.getFormula(ws, "C6")).toBe("@A2:A4+SUM(@$A$2:$A$4)");
  });

  it("the referenced rows deleted", () => {
    const { ws } = book();
    Cell.setValue(ws, "C5", { formula: "@A1:A3" });
    Row.splice(ws, 1, 3);
    expect(Cell.getFormula(ws, "C2")).toBe("@#REF!");
  });
});
