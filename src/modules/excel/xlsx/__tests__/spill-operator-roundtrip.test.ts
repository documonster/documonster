/**
 * The spill-range operator through both containers. A file stores `A1#` as `_xlfn.ANCHORARRAY(A1)` (MS-XLSX's formula
 * grammar has no `#` operator; Excel's XLSB save of `future-functions.xlsx` calls `_xlfn.ANCHORARRAY` through its stub name),
 * and it reads back as the operator, so a formula set with it comes back exactly as set.
 */
import { Archive } from "@archive";
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Row,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

async function sheetXml(bytes: Uint8Array, part = "xl/worksheets/sheet1.xml"): Promise<string> {
  const files = await Archive.unpackZip(bytes);
  return new TextDecoder().decode(files.get(part)!.data);
}

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const other = Workbook.addWorksheet(wb, "My Sheet");
  Row.addMany(ws, [[1], [2], [3]]);
  Row.add(other, [4]);
  return { wb, ws, other };
}

const FORMULAS = [
  "A1#",
  "SUM(A1#)",
  "SUM(S!$A$1#)*2",
  "ROWS('My Sheet'!A1#)+1",
  "IFERROR(INDEX(A1#,2),#N/A)"
];

describe.each(["xlsx", "xlsb"] as const)("the spill-range operator through %s", format => {
  it("round-trips a cell formula exactly", async () => {
    const { wb, ws } = book();
    FORMULAS.forEach((formula, i) => Cell.setValue(ws, i + 1, 3, { formula: `=${formula}` }));
    const back = Workbook.getWorksheet(
      await Workbook.read(await Workbook.toBuffer(wb, { format })),
      "S"
    )!;
    expect(FORMULAS.map((_, i) => Cell.getFormula(back, i + 1, 3))).toEqual(FORMULAS);
  });

  it("round-trips a defined name, a conditional format and a validation", async () => {
    const { wb, ws } = book();
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Spilled", "S!$A$1#");
    ConditionalFormatting.add(ws, {
      ref: "B1:B3",
      rules: [{ type: "expression", formulae: ["COUNT($A$1#)>2"], style: { font: { bold: true } } }]
    });
    DataValidation.add(ws, "D1", { type: "list", formulae: ["$A$1#"] });
    const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
    expect(DefinedNames.get(Workbook.getDefinedNames(read), "Spilled")?.formulaExpression).toBe(
      "S!$A$1#"
    );
    const sheet = Workbook.getWorksheet(read, "S")!;
    expect(ConditionalFormatting.list(sheet)[0]!.rules[0]).toMatchObject({
      formulae: ["COUNT($A$1#)>2"]
    });
    expect(DataValidation.find(sheet, "D1")).toMatchObject({ formulae: ["$A$1#"] });
  });
});

describe("the XLSX file holds the future-function form", () => {
  it("in cells, shared formulas, defined names and rules", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "SUM(A1#)" });
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Spilled", "S!$A$1#");
    ConditionalFormatting.add(ws, {
      ref: "B1",
      rules: [{ type: "expression", formulae: ["A1#>1"], style: {} }]
    });
    DataValidation.add(ws, "D1", { type: "list", formulae: ["$A$1#"] });
    const bytes = await Workbook.toBuffer(wb);
    const sheet = await sheetXml(bytes);
    const workbook = await sheetXml(bytes, "xl/workbook.xml");
    expect(sheet).toContain("SUM(_xlfn.ANCHORARRAY(A1))</f>");
    expect(sheet).toContain("<formula>_xlfn.ANCHORARRAY(A1)&gt;1</formula>");
    expect(sheet).toContain("<formula1>_xlfn.ANCHORARRAY($A$1)</formula1>");
    expect(workbook).toContain(">_xlfn.ANCHORARRAY(S!$A$1)</definedName>");
    expect(sheet + workbook).not.toMatch(/[A-Z0-9]#/);
  });

  it("reads Excel's spelling back as the operator, and a call it cannot be as written", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "_xlfn.ANCHORARRAY(A1)" });
    Cell.setValue(ws, "C2", { formula: "_xlfn.ANCHORARRAY(INDEX(A1:A3,1))" });
    const back = Workbook.listWorksheets(await Workbook.read(await Workbook.toBuffer(wb)))[0]!;
    expect(Cell.getFormula(back, "C1")).toBe("A1#");
    expect(Cell.getFormula(back, "C2")).toBe("_xlfn.ANCHORARRAY(INDEX(A1:A3,1))");
  });

  it("keeps the operator in a shared formula's dependents", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "SUM(A1#)", shareType: "shared", ref: "C1:C2" } as never);
    Cell.setValue(ws, "C2", { sharedFormula: "C1" });
    const back = Workbook.listWorksheets(await Workbook.read(await Workbook.toBuffer(wb)))[0]!;
    expect(Cell.getFormula(back, "C1")).toBe("SUM(A1#)");
    expect(Cell.getFormula(back, "C2")).toBe("SUM(A2#)");
  });

  it("is what Worksheet.getModel's cells keep, unchanged by a read and a write", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "C1", { formula: "A1#" });
    const first = await sheetXml(await Workbook.toBuffer(wb));
    const second = await sheetXml(
      await Workbook.toBuffer(await Workbook.read(await Workbook.toBuffer(wb)))
    );
    const formula = (xml: string) => /<c r="C1"[^>]*>(?:<f[^>]*>([^<]*)<\/f>)/.exec(xml)?.[1];
    expect(formula(first)).toBe("_xlfn.ANCHORARRAY(A1)");
    expect(formula(second)).toBe(formula(first));
    expect(Worksheet.getName(ws)).toBe("S");
  });
});
