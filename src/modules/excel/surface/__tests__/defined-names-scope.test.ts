/**
 * `DefinedNames.add` and `addFormula` take the scope `get` and `remove` take, and return what they attached — the
 * copy `get` returns. A sheet-scoped name could only be made through `setModel` over the whole registry.
 */
import { Cell, DefinedNames, ExcelArgumentError, Workbook } from "@excel";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  for (const name of ["One", "Two"]) {
    const ws = Workbook.addWorksheet(wb, name);
    Cell.setValue(ws, "A1", name === "One" ? 10 : 20);
  }
  return wb;
}

describe("DefinedNames.add / addFormula", () => {
  it("return what they attached", () => {
    const dn = Workbook.getDefinedNames(book());
    expect(DefinedNames.add(dn, "One!$A$1", "First")).toEqual({
      name: "First",
      ranges: ["One!$A$1"]
    });
    expect(DefinedNames.addFormula(dn, "Rate", "0.2")).toEqual({
      name: "Rate",
      ranges: ["0.2"],
      formulaExpression: "0.2"
    });
    expect(DefinedNames.add(dn, "One!$A$1", "First")).toEqual(DefinedNames.get(dn, "First"));
  });

  it("make a sheet-scoped name beside a workbook one of the same name", () => {
    const wb = book();
    const dn = Workbook.getDefinedNames(wb);
    DefinedNames.add(dn, "One!$A$1", "Total");
    expect(DefinedNames.add(dn, "Two!$A$1", "Total", { localSheetId: 1 })).toEqual({
      name: "Total",
      ranges: ["Two!$A$1"],
      localSheetId: 1
    });
    expect(DefinedNames.addFormula(dn, "Rate", "Two!$A$1*2", { localSheetId: 1 })).toMatchObject({
      localSheetId: 1
    });
    expect(DefinedNames.get(dn, "Total")?.ranges).toEqual(["One!$A$1"]);
    expect(DefinedNames.get(dn, "Total", { localSheetId: 1 })?.ranges).toEqual(["Two!$A$1"]);
    expect(DefinedNames.get(dn, "Rate")).toBeUndefined();
    // The sheet's own name wins on that sheet.
    Cell.setValue(Workbook.getWorksheet(wb, "Two")!, "B1", { formula: "Total+Rate" });
    Cell.setValue(Workbook.getWorksheet(wb, "One")!, "B1", { formula: "Total" });
    calculateFormulas(wb);
    expect(Cell.getResult(Workbook.getWorksheet(wb, "Two")!, "B1")).toBe(60);
    expect(Cell.getResult(Workbook.getWorksheet(wb, "One")!, "B1")).toBe(10);
  });

  it.each(["xlsx", "xlsb"] as const)("a sheet-scoped name survives %s", async format => {
    const wb = book();
    const dn = Workbook.getDefinedNames(wb);
    DefinedNames.add(dn, "Two!$A$1", "Total", { localSheetId: 1 });
    DefinedNames.addFormula(dn, "Rate", "0.5", { localSheetId: 0 });
    const read = Workbook.getDefinedNames(
      await Workbook.read(await Workbook.toBuffer(wb, { format }))
    );
    expect(DefinedNames.get(read, "Total", { localSheetId: 1 })?.ranges).toEqual(["Two!$A$1"]);
    expect(DefinedNames.get(read, "Rate", { localSheetId: 0 })?.formulaExpression).toBe("0.5");
    expect(DefinedNames.get(read, "Total")).toBeUndefined();
  });

  it("check the scope as get and remove do, before changing anything", () => {
    const dn = Workbook.getDefinedNames(book());
    for (const localSheetId of [-1, 1.5, "1"]) {
      expect(() => DefinedNames.add(dn, "One!$A$1", "X", { localSheetId } as never)).toThrow(
        expect.objectContaining({ field: "options.localSheetId" })
      );
      expect(() => DefinedNames.addFormula(dn, "Y", "1", { localSheetId } as never)).toThrow(
        ExcelArgumentError
      );
    }
    expect(DefinedNames.list(dn)).toEqual([]);
  });
});

describe("adding many names", () => {
  // `add` reports the name it added by looking it up, and the lookup built every name's entry to find one: adding
  // names was quadratic, 39 seconds for 4,000.
  it("takes time linear in their number", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    const names = Workbook.getDefinedNames(wb);
    const started = performance.now();
    for (let i = 0; i < 20_000; i++) {
      DefinedNames.add(names, `S!$A$${i + 1}`, `name_${i}`);
    }
    // Measured ~110 ms.
    expect(performance.now() - started).toBeLessThan(3000);
    expect(DefinedNames.get(names, "NAME_19999")).toEqual({
      name: "name_19999",
      ranges: ["S!$A$20000"]
    });
  });
});
