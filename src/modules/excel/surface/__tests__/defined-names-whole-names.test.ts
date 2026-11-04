/**
 * Whole names: a name can be looked up and deleted however it is defined — a reference, a formula, a sheet-scoped
 * name — and `list` says when a name is hidden.
 */
import { Cell, DefinedNameError, DefinedNames, ExcelArgumentError, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function workbook() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const dn = Workbook.getDefinedNames(wb);
  DefinedNames.add(dn, "S!$A$1:$A$3", "Block");
  DefinedNames.addFormula(dn, "Tax", "0.2");
  DefinedNames.setModel(dn, [
    ...DefinedNames.getModel(dn),
    { name: "Local", ranges: ["S!$B$1"], localSheetId: 0 },
    { name: "Secret", ranges: ["S!$C$1"], hidden: true }
  ]);
  return { wb, ws, dn };
}

describe("DefinedNames.remove(dn, name)", () => {
  it("deletes a formula-defined name", async () => {
    const { wb, dn } = workbook();
    expect(DefinedNames.remove(dn, "tax")).toBe(true);
    expect(DefinedNames.get(dn, "Tax")).toBeUndefined();
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(DefinedNames.list(Workbook.getDefinedNames(back)).map(n => n.name)).not.toContain("Tax");
  });

  it("deletes every cell of a reference name", () => {
    const { ws, dn } = workbook();
    expect(DefinedNames.remove(dn, "Block")).toBe(true);
    expect(Cell.getNames(ws, "A2")).toEqual([]);
    expect(DefinedNames.remove(dn, "Block")).toBe(false);
  });

  it("deletes a sheet-scoped name only in its scope", () => {
    const { dn } = workbook();
    expect(DefinedNames.remove(dn, "Local")).toBe(false);
    expect(DefinedNames.remove(dn, "Local", { localSheetId: 0 })).toBe(true);
    expect(DefinedNames.get(dn, "Local", { localSheetId: 0 })).toBeUndefined();
  });

  it("keeps the location form", () => {
    const { ws, dn } = workbook();
    expect(DefinedNames.remove(dn, "S!$A$2", "Block")).toBe(true);
    expect(Cell.getNames(ws, "A1")).toEqual(["Block"]);
    expect(Cell.getNames(ws, "A2")).toEqual([]);
  });

  it("refuses what is not a name or a scope", () => {
    const { dn } = workbook();
    expect(() => DefinedNames.remove(dn, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => DefinedNames.remove(dn, "Tax", { localSheetId: -1 })).toThrow(ExcelArgumentError);
  });
});

describe("DefinedNames.get / list", () => {
  it("looks a name up as Excel compares names, in its scope", () => {
    const { dn } = workbook();
    expect(DefinedNames.get(dn, "block")).toEqual({ name: "Block", ranges: ["S!$A$1:$A$3"] });
    expect(DefinedNames.get(dn, "Local")).toBeUndefined();
    expect(DefinedNames.get(dn, "Local", { localSheetId: 0 })).toMatchObject({ name: "Local" });
  });

  it("reports a hidden name as hidden", () => {
    const { dn } = workbook();
    expect(DefinedNames.list(dn).find(n => n.name === "Secret")).toMatchObject({ hidden: true });
  });
});

describe("DefinedNames.setModel", () => {
  it("refuses two entries that are one name in one scope, and changes nothing", () => {
    const { dn } = workbook();
    const before = DefinedNames.getModel(dn);
    expect(() =>
      DefinedNames.setModel(dn, [
        { name: "N", ranges: ["S!$A$1"] },
        { name: "n", ranges: ["S!$A$2"] }
      ])
    ).toThrow(DefinedNameError);
    expect(DefinedNames.getModel(dn)).toEqual(before);
  });

  it("accepts one name in two scopes", () => {
    const { dn } = workbook();
    DefinedNames.setModel(dn, [
      { name: "N", ranges: ["S!$A$1"] },
      { name: "n", ranges: ["S!$A$2"], localSheetId: 0 }
    ]);
    expect(DefinedNames.list(dn)).toHaveLength(2);
  });
});
