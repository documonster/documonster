import {
  createDefinedNames,
  definedNamesAdd,
  definedNamesGetAllNames,
  definedNamesGetNames,
  definedNamesGetRanges,
  definedNamesGetRangesScoped,
  definedNamesModel,
  definedNamesRemove,
  definedNamesSetModel,
  definedNamesSpliceColumns,
  definedNamesSpliceRows
} from "@excel/defined-names";
import { Cell, Workbook } from "@excel/index";
import { getDefinedNames } from "@excel/workbook";
import { describe, it, expect } from "vitest";

describe("DefinedNames", () => {
  // ===========================================================================
  // Add / Get Names
  // ===========================================================================

  it("adds names for cells (absolute and relative references)", () => {
    const dn = createDefinedNames();

    definedNamesAdd(dn, "blort!A1", "foo");
    expect(definedNamesGetNames(dn, "blort!A1")).toEqual(["foo"]);
    expect(definedNamesGetNames(dn, "blort!$A$1")).toEqual(["foo"]);

    definedNamesAdd(dn, "blort!$B$4", "bar");
    expect(definedNamesGetNames(dn, "blort!B4")).toEqual(["bar"]);
    expect(definedNamesGetNames(dn, "blort!$B$4")).toEqual(["bar"]);

    definedNamesAdd(dn, "'blo rt'!$B$4", "bar");
    expect(definedNamesGetNames(dn, "'blo rt'!$B$4")).toEqual(["bar"]);
    definedNamesAdd(dn, "'blo ,!rt'!$B$4", "bar");
    expect(definedNamesGetNames(dn, "'blo ,!rt'!$B$4")).toEqual(["bar"]);
  });

  it("adds multiple names to the same cell", () => {
    const dn = createDefinedNames();

    definedNamesAdd(dn, "Sheet1!A1", "name1");
    definedNamesAdd(dn, "Sheet1!A1", "name2");

    const names = definedNamesGetNames(dn, "Sheet1!A1");
    expect(names).toContain("name1");
    expect(names).toContain("name2");
    expect(names.length).toBe(2);
  });

  it("getNames returns empty array for unnamed cell", () => {
    const dn = createDefinedNames();
    expect(definedNamesGetNames(dn, "Sheet1!Z99")).toEqual([]);
  });

  // ===========================================================================
  // Remove Names
  // ===========================================================================

  it("removes names for cells", () => {
    const dn = createDefinedNames();

    definedNamesAdd(dn, "blort!A1", "foo");
    definedNamesAdd(dn, "blort!A1", "bar");
    definedNamesRemove(dn, "blort!A1", "foo");

    expect(definedNamesGetNames(dn, "blort!A1")).toEqual(["bar"]);
  });

  it("removes non-existent name without error", () => {
    const dn = createDefinedNames();
    definedNamesAdd(dn, "blort!A1", "bar");
    definedNamesRemove(dn, "blort!A1", "foo"); // foo doesn't exist

    expect(definedNamesGetNames(dn, "blort!A1")).toEqual(["bar"]);
  });

  // ===========================================================================
  // Get Ranges
  // ===========================================================================

  it("gets the right ranges for a name (vertical, horizontal, square, single)", () => {
    const dn = createDefinedNames();

    definedNamesAdd(dn, "blort!A1", "vertical");
    definedNamesAdd(dn, "blort!A2", "vertical");
    definedNamesAdd(dn, "blort!A3", "vertical");

    definedNamesAdd(dn, "blort!C1", "horizontal");
    definedNamesAdd(dn, "blort!D1", "horizontal");
    definedNamesAdd(dn, "blort!E1", "horizontal");

    definedNamesAdd(dn, "blort!C3", "square");
    definedNamesAdd(dn, "blort!D3", "square");
    definedNamesAdd(dn, "blort!C4", "square");
    definedNamesAdd(dn, "blort!D4", "square");

    definedNamesAdd(dn, "other!A1", "single");

    expect(definedNamesGetRanges(dn, "vertical")).toEqual({
      name: "vertical",
      ranges: ["blort!$A$1:$A$3"]
    });
    expect(definedNamesGetRanges(dn, "horizontal")).toEqual({
      name: "horizontal",
      ranges: ["blort!$C$1:$E$1"]
    });
    expect(definedNamesGetRanges(dn, "square")).toEqual({
      name: "square",
      ranges: ["blort!$C$3:$D$4"]
    });
    expect(definedNamesGetRanges(dn, "single")).toEqual({
      name: "single",
      ranges: ["other!$A$1"]
    });
  });

  it("getRanges returns empty ranges for non-existent name", () => {
    const dn = createDefinedNames();
    const result = definedNamesGetRanges(dn, "nonexistent");
    expect(result.name).toBe("nonexistent");
    expect(result.ranges).toEqual([]);
  });

  // ===========================================================================
  // Splice
  // ===========================================================================

  it("splices rows and columns correctly", () => {
    const dn = createDefinedNames();
    definedNamesAdd(dn, "vertical!A1", "vertical");
    definedNamesAdd(dn, "vertical!A2", "vertical");
    definedNamesAdd(dn, "vertical!A3", "vertical");
    definedNamesAdd(dn, "vertical!A4", "vertical");

    definedNamesAdd(dn, "horizontal!A1", "horizontal");
    definedNamesAdd(dn, "horizontal!B1", "horizontal");
    definedNamesAdd(dn, "horizontal!C1", "horizontal");
    definedNamesAdd(dn, "horizontal!D1", "horizontal");

    ["A", "B", "C", "D"].forEach(col => {
      [1, 2, 3, 4].forEach(row => {
        definedNamesAdd(dn, `square!${col}${row}`, "square");
      });
    });

    definedNamesAdd(dn, "single!A1", "singleA1");
    definedNamesAdd(dn, "single!D1", "singleD1");
    definedNamesAdd(dn, "single!A4", "singleA4");
    definedNamesAdd(dn, "single!D4", "singleD4");

    definedNamesSpliceRows(dn, "vertical", 2, 2, 1);
    definedNamesSpliceColumns(dn, "horizontal", 2, 2, 1);
    definedNamesSpliceRows(dn, "square", 2, 2, 1);
    definedNamesSpliceColumns(dn, "square", 2, 2, 1);
    definedNamesSpliceRows(dn, "single", 2, 2, 1);
    definedNamesSpliceColumns(dn, "single", 2, 2, 1);

    expect(definedNamesGetRanges(dn, "vertical")).toEqual({
      name: "vertical",
      ranges: ["vertical!$A$1", "vertical!$A$3"]
    });
    expect(definedNamesGetRanges(dn, "horizontal")).toEqual({
      name: "horizontal",
      ranges: ["horizontal!$A$1", "horizontal!$C$1"]
    });
    expect(definedNamesGetRanges(dn, "square")).toEqual({
      name: "square",
      ranges: ["square!$A$1", "square!$C$1", "square!$A$3", "square!$C$3"]
    });
    expect(definedNamesGetRanges(dn, "singleA1")).toEqual({
      name: "singleA1",
      ranges: ["single!$A$1"]
    });
    expect(definedNamesGetRanges(dn, "singleD1")).toEqual({
      name: "singleD1",
      ranges: ["single!$C$1"]
    });
    expect(definedNamesGetRanges(dn, "singleA4")).toEqual({
      name: "singleA4",
      ranges: ["single!$A$3"]
    });
    expect(definedNamesGetRanges(dn, "singleD4")).toEqual({
      name: "singleD4",
      ranges: ["single!$C$3"]
    });
  });

  // ===========================================================================
  // Model Serialization
  // ===========================================================================

  it("creates matrix from model", () => {
    const dn = createDefinedNames();

    definedNamesSetModel(dn, []);
    definedNamesAdd(dn, "blort!A1", "bar");
    definedNamesRemove(dn, "blort!A1", "foo");

    expect(definedNamesGetNames(dn, "blort!A1")).toEqual(["bar"]);
  });

  it("skips values with invalid ranges (formulas and #REF!)", () => {
    const dn = createDefinedNames();
    definedNamesSetModel(dn, [
      { name: "eq", ranges: ['"="'] },
      { name: "ref", ranges: ["#REF!"] },
      { name: "single", ranges: ["Sheet3!$A$1"] },
      { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] }
    ]);

    expect(definedNamesModel(dn)).toEqual([
      { name: "single", ranges: ["Sheet3!$A$1"] },
      { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] }
    ]);
  });

  it("model getter returns all defined names as array", () => {
    const dn = createDefinedNames();
    definedNamesAdd(dn, "Sheet1!A1", "alpha");
    definedNamesAdd(dn, "Sheet1!A2", "alpha");
    definedNamesAdd(dn, "Sheet1!B1", "beta");

    const model = definedNamesModel(dn);
    expect(model.length).toBe(2);

    const alpha = model.find(m => m.name === "alpha");
    expect(alpha).toBeDefined();
    expect(alpha!.ranges.length).toBe(1); // A1:A2 merged into single range

    const beta = model.find(m => m.name === "beta");
    expect(beta).toBeDefined();
    expect(beta!.ranges).toEqual(["Sheet1!$B$1"]);
  });

  it("model round-trip preserves names", () => {
    const dn1 = createDefinedNames();
    definedNamesAdd(dn1, "Sheet1!A1", "myName");
    definedNamesAdd(dn1, "Sheet1!A2", "myName");
    definedNamesAdd(dn1, "Sheet2!B3", "otherName");

    const model = definedNamesModel(dn1);

    const dn2 = createDefinedNames();
    definedNamesSetModel(dn2, model);

    expect(definedNamesGetRanges(dn2, "myName")).toEqual(definedNamesGetRanges(dn1, "myName"));
    expect(definedNamesGetRanges(dn2, "otherName")).toEqual(
      definedNamesGetRanges(dn1, "otherName")
    );
  });

  // ===========================================================================
  // XLSX Round-Trip
  // ===========================================================================

  it("defined names survive XLSX round-trip", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Data");
    Cell.setValue(ws, "A1", 100);
    Cell.setValue(ws, "A2", 200);
    Cell.setValue(ws, "A3", 300);

    definedNamesAdd(getDefinedNames(wb), "Data!A1:A3", "myRange");

    const buffer = await Workbook.toXlsxBuffer(wb);
    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, buffer);

    const ranges = definedNamesGetRanges(getDefinedNames(wb2), "myRange");
    expect(ranges.name).toBe("myRange");
    expect(ranges.ranges.length).toBe(1);
    expect(ranges.ranges[0]).toContain("Data");
  });

  // ===========================================================================
  // Two-Phase Classifier
  // ===========================================================================

  describe("two-phase classifier", () => {
    it("classifies sheet names with parentheses as reference, not formula", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "MyCell", ranges: [], rawText: "'Budget (2024)'!$A$1" }]);

      // Should be in matrixMap, not formulaMap or opaqueMap
      expect(dn.matrixMap["MyCell"]).toBeDefined();
      expect(dn.formulaMap["MyCell"]).toBeUndefined();
      expect(dn.opaqueMap["MyCell"]).toBeUndefined();
    });

    it("classifies OFFSET formula as formula", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        { name: "MyFormula", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }
      ]);

      expect(dn.formulaMap["MyFormula"]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
      expect(dn.matrixMap["MyFormula"]).toBeUndefined();
      expect(dn.opaqueMap["MyFormula"]).toBeUndefined();
    });

    it("classifies LAMBDA formula as formula", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "MyLambda", ranges: [], rawText: "LAMBDA(x,y,x+y)" }]);

      expect(dn.formulaMap["MyLambda"]).toBe("LAMBDA(x,y,x+y)");
      expect(dn.matrixMap["MyLambda"]).toBeUndefined();
      expect(dn.opaqueMap["MyLambda"]).toBeUndefined();
    });

    it("classifies #REF! as opaque", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "BadRef", ranges: [], rawText: "#REF!" }]);

      expect(dn.opaqueMap["BadRef"]).toBeDefined();
      expect(dn.opaqueMap["BadRef"].rawText).toBe("#REF!");
      expect(dn.matrixMap["BadRef"]).toBeUndefined();
      expect(dn.formulaMap["BadRef"]).toBeUndefined();
    });

    it("classifies string literal as opaque", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "MyStr", ranges: [], rawText: '"hello world"' }]);

      expect(dn.opaqueMap["MyStr"]).toBeDefined();
      expect(dn.opaqueMap["MyStr"].rawText).toBe('"hello world"');
      expect(dn.matrixMap["MyStr"]).toBeUndefined();
      expect(dn.formulaMap["MyStr"]).toBeUndefined();
    });

    it("classifies array constant as opaque", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "MyArr", ranges: [], rawText: "{1,2;3,4}" }]);

      expect(dn.opaqueMap["MyArr"]).toBeDefined();
      expect(dn.opaqueMap["MyArr"].rawText).toBe("{1,2;3,4}");
      expect(dn.matrixMap["MyArr"]).toBeUndefined();
      expect(dn.formulaMap["MyArr"]).toBeUndefined();
    });

    it("classifies plain cell reference as reference", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [{ name: "SingleCell", ranges: [], rawText: "Sheet1!$A$1" }]);

      expect(dn.matrixMap["SingleCell"]).toBeDefined();
      expect(dn.formulaMap["SingleCell"]).toBeUndefined();
      expect(dn.opaqueMap["SingleCell"]).toBeUndefined();
    });

    it("classifies comma-separated ranges as reference", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        { name: "MultiRange", ranges: [], rawText: "Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2" }
      ]);

      expect(dn.matrixMap["MultiRange"]).toBeDefined();
      expect(dn.formulaMap["MultiRange"]).toBeUndefined();
      expect(dn.opaqueMap["MultiRange"]).toBeUndefined();
    });

    it("preserves opaque names in model getter output", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        { name: "Good", ranges: [], rawText: "Sheet1!$A$1" },
        { name: "BadRef", ranges: [], rawText: "#REF!" },
        { name: "MyFormula", ranges: [], rawText: "SUM(Sheet1!$A$1:$A$3)" }
      ]);

      const model = definedNamesModel(dn);
      // Should have all three: reference + formula + opaque
      expect(model.length).toBe(3);

      const good = model.find(m => m.name === "Good");
      expect(good).toBeDefined();
      expect(good!.ranges.length).toBeGreaterThan(0);

      const formula = model.find(m => m.name === "MyFormula");
      expect(formula).toBeDefined();
      expect(formula!.formulaExpression).toBe("SUM(Sheet1!$A$1:$A$3)");

      const opaque = model.find(m => m.name === "BadRef");
      expect(opaque).toBeDefined();
      expect(opaque!.kind).toBe("opaque");
      expect(opaque!.rawText).toBe("#REF!");
      expect(opaque!.ranges).toEqual([]);
    });
  });

  // ===========================================================================
  // Opaque Round-Trip
  // ===========================================================================

  it("opaque defined names survive XLSX round-trip", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);

    // Inject opaque names via model setter
    const existingModel = definedNamesModel(getDefinedNames(wb));
    definedNamesSetModel(getDefinedNames(wb), [
      ...existingModel,
      { name: "OpaqueError", ranges: [], rawText: "#REF!" },
      { name: "OpaqueStr", ranges: [], rawText: '"hello"' }
    ]);

    const buffer = await Workbook.toXlsxBuffer(wb);
    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, buffer);

    // Opaque names should survive round-trip
    expect(getDefinedNames(wb2).opaqueMap["OpaqueError"]).toBeDefined();
    expect(getDefinedNames(wb2).opaqueMap["OpaqueError"].rawText).toBe("#REF!");
    expect(getDefinedNames(wb2).opaqueMap["OpaqueStr"]).toBeDefined();
    expect(getDefinedNames(wb2).opaqueMap["OpaqueStr"].rawText).toBe('"hello"');
  });

  // ===========================================================================
  // Scoped names: workbook-scope + sheet-scope coexistence
  // ===========================================================================

  describe("scoped defined names", () => {
    it("should store same-name entries with different scopes independently", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        { name: "Total", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
        {
          name: "Total",
          ranges: ["Sheet2!$B$2"],
          rawText: "Sheet2!$B$2",
          localSheetId: 0
        }
      ]);

      // getAllNames should return two entries
      const allNames = definedNamesGetAllNames(dn);
      expect(allNames.length).toBe(2);
      const global = allNames.find(e => e.localSheetId === undefined);
      const scoped = allNames.find(e => e.localSheetId === 0);
      expect(global).toBeDefined();
      expect(scoped).toBeDefined();
      expect(global!.name).toBe("Total");
      expect(scoped!.name).toBe("Total");

      // getRangesScoped should return different ranges for each scope
      const globalRanges = definedNamesGetRangesScoped(dn, "Total");
      expect(globalRanges.ranges).toEqual(["Sheet1!$A$1"]);

      const scopedRanges = definedNamesGetRangesScoped(dn, "Total", 0);
      expect(scopedRanges.ranges).toEqual(["Sheet2!$B$2"]);
    });

    it("should round-trip scoped names through model getter/setter", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        { name: "Rate", ranges: ["Sheet1!$C$1"], rawText: "Sheet1!$C$1" },
        {
          name: "Rate",
          ranges: ["Sheet2!$D$1"],
          rawText: "Sheet2!$D$1",
          localSheetId: 1
        }
      ]);

      const model = definedNamesModel(dn);
      expect(model.length).toBe(2);

      const globalEntry = model.find(m => m.localSheetId === undefined);
      const scopedEntry = model.find(m => m.localSheetId === 1);
      expect(globalEntry).toBeDefined();
      expect(globalEntry!.name).toBe("Rate");
      expect(scopedEntry).toBeDefined();
      expect(scopedEntry!.name).toBe("Rate");
      expect(scopedEntry!.localSheetId).toBe(1);
    });

    it("should handle formula-based scoped names", () => {
      const dn = createDefinedNames();
      definedNamesSetModel(dn, [
        {
          name: "MyLambda",
          ranges: ["LAMBDA(x,x+1)"],
          rawText: "LAMBDA(x,x+1)",
          formulaExpression: "LAMBDA(x,x+1)"
        },
        {
          name: "MyLambda",
          ranges: ["LAMBDA(x,x*2)"],
          rawText: "LAMBDA(x,x*2)",
          formulaExpression: "LAMBDA(x,x*2)",
          localSheetId: 0
        }
      ]);

      const globalRanges = definedNamesGetRangesScoped(dn, "MyLambda");
      expect(globalRanges.ranges).toEqual(["LAMBDA(x,x+1)"]);

      const scopedRanges = definedNamesGetRangesScoped(dn, "MyLambda", 0);
      expect(scopedRanges.ranges).toEqual(["LAMBDA(x,x*2)"]);
    });
  });
});
