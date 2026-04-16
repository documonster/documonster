import { DefinedNames } from "@excel/defined-names";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("DefinedNames", () => {
  // ===========================================================================
  // Add / Get Names
  // ===========================================================================

  it("adds names for cells (absolute and relative references)", () => {
    const dn = new DefinedNames();

    dn.add("blort!A1", "foo");
    expect(dn.getNames("blort!A1")).toEqual(["foo"]);
    expect(dn.getNames("blort!$A$1")).toEqual(["foo"]);

    dn.add("blort!$B$4", "bar");
    expect(dn.getNames("blort!B4")).toEqual(["bar"]);
    expect(dn.getNames("blort!$B$4")).toEqual(["bar"]);

    dn.add("'blo rt'!$B$4", "bar");
    expect(dn.getNames("'blo rt'!$B$4")).toEqual(["bar"]);
    dn.add("'blo ,!rt'!$B$4", "bar");
    expect(dn.getNames("'blo ,!rt'!$B$4")).toEqual(["bar"]);
  });

  it("adds multiple names to the same cell", () => {
    const dn = new DefinedNames();

    dn.add("Sheet1!A1", "name1");
    dn.add("Sheet1!A1", "name2");

    const names = dn.getNames("Sheet1!A1");
    expect(names).toContain("name1");
    expect(names).toContain("name2");
    expect(names.length).toBe(2);
  });

  it("getNames returns empty array for unnamed cell", () => {
    const dn = new DefinedNames();
    expect(dn.getNames("Sheet1!Z99")).toEqual([]);
  });

  // ===========================================================================
  // Remove Names
  // ===========================================================================

  it("removes names for cells", () => {
    const dn = new DefinedNames();

    dn.add("blort!A1", "foo");
    dn.add("blort!A1", "bar");
    dn.remove("blort!A1", "foo");

    expect(dn.getNames("blort!A1")).toEqual(["bar"]);
  });

  it("removes non-existent name without error", () => {
    const dn = new DefinedNames();
    dn.add("blort!A1", "bar");
    dn.remove("blort!A1", "foo"); // foo doesn't exist

    expect(dn.getNames("blort!A1")).toEqual(["bar"]);
  });

  // ===========================================================================
  // Get Ranges
  // ===========================================================================

  it("gets the right ranges for a name (vertical, horizontal, square, single)", () => {
    const dn = new DefinedNames();

    dn.add("blort!A1", "vertical");
    dn.add("blort!A2", "vertical");
    dn.add("blort!A3", "vertical");

    dn.add("blort!C1", "horizontal");
    dn.add("blort!D1", "horizontal");
    dn.add("blort!E1", "horizontal");

    dn.add("blort!C3", "square");
    dn.add("blort!D3", "square");
    dn.add("blort!C4", "square");
    dn.add("blort!D4", "square");

    dn.add("other!A1", "single");

    expect(dn.getRanges("vertical")).toEqual({
      name: "vertical",
      ranges: ["blort!$A$1:$A$3"]
    });
    expect(dn.getRanges("horizontal")).toEqual({
      name: "horizontal",
      ranges: ["blort!$C$1:$E$1"]
    });
    expect(dn.getRanges("square")).toEqual({
      name: "square",
      ranges: ["blort!$C$3:$D$4"]
    });
    expect(dn.getRanges("single")).toEqual({
      name: "single",
      ranges: ["other!$A$1"]
    });
  });

  it("getRanges returns empty ranges for non-existent name", () => {
    const dn = new DefinedNames();
    const result = dn.getRanges("nonexistent");
    expect(result.name).toBe("nonexistent");
    expect(result.ranges).toEqual([]);
  });

  // ===========================================================================
  // Splice
  // ===========================================================================

  it("splices rows and columns correctly", () => {
    const dn = new DefinedNames();
    dn.add("vertical!A1", "vertical");
    dn.add("vertical!A2", "vertical");
    dn.add("vertical!A3", "vertical");
    dn.add("vertical!A4", "vertical");

    dn.add("horizontal!A1", "horizontal");
    dn.add("horizontal!B1", "horizontal");
    dn.add("horizontal!C1", "horizontal");
    dn.add("horizontal!D1", "horizontal");

    ["A", "B", "C", "D"].forEach(col => {
      [1, 2, 3, 4].forEach(row => {
        dn.add(`square!${col}${row}`, "square");
      });
    });

    dn.add("single!A1", "singleA1");
    dn.add("single!D1", "singleD1");
    dn.add("single!A4", "singleA4");
    dn.add("single!D4", "singleD4");

    dn.spliceRows("vertical", 2, 2, 1);
    dn.spliceColumns("horizontal", 2, 2, 1);
    dn.spliceRows("square", 2, 2, 1);
    dn.spliceColumns("square", 2, 2, 1);
    dn.spliceRows("single", 2, 2, 1);
    dn.spliceColumns("single", 2, 2, 1);

    expect(dn.getRanges("vertical")).toEqual({
      name: "vertical",
      ranges: ["vertical!$A$1", "vertical!$A$3"]
    });
    expect(dn.getRanges("horizontal")).toEqual({
      name: "horizontal",
      ranges: ["horizontal!$A$1", "horizontal!$C$1"]
    });
    expect(dn.getRanges("square")).toEqual({
      name: "square",
      ranges: ["square!$A$1", "square!$C$1", "square!$A$3", "square!$C$3"]
    });
    expect(dn.getRanges("singleA1")).toEqual({ name: "singleA1", ranges: ["single!$A$1"] });
    expect(dn.getRanges("singleD1")).toEqual({ name: "singleD1", ranges: ["single!$C$1"] });
    expect(dn.getRanges("singleA4")).toEqual({ name: "singleA4", ranges: ["single!$A$3"] });
    expect(dn.getRanges("singleD4")).toEqual({ name: "singleD4", ranges: ["single!$C$3"] });
  });

  // ===========================================================================
  // Model Serialization
  // ===========================================================================

  it("creates matrix from model", () => {
    const dn = new DefinedNames();

    dn.model = [];
    dn.add("blort!A1", "bar");
    dn.remove("blort!A1", "foo");

    expect(dn.getNames("blort!A1")).toEqual(["bar"]);
  });

  it("skips values with invalid ranges (formulas and #REF!)", () => {
    const dn = new DefinedNames();
    dn.model = [
      { name: "eq", ranges: ['"="'] },
      { name: "ref", ranges: ["#REF!"] },
      { name: "single", ranges: ["Sheet3!$A$1"] },
      { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] }
    ];

    expect(dn.model).toEqual([
      { name: "single", ranges: ["Sheet3!$A$1"] },
      { name: "range", ranges: ["Sheet3!$A$2:$F$2228"] }
    ]);
  });

  it("model getter returns all defined names as array", () => {
    const dn = new DefinedNames();
    dn.add("Sheet1!A1", "alpha");
    dn.add("Sheet1!A2", "alpha");
    dn.add("Sheet1!B1", "beta");

    const model = dn.model;
    expect(model.length).toBe(2);

    const alpha = model.find(m => m.name === "alpha");
    expect(alpha).toBeDefined();
    expect(alpha!.ranges.length).toBe(1); // A1:A2 merged into single range

    const beta = model.find(m => m.name === "beta");
    expect(beta).toBeDefined();
    expect(beta!.ranges).toEqual(["Sheet1!$B$1"]);
  });

  it("model round-trip preserves names", () => {
    const dn1 = new DefinedNames();
    dn1.add("Sheet1!A1", "myName");
    dn1.add("Sheet1!A2", "myName");
    dn1.add("Sheet2!B3", "otherName");

    const model = dn1.model;

    const dn2 = new DefinedNames();
    dn2.model = model;

    expect(dn2.getRanges("myName")).toEqual(dn1.getRanges("myName"));
    expect(dn2.getRanges("otherName")).toEqual(dn1.getRanges("otherName"));
  });

  // ===========================================================================
  // XLSX Round-Trip
  // ===========================================================================

  it("defined names survive XLSX round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = 100;
    ws.getCell("A2").value = 200;
    ws.getCell("A3").value = 300;

    wb.definedNames.add("Data!A1:A3", "myRange");

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);

    const ranges = wb2.definedNames.getRanges("myRange");
    expect(ranges.name).toBe("myRange");
    expect(ranges.ranges.length).toBe(1);
    expect(ranges.ranges[0]).toContain("Data");
  });

  // ===========================================================================
  // Two-Phase Classifier
  // ===========================================================================

  describe("two-phase classifier", () => {
    it("classifies sheet names with parentheses as reference, not formula", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MyCell", ranges: [], rawText: "'Budget (2024)'!$A$1" }];

      // Should be in matrixMap, not formulaMap or opaqueMap
      expect(dn.matrixMap["MyCell"]).toBeDefined();
      expect(dn.formulaMap["MyCell"]).toBeUndefined();
      expect(dn.opaqueMap["MyCell"]).toBeUndefined();
    });

    it("classifies OFFSET formula as formula", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MyFormula", ranges: [], rawText: "OFFSET(Sheet1!$A$1,0,0,3,1)" }];

      expect(dn.formulaMap["MyFormula"]).toBe("OFFSET(Sheet1!$A$1,0,0,3,1)");
      expect(dn.matrixMap["MyFormula"]).toBeUndefined();
      expect(dn.opaqueMap["MyFormula"]).toBeUndefined();
    });

    it("classifies LAMBDA formula as formula", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MyLambda", ranges: [], rawText: "LAMBDA(x,y,x+y)" }];

      expect(dn.formulaMap["MyLambda"]).toBe("LAMBDA(x,y,x+y)");
      expect(dn.matrixMap["MyLambda"]).toBeUndefined();
      expect(dn.opaqueMap["MyLambda"]).toBeUndefined();
    });

    it("classifies #REF! as opaque", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "BadRef", ranges: [], rawText: "#REF!" }];

      expect(dn.opaqueMap["BadRef"]).toBeDefined();
      expect(dn.opaqueMap["BadRef"].rawText).toBe("#REF!");
      expect(dn.matrixMap["BadRef"]).toBeUndefined();
      expect(dn.formulaMap["BadRef"]).toBeUndefined();
    });

    it("classifies string literal as opaque", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MyStr", ranges: [], rawText: '"hello world"' }];

      expect(dn.opaqueMap["MyStr"]).toBeDefined();
      expect(dn.opaqueMap["MyStr"].rawText).toBe('"hello world"');
      expect(dn.matrixMap["MyStr"]).toBeUndefined();
      expect(dn.formulaMap["MyStr"]).toBeUndefined();
    });

    it("classifies array constant as opaque", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MyArr", ranges: [], rawText: "{1,2;3,4}" }];

      expect(dn.opaqueMap["MyArr"]).toBeDefined();
      expect(dn.opaqueMap["MyArr"].rawText).toBe("{1,2;3,4}");
      expect(dn.matrixMap["MyArr"]).toBeUndefined();
      expect(dn.formulaMap["MyArr"]).toBeUndefined();
    });

    it("classifies plain cell reference as reference", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "SingleCell", ranges: [], rawText: "Sheet1!$A$1" }];

      expect(dn.matrixMap["SingleCell"]).toBeDefined();
      expect(dn.formulaMap["SingleCell"]).toBeUndefined();
      expect(dn.opaqueMap["SingleCell"]).toBeUndefined();
    });

    it("classifies comma-separated ranges as reference", () => {
      const dn = new DefinedNames();
      dn.model = [{ name: "MultiRange", ranges: [], rawText: "Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2" }];

      expect(dn.matrixMap["MultiRange"]).toBeDefined();
      expect(dn.formulaMap["MultiRange"]).toBeUndefined();
      expect(dn.opaqueMap["MultiRange"]).toBeUndefined();
    });

    it("preserves opaque names in model getter output", () => {
      const dn = new DefinedNames();
      dn.model = [
        { name: "Good", ranges: [], rawText: "Sheet1!$A$1" },
        { name: "BadRef", ranges: [], rawText: "#REF!" },
        { name: "MyFormula", ranges: [], rawText: "SUM(Sheet1!$A$1:$A$3)" }
      ];

      const model = dn.model;
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
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;

    // Inject opaque names via model setter
    const existingModel = wb.definedNames.model;
    wb.definedNames.model = [
      ...existingModel,
      { name: "OpaqueError", ranges: [], rawText: "#REF!" },
      { name: "OpaqueStr", ranges: [], rawText: '"hello"' }
    ];

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);

    // Opaque names should survive round-trip
    expect(wb2.definedNames.opaqueMap["OpaqueError"]).toBeDefined();
    expect(wb2.definedNames.opaqueMap["OpaqueError"].rawText).toBe("#REF!");
    expect(wb2.definedNames.opaqueMap["OpaqueStr"]).toBeDefined();
    expect(wb2.definedNames.opaqueMap["OpaqueStr"].rawText).toBe('"hello"');
  });

  // ===========================================================================
  // Scoped names: workbook-scope + sheet-scope coexistence
  // ===========================================================================

  describe("scoped defined names", () => {
    it("should store same-name entries with different scopes independently", () => {
      const dn = new DefinedNames();
      dn.model = [
        { name: "Total", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
        {
          name: "Total",
          ranges: ["Sheet2!$B$2"],
          rawText: "Sheet2!$B$2",
          localSheetId: 0
        }
      ];

      // getAllNames should return two entries
      const allNames = dn.getAllNames();
      expect(allNames.length).toBe(2);
      const global = allNames.find(e => e.localSheetId === undefined);
      const scoped = allNames.find(e => e.localSheetId === 0);
      expect(global).toBeDefined();
      expect(scoped).toBeDefined();
      expect(global!.name).toBe("Total");
      expect(scoped!.name).toBe("Total");

      // getRangesScoped should return different ranges for each scope
      const globalRanges = dn.getRangesScoped("Total");
      expect(globalRanges.ranges).toEqual(["Sheet1!$A$1"]);

      const scopedRanges = dn.getRangesScoped("Total", 0);
      expect(scopedRanges.ranges).toEqual(["Sheet2!$B$2"]);
    });

    it("should round-trip scoped names through model getter/setter", () => {
      const dn = new DefinedNames();
      dn.model = [
        { name: "Rate", ranges: ["Sheet1!$C$1"], rawText: "Sheet1!$C$1" },
        {
          name: "Rate",
          ranges: ["Sheet2!$D$1"],
          rawText: "Sheet2!$D$1",
          localSheetId: 1
        }
      ];

      const model = dn.model;
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
      const dn = new DefinedNames();
      dn.model = [
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
      ];

      const globalRanges = dn.getRangesScoped("MyLambda");
      expect(globalRanges.ranges).toEqual(["LAMBDA(x,x+1)"]);

      const scopedRanges = dn.getRangesScoped("MyLambda", 0);
      expect(scopedRanges.ranges).toEqual(["LAMBDA(x,x*2)"]);
    });
  });
});
