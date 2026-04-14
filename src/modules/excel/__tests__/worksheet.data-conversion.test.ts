import type { CellValue, WorksheetViewFrozen } from "@excel/types";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Worksheet", () => {
  // ===========================================================================
  // Helpers
  // ===========================================================================

  /** Create a worksheet from AOA data for quick test setup */
  function createWsFromAOA(data: CellValue[][]) {
    const wb = new Workbook();
    return wb.addWorksheet("Sheet1").addAOA(data);
  }

  // ===========================================================================
  // addJSON
  // ===========================================================================

  describe("addJSON", () => {
    it("adds JSON data with headers", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addJSON([
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 }
      ]);

      expect(ws.getCell("A1").value).toBe("name");
      expect(ws.getCell("B1").value).toBe("age");
      expect(ws.getCell("A2").value).toBe("Alice");
      expect(ws.getCell("B2").value).toBe(30);
      expect(ws.getCell("A3").value).toBe("Bob");
      expect(ws.getCell("B3").value).toBe(25);
    });

    it("respects header option for ordering", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addJSON([{ name: "Alice", age: 30, city: "NYC" }], { header: ["age", "name"] });

      expect(ws.getCell("A1").value).toBe("age");
      expect(ws.getCell("B1").value).toBe("name");
      expect(ws.getCell("C1").value).toBe("city");
      expect(ws.getCell("A2").value).toBe(30);
      expect(ws.getCell("B2").value).toBe("Alice");
      expect(ws.getCell("C2").value).toBe("NYC");
    });

    it("skips header when skipHeader is true", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addJSON([{ name: "Alice", age: 30 }], { skipHeader: true });

      expect(ws.getCell("A1").value).toBe("Alice");
      expect(ws.getCell("B1").value).toBe(30);
    });

    it("adds JSON data with origin offset", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addAOA([["Header1", "Header2"]]);
      ws.addJSON([{ a: 1, b: 2 }], { origin: "A2", skipHeader: true });

      expect(ws.getCell("A1").value).toBe("Header1");
      expect(ws.getCell("A2").value).toBe(1);
      expect(ws.getCell("B2").value).toBe(2);
    });

    it("appends to bottom with origin: -1", () => {
      const ws = createWsFromAOA([
        ["a", "b"],
        [1, 2]
      ]);
      ws.addJSON([{ c: 3, d: 4 }], { origin: -1 });

      expect(ws.getCell("A3").value).toBe("c");
      expect(ws.getCell("B3").value).toBe("d");
      expect(ws.getCell("A4").value).toBe(3);
      expect(ws.getCell("B4").value).toBe(4);
    });

    it("returns this for chaining", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const result = ws.addJSON([{ a: 1 }]);
      expect(result).toBe(ws);
    });

    it("handles empty data array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const result = ws.addJSON([]);
      expect(result).toBe(ws);
      expect(ws.rowCount).toBe(0);
    });
  });

  // ===========================================================================
  // toJSON
  // ===========================================================================

  describe("toJSON", () => {
    it("converts worksheet to JSON array (default: first row as header)", () => {
      const ws = createWsFromAOA([
        ["name", "age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = ws.toJSON();

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice", age: 30 });
      expect(result[1]).toMatchObject({ name: "Bob", age: 25 });
    });

    it("returns array of arrays with header: 1", () => {
      const ws = createWsFromAOA([
        ["name", "age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = ws.toJSON({ header: 1 });

      expect(result).toHaveLength(3);
      expect(result[0]).toEqual(["name", "age"]);
      expect(result[1]).toEqual(["Alice", 30]);
      expect(result[2]).toEqual(["Bob", 25]);
    });

    it("uses column letters as keys with header: 'A'", () => {
      const ws = createWsFromAOA([
        ["name", "age"],
        ["Alice", 30]
      ]);

      const result = ws.toJSON({ header: "A" });

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ A: "name", B: "age" });
      expect(result[1]).toMatchObject({ A: "Alice", B: 30 });
    });

    it("uses custom keys with header: string[]", () => {
      const ws = createWsFromAOA([
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const result = ws.toJSON({ header: ["person", "years"] });

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ person: "Alice", years: 30 });
      expect(result[1]).toMatchObject({ person: "Bob", years: 25 });
    });

    it("handles empty cells with defaultValue", () => {
      const ws = createWsFromAOA([
        ["col1", "col2"],
        ["value", null]
      ]);

      const result = ws.toJSON({ defaultValue: "" });

      expect(result[0]).toMatchObject({ col1: "value", col2: "" });
    });

    it("skips blank rows by default for objects", () => {
      const ws = createWsFromAOA([["name"], ["Alice"], [null], ["Bob"]]);

      const result = ws.toJSON();

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice" });
      expect(result[1]).toMatchObject({ name: "Bob" });
    });

    it("includes blank rows with blankRows: true for objects", () => {
      const ws = createWsFromAOA([["name"], ["Alice"], [null], ["Bob"]]);

      const result = ws.toJSON({ blankRows: true });

      expect(result).toHaveLength(3);
    });

    it("includes blank rows by default with header: 1", () => {
      const ws = createWsFromAOA([["name"], ["Alice"], [null], ["Bob"]]);

      const result = ws.toJSON({ header: 1 });

      expect(result).toHaveLength(4);
    });

    it("disambiguates duplicate headers", () => {
      const ws = createWsFromAOA([
        ["name", "name", "name"],
        ["Alice", "Bob", "Charlie"]
      ]);

      const result = ws.toJSON();

      expect(result[0]).toHaveProperty("name", "Alice");
      expect(result[0]).toHaveProperty("name_1", "Bob");
      expect(result[0]).toHaveProperty("name_2", "Charlie");
    });

    it("returns formatted text with raw: false", () => {
      const ws = createWsFromAOA([
        ["name", "age", "birthday"],
        ["Alice", 30, new Date(1994, 5, 15)],
        ["Bob", 25, new Date(1999, 11, 25)]
      ]);

      const result = ws.toJSON({ raw: false });

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice", age: "30" });
      expect(result[1]).toMatchObject({ name: "Bob", age: "25" });
      expect(typeof result[0].birthday).toBe("string");
      expect(typeof result[1].birthday).toBe("string");
    });

    it("returns raw values by default (raw: true/undefined)", () => {
      const date1 = new Date(1994, 5, 15);
      const date2 = new Date(1999, 11, 25);
      const ws = createWsFromAOA([
        ["name", "age", "birthday"],
        ["Alice", 30, date1],
        ["Bob", 25, date2]
      ]);

      const result = ws.toJSON();

      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice", age: 30 });
      expect(result[1]).toMatchObject({ name: "Bob", age: 25 });
      expect(result[0].birthday).toEqual(date1);
      expect(result[1].birthday).toEqual(date2);
    });

    it("formats time values correctly with raw: false (timezone-independent)", () => {
      const timeSerial = 32 / 86400;
      const timeAsDate = new Date(Math.round((timeSerial - 25569) * 86400000));

      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "time";
      ws.getCell("A2").value = timeAsDate;
      ws.getCell("A2").numFmt = "h:mm:ss";

      const result = ws.toJSON({ raw: false });

      expect(result).toHaveLength(1);
      expect(result[0].time).toBe("0:00:32");
    });

    it("formats formula result with number result correctly", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "total";
      ws.getCell("A2").value = { formula: "SUM(B1:B10)", result: 1234.567 };
      ws.getCell("A2").numFmt = "#,##0.00";

      const result = ws.toJSON({ raw: false });

      expect(result[0].total).toBe("1,234.57");
    });

    it("formats formula result with elapsed time format", () => {
      const durationSerial = 1.5;
      const durationAsDate = new Date(Math.round((durationSerial - 25569) * 86400000));

      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "elapsed";
      ws.getCell("A2").value = { formula: "B1-C1", result: durationAsDate };
      ws.getCell("A2").numFmt = "[h]:mm:ss";

      const result = ws.toJSON({ raw: false });

      expect(result[0].elapsed).toBe("36:00:00");
    });

    it("respects string range option", () => {
      const ws = createWsFromAOA([
        ["A", "B", "C", "D"],
        [1, 2, 3, 4],
        [5, 6, 7, 8]
      ]);

      const result = ws.toJSON({ header: 1, range: "B1:C3" });

      expect(result).toHaveLength(3);
      expect(result[0]).toEqual(["B", "C"]);
      expect(result[1]).toEqual([2, 3]);
      expect(result[2]).toEqual([6, 7]);
    });

    it("respects numeric range option (starting row)", () => {
      const ws = createWsFromAOA([
        ["header1", "header2"],
        ["skip1", "skip2"],
        ["name", "age"],
        ["Alice", 30]
      ]);

      // range: 2 means start at 0-indexed row 2 (which is 1-indexed row 3)
      const result = ws.toJSON({ range: 2 });

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({ name: "Alice", age: 30 });
    });

    it("returns empty array for empty worksheet", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Empty");

      expect(ws.toJSON()).toEqual([]);
      expect(ws.toJSON({ header: 1 })).toEqual([]);
    });
  });

  // ===========================================================================
  // addAOA
  // ===========================================================================

  describe("addAOA", () => {
    it("adds array of arrays to worksheet", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addAOA([
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      expect(ws.getCell("A1").value).toBe("Name");
      expect(ws.getCell("B1").value).toBe("Age");
      expect(ws.getCell("A2").value).toBe("Alice");
      expect(ws.getCell("B2").value).toBe(30);
    });

    it("handles origin option", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addAOA([["a", "b"]], { origin: "C3" });

      expect(ws.getCell("C3").value).toBe("a");
      expect(ws.getCell("D3").value).toBe("b");
    });

    it("handles different data types", () => {
      const date = new Date("2024-01-01");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addAOA([["string", 123, true, date, null]]);

      expect(ws.getCell("A1").value).toBe("string");
      expect(ws.getCell("B1").value).toBe(123);
      expect(ws.getCell("C1").value).toBe(true);
      expect(ws.getCell("D1").value).toEqual(date);
    });

    it("appends with origin: -1", () => {
      const ws = createWsFromAOA([["Row1"], ["Row2"]]);
      ws.addAOA([["Row3"]], { origin: -1 });

      expect(ws.getCell("A3").value).toBe("Row3");
    });

    it("returns this for chaining", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const result = ws.addAOA([["a"]]);
      expect(result).toBe(ws);
    });

    it("handles empty data array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const result = ws.addAOA([]);
      expect(result).toBe(ws);
      expect(ws.rowCount).toBe(0);
    });
  });

  // ===========================================================================
  // toAOA
  // ===========================================================================

  describe("toAOA", () => {
    it("converts worksheet to array of arrays", () => {
      const ws = createWsFromAOA([
        ["Name", "Age"],
        ["Alice", 30]
      ]);

      const result = ws.toAOA();

      expect(result[0]).toEqual(["Name", "Age"]);
      expect(result[1]).toEqual(["Alice", 30]);
    });
  });

  // ===========================================================================
  // XLSX Round-Trip
  // ===========================================================================

  describe("XLSX round-trip", () => {
    it("addJSON → writeBuffer → load → toJSON preserves data", async () => {
      const wb1 = new Workbook();
      wb1.addWorksheet("Data").addJSON([
        { name: "Alice", age: 30, active: true },
        { name: "Bob", age: 25, active: false }
      ]);

      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const result = wb2.getWorksheet("Data")!.toJSON();
      expect(result).toHaveLength(2);
      expect(result[0]).toMatchObject({ name: "Alice", age: 30, active: true });
      expect(result[1]).toMatchObject({ name: "Bob", age: 25, active: false });
    });

    it("addAOA → writeBuffer → load → toAOA preserves data", async () => {
      const wb1 = new Workbook();
      wb1.addWorksheet("Data").addAOA([
        ["X", "Y"],
        [1, 2],
        [3, 4]
      ]);

      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const result = wb2.getWorksheet("Data")!.toAOA();
      expect(result[0]).toEqual(["X", "Y"]);
      expect(result[1]).toEqual([1, 2]);
      expect(result[2]).toEqual([3, 4]);
    });
  });

  // ===========================================================================
  // importSheet
  // ===========================================================================

  describe("importSheet", () => {
    it("imports worksheet data to a new workbook", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source").addJSON([{ name: "Alice", age: 30 }]);

      const wb2 = new Workbook();
      wb2.importSheet(ws1, "Imported");

      expect(wb2.worksheets).toHaveLength(1);
      expect(wb2.worksheets[0].name).toBe("Imported");
      expect(wb2.worksheets[0].getCell("A1").value).toBe("name");
      expect(wb2.worksheets[0].getCell("A2").value).toBe("Alice");
    });

    it("uses source name when no name provided", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("OriginalName").addAOA([["data"]]);

      const wb2 = new Workbook();
      wb2.importSheet(ws1);

      expect(wb2.worksheets[0].name).toBe("OriginalName");
    });

    it("copies column widths", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getColumn(1).width = 20;
      ws1.getColumn(2).width = 40;
      ws1.getCell("A1").value = "data";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.getColumn(1).width).toBe(20);
      expect(ws2.getColumn(2).width).toBe(40);
    });

    it("copies cell styles", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getCell("A1").value = "styled";
      ws1.getCell("A1").font = { bold: true, size: 16 };

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.getCell("A1").value).toBe("styled");
      expect(ws2.getCell("A1").font!.bold).toBe(true);
      expect(ws2.getCell("A1").font!.size).toBe(16);
    });

    it("copies merged cells", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getCell("A1").value = "merged";
      ws1.mergeCells("A1:C3");

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.model.mergeCells).toContain("A1:C3");
      expect(ws2.getCell("A1").value).toBe("merged");
    });

    it("copies row heights", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getRow(1).height = 30;
      ws1.getRow(2).height = 50;
      ws1.getCell("A1").value = "tall row";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.getRow(1).height).toBe(30);
      expect(ws2.getRow(2).height).toBe(50);
    });

    it("copies data validations", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getCell("A1").dataValidation = {
        type: "list",
        formulae: ['"Yes,No"']
      };

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.getCell("A1").dataValidation).toEqual({
        type: "list",
        formulae: ['"Yes,No"']
      });
    });

    it("copies views (frozen panes)", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
      ws1.getCell("A1").value = "header";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.views).toHaveLength(1);
      const view = ws2.views[0] as WorksheetViewFrozen;
      expect(view.state).toBe("frozen");
      expect(view.xSplit).toBe(1);
      expect(view.ySplit).toBe(1);
    });

    it("copies worksheet state (hidden)", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.state = "hidden";
      ws1.getCell("A1").value = "data";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.state).toBe("hidden");
    });

    it("copies auto filter", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.getCell("A1").value = "Name";
      ws1.getCell("B1").value = "Age";
      ws1.autoFilter = "A1:B1";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.autoFilter).toBe("A1:B1");
    });

    it("copies page setup and header/footer", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Source");
      ws1.pageSetup.orientation = "landscape";
      ws1.headerFooter.oddHeader = "Page &P";
      ws1.getCell("A1").value = "data";

      const wb2 = new Workbook();
      const ws2 = wb2.importSheet(ws1);

      expect(ws2.pageSetup.orientation).toBe("landscape");
      expect(ws2.headerFooter.oddHeader).toBe("Page &P");
    });
  });
});
