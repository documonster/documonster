import { testUtils } from "@excel/__tests__/shared";
import { Column } from "@excel/column";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Column", () => {
  // ===========================================================================
  // Creation & Basic Properties
  // ===========================================================================

  it("creates by definition with header, key, and width", () => {
    const sheet = testUtils.createSheetMock();

    sheet.addColumn(1, {
      header: "Col 1",
      key: "id1",
      width: 10
    });

    expect(sheet.getColumn(1).header).toBe("Col 1");
    expect(sheet.getColumn(1).headers).toEqual(["Col 1"]);
    expect(sheet.getCell(1, 1).value).toBe("Col 1");
    expect(sheet.getColumn("id1")).toBe(sheet.getColumn(1));

    sheet.getRow(2).values = { id1: "Hello, World!" };
    expect(sheet.getCell(2, 1).value).toBe("Hello, World!");
  });

  it("maintains properties (key, number, letter, header)", () => {
    const sheet = testUtils.createSheetMock();

    const column = sheet.addColumn(1);

    column.key = "id1";
    expect(sheet._keys.id1).toBe(column);

    expect(column.number).toBe(1);
    expect(column.letter).toBe("A");

    column.header = "Col 1";
    expect(sheet.getColumn(1).header).toBe("Col 1");
    expect(sheet.getColumn(1).headers).toEqual(["Col 1"]);
    expect(sheet.getCell(1, 1).value).toBe("Col 1");

    column.header = ["Col A1", "Col A2"];
    expect(sheet.getColumn(1).header).toEqual(["Col A1", "Col A2"]);
    expect(sheet.getColumn(1).headers).toEqual(["Col A1", "Col A2"]);
    expect(sheet.getCell(1, 1).value).toBe("Col A1");
    expect(sheet.getCell(2, 1).value).toBe("Col A2");

    sheet.getRow(3).values = { id1: "Hello, World!" };
    expect(sheet.getCell(3, 1).value).toBe("Hello, World!");
  });

  // ===========================================================================
  // Model Serialization
  // ===========================================================================

  it("creates model from columns with outlineLevels", () => {
    const sheet = testUtils.createSheetMock();

    sheet.addColumn(1, { header: "Col 1", key: "id1", width: 10 });
    sheet.addColumn(2, { header: "Col 2", key: "name", width: 10 });
    sheet.addColumn(3, { header: "Col 2", key: "dob", width: 10, outlineLevel: 1 });

    const model = Column.toModel(sheet.columns);
    expect(model!.length).toBe(2);

    expect(model![0].width).toBe(10);
    expect(model![0].outlineLevel).toBe(0);
    expect(model![0].collapsed).toBe(false);

    expect(model![1].width).toBe(10);
    expect(model![1].outlineLevel).toBe(1);
    expect(model![1].collapsed).toBe(true);
  });

  // ===========================================================================
  // Column Values
  // ===========================================================================

  it("gets column values", () => {
    const sheet = testUtils.createSheetMock();
    sheet.getCell(1, 1).value = "a";
    sheet.getCell(2, 1).value = "b";
    sheet.getCell(4, 1).value = "d";

    expect(sheet.getColumn(1).values).toEqual([, "a", "b", , "d"]);
  });

  it("sets column values from dense array", () => {
    const sheet = testUtils.createSheetMock();

    sheet.getColumn(1).values = [2, 3, 5, 7, 11];

    expect(sheet.getCell(1, 1).value).toBe(2);
    expect(sheet.getCell(2, 1).value).toBe(3);
    expect(sheet.getCell(3, 1).value).toBe(5);
    expect(sheet.getCell(4, 1).value).toBe(7);
    expect(sheet.getCell(5, 1).value).toBe(11);
    expect(sheet.getCell(6, 1).value).toBe(null);
  });

  it("sets column values from explicit sparse array", () => {
    const sheet = testUtils.createSheetMock();
    const values: (number | undefined)[] = [];
    values[2] = 2;
    values[3] = 3;
    values[5] = 5;
    values[11] = 11;
    sheet.getColumn(1).values = values;

    expect(sheet.getCell(1, 1).value).toBe(null);
    expect(sheet.getCell(2, 1).value).toBe(2);
    expect(sheet.getCell(3, 1).value).toBe(3);
    expect(sheet.getCell(4, 1).value).toBe(null);
    expect(sheet.getCell(5, 1).value).toBe(5);
    expect(sheet.getCell(6, 1).value).toBe(null);
    expect(sheet.getCell(7, 1).value).toBe(null);
    expect(sheet.getCell(8, 1).value).toBe(null);
    expect(sheet.getCell(9, 1).value).toBe(null);
    expect(sheet.getCell(10, 1).value).toBe(null);
    expect(sheet.getCell(11, 1).value).toBe(11);
    expect(sheet.getCell(12, 1).value).toBe(null);
  });

  it("sets column values from elision-style sparse array", () => {
    const sheet = testUtils.createSheetMock();
    sheet.getColumn(1).values = [, , 2, 3, , 5, , 7, , , , 11];

    expect(sheet.getCell(1, 1).value).toBe(null);
    expect(sheet.getCell(2, 1).value).toBe(2);
    expect(sheet.getCell(3, 1).value).toBe(3);
    expect(sheet.getCell(4, 1).value).toBe(null);
    expect(sheet.getCell(5, 1).value).toBe(5);
    expect(sheet.getCell(6, 1).value).toBe(null);
    expect(sheet.getCell(7, 1).value).toBe(7);
    expect(sheet.getCell(8, 1).value).toBe(null);
    expect(sheet.getCell(9, 1).value).toBe(null);
    expect(sheet.getCell(10, 1).value).toBe(null);
    expect(sheet.getCell(11, 1).value).toBe(11);
    expect(sheet.getCell(12, 1).value).toBe(null);
  });

  // ===========================================================================
  // Width & Defaults
  // ===========================================================================

  it("sets default column width when no explicit width given", () => {
    const sheet = testUtils.createSheetMock();

    sheet.addColumn(1, { header: "Col 1", key: "id1", style: { numFmt: "0.00%" } });
    sheet.addColumn(2, { header: "Col 2", key: "id2", style: { numFmt: "0.00%" }, width: 10 });
    sheet.getColumn(3).numFmt = "0.00%";

    const model = Column.toModel(sheet.columns);
    expect(model!.length).toBe(3);
    expect(model![0].width).toBe(9); // default
    expect(model![1].width).toBe(10); // explicit
    expect(model![2].width).toBe(9); // default
  });

  it("isCustomWidth is true when width differs from default", () => {
    const sheet = testUtils.createSheetMock();
    sheet.addColumn(1, { header: "Col 1", width: 20 });
    sheet.addColumn(2, { header: "Col 2" });

    expect(sheet.getColumn(1).isCustomWidth).toBe(true);
    expect(sheet.getColumn(2).isCustomWidth).toBe(false);
  });

  // ===========================================================================
  // Hidden & Outline
  // ===========================================================================

  it("hidden property can be set and read", () => {
    const sheet = testUtils.createSheetMock();
    const col = sheet.addColumn(1);

    expect(col.hidden).toBe(false);
    col.hidden = true;
    expect(col.hidden).toBe(true);
    col.hidden = false;
    expect(col.hidden).toBe(false);
  });

  it("outlineLevel and collapsed interact correctly", () => {
    const sheet = testUtils.createSheetMock();
    const col = sheet.addColumn(1);

    expect(col.outlineLevel).toBe(0);
    expect(col.collapsed).toBe(false);

    col.outlineLevel = 1;
    expect(col.outlineLevel).toBe(1);
    // collapsed depends on worksheet.properties.outlineLevelCol
    expect(col.collapsed).toBe(true);
  });

  // ===========================================================================
  // eachCell
  // ===========================================================================

  it("eachCell iterates over non-empty cells", () => {
    const sheet = testUtils.createSheetMock();
    sheet.getCell(1, 1).value = "a";
    sheet.getCell(3, 1).value = "c";
    sheet.getCell(5, 1).value = "e";

    const collected: Array<{ row: number; value: unknown }> = [];
    sheet.getColumn(1).eachCell((cell, rowNumber) => {
      collected.push({ row: rowNumber, value: cell.value });
    });

    expect(collected).toEqual([
      { row: 1, value: "a" },
      { row: 3, value: "c" },
      { row: 5, value: "e" }
    ]);
  });

  it("eachCell with includeEmpty iterates all rows up to last", () => {
    const sheet = testUtils.createSheetMock();
    sheet.getCell(1, 1).value = "a";
    sheet.getCell(3, 1).value = "c";

    const rows: number[] = [];
    sheet.getColumn(1).eachCell({ includeEmpty: true }, (_cell, rowNumber) => {
      rows.push(rowNumber);
    });

    // Should include rows 1, 2, 3 (up to last non-empty)
    expect(rows).toContain(1);
    expect(rows).toContain(2);
    expect(rows).toContain(3);
  });

  // ===========================================================================
  // Style Properties
  // ===========================================================================

  it("column style properties propagate to cells", () => {
    const sheet = testUtils.createSheetMock();
    sheet.addColumn(1, { header: "Col 1" });
    sheet.getCell(2, 1).value = "data";

    sheet.getColumn(1).font = { bold: true };
    expect(sheet.getCell(1, 1).font).toEqual({ bold: true });
    expect(sheet.getCell(2, 1).font).toEqual({ bold: true });
  });

  it("column numFmt propagates to cells", () => {
    const sheet = testUtils.createSheetMock();
    sheet.getCell(1, 1).value = 0.5;
    sheet.getColumn(1).numFmt = "0.00%";

    expect(sheet.getCell(1, 1).numFmt).toBe("0.00%");
  });

  // ===========================================================================
  // headerCount
  // ===========================================================================

  it("headerCount reflects number of header rows", () => {
    const sheet = testUtils.createSheetMock();
    const col = sheet.addColumn(1, { header: "Single" });
    expect(col.headerCount).toBe(1);

    col.header = ["Row1", "Row2", "Row3"];
    expect(col.headerCount).toBe(3);
  });

  // ===========================================================================
  // equivalentTo
  // ===========================================================================

  it("equivalentTo returns true for columns with same properties", () => {
    const sheet = testUtils.createSheetMock();
    sheet.addColumn(1, { width: 10 });
    sheet.addColumn(2, { width: 10 });

    expect(sheet.getColumn(1).equivalentTo(sheet.getColumn(2))).toBe(true);
  });

  it("equivalentTo returns false for columns with different width", () => {
    const sheet = testUtils.createSheetMock();
    sheet.addColumn(1, { width: 10 });
    sheet.addColumn(2, { width: 20 });

    expect(sheet.getColumn(1).equivalentTo(sheet.getColumn(2))).toBe(false);
  });

  // ===========================================================================
  // XLSX Round-Trip (via Workbook)
  // ===========================================================================

  it("column properties survive XLSX round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("test");
    ws.columns = [
      { header: "ID", key: "id", width: 10 },
      { header: "Name", key: "name", width: 32 },
      { header: "DOB", key: "dob", width: 15, outlineLevel: 1 }
    ];
    ws.addRow({ id: 1, name: "Alice", dob: new Date(1990, 0, 1) });

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);

    const ws2 = wb2.getWorksheet("test")!;
    expect(ws2.getColumn(1).width).toBe(10);
    expect(ws2.getColumn(2).width).toBe(32);
    expect(ws2.getColumn(3).width).toBe(15);
    expect(ws2.getColumn(3).outlineLevel).toBe(1);
  });

  describe("style isolation", () => {
    it("mutating a cell border after col.border broadcast does not leak to other cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      ws.getCell("A1").value = "row1";
      ws.getCell("A2").value = "row2";

      ws.getColumn(1).border = { top: { style: "thin" }, bottom: { style: "thin" } };

      ws.getCell("A1").border!.top = { style: "thick" };

      expect(ws.getCell("A1").border!.top).toEqual({ style: "thick" });
      expect(ws.getCell("A2").border!.top).toEqual({ style: "thin" });
    });

    it("mutating a cell fill after col.fill broadcast does not leak to other cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      ws.getCell("A1").value = "row1";
      ws.getCell("A2").value = "row2";

      ws.getColumn(1).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" }
      };

      (ws.getCell("A1").fill as any).fgColor = { argb: "FF00FF00" };

      expect((ws.getCell("A1").fill as any).fgColor).toEqual({ argb: "FF00FF00" });
      expect((ws.getCell("A2").fill as any).fgColor).toEqual({ argb: "FFFF0000" });
    });
  });
});
