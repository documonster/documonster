import { testUtils } from "@excel/__tests__/shared";
import {
  cellAlignment,
  cellBorder,
  cellFill,
  cellFont,
  cellFormula,
  cellGetValue,
  cellNumFmt,
  cellResult,
  cellSetAlignment,
  cellSetBorder,
  cellSetFill,
  cellSetFont,
  cellSetNumFmt,
  cellSetProtection,
  cellSetValue,
  cellType
} from "@excel/cell";
import { columnHeaders } from "@excel/column";
import { ValueType } from "@excel/enums";
import { Cell, Workbook } from "@excel/index";
import {
  rowAddPageBreak,
  rowSetAlignment,
  rowSetBorder,
  rowSetFill,
  rowSetFont,
  rowSetNumFmt,
  rowValues
} from "@excel/row";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
import type { CellFormulaValue } from "@excel/types";
import {
  columnSetAlignment,
  columnSetBorder,
  columnSetDefn,
  columnSetFill,
  columnSetFont,
  columnSetHeader,
  columnSetKey,
  columnSetNumFmt,
  getCell,
  rowCommit,
  rowGetCell
} from "@excel/worksheet";
import { testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

const CONCATENATE_HELLO_WORLD = 'CONCATENATE("Hello", ", ", "World!")';

describe("WorksheetWriter", () => {
  describe("Values", () => {
    it("stores values properly", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      const now = new Date();

      // plain number
      cellSetValue(ws.getCell("A1"), 7);

      // simple string
      cellSetValue(ws.getCell("B1"), "Hello, World!");

      // floating point
      cellSetValue(ws.getCell("C1"), 3.14);

      // 5 will be overwritten by the current date-time
      cellSetValue(ws.getCell("D1"), 5);
      cellSetValue(ws.getCell("D1"), now);

      // constructed string - will share recored with B1
      cellSetValue(ws.getCell("E1"), `${["Hello", "World"].join(", ")}!`);

      // hyperlink
      cellSetValue(ws.getCell("F1"), {
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      });

      // number formula
      cellSetValue(ws.getCell("A2"), { formula: "A1", result: 7 });

      // string formula
      cellSetValue(ws.getCell("B2"), {
        formula: CONCATENATE_HELLO_WORLD,
        result: "Hello, World!"
      });

      // date formula
      cellSetValue(ws.getCell("C2"), { formula: "D1", result: now });

      expect(cellGetValue(ws.getCell("A1"))).toBe(7);
      expect(cellGetValue(ws.getCell("B1"))).toBe("Hello, World!");
      expect(cellGetValue(ws.getCell("C1"))).toBe(3.14);
      expect(cellGetValue(ws.getCell("D1"))).toBe(now);
      expect(cellGetValue(ws.getCell("E1"))).toBe("Hello, World!");
      expect(cellGetValue(ws.getCell("F1"))).toEqual({
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      });

      expect(cellGetValue(ws.getCell("A2"))).toEqual({ formula: "A1", result: 7 });

      expect(cellGetValue(ws.getCell("B2"))).toEqual({
        formula: CONCATENATE_HELLO_WORLD,
        result: "Hello, World!"
      });

      expect(cellGetValue(ws.getCell("C2"))).toEqual({ formula: "D1", result: now });
    });

    it("stores shared string values properly", () => {
      const wb = new WorkbookWriter({
        useSharedStrings: true
      });
      const ws = wb.addWorksheet("blort");

      cellSetValue(ws.getCell("A1"), "Hello, World!");

      cellSetValue(ws.getCell("A2"), "Hello");
      cellSetValue(ws.getCell("B2"), "World");
      cellSetValue(ws.getCell("C2"), {
        formula: 'CONCATENATE(A2, ", ", B2, "!")',
        result: "Hello, World!"
      });

      cellSetValue(ws.getCell("A3"), `${["Hello", "World"].join(", ")}!`);

      // A1 and A3 should reference the same string object
      expect(cellGetValue(ws.getCell("A1"))).toBe(cellGetValue(ws.getCell("A3")));

      // A1 and C2 should not reference the same object
      expect(cellGetValue(ws.getCell("A1"))).toBe(
        (cellGetValue(ws.getCell("C2")) as CellFormulaValue).result
      );
    });

    it("assigns cell types properly", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // plain number
      cellSetValue(ws.getCell("A1"), 7);

      // simple string
      cellSetValue(ws.getCell("B1"), "Hello, World!");

      // floating point
      cellSetValue(ws.getCell("C1"), 3.14);

      // date-time
      cellSetValue(ws.getCell("D1"), new Date());

      // hyperlink
      cellSetValue(ws.getCell("E1"), {
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      });

      // number formula
      cellSetValue(ws.getCell("A2"), { formula: "A1", result: 7 });

      // string formula
      cellSetValue(ws.getCell("B2"), {
        formula: CONCATENATE_HELLO_WORLD,
        result: "Hello, World!"
      });

      // date formula
      cellSetValue(ws.getCell("C2"), { formula: "D1", result: new Date() });

      expect(cellType(ws.getCell("A1"))).toBe(ValueType.Number);
      expect(cellType(ws.getCell("B1"))).toBe(ValueType.String);
      expect(cellType(ws.getCell("C1"))).toBe(ValueType.Number);
      expect(cellType(ws.getCell("D1"))).toBe(ValueType.Date);
      expect(cellType(ws.getCell("E1"))).toBe(ValueType.Hyperlink);

      expect(cellType(ws.getCell("A2"))).toBe(ValueType.Formula);
      expect(cellType(ws.getCell("B2"))).toBe(ValueType.Formula);
      expect(cellType(ws.getCell("C2"))).toBe(ValueType.Formula);
    });

    it("adds columns", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      ws.columns = [
        { key: "id", width: 10 },
        { key: "name", width: 32 },
        { key: "dob", width: 10 }
      ];

      expect(ws.getColumn("id").number).toBe(1);
      expect(ws.getColumn("id").width).toBe(10);
      expect(ws.getColumn("A")).toBe(ws.getColumn("id"));
      expect(ws.getColumn(1)).toBe(ws.getColumn("id"));

      expect(ws.getColumn("name").number).toBe(2);
      expect(ws.getColumn("name").width).toBe(32);
      expect(ws.getColumn("B")).toBe(ws.getColumn("name"));
      expect(ws.getColumn(2)).toBe(ws.getColumn("name"));

      expect(ws.getColumn("dob").number).toBe(3);
      expect(ws.getColumn("dob").width).toBe(10);
      expect(ws.getColumn("C")).toBe(ws.getColumn("dob"));
      expect(ws.getColumn(3)).toBe(ws.getColumn("dob"));
    });

    it("adds column headers", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      ws.columns = [
        { header: "Id", width: 10 },
        { header: "Name", width: 32 },
        { header: "D.O.B.", width: 10 }
      ];

      expect(cellGetValue(ws.getCell("A1"))).toBe("Id");
      expect(cellGetValue(ws.getCell("B1"))).toBe("Name");
      expect(cellGetValue(ws.getCell("C1"))).toBe("D.O.B.");
    });

    it("adds column headers by number", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // by defn
      columnSetDefn(ws.getColumn(1), { key: "id", header: "Id", width: 10 });

      // by property
      columnSetKey(ws.getColumn(2), "name");
      columnSetHeader(ws.getColumn(2), "Name");
      ws.getColumn(2).width = 32;

      expect(cellGetValue(ws.getCell("A1"))).toBe("Id");
      expect(cellGetValue(ws.getCell("B1"))).toBe("Name");

      expect(ws.getColumn("A").key).toBe("id");
      expect(ws.getColumn(1).key).toBe("id");
      expect(ws.getColumn(1).header).toBe("Id");
      expect(columnHeaders(ws.getColumn(1))).toEqual(["Id"]);
      expect(ws.getColumn(1).width).toBe(10);

      expect(ws.getColumn(2).key).toBe("name");
      expect(ws.getColumn(2).header).toBe("Name");
      expect(columnHeaders(ws.getColumn(2))).toEqual(["Name"]);
      expect(ws.getColumn(2).width).toBe(32);
    });

    it("adds column headers by letter", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // by defn
      columnSetDefn(ws.getColumn("A"), { key: "id", header: "Id", width: 10 });

      // by property
      columnSetKey(ws.getColumn("B"), "name");
      columnSetHeader(ws.getColumn("B"), "Name");
      ws.getColumn("B").width = 32;

      expect(cellGetValue(ws.getCell("A1"))).toBe("Id");
      expect(cellGetValue(ws.getCell("B1"))).toBe("Name");

      expect(ws.getColumn("A").key).toBe("id");
      expect(ws.getColumn(1).key).toBe("id");
      expect(ws.getColumn("A").header).toBe("Id");
      expect(columnHeaders(ws.getColumn("A"))).toEqual(["Id"]);
      expect(ws.getColumn("A").width).toBe(10);

      expect(ws.getColumn("B").key).toBe("name");
      expect(ws.getColumn("B").header).toBe("Name");
      expect(columnHeaders(ws.getColumn("B"))).toEqual(["Name"]);
      expect(ws.getColumn("B").width).toBe(32);
    });

    it("adds rows by object", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // add columns to define column keys
      ws.columns = [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 32 },
        { header: "D.O.B.", key: "dob", width: 10 }
      ];

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);

      ws.addRow({ id: 1, name: "John Doe", dob: dateValue1 });
      ws.addRow({ id: 2, name: "Jane Doe", dob: dateValue2 });

      expect(cellGetValue(ws.getCell("A2"))).toBe(1);
      expect(cellGetValue(ws.getCell("B2"))).toBe("John Doe");
      expect(cellGetValue(ws.getCell("C2"))).toBe(dateValue1);

      expect(cellGetValue(ws.getCell("A3"))).toBe(2);
      expect(cellGetValue(ws.getCell("B3"))).toBe("Jane Doe");
      expect(cellGetValue(ws.getCell("C3"))).toBe(dateValue2);

      expect(rowValues(ws.getRow(2))).toEqual([, 1, "John Doe", dateValue1]);
      expect(rowValues(ws.getRow(3))).toEqual([, 2, "Jane Doe", dateValue2]);
    });

    it("adds rows by contiguous array", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);

      ws.addRow([1, "John Doe", dateValue1]);
      ws.addRow([2, "Jane Doe", dateValue2]);

      expect(cellGetValue(ws.getCell("A1"))).toBe(1);
      expect(cellGetValue(ws.getCell("B1"))).toBe("John Doe");
      expect(cellGetValue(ws.getCell("C1"))).toBe(dateValue1);

      expect(cellGetValue(ws.getCell("A2"))).toBe(2);
      expect(cellGetValue(ws.getCell("B2"))).toBe("Jane Doe");
      expect(cellGetValue(ws.getCell("C2"))).toBe(dateValue2);

      expect(rowValues(ws.getRow(1))).toEqual([, 1, "John Doe", dateValue1]);
      expect(rowValues(ws.getRow(2))).toEqual([, 2, "Jane Doe", dateValue2]);
    });

    it("adds rows by sparse array", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const rows = [, [, 1, "John Doe", , dateValue1], [, 2, "Jane Doe", , dateValue2]];
      const row3: any[] = [];
      row3[1] = 3;
      row3[3] = "Sam";
      row3[5] = dateValue1;
      rows.push(row3);
      rows.forEach(row => {
        if (row) {
          ws.addRow(row);
        }
      });

      expect(cellGetValue(ws.getCell("A1"))).toBe(1);
      expect(cellGetValue(ws.getCell("B1"))).toBe("John Doe");
      expect(cellGetValue(ws.getCell("D1"))).toBe(dateValue1);

      expect(cellGetValue(ws.getCell("A2"))).toBe(2);
      expect(cellGetValue(ws.getCell("B2"))).toBe("Jane Doe");
      expect(cellGetValue(ws.getCell("D2"))).toBe(dateValue2);

      expect(cellGetValue(ws.getCell("A3"))).toBe(3);
      expect(cellGetValue(ws.getCell("C3"))).toBe("Sam");
      expect(cellGetValue(ws.getCell("E3"))).toBe(dateValue1);

      expect(rowValues(ws.getRow(1))).toEqual(rows[1]);
      expect(rowValues(ws.getRow(2))).toEqual(rows[2]);
      expect(rowValues(ws.getRow(3))).toEqual(rows[3]);
    });

    it("sets row styles", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("basket");

      cellSetValue(ws.getCell("A1"), 5);
      cellSetNumFmt(ws.getCell("A1"), testUtils.styles.numFmts.numFmt1);
      cellSetFont(ws.getCell("A1"), testUtils.styles.fonts.arialBlackUI14);

      cellSetValue(ws.getCell("C1"), "Hello, World!");
      cellSetAlignment(ws.getCell("C1"), testUtils.styles.namedAlignments.bottomRight);
      cellSetBorder(ws.getCell("C1"), testUtils.styles.borders.doubleRed);
      cellSetFill(ws.getCell("C1"), testUtils.styles.fills.redDarkVertical);

      rowSetNumFmt(ws.getRow(1), testUtils.styles.numFmts.numFmt2);
      rowSetFont(ws.getRow(1), testUtils.styles.fonts.comicSansUdB16);
      rowSetAlignment(ws.getRow(1), testUtils.styles.namedAlignments.middleCentre);
      rowSetBorder(ws.getRow(1), testUtils.styles.borders.thin);
      rowSetFill(ws.getRow(1), testUtils.styles.fills.redGreenDarkTrellis);

      expect(cellNumFmt(ws.getCell("A1"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("A1"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("A1"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("A1"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("A1"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      expect(ws.findCell("B1")).toBeUndefined();

      expect(cellNumFmt(ws.getCell("C1"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("C1"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("C1"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("C1"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("C1"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // when we 'get' the previously null cell, it should inherit the row styles
      expect(cellNumFmt(ws.getCell("B1"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("B1"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("B1"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("B1"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("B1"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });

    it("sets col styles", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("basket");

      cellSetValue(ws.getCell("A1"), 5);
      cellSetNumFmt(ws.getCell("A1"), testUtils.styles.numFmts.numFmt1);
      cellSetFont(ws.getCell("A1"), testUtils.styles.fonts.arialBlackUI14);

      cellSetValue(ws.getCell("A3"), "Hello, World!");
      cellSetAlignment(ws.getCell("A3"), testUtils.styles.namedAlignments.bottomRight);
      cellSetBorder(ws.getCell("A3"), testUtils.styles.borders.doubleRed);
      cellSetFill(ws.getCell("A3"), testUtils.styles.fills.redDarkVertical);

      columnSetNumFmt(ws.getColumn("A"), testUtils.styles.numFmts.numFmt2);
      columnSetFont(ws.getColumn("A"), testUtils.styles.fonts.comicSansUdB16);
      columnSetAlignment(ws.getColumn("A"), testUtils.styles.namedAlignments.middleCentre);
      columnSetBorder(ws.getColumn("A"), testUtils.styles.borders.thin);
      columnSetFill(ws.getColumn("A"), testUtils.styles.fills.redGreenDarkTrellis);

      expect(cellNumFmt(ws.getCell("A1"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("A1"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("A1"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("A1"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("A1"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      expect(ws.findRow(2)).toBeUndefined();

      expect(cellNumFmt(ws.getCell("A3"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("A3"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("A3"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("A3"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("A3"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // when we 'get' the previously null cell, it should inherit the column styles
      expect(cellNumFmt(ws.getCell("A2"))).toBe(testUtils.styles.numFmts.numFmt2);
      expect(cellFont(ws.getCell("A2"))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(cellAlignment(ws.getCell("A2"))).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(cellBorder(ws.getCell("A2"))).toEqual(testUtils.styles.borders.thin);
      expect(cellFill(ws.getCell("A2"))).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });
  });

  describe("Merge Cells", () => {
    it("references the same top-left value", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // initial values
      cellSetValue(ws.getCell("A1"), "A1");
      cellSetValue(ws.getCell("B1"), "B1");
      cellSetValue(ws.getCell("A2"), "A2");
      cellSetValue(ws.getCell("B2"), "B2");

      ws.mergeCells("A1:B2");

      expect(cellGetValue(ws.getCell("A1"))).toBe("A1");
      expect(cellGetValue(ws.getCell("B1"))).toBe("A1");
      expect(cellGetValue(ws.getCell("A2"))).toBe("A1");
      expect(cellGetValue(ws.getCell("B2"))).toBe("A1");

      expect(cellType(ws.getCell("A1"))).toBe(ValueType.String);
      expect(cellType(ws.getCell("B1"))).toBe(ValueType.Merge);
      expect(cellType(ws.getCell("A2"))).toBe(ValueType.Merge);
      expect(cellType(ws.getCell("B2"))).toBe(ValueType.Merge);
    });

    it("does not allow overlapping merges", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      ws.mergeCells("B2:C3");

      // intersect four corners
      expect(() => {
        ws.mergeCells("A1:B2");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("C1:D2");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("C3:D4");
      }).toThrow(Error);
      expect(() => {
        ws.mergeCells("A3:B4");
      }).toThrow(Error);

      // enclosing
      expect(() => {
        ws.mergeCells("A1:D4");
      }).toThrow(Error);
    });
  });

  describe("Page Breaks", () => {
    it("adds multiple row breaks", () => {
      const wb = new WorkbookWriter();
      const ws = wb.addWorksheet("blort");

      // initial values
      cellSetValue(ws.getCell("A1"), "A1");
      cellSetValue(ws.getCell("B1"), "B1");
      cellSetValue(ws.getCell("A2"), "A2");
      cellSetValue(ws.getCell("B2"), "B2");
      cellSetValue(ws.getCell("A3"), "A3");
      cellSetValue(ws.getCell("B3"), "B3");

      let row = ws.getRow(1);
      rowAddPageBreak(row);
      row = ws.getRow(2);
      rowAddPageBreak(row);
      expect(ws.rowBreaks.length).toBe(2);
    });
  });

  // String formula result with date format should not be converted to date
  describe("String formula result with date format", () => {
    it("preserves string formula result with date format", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Set up a cell with text and date format (mmm-yy)
      Cell.setValue(ws, "A1", "test");
      Cell.setStyle(ws, "A1", { numFmt: "mmm-yy" });

      // Set up a formula that references the text cell
      Cell.setValue(ws, "A2", { formula: "A1", result: "test" });
      Cell.setStyle(ws, "A2", { numFmt: "mmm-yy" });

      // Write to buffer
      const buffer = await Workbook.toBuffer(wb);

      // Read back
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

      // Verify A1 is preserved as string
      expect(Cell.getValue(ws2, "A1")).toBe("test");
      expect(Cell.getStyle(ws2, "A1").numFmt).toBe("mmm-yy");

      // Verify A2 formula result is preserved as string, not converted to Invalid Date
      const cellA2 = getCell(ws2, "A2");
      expect(cellFormula(cellA2)).toBe("A1");
      expect(cellResult(cellA2)).toBe("test");
      expect(typeof cellResult(cellA2)).toBe("string");
    });
  });

  describe("XML Element Order", () => {
    it("writes sheetProtection before autoFilter for valid Excel files", async () => {
      // Per OOXML spec (http://www.datypic.com/sc/ooxml/e-ssml_worksheet.html),
      // sheetProtection (#8) must come before autoFilter (#11) in the XML sequence.
      // When the order is wrong, Excel cannot open the file.
      const testFile = testFilePath("xml-element-order.test");
      const wb = new WorkbookWriter({
        filename: testFile,
        useStyles: true
      });
      const ws = wb.addWorksheet("test");

      // Add some data
      cellSetValue(ws.getCell("A1"), "Name");
      cellSetValue(ws.getCell("B1"), "Age");
      cellSetValue(ws.getCell("C1"), "City");
      rowCommit(ws.getRow(1));

      cellSetValue(ws.getCell("A2"), "John");
      cellSetValue(ws.getCell("B2"), 30);
      cellSetValue(ws.getCell("C2"), "NYC");
      rowCommit(ws.getRow(2));

      // Set autoFilter
      ws.autoFilter = { from: "A1", to: "C1" };

      // Set sheet protection
      await ws.protect("test", { formatColumns: true, formatRows: true, autoFilter: true });

      ws.commit();
      await wb.commit();

      // Read it back and verify it's valid
      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, testFile);

      const ws2 = Workbook.getWorksheet(wb2, "test")!;
      expect(ws2).toBeDefined();

      // Verify autoFilter is present and correct
      expect(ws2.autoFilter).toBe("A1:C1");

      // Verify protection is present with correct options
      expect(ws2.sheetProtection!.sheet).toBe(true);
      expect(ws2.sheetProtection!.formatColumns).toBe(true);
      expect(ws2.sheetProtection!.formatRows).toBe(true);
      expect(ws2.sheetProtection!.autoFilter).toBe(true);

      // Verify data is intact
      expect(Cell.getValue(ws2, "A1")).toBe("Name");
      expect(Cell.getValue(ws2, "B1")).toBe("Age");
      expect(Cell.getValue(ws2, "C1")).toBe("City");
      expect(Cell.getValue(ws2, "A2")).toBe("John");
      expect(Cell.getValue(ws2, "B2")).toBe(30);
      expect(Cell.getValue(ws2, "C2")).toBe("NYC");
    });

    it("protect should work with streaming workbook writer", async () => {
      const testFile = testFilePath("pr-1262.stream-writer.test");
      const workbook = new WorkbookWriter({
        filename: testFile
      });

      const sheet = workbook.addWorksheet("data");
      const row = sheet.addRow(["readonly cell"]);
      cellSetProtection(rowGetCell(row, 1), {
        locked: true
      });
      rowCommit(row);

      expect(sheet.protect).toBeDefined();

      await sheet.protect("password", {
        spinCount: 1
      });

      sheet.commit();
      await workbook.commit();

      const checkBook = Workbook.create();
      await Workbook.readFile(checkBook, testFile);
      const checkSheet = Workbook.getWorksheet(checkBook, "data")!;
      expect(checkSheet).toBeDefined();
      expect(checkSheet!.sheetProtection!.spinCount).toBe(1);
    });
  });
});
