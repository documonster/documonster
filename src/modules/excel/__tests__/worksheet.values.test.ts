import fs from "node:fs";

import {
  cellFont,
  cellGetValue,
  cellNote,
  cellNumFmt,
  cellSetNote,
  cellName,
  cellSetName
} from "@excel/cell";
import { columnHeaders } from "@excel/column";
import { Enums } from "@excel/enums";
import { Cell, Column, Workbook, Worksheet } from "@excel/index";
import { rowSetAlignment, rowSetBorder, rowValues } from "@excel/row";
import {
  columnEachCell,
  columnSetAlignment,
  columnSetBorder,
  columnSetDefn,
  columnSetHeader,
  columnSetKey,
  getCell,
  getColumn,
  getRow,
  rowEachCell,
  rowGetCell
} from "@excel/worksheet";
import { getUniqueTestFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Values", () => {
    it("stores values properly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      const now = new Date();

      // plain number
      Cell.setValue(ws, "A1", 7);

      // simple string
      Cell.setValue(ws, "B1", "Hello, World!");

      // floating point
      Cell.setValue(ws, "C1", 3.14);

      // 5 will be overwritten by the current date-time
      Cell.setValue(ws, "D1", 5);
      Cell.setValue(ws, "D1", now);

      // constructed string - will share recorded with B1
      Cell.setValue(ws, "E1", `${["Hello", "World"].join(", ")}!`);

      // hyperlink
      Cell.setValue(ws, "F1", {
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      });

      // number formula
      Cell.setValue(ws, "A2", { formula: "A1", result: 7 });

      // string formula
      Cell.setValue(ws, "B2", {
        formula: 'CONCATENATE("Hello", ", ", "World!")',
        result: "Hello, World!"
      });

      // date formula
      Cell.setValue(ws, "C2", { formula: "D1", result: now });

      expect(Cell.getValue(ws, "A1")).toBe(7);
      expect(Cell.getValue(ws, "B1")).toBe("Hello, World!");
      expect(Cell.getValue(ws, "C1")).toBe(3.14);
      expect(Cell.getValue(ws, "D1")).toBe(now);
      expect(Cell.getValue(ws, "E1")).toBe("Hello, World!");
      expect(Cell.getText(ws, "F1")).toBe("www.google.com");
      expect(Cell.getHyperlink(ws, "F1")).toBe("http://www.google.com");

      expect(Cell.getFormula(ws, "A2")).toBe("A1");
      expect(Cell.getResult(ws, "A2")).toBe(7);

      expect(Cell.getFormula(ws, "B2")).toBe('CONCATENATE("Hello", ", ", "World!")');
      expect(Cell.getResult(ws, "B2")).toBe("Hello, World!");

      expect(Cell.getFormula(ws, "C2")).toBe("D1");
      expect(Cell.getResult(ws, "C2")).toBe(now);
    });

    it("stores shared string values properly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Cell.setValue(ws, "A1", "Hello, World!");

      Cell.setValue(ws, "A2", "Hello");
      Cell.setValue(ws, "B2", "World");
      Cell.setValue(ws, "C2", {
        formula: 'CONCATENATE(A2, ", ", B2, "!")',
        result: "Hello, World!"
      });

      Cell.setValue(ws, "A3", `${["Hello", "World"].join(", ")}!`);

      // A1 and A3 should reference the same string object
      expect(Cell.getValue(ws, "A1")).toBe(Cell.getValue(ws, "A3"));

      // A1 and C2 should not reference the same object
      expect(Cell.getValue(ws, "A1")).toBe(Cell.getResult(ws, "C2"));
    });

    it("assigns cell types properly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // plain number
      Cell.setValue(ws, "A1", 7);

      // simple string
      Cell.setValue(ws, "B1", "Hello, World!");

      // floating point
      Cell.setValue(ws, "C1", 3.14);

      // date-time
      Cell.setValue(ws, "D1", new Date());

      // hyperlink
      Cell.setValue(ws, "E1", {
        text: "www.google.com",
        hyperlink: "http://www.google.com"
      });

      // number formula
      Cell.setValue(ws, "A2", { formula: "A1", result: 7 });

      // string formula
      Cell.setValue(ws, "B2", {
        formula: 'CONCATENATE("Hello", ", ", "World!")',
        result: "Hello, World!"
      });

      // date formula
      Cell.setValue(ws, "C2", { formula: "D1", result: new Date() });

      expect(Cell.getType(ws, "A1")).toBe(Enums.ValueType.Number);
      expect(Cell.getType(ws, "B1")).toBe(Enums.ValueType.String);
      expect(Cell.getType(ws, "C1")).toBe(Enums.ValueType.Number);
      expect(Cell.getType(ws, "D1")).toBe(Enums.ValueType.Date);
      expect(Cell.getType(ws, "E1")).toBe(Enums.ValueType.Hyperlink);

      expect(Cell.getType(ws, "A2")).toBe(Enums.ValueType.Formula);
      expect(Cell.getType(ws, "B2")).toBe(Enums.ValueType.Formula);
      expect(Cell.getType(ws, "C2")).toBe(Enums.ValueType.Formula);
    });

    it("adds columns", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Worksheet.setColumns(ws, [
        { key: "id", width: 10 },
        { key: "name", width: 32 },
        { key: "dob", width: 10 }
      ]);

      expect(getColumn(ws, "id").number).toBe(1);
      expect(Column.getWidth(ws, "id")).toBe(10);
      expect(getColumn(ws, "A")).toBe(getColumn(ws, "id"));
      expect(getColumn(ws, 1)).toBe(getColumn(ws, "id"));

      expect(getColumn(ws, "name").number).toBe(2);
      expect(Column.getWidth(ws, "name")).toBe(32);
      expect(getColumn(ws, "B")).toBe(getColumn(ws, "name"));
      expect(getColumn(ws, 2)).toBe(getColumn(ws, "name"));

      expect(getColumn(ws, "dob").number).toBe(3);
      expect(Column.getWidth(ws, "dob")).toBe(10);
      expect(getColumn(ws, "C")).toBe(getColumn(ws, "dob"));
      expect(getColumn(ws, 3)).toBe(getColumn(ws, "dob"));
    });

    it("adds column headers", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Worksheet.setColumns(ws, [
        { header: "Id", width: 10 },
        { header: "Name", width: 32 },
        { header: "D.O.B.", width: 10 }
      ]);

      expect(Cell.getValue(ws, "A1")).toBe("Id");
      expect(Cell.getValue(ws, "B1")).toBe("Name");
      expect(Cell.getValue(ws, "C1")).toBe("D.O.B.");
    });

    describe("date column headers", () => {
      it("supports Date type as column header", async () => {
        const workbook = Workbook.create();
        const ws = Workbook.addWorksheet(workbook, "Sheet1");
        const dateValue = new Date("2024-02-02");

        Worksheet.setColumns(ws, [
          { header: dateValue, key: "date1", width: 15, style: { numFmt: "yyyy/mm/dd" } },
          { header: "Name", key: "name", width: 20 }
        ]);

        Worksheet.addRow(ws, { date1: new Date("2024-03-15"), name: "Test 1" });

        const headerCell = getCell(ws, "A1");
        expect(cellGetValue(headerCell)).toBeInstanceOf(Date);
        expect((cellGetValue(headerCell) as Date).toISOString()).toBe("2024-02-02T00:00:00.000Z");
        expect(cellNumFmt(headerCell)).toBe("yyyy/mm/dd");

        const buffer = await Workbook.toBuffer(workbook);
        const wb2 = Workbook.create();
        await Workbook.read(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;
        const headerCell2 = getCell(ws2!, "A1");
        expect(cellGetValue(headerCell2)).toBeInstanceOf(Date);
      });

      it("supports number type as column header", () => {
        const workbook = Workbook.create();
        const ws = Workbook.addWorksheet(workbook, "Sheet1");

        Worksheet.setColumns(ws, [
          { header: 12345, key: "num", width: 15 },
          { header: "Name", key: "name", width: 20 }
        ]);

        const headerCell = getCell(ws, "A1");
        expect(cellGetValue(headerCell)).toBe(12345);
        expect(typeof cellGetValue(headerCell)).toBe("number");
      });

      it("supports mixed types in multi-row headers", () => {
        const workbook = Workbook.create();
        const ws = Workbook.addWorksheet(workbook, "Sheet1");
        const date1 = new Date("2024-01-01");
        const date2 = new Date("2024-01-31");

        Worksheet.setColumns(ws, [
          { header: [date1, "January"], key: "jan", width: 15 },
          { header: ["Q1", date2], key: "q1", width: 15 }
        ]);

        expect(Cell.getValue(ws, "A1")).toBeInstanceOf(Date);
        expect(Cell.getValue(ws, "A2")).toBe("January");
        expect(Cell.getValue(ws, "B1")).toBe("Q1");
        expect(Cell.getValue(ws, "B2")).toBeInstanceOf(Date);
      });

      it("applies column style to date headers", () => {
        const workbook = Workbook.create();
        const ws = Workbook.addWorksheet(workbook, "Sheet1");

        Worksheet.setColumns(ws, [
          {
            header: new Date("2024-06-15"),
            key: "date",
            width: 15,
            style: {
              numFmt: "dd-mmm-yyyy",
              font: { bold: true }
            }
          }
        ]);

        const headerCell = getCell(ws, "A1");
        expect(cellGetValue(headerCell)).toBeInstanceOf(Date);
        expect(cellNumFmt(headerCell)).toBe("dd-mmm-yyyy");
        expect(cellFont(headerCell)?.bold).toBe(true);
      });
    });

    it("adds column headers by number", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // by defn
      columnSetDefn(getColumn(ws, 1), { key: "id", header: "Id", width: 10 });

      // by property
      columnSetKey(getColumn(ws, 2), "name");
      columnSetHeader(getColumn(ws, 2), "Name");
      Column.setWidth(ws, 2, 32);

      expect(Cell.getValue(ws, "A1")).toBe("Id");
      expect(Cell.getValue(ws, "B1")).toBe("Name");

      expect(Column.getKey(ws, "A")).toBe("id");
      expect(Column.getKey(ws, 1)).toBe("id");
      expect(Column.getHeader(ws, 1)).toBe("Id");
      expect(columnHeaders(getColumn(ws, 1))).toEqual(["Id"]);
      expect(Column.getWidth(ws, 1)).toBe(10);

      expect(Column.getKey(ws, 2)).toBe("name");
      expect(Column.getHeader(ws, 2)).toBe("Name");
      expect(columnHeaders(getColumn(ws, 2))).toEqual(["Name"]);
      expect(Column.getWidth(ws, 2)).toBe(32);
    });

    it("adds column headers by letter", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // by defn
      columnSetDefn(getColumn(ws, "A"), { key: "id", header: "Id", width: 10 });

      // by property
      columnSetKey(getColumn(ws, "B"), "name");
      columnSetHeader(getColumn(ws, "B"), "Name");
      Column.setWidth(ws, "B", 32);

      expect(Cell.getValue(ws, "A1")).toBe("Id");
      expect(Cell.getValue(ws, "B1")).toBe("Name");

      expect(Column.getKey(ws, "A")).toBe("id");
      expect(Column.getKey(ws, 1)).toBe("id");
      expect(Column.getHeader(ws, "A")).toBe("Id");
      expect(columnHeaders(getColumn(ws, "A"))).toEqual(["Id"]);
      expect(Column.getWidth(ws, "A")).toBe(10);

      expect(Column.getKey(ws, "B")).toBe("name");
      expect(Column.getHeader(ws, "B")).toBe("Name");
      expect(columnHeaders(getColumn(ws, "B"))).toEqual(["Name"]);
      expect(Column.getWidth(ws, "B")).toBe(32);
    });

    it("adds rows by object", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      // add columns to define column keys
      Worksheet.setColumns(ws, [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 32 },
        { header: "D.O.B.", key: "dob", width: 10 }
      ]);

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);

      Worksheet.addRow(ws, { id: 1, name: "John Doe", dob: dateValue1 });
      Worksheet.addRow(ws, { id: 2, name: "Jane Doe", dob: dateValue2 });

      expect(Cell.getValue(ws, "A2")).toBe(1);
      expect(Cell.getValue(ws, "B2")).toBe("John Doe");
      expect(Cell.getValue(ws, "C2")).toBe(dateValue1);

      expect(Cell.getValue(ws, "A3")).toBe(2);
      expect(Cell.getValue(ws, "B3")).toBe("Jane Doe");
      expect(Cell.getValue(ws, "C3")).toBe(dateValue2);

      expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, 1, "John Doe", dateValue1]);
      expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, 2, "Jane Doe", dateValue2]);

      const values = [
        ,
        [, "Id", "Name", "D.O.B."],
        [, 1, "John Doe", dateValue1],
        [, 2, "Jane Doe", dateValue2]
      ];
      Worksheet.eachRow(ws, (row, rowNumber) => {
        expect(rowValues(row)).toEqual(values[rowNumber]);
        rowEachCell(row, (cell: any, colNumber: any) => {
          expect(cellGetValue(cell)).toBe(values[rowNumber]![colNumber]);
        });
      });
    });

    it("adds rows by contiguous array", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);

      Worksheet.addRow(ws, [1, "John Doe", dateValue1]);
      Worksheet.addRow(ws, [2, "Jane Doe", dateValue2]);

      expect(Cell.getValue(ws, "A1")).toBe(1);
      expect(Cell.getValue(ws, "B1")).toBe("John Doe");
      expect(Cell.getValue(ws, "C1")).toBe(dateValue1);

      expect(Cell.getValue(ws, "A2")).toBe(2);
      expect(Cell.getValue(ws, "B2")).toBe("Jane Doe");
      expect(Cell.getValue(ws, "C2")).toBe(dateValue2);

      expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, 1, "John Doe", dateValue1]);
      expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, 2, "Jane Doe", dateValue2]);
    });

    it("adds rows by sparse array", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const rows = [, [, 1, "John Doe", , dateValue1], [, 2, "Jane Doe", , dateValue2]];
      const row3: (number | string | Date | undefined)[] = [];
      row3[1] = 3;
      row3[3] = "Sam";
      row3[5] = dateValue1;
      rows.push(row3);
      rows.forEach(row => {
        if (row) {
          Worksheet.addRow(ws, row);
        }
      });

      expect(Cell.getValue(ws, "A1")).toBe(1);
      expect(Cell.getValue(ws, "B1")).toBe("John Doe");
      expect(Cell.getValue(ws, "D1")).toBe(dateValue1);

      expect(Cell.getValue(ws, "A2")).toBe(2);
      expect(Cell.getValue(ws, "B2")).toBe("Jane Doe");
      expect(Cell.getValue(ws, "D2")).toBe(dateValue2);

      expect(Cell.getValue(ws, "A3")).toBe(3);
      expect(Cell.getValue(ws, "C3")).toBe("Sam");
      expect(Cell.getValue(ws, "E3")).toBe(dateValue1);

      expect(rowValues(Worksheet.getRow(ws, 1))).toEqual(rows[1]);
      expect(rowValues(Worksheet.getRow(ws, 2))).toEqual(rows[2]);
      expect(rowValues(Worksheet.getRow(ws, 3))).toEqual(rows[3]);

      Worksheet.eachRow(ws, (row, rowNumber) => {
        expect(rowValues(row)).toEqual(rows[rowNumber]);
        rowEachCell(row, (cell: any, colNumber: any) => {
          expect(cellGetValue(cell)).toBe(rows[rowNumber]![colNumber]);
        });
      });
    });

    it("supports sparse array rows beyond declared columns with XLSX round-trip", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "ExampleWS");

      Worksheet.setColumns(worksheet, [
        { header: "Id", key: "id", width: 10 },
        { header: "Name", key: "name", width: 32 },
        { header: "D.O.B.", key: "dob", width: 10 }
      ]);

      const sparse = [] as any[];
      sparse[1] = 4;
      sparse[5] = "Kyle";
      sparse[9] = new Date("2020-01-02T00:00:00.000Z");

      Worksheet.addRow(worksheet, sparse);

      Worksheet.addRows(worksheet, [
        [5, "Bob", new Date("2020-01-03T00:00:00.000Z")],
        { id: 6, name: "Barbara", dob: new Date("2020-01-04T00:00:00.000Z") }
      ]);

      // worksheet.columns creates a header row at row 1
      expect(Cell.getValue(worksheet, 2, 1)).toBe(4);
      expect(Cell.getValue(worksheet, 2, 5)).toBe("Kyle");
      expect(Cell.getValue(worksheet, 3, 1)).toBe(5);
      expect(Cell.getValue(worksheet, 4, 1)).toBe(6);

      const filename = getUniqueTestFilePath(import.meta.url);
      await Workbook.writeFile(workbook, filename);
      expect(fs.existsSync(filename)).toBe(true);

      const readBack = Workbook.create();
      await Workbook.readFile(readBack, filename);

      const ws2 = Workbook.getWorksheet(readBack, "ExampleWS")!;
      expect(ws2).toBeTruthy();

      expect(Cell.getValue(ws2, 2, 1)).toBe(4);
      expect(Cell.getValue(ws2, 2, 5)).toBe("Kyle");

      const v = Cell.getValue(ws2, 2, 9);
      expect(v).not.toBeNull();
      expect(v instanceof Date || typeof v === "number").toBe(true);

      expect(Cell.getValue(ws2, 3, 1)).toBe(5);
      expect(Cell.getValue(ws2, 3, 2)).toBe("Bob");

      expect(Cell.getValue(ws2, 4, 1)).toBe(6);
      expect(Cell.getValue(ws2, 4, 2)).toBe("Barbara");
    });

    describe("Splice", () => {
      describe("Rows", () => {
        it("Remove only", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-remove-only");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

          Worksheet.spliceRows(ws, 2, 1);

          expect(ws).not.toBeUndefined();
          expect(Cell.getValue(ws, "A1")).toBe("1,1");
          expect(Cell.getType(ws, "A1")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "B1")).toBe("1,2");
          expect(Cell.getType(ws, "B1")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "C1")).toBe("1,3");
          expect(Cell.getType(ws, "C1")).toBe(Enums.ValueType.String);

          expect(Cell.getType(ws, "A2")).toBe(Enums.ValueType.Null);
          expect(Cell.getType(ws, "B2")).toBe(Enums.ValueType.Null);
          expect(Cell.getType(ws, "C2")).toBe(Enums.ValueType.Null);

          expect(Cell.getValue(ws, "A3")).toBe(4.1);
          expect(Cell.getType(ws, "A3")).toBe(Enums.ValueType.Number);
          expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Null);
          expect(Cell.getValue(ws, "C3")).toBe(4.3);
          expect(Cell.getType(ws, "C3")).toBe(Enums.ValueType.Number);

          expect(Cell.getValue(ws, "A4")).toBe("5,1");
          expect(Cell.getType(ws, "A4")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "B4")).toBe("5,2");
          expect(Cell.getType(ws, "B4")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "C4")).toBe("5,3");
          expect(Cell.getType(ws, "C4")).toBe(Enums.ValueType.String);

          Worksheet.addRow(ws, ["5,1b", "5,2b", "5,3b"]);
          expect(Cell.getValue(ws, "A5")).toBe("5,1b");
          expect(Cell.getType(ws, "A5")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "B5")).toBe("5,2b");
          expect(Cell.getType(ws, "B5")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "C5")).toBe("5,3b");
          expect(Cell.getType(ws, "C5")).toBe(Enums.ValueType.String);
        });

        it("spliceRows remove last row", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-rows-remove-last");
          Worksheet.addRows(ws, [["1st"], ["2nd"], ["3rd"]]);

          Worksheet.spliceRows(ws, Worksheet.rowCount(ws), 1);

          expect(cellGetValue(rowGetCell(Worksheet.getRow(ws, Worksheet.rowCount(ws)), 1))).toBe(
            "2nd"
          );
        });
        it("Remove and insert fewer", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-insert-fewer");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

          Worksheet.spliceRows(ws, 2, 2, ["one", "two", "three"]);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, 4.1, , 4.3]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, "5,1", "5,2", "5,3"]);
        });
        it("Remove and insert same", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-insert-same");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

          Worksheet.spliceRows(ws, 2, 2, ["one", "two", "three"], ["une", "deux", "trois"]);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "une", "deux", "trois"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, 4.1, , 4.3]);
          expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([, "5,1", "5,2", "5,3"]);
        });
        it("Remove and insert more", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-insert-more");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3"]);

          Worksheet.spliceRows(
            ws,
            2,
            2,
            ["one", "two", "three"],
            ["une", "deux", "trois"],
            ["uno", "due", "tre"]
          );

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "une", "deux", "trois"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, "uno", "due", "tre"]);
          expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([, 4.1, , 4.3]);
          expect(rowValues(Worksheet.getRow(ws, 6))).toEqual([, "5,1", "5,2", "5,3"]);
        });
        it("Remove style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-remove-style");
          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(ws, ["4,1", "4,2", "4,3", "4,4"]);

          Cell.setStyle(ws, "A1", { numFmt: "# ?/?" });
          Cell.setStyle(ws, "B2", {
            fill: {
              type: "pattern",
              pattern: "darkVertical",
              fgColor: { argb: "FFFF0000" }
            }
          });
          rowSetBorder(Worksheet.getRow(ws, 3), {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          });
          rowSetAlignment(Worksheet.getRow(ws, 4), { horizontal: "left", vertical: "middle" });

          Worksheet.spliceRows(ws, 2, 2);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "4,1", "4,2", "4,3", "4,4"]);
          expect(Cell.getStyle(ws, "A1")).toEqual({ numFmt: "# ?/?" });
          expect(getRow(ws, 2).style).toEqual({
            alignment: { horizontal: "left", vertical: "middle" }
          });
        });
        it("Insert style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-insert-style");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Cell.setStyle(ws, "A2", {
            fill: {
              type: "pattern",
              pattern: "darkVertical",
              fgColor: { argb: "FFFF0000" }
            }
          });
          rowSetAlignment(Worksheet.getRow(ws, 2), { horizontal: "left", vertical: "middle" });

          Worksheet.spliceRows(ws, 2, 0, ["one", "two", "three"]);
          Cell.setStyle(ws, "A2", {
            border: {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" }
            }
          });

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "one", "two", "three"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "2,1", "2,2", "2,3"]);
          expect(getRow(ws, 3).style.alignment).toEqual({ horizontal: "left", vertical: "middle" });
          expect(Cell.getStyle(ws, "A2").border).toEqual({
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          });
          expect(Cell.getStyle(ws, "A3").alignment).toEqual({
            horizontal: "left",
            vertical: "middle"
          });
          expect(Cell.getStyle(ws, "A3").fill).toEqual({
            type: "pattern",
            pattern: "darkVertical",
            fgColor: { argb: "FFFF0000" }
          });
        });
        it("Replace style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-row-replace-style");
          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);

          Cell.setStyle(ws, "B1", { numFmt: "top" });
          Cell.setStyle(ws, "B2", { numFmt: "middle" });
          Cell.setStyle(ws, "B3", { numFmt: "bottom" });
          rowSetAlignment(Worksheet.getRow(ws, 1), { horizontal: "left", vertical: "top" });
          rowSetAlignment(Worksheet.getRow(ws, 2), { horizontal: "center", vertical: "middle" });
          rowSetAlignment(Worksheet.getRow(ws, 3), { horizontal: "right", vertical: "bottom" });

          Worksheet.spliceRows(ws, 2, 1, ["two-one", "two-two", "two-three", "two-four"]);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,2", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([
            ,
            "two-one",
            "two-two",
            "two-three",
            "two-four"
          ]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "3,1", "3,2", "3,3", "3,4"]);
          expect(Cell.getStyle(ws, "B1")).toEqual({
            numFmt: "top",
            alignment: { horizontal: "left", vertical: "top" }
          });
          expect(Cell.getStyle(ws, "B2")).toEqual({});
          expect(Cell.getStyle(ws, "B3")).toEqual({
            numFmt: "bottom",
            alignment: { horizontal: "right", vertical: "bottom" }
          });
          expect(getRow(ws, 1).style).toEqual({
            alignment: { horizontal: "left", vertical: "top" }
          });
          expect(getRow(ws, 2).style).toEqual({});
          expect(getRow(ws, 3).style).toEqual({
            alignment: { horizontal: "right", vertical: "bottom" }
          });
        });
        it("Remove defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-row-remove-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceRows(wsSquare, 2, 2);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([, "1,1", "1,2", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([, "4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3].forEach(row => {
              if (row === 3) {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          const wsSingles = Workbook.addWorksheet(wb, "splice-row-remove-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceRows(wsSingles, 2, 2);

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 2))).toEqual([, "4,1", , , "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A2"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "D2"))).toBe("single-D4");
        });
        it("Insert defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-row-insert-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceRows(wsSquare, 3, 0, ["foo", "bar", "baz", "qux"]);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([, "1,1", "1,2", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([, "2,1", "2,2", "2,3", "2,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 3))).toEqual([, "foo", "bar", "baz", "qux"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 4))).toEqual([, "3,1", "3,2", "3,3", "3,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 5))).toEqual([, "4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4, 5].forEach(row => {
              if (row === 3) {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          const wsSingles = Workbook.addWorksheet(wb, "splice-row-insert-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceRows(wsSingles, 3, 0, ["foo", "bar", "baz", "qux"]);

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 3))).toEqual([, "foo", "bar", "baz", "qux"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 5))).toEqual([, "4,1", , , "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A5"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "D5"))).toBe("single-D4");
        });
        it("Replace defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-row-replace-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceRows(wsSquare, 2, 1, ["foo", "bar", "baz", "qux"]);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([, "1,1", "1,2", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([, "foo", "bar", "baz", "qux"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 3))).toEqual([, "3,1", "3,2", "3,3", "3,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 4))).toEqual([, "4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              if (row === 2) {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          const wsSingles = Workbook.addWorksheet(wb, "splice-row-replace-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceRows(wsSingles, 2, 1, ["foo", "bar", "baz", "qux"]);

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , , "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 2))).toEqual([, "foo", "bar", "baz", "qux"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 4))).toEqual([, "4,1", , , "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A4"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "D4"))).toBe("single-D4");
        });
      });
      describe("Columns", () => {
        it("splices columns", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-column-remove-only");

          Worksheet.setColumns(ws, [
            { key: "id", width: 10 },
            { key: "name", width: 32 },
            { key: "dob", width: 10 }
          ]);
          Worksheet.addRow(ws, { id: "id1", name: "name1", dob: "dob1" });
          Worksheet.addRow(ws, { id: 2, dob: "dob2" });
          Worksheet.addRow(ws, { name: "name3", dob: 3 });

          Worksheet.spliceColumns(ws, 2, 1);

          expect(ws).not.toBeUndefined();
          expect(Cell.getValue(ws, "A1")).toBe("id1");
          expect(Cell.getType(ws, "A1")).toBe(Enums.ValueType.String);
          expect(Cell.getValue(ws, "B1")).toBe("dob1");
          expect(Cell.getType(ws, "B1")).toBe(Enums.ValueType.String);
          expect(Cell.getType(ws, "C1")).toBe(Enums.ValueType.Null);

          expect(Cell.getValue(ws, "A2")).toBe(2);
          expect(Cell.getType(ws, "A2")).toBe(Enums.ValueType.Number);
          expect(Cell.getValue(ws, "B2")).toBe("dob2");
          expect(Cell.getType(ws, "B2")).toBe(Enums.ValueType.String);
          expect(Cell.getType(ws, "C2")).toBe(Enums.ValueType.Null);

          expect(Cell.getType(ws, "A3")).toBe(Enums.ValueType.Null);
          expect(Cell.getValue(ws, "B3")).toBe(3);
          expect(Cell.getType(ws, "B3")).toBe(Enums.ValueType.Number);
          expect(Cell.getType(ws, "C3")).toBe(Enums.ValueType.Null);
        });

        it("comments shift correctly on spliceColumns", async () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "testSheet");

          Worksheet.addRow(ws, [
            "test1",
            "test2",
            "test3",
            "test4",
            "test5",
            "test6",
            "test7",
            "test8"
          ]);

          const row = Worksheet.getRow(ws, 1);
          cellSetNote(rowGetCell(row, 1), "test1");
          cellSetNote(rowGetCell(row, 2), "test2");
          cellSetNote(rowGetCell(row, 3), "test3");
          cellSetNote(rowGetCell(row, 4), "test4");

          Worksheet.spliceColumns(ws, 2, 1);

          expect(cellNote(rowGetCell(row, 1))).toBe("test1");
          expect(cellNote(rowGetCell(row, 2))).toBe("test3");
          expect(cellNote(rowGetCell(row, 3))).toBe("test4");
          expect(cellNote(rowGetCell(row, 4))).toBe(undefined);

          const buffer = await Workbook.toBuffer(wb);
          expect(buffer.byteLength).toBeGreaterThan(0);
        });
        it("Remove and insert fewer", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-column-insert-fewer");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Cell.setValue(ws, "E4", 4.5);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

          Worksheet.spliceColumns(ws, 2, 2, ["one", "two", "three", "four", "five"]);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "one", "1,4", "1,5"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "2,1", "two", "2,4", "2,5"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, , "three"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, 4.1, "four", , 4.5]);
          expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([, "5,1", "five", "5,4", "5,5"]);
        });
        it("Remove and insert same", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-column-insert-same");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Cell.setValue(ws, "E4", 4.5);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

          Worksheet.spliceColumns(
            ws,
            2,
            2,
            ["one", "two", "three", "four", "five"],
            ["une", "deux", "trois", "quatre", "cinq"]
          );

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "one", "une", "1,4", "1,5"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([
            ,
            "2,1",
            "two",
            "deux",
            "2,4",
            "2,5"
          ]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, , "three", "trois"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, 4.1, "four", "quatre", , 4.5]);
          expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([
            ,
            "5,1",
            "five",
            "cinq",
            "5,4",
            "5,5"
          ]);
        });
        it("Remove and insert more", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-column-insert-more");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4", "1,5"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4", "2,5"]);
          Cell.setValue(ws, "A4", 4.1);
          Cell.setValue(ws, "C4", 4.3);
          Cell.setValue(ws, "E4", 4.5);
          Worksheet.addRow(ws, ["5,1", "5,2", "5,3", "5,4", "5,5"]);

          Worksheet.spliceColumns(
            ws,
            2,
            2,
            ["one", "two", "three", "four", "five"],
            ["une", "deux", "trois", "quatre", "cinq"],
            ["uno", "due", "tre", "quatro", "cinque"]
          );

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([
            ,
            "1,1",
            "one",
            "une",
            "uno",
            "1,4",
            "1,5"
          ]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([
            ,
            "2,1",
            "two",
            "deux",
            "due",
            "2,4",
            "2,5"
          ]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, , "three", "trois", "tre"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([
            ,
            4.1,
            "four",
            "quatre",
            "quatro",
            ,
            4.5
          ]);
          expect(rowValues(Worksheet.getRow(ws, 5))).toEqual([
            ,
            "5,1",
            "five",
            "cinq",
            "cinque",
            "5,4",
            "5,5"
          ]);
        });
        it("handles column keys", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-column-insert-fewer");
          Worksheet.setColumns(ws, [
            { key: "id", width: 10 },
            { key: "dob", width: 20 },
            { key: "name", width: 30 },
            { key: "age", width: 40 }
          ]);

          const values = [
            { id: "123", name: "Jack", dob: new Date(), age: 0 },
            { id: "124", name: "Jill", dob: new Date(), age: 0 }
          ];
          values.forEach(value => {
            Worksheet.addRow(ws, value);
          });

          Worksheet.spliceColumns(ws, 2, 1, ["B1", "B2"], ["C1", "C2"]);

          values.forEach((rowValues, index) => {
            const row = Worksheet.getRow(ws, index + 1);
            Object.entries(rowValues).forEach(([key, value]) => {
              if (key !== "dob") {
                expect(cellGetValue(rowGetCell(row, key))).toBe(value);
              }
            });
          });

          expect(Column.getWidth(ws, 1)).toBe(10);
          expect(Column.getWidth(ws, 2)).toBeUndefined();
          expect(Column.getWidth(ws, 3)).toBeUndefined();
          expect(Column.getWidth(ws, 4)).toBe(30);
          expect(Column.getWidth(ws, 5)).toBe(40);
        });

        it("Splices to end", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-to-end");
          Worksheet.setColumns(ws, [
            { header: "Col-1", width: 10 },
            { header: "Col-2", width: 10 },
            { header: "Col-3", width: 10 },
            { header: "Col-4", width: 10 },
            { header: "Col-5", width: 10 },
            { header: "Col-6", width: 10 }
          ]);

          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);
          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);

          // splice last 3 columns
          Worksheet.spliceColumns(ws, 4, 3);
          expect(Cell.getValue(ws, 1, 1)).toBe("Col-1");
          expect(Cell.getValue(ws, 1, 2)).toBe("Col-2");
          expect(Cell.getValue(ws, 1, 3)).toBe("Col-3");
          expect(Cell.getValue(ws, 1, 4)).toBeNull();
          expect(Cell.getValue(ws, 1, 5)).toBeNull();
          expect(Cell.getValue(ws, 1, 6)).toBeNull();
          expect(Cell.getValue(ws, 1, 7)).toBeNull();
          expect(Cell.getValue(ws, 2, 1)).toBe(1);
          expect(Cell.getValue(ws, 2, 2)).toBe(2);
          expect(Cell.getValue(ws, 2, 3)).toBe(3);
          expect(Cell.getValue(ws, 2, 4)).toBeNull();
          expect(Cell.getValue(ws, 2, 5)).toBeNull();
          expect(Cell.getValue(ws, 2, 6)).toBeNull();
          expect(Cell.getValue(ws, 2, 7)).toBeNull();
          expect(Cell.getValue(ws, 3, 1)).toBe(1);
          expect(Cell.getValue(ws, 3, 2)).toBe(2);
          expect(Cell.getValue(ws, 3, 3)).toBe(3);
          expect(Cell.getValue(ws, 3, 4)).toBeNull();
          expect(Cell.getValue(ws, 3, 5)).toBeNull();
          expect(Cell.getValue(ws, 3, 6)).toBeNull();
          expect(Cell.getValue(ws, 3, 7)).toBeNull();

          expect(Column.getHeader(ws, 1)).toBe("Col-1");
          expect(Column.getHeader(ws, 2)).toBe("Col-2");
          expect(Column.getHeader(ws, 3)).toBe("Col-3");
          expect(Column.getHeader(ws, 4)).toBeUndefined();
          expect(Column.getHeader(ws, 5)).toBeUndefined();
          expect(Column.getHeader(ws, 6)).toBeUndefined();
        });
        it("Splices past end", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-to-end");
          Worksheet.setColumns(ws, [
            { header: "Col-1", width: 10 },
            { header: "Col-2", width: 10 },
            { header: "Col-3", width: 10 },
            { header: "Col-4", width: 10 },
            { header: "Col-5", width: 10 },
            { header: "Col-6", width: 10 }
          ]);

          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);
          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);

          // splice last 3 columns
          Worksheet.spliceColumns(ws, 4, 4);
          expect(Cell.getValue(ws, 1, 1)).toBe("Col-1");
          expect(Cell.getValue(ws, 1, 2)).toBe("Col-2");
          expect(Cell.getValue(ws, 1, 3)).toBe("Col-3");
          expect(Cell.getValue(ws, 1, 4)).toBeNull();
          expect(Cell.getValue(ws, 1, 5)).toBeNull();
          expect(Cell.getValue(ws, 1, 6)).toBeNull();
          expect(Cell.getValue(ws, 1, 7)).toBeNull();
          expect(Cell.getValue(ws, 2, 1)).toBe(1);
          expect(Cell.getValue(ws, 2, 2)).toBe(2);
          expect(Cell.getValue(ws, 2, 3)).toBe(3);
          expect(Cell.getValue(ws, 2, 4)).toBeNull();
          expect(Cell.getValue(ws, 2, 5)).toBeNull();
          expect(Cell.getValue(ws, 2, 6)).toBeNull();
          expect(Cell.getValue(ws, 2, 7)).toBeNull();
          expect(Cell.getValue(ws, 3, 1)).toBe(1);
          expect(Cell.getValue(ws, 3, 2)).toBe(2);
          expect(Cell.getValue(ws, 3, 3)).toBe(3);
          expect(Cell.getValue(ws, 3, 4)).toBeNull();
          expect(Cell.getValue(ws, 3, 5)).toBeNull();
          expect(Cell.getValue(ws, 3, 6)).toBeNull();
          expect(Cell.getValue(ws, 3, 7)).toBeNull();

          expect(Column.getHeader(ws, 1)).toBe("Col-1");
          expect(Column.getHeader(ws, 2)).toBe("Col-2");
          expect(Column.getHeader(ws, 3)).toBe("Col-3");
          expect(Column.getHeader(ws, 4)).toBeUndefined();
          expect(Column.getHeader(ws, 5)).toBeUndefined();
          expect(Column.getHeader(ws, 6)).toBeUndefined();
        });
        it("Splices almost to end", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-to-end");
          Worksheet.setColumns(ws, [
            { header: "Col-1", width: 10 },
            { header: "Col-2", width: 10 },
            { header: "Col-3", width: 10 },
            { header: "Col-4", width: 10 },
            { header: "Col-5", width: 10 },
            { header: "Col-6", width: 10 }
          ]);

          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);
          Worksheet.addRow(ws, [1, 2, 3, 4, 5, 6]);

          // splice last 3 columns
          Worksheet.spliceColumns(ws, 4, 2);
          expect(Cell.getValue(ws, 1, 1)).toBe("Col-1");
          expect(Cell.getValue(ws, 1, 2)).toBe("Col-2");
          expect(Cell.getValue(ws, 1, 3)).toBe("Col-3");
          expect(Cell.getValue(ws, 1, 4)).toBe("Col-6");
          expect(Cell.getValue(ws, 1, 5)).toBeNull();
          expect(Cell.getValue(ws, 1, 6)).toBeNull();
          expect(Cell.getValue(ws, 1, 7)).toBeNull();
          expect(Cell.getValue(ws, 2, 1)).toBe(1);
          expect(Cell.getValue(ws, 2, 2)).toBe(2);
          expect(Cell.getValue(ws, 2, 3)).toBe(3);
          expect(Cell.getValue(ws, 2, 4)).toBe(6);
          expect(Cell.getValue(ws, 2, 5)).toBeNull();
          expect(Cell.getValue(ws, 2, 6)).toBeNull();
          expect(Cell.getValue(ws, 2, 7)).toBeNull();
          expect(Cell.getValue(ws, 3, 1)).toBe(1);
          expect(Cell.getValue(ws, 3, 2)).toBe(2);
          expect(Cell.getValue(ws, 3, 3)).toBe(3);
          expect(Cell.getValue(ws, 3, 4)).toBe(6);
          expect(Cell.getValue(ws, 3, 5)).toBeNull();
          expect(Cell.getValue(ws, 3, 6)).toBeNull();
          expect(Cell.getValue(ws, 3, 7)).toBeNull();

          expect(Column.getHeader(ws, 1)).toBe("Col-1");
          expect(Column.getHeader(ws, 2)).toBe("Col-2");
          expect(Column.getHeader(ws, 3)).toBe("Col-3");
          expect(Column.getHeader(ws, 4)).toBe("Col-6");
          expect(Column.getHeader(ws, 5)).toBeUndefined();
          expect(Column.getHeader(ws, 6)).toBeUndefined();
        });

        it("Remove style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-col-remove-style");
          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(ws, ["4,1", "4,2", "4,3", "4,4"]);

          Cell.setStyle(ws, "A1", { numFmt: "# ?/?" });
          Cell.setStyle(ws, "B2", {
            fill: {
              type: "pattern",
              pattern: "darkVertical",
              fgColor: { argb: "FFFF0000" }
            }
          });
          columnSetBorder(getColumn(ws, 3), {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          });
          columnSetAlignment(getColumn(ws, 4), { horizontal: "left", vertical: "middle" });

          Worksheet.spliceColumns(ws, 2, 2);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "1,4"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "2,1", "2,4"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "3,1", "3,4"]);
          expect(rowValues(Worksheet.getRow(ws, 4))).toEqual([, "4,1", "4,4"]);
          expect(Cell.getStyle(ws, "A1")).toEqual({ numFmt: "# ?/?" });
          expect(getColumn(ws, 2).style).toEqual({
            alignment: { horizontal: "left", vertical: "middle" }
          });
          expect(Cell.getStyle(ws, "B4")).toEqual({
            alignment: { horizontal: "left", vertical: "middle" }
          });
        });
        it("Insert style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-col-insert-style");

          Worksheet.addRow(ws, ["1,1", "1,2", "1,3"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3"]);
          Worksheet.addRow(ws, ["3,1", "3,2", "3,3"]);
          Cell.setStyle(ws, "B2", {
            fill: {
              type: "pattern",
              pattern: "darkVertical",
              fgColor: { argb: "FFFF0000" }
            }
          });
          columnSetAlignment(getColumn(ws, 2), { horizontal: "left", vertical: "middle" });

          Worksheet.spliceColumns(ws, 2, 0, ["one", "two", "three"]);
          Cell.setStyle(ws, "B2", {
            border: {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" }
            }
          });

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "one", "1,2", "1,3"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "2,1", "two", "2,2", "2,3"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "3,1", "three", "3,2", "3,3"]);
          expect(getColumn(ws, 3).style).toEqual({
            alignment: { horizontal: "left", vertical: "middle" }
          });
          expect(Cell.getStyle(ws, "B2")).toEqual({
            border: {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" }
            }
          });
          expect(Cell.getStyle(ws, "C2")).toEqual({
            alignment: { horizontal: "left", vertical: "middle" },
            fill: { type: "pattern", pattern: "darkVertical", fgColor: { argb: "FFFF0000" } }
          });
        });
        it("Replace style", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "splice-col-replace-style");
          Worksheet.addRow(ws, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(ws, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(ws, ["3,1", "3,2", "3,3", "3,4"]);

          Cell.setStyle(ws, "A2", { numFmt: "left" });
          Cell.setStyle(ws, "B2", { numFmt: "center" });
          Cell.setStyle(ws, "C2", { numFmt: "right" });
          columnSetAlignment(getColumn(ws, 1), { horizontal: "left", vertical: "top" });
          columnSetAlignment(getColumn(ws, 2), { horizontal: "center", vertical: "middle" });
          columnSetAlignment(getColumn(ws, 3), { horizontal: "right", vertical: "bottom" });

          Worksheet.spliceColumns(ws, 2, 1, ["one-two", "two-two", "three-two"]);

          expect(ws).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(ws, 1))).toEqual([, "1,1", "one-two", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(ws, 2))).toEqual([, "2,1", "two-two", "2,3", "2,4"]);
          expect(rowValues(Worksheet.getRow(ws, 3))).toEqual([, "3,1", "three-two", "3,3", "3,4"]);
          expect(Cell.getStyle(ws, "A2")).toEqual({
            numFmt: "left",
            alignment: { horizontal: "left", vertical: "top" }
          });
          expect(Cell.getStyle(ws, "B2")).toEqual({});
          expect(Cell.getStyle(ws, "C2")).toEqual({
            numFmt: "right",
            alignment: { horizontal: "right", vertical: "bottom" }
          });
          expect(getColumn(ws, 1).style).toEqual({
            alignment: { horizontal: "left", vertical: "top" }
          });
          expect(getColumn(ws, 2).style).toEqual({});
          expect(getColumn(ws, 3).style).toEqual({
            alignment: { horizontal: "right", vertical: "bottom" }
          });
        });
        it("Remove defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-col-remove-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceColumns(wsSquare, 2, 2);

          const wsSingles = Workbook.addWorksheet(wb, "splice-col-remove-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceColumns(wsSingles, 2, 2);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([, "1,1", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([, "2,1", "2,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 3))).toEqual([, "3,1", "3,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 4))).toEqual([, "4,1", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3].forEach(row => {
              if (["C", "D"].includes(col)) {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 4))).toEqual([, "4,1", "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A4"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "B1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "B4"))).toBe("single-D4");
        });
        it("Insert defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-col-insert-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceColumns(wsSquare, 3, 0, ["foo", "bar", "baz", "qux"]);

          const wsSingles = Workbook.addWorksheet(wb, "splice-col-insert-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceColumns(wsSingles, 3, 0, ["foo", "bar", "baz", "qux"]);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([
            ,
            "1,1",
            "1,2",
            "foo",
            "1,3",
            "1,4"
          ]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([
            ,
            "2,1",
            "2,2",
            "bar",
            "2,3",
            "2,4"
          ]);
          expect(rowValues(Worksheet.getRow(wsSquare, 3))).toEqual([
            ,
            "3,1",
            "3,2",
            "baz",
            "3,3",
            "3,4"
          ]);
          expect(rowValues(Worksheet.getRow(wsSquare, 4))).toEqual([
            ,
            "4,1",
            "4,2",
            "qux",
            "4,3",
            "4,4"
          ]);

          ["A", "B", "C", "D", "E"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              if (col === "C") {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", , "foo", , "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 4))).toEqual([, "4,1", , "qux", , "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A4"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "E1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "E4"))).toBe("single-D4");
        });
        it("Replace defined names", () => {
          const wb = Workbook.create();
          const wsSquare = Workbook.addWorksheet(wb, "splice-col-replace-name-square");
          Worksheet.addRow(wsSquare, ["1,1", "1,2", "1,3", "1,4"]);
          Worksheet.addRow(wsSquare, ["2,1", "2,2", "2,3", "2,4"]);
          Worksheet.addRow(wsSquare, ["3,1", "3,2", "3,3", "3,4"]);
          Worksheet.addRow(wsSquare, ["4,1", "4,2", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              cellSetName(getCell(wsSquare, col + row), "square");
            });
          });

          Worksheet.spliceColumns(wsSquare, 2, 1, ["foo", "bar", "baz", "qux"]);

          const wsSingles = Workbook.addWorksheet(wb, "splice-col-replace-name-singles");
          Cell.setValue(wsSingles, "A1", "1,1");
          Cell.setValue(wsSingles, "A4", "4,1");
          Cell.setValue(wsSingles, "D1", "1,4");
          Cell.setValue(wsSingles, "D4", "4,4");

          ["A", "D"].forEach(col => {
            [1, 4].forEach(row => {
              cellSetName(getCell(wsSingles, col + row), `single-${col}${row}`);
            });
          });

          Worksheet.spliceColumns(wsSingles, 2, 1, ["foo", "bar", "baz", "qux"]);

          expect(wsSquare).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSquare, 1))).toEqual([, "1,1", "foo", "1,3", "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 2))).toEqual([, "2,1", "bar", "2,3", "2,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 3))).toEqual([, "3,1", "baz", "3,3", "3,4"]);
          expect(rowValues(Worksheet.getRow(wsSquare, 4))).toEqual([, "4,1", "qux", "4,3", "4,4"]);

          ["A", "B", "C", "D"].forEach(col => {
            [1, 2, 3, 4].forEach(row => {
              if (col === "B") {
                expect(cellName(getCell(wsSquare, col + row))).toBeUndefined();
              } else {
                expect(cellName(getCell(wsSquare, col + row))).toBe("square");
              }
            });
          });

          expect(wsSingles).not.toBeUndefined();
          expect(rowValues(Worksheet.getRow(wsSingles, 1))).toEqual([, "1,1", "foo", , "1,4"]);
          expect(rowValues(Worksheet.getRow(wsSingles, 4))).toEqual([, "4,1", "qux", , "4,4"]);
          expect(cellName(getCell(wsSingles, "A1"))).toBe("single-A1");
          expect(cellName(getCell(wsSingles, "A4"))).toBe("single-A4");
          expect(cellName(getCell(wsSingles, "D1"))).toBe("single-D1");
          expect(cellName(getCell(wsSingles, "D4"))).toBe("single-D4");
        });
      });
    });

    it("iterates over rows", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "B2", 2);
      Cell.setValue(ws, "D4", 4);
      Cell.setValue(ws, "F6", 6);
      Worksheet.eachRow(ws, (row, rowNumber) => {
        expect(rowNumber).not.toBe(3);
        expect(rowNumber).not.toBe(5);
      });

      let count = 1;
      Worksheet.eachRow(ws, { includeEmpty: true }, (row, rowNumber) => {
        expect(rowNumber).toBe(count++);
      });
    });

    it("iterates over column cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Cell.setValue(ws, "A4", 4);
      Cell.setValue(ws, "A6", 6);
      const colA = getColumn(ws, "A");
      columnEachCell(colA, (cell: any, rowNumber: any) => {
        expect(rowNumber).not.toBe(3);
        expect(rowNumber).not.toBe(5);
        expect(cellGetValue(cell)).toBe(rowNumber);
      });

      let count = 1;
      columnEachCell(colA, { includeEmpty: true }, (cell: any, rowNumber: any) => {
        expect(rowNumber).toBe(count++);
      });
      expect(count).toBe(7);
    });

    it("returns sheet values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Cell.setValue(ws, "A1", 11);
      Cell.setValue(ws, "C1", "C1");
      Cell.setValue(ws, "A2", 21);
      Cell.setValue(ws, "B2", "B2");
      Cell.setValue(ws, "A4", "end");

      expect(Worksheet.getValues(ws)).toEqual([, [, 11, , "C1"], [, 21, "B2"], , [, "end"]]);
    });

    it("calculates rowCount and actualRowCount", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Cell.setValue(ws, "A1", "A1");
      Cell.setValue(ws, "C1", "C1");
      Cell.setValue(ws, "A3", "A3");
      Cell.setValue(ws, "D3", "D3");
      Cell.setValue(ws, "A4", null);
      Cell.setValue(ws, "B5", "B5");

      expect(Worksheet.rowCount(ws)).toBe(5);
      expect(Worksheet.actualRowCount(ws)).toBe(3);
    });

    it("calculates columnCount and actualColumnCount", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Cell.setValue(ws, "A1", "A1");
      Cell.setValue(ws, "C1", "C1");
      Cell.setValue(ws, "A3", "A3");
      Cell.setValue(ws, "D3", "D3");
      Cell.setValue(ws, "E4", null);
      Cell.setValue(ws, "F5", "F5");

      expect(Worksheet.columnCount(ws)).toBe(6);
      expect(Worksheet.actualColumnCount(ws)).toBe(4);
    });
  });
});
