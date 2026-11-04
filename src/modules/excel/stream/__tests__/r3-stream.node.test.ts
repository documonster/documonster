/**
 * Round-3 review fixes in the streaming writer: the boundary checks the document model applies, committed rows and
 * workbooks refusing edits, and checked row height and column width.
 */
import { Writable } from "node:stream";

import {
  ExcelArgumentError,
  ExcelStreamStateError,
  InvalidAddressError,
  Stream,
  Workbook
} from "@excel";
import { describe, expect, it } from "vitest";

function writer() {
  const chunks: Uint8Array[] = [];
  const stream = new Writable({
    write(chunk, _encoding, callback) {
      chunks.push(chunk);
      callback();
    }
  });
  const wb = new Stream.WorkbookWriter({ stream });
  return { wb, ws: wb.addWorksheet("S"), bytes: () => Buffer.concat(chunks) };
}

describe("WorksheetWriter#mergeCells", () => {
  it("refuses a sheet qualifier and a non-reference; a single cell merges nothing", () => {
    const { ws } = writer();
    expect(() => ws.mergeCells("Other!A30:B31")).toThrow(InvalidAddressError);
    expect(() => ws.mergeCells("zz")).toThrow(InvalidAddressError);
    expect(() => ws.mergeCells("A10:B10", "x" as never)).toThrow(InvalidAddressError);
    ws.mergeCells("A20");
    ws.mergeCells("a1:b2");
    expect(() => ws.mergeCells("B2:C3")).toThrow();
  });
});

describe("protection", () => {
  it("is checked as Worksheet.protect and Workbook.protect check it", async () => {
    const { wb, ws } = writer();
    await expect(ws.protect({ password: 42 } as never)).rejects.toThrow(ExcelArgumentError);
    await expect(wb.protect({ lockStructure: "yes" } as never)).rejects.toThrow(ExcelArgumentError);
  });
});

describe("rows and cells", () => {
  it("addRow refuses values that are not a row", () => {
    const { ws } = writer();
    expect(() => ws.addRow(42 as never)).toThrow(ExcelArgumentError);
    expect(() => ws.addRow("abc" as never)).toThrow(ExcelArgumentError);
  });

  it("getCell and getRow refuse an address off the grid with an address error", () => {
    const { ws } = writer();
    expect(() => ws.getCell("zz")).toThrow(InvalidAddressError);
    expect(() => ws.getRow(0)).toThrow(/whole numbers from 1/);
  });

  it("a committed row's handles refuse edits", () => {
    const { ws } = writer();
    const row = ws.addRow(["a", 1]);
    Stream.commitRow(row);
    const cell = Stream.getRowCell(row, 1);
    expect(() => Stream.setCellFont(cell, { bold: true })).toThrow(ExcelStreamStateError);
    expect(() => Stream.setCellNote(cell, "late")).toThrow(ExcelStreamStateError);
    expect(() => Stream.setCellValue(cell, 2)).toThrow(ExcelStreamStateError);
    expect(() => Stream.setRowFont(row, { bold: true })).toThrow(ExcelStreamStateError);
    expect(() => Stream.addRowPageBreak(row)).toThrow(ExcelStreamStateError);
    expect(() => Stream.setRowHeight(row, 20)).toThrow(ExcelStreamStateError);
  });

  it("row height, hidden and column width are checked as the document model checks them", () => {
    const { ws } = writer();
    const row = ws.addRow([1]);
    expect(() => Stream.setRowHeight(row, 1e6)).toThrow(ExcelArgumentError);
    expect(() => Stream.setRowHeight(row, -1)).toThrow(ExcelArgumentError);
    Stream.setRowHeight(row, 30);
    expect(row.height).toBe(30);
    expect(() => Stream.setRowHidden(row, "yes" as never)).toThrow(ExcelArgumentError);
    Stream.setRowHidden(row, true);
    expect(row.hidden).toBe(true);
    const column = ws.getColumn(1);
    expect(() => Stream.setColumnWidth(column, -5)).toThrow(ExcelArgumentError);
    expect(() => Stream.setColumnWidth(column, 300)).toThrow(ExcelArgumentError);
    Stream.setColumnWidth(column, 12.5);
    expect(column.width).toBe(12.5);
  });

  it("setCellNumFmt takes a NumFmt as the document model does", () => {
    const { ws } = writer();
    const cell = Stream.getRowCell(ws.addRow([1]), 1);
    Stream.setCellNumFmt(cell, { id: 2, formatCode: "0.00" });
    expect(() => Stream.setCellNumFmt(cell, 42 as never)).toThrow(ExcelArgumentError);
  });
});

describe("cell styles", () => {
  it("addCellStyle and useBuiltInCellStyle are checked as the Workbook ones are", () => {
    const { wb } = writer();
    expect(() => wb.addCellStyle("X", null as never)).toThrow(ExcelArgumentError);
    expect(() => wb.useBuiltInCellStyle("Bogus" as never)).toThrow(ExcelArgumentError);
    const doc = Workbook.create();
    expect(() => Workbook.addCellStyle(doc, "X", null as never)).toThrow(ExcelArgumentError);
  });
});

describe("a committed workbook", () => {
  it("refuses another commit and another sheet with ExcelStreamStateError", async () => {
    const { wb, ws } = writer();
    ws.addRow([1]);
    await wb.commit();
    await expect(wb.commit()).rejects.toThrow(ExcelStreamStateError);
    expect(() => wb.addWorksheet("Late")).toThrow(ExcelStreamStateError);
  });
});

describe("X7 every edit of a committed workbook is an ExcelStreamStateError", () => {
  it.each([
    ["addCellStyle", (wb: Stream.WorkbookWriter) => wb.addCellStyle("X", { font: { bold: true } })],
    ["useBuiltInCellStyle", (wb: Stream.WorkbookWriter) => wb.useBuiltInCellStyle("Heading1")],
    [
      "addImage",
      (wb: Stream.WorkbookWriter) =>
        wb.addImage({ buffer: new Uint8Array([137, 80, 78, 71]), extension: "png" })
    ],
    ["unprotect", (wb: Stream.WorkbookWriter) => wb.unprotect()],
    [
      "defaultFont",
      (wb: Stream.WorkbookWriter) => {
        wb.defaultFont = { name: "Arial" };
      }
    ]
  ])("%s", async (_name, edit) => {
    const { wb, ws } = writer();
    ws.addRow([1]);
    await wb.commit();
    expect(() => edit(wb)).toThrow(ExcelStreamStateError);
  });

  it("protect", async () => {
    const { wb } = writer();
    await wb.commit();
    await expect(wb.protect({ password: "x" })).rejects.toThrow(ExcelStreamStateError);
  });
});
