/**
 * Positions are checked before anything changes, and a bad one is an argument error.
 *
 * - Every `Column` setter (and `Stream`'s `WorksheetWriter#getColumn`) refuses a column number that is not a whole
 *   number from 1 to 16 384 with `ExcelColumnOutOfBoundsError`, as the getters do — they used to throw a raw `TypeError` for
 *   0 or 1.5, and accept 16 385 (materialising every column up to it).
 * - `ExcelRowOutOfBoundsError`, `ExcelColumnOutOfBoundsError` and `ExcelInvalidAddressError` are `ExcelArgumentError`s carrying the
 *   `field` of the parameter at fault, so the one `catch` a caller writes for bad input catches them.
 */
import {
  Address,
  Cell,
  Column,
  ExcelColumnOutOfBoundsError,
  ExcelArgumentError,
  ExcelError,
  ExcelInvalidAddressError,
  isExcelError,
  Row,
  ExcelRowOutOfBoundsError,
  Stream,
  Workbook
} from "@excel";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return Workbook.addWorksheet(wb, "S");
}

const SETTERS: [string, (ws: ReturnType<typeof sheet>, col: number) => unknown][] = [
  ["setWidth", (ws, col) => Column.setWidth(ws, col, 10)],
  ["setHidden", (ws, col) => Column.setHidden(ws, col, true)],
  ["setOutlineLevel", (ws, col) => Column.setOutlineLevel(ws, col, 1)],
  ["setKey", (ws, col) => Column.setKey(ws, col, "k")],
  ["setHeader", (ws, col) => Column.setHeader(ws, col, "h")],
  ["setStyle", (ws, col) => Column.setStyle(ws, col, { font: { bold: true } })],
  ["setNumFmt", (ws, col) => Column.setNumFmt(ws, col, "0.00")],
  ["setFont", (ws, col) => Column.setFont(ws, col, { bold: true })],
  ["addPageBreak", (ws, col) => Column.addPageBreak(ws, col)],
  ["getStyle", (ws, col) => Column.getStyle(ws, col)]
];

describe("Column setters check the column first", () => {
  describe.each(SETTERS)("Column.%s", (_name, call) => {
    it.each([0, -1, 1.5, 16385, 1e7, Number.NaN])("refuses %s and changes nothing", col => {
      const ws = sheet();
      expect(() => call(ws, col)).toThrow(ExcelColumnOutOfBoundsError);
      expect(Column.getDefinitions(ws)).toEqual([]);
      expect(Column.has(ws, 1)).toBe(false);
    });

    it("accepts 16384", () => {
      const ws = sheet();
      expect(() => call(ws, 16384)).not.toThrow();
    });
  });

  it("Stream's WorksheetWriter#getColumn checks the same way", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    for (const col of [0, 1.5, 16385]) {
      expect(() => ws.getColumn(col)).toThrow(ExcelColumnOutOfBoundsError);
    }
    expect(ws.getColumn(16384).number).toBe(16384);
    // A built-in property name is not a column key.
    expect(() => ws.getColumn("toString")).toThrow(ExcelColumnOutOfBoundsError);
  });
});

describe("position errors are argument errors", () => {
  it.each<[string, () => unknown, abstract new (...args: never[]) => Error, string]>([
    ["Row.setHeight row 0", () => Row.setHeight(sheet(), 0, 5), ExcelRowOutOfBoundsError, "row"],
    ["Cell.getValue ZZ0", () => Cell.getValue(sheet(), "ZZ0"), ExcelInvalidAddressError, "addr"],
    ["Cell.getValue col 0", () => Cell.getValue(sheet(), 1, 0), ExcelColumnOutOfBoundsError, "col"],
    ["Column.setWidth 0", () => Column.setWidth(sheet(), 0, 5), ExcelColumnOutOfBoundsError, "col"],
    ["Address.encodeCol -1", () => Address.encodeCol(-1), ExcelColumnOutOfBoundsError, "col"],
    ["Address.encodeRow -1", () => Address.encodeRow(-1), ExcelRowOutOfBoundsError, "row"],
    ["Address.decodeCell nope", () => Address.decodeCell("nope"), ExcelInvalidAddressError, "cstr"]
  ])("%s", (_name, call, type, field) => {
    let error: unknown;
    try {
      call();
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(type);
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect(error).toBeInstanceOf(ExcelError);
    expect(isExcelError(error)).toBe(true);
    expect((error as ExcelArgumentError).field).toBe(field);
  });

  it("Address names the call in the message", () => {
    expect(() => Address.encodeCol(-1)).toThrow(/^Address\.encodeCol: /);
    expect(() => Address.encodeCell({ c: 0, r: -1 })).toThrow(
      expect.objectContaining({
        field: "cell.r",
        message: expect.stringMatching(/^Address\.encodeCell: /)
      })
    );
  });

  it("the constructors keep their positional forms, and take a field in the options", () => {
    const row = new ExcelRowOutOfBoundsError(0);
    expect([row.row, row.field, row.message]).toEqual([0, "row", "Row 0 is out of bounds"]);
    const col = new ExcelColumnOutOfBoundsError("XFE", "too far", { field: "options.col" });
    expect([col.column, col.field, col.message]).toEqual([
      "XFE",
      "options.col",
      "Column XFE is out of bounds: too far"
    ]);
    const cause = new Error("x");
    const address = new ExcelInvalidAddressError("Q", undefined, { cause });
    expect([address.address, address.field, address.cause]).toEqual(["Q", "addr", cause]);
  });
});
