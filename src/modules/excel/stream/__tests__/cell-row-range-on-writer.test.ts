/**
 * `Cell.*`, `Row.*` and `Range.*` address a streamed sheet by position as they address a worksheet: an edit reaches the
 * file, an edit or read of a row the writer has written is refused with `ExcelStreamStateError` (as `Stream.setCell*`
 * refuses one through a committed row's handle), what needs the rows already written is refused by name, and a range
 * write that cannot be made changes nothing.
 */
import {
  Cell,
  ExcelStreamStateError,
  ExcelUnsupportedFeatureError,
  Range,
  Row,
  Stream,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

function writer(): { wb: Stream.WorkbookWriter; ws: Stream.WorksheetWriter } {
  const wb = new Stream.WorkbookWriter({ useStyles: true });
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

async function readBack(wb: Stream.WorkbookWriter): Promise<Workbook.Handle> {
  return Workbook.read(await wb.toBuffer());
}

function sheetOf(book: Workbook.Handle) {
  return Workbook.getWorksheet(book, "S")!;
}

describe("Cell functions on a streaming worksheet writer", () => {
  it("write by address, read back the same, and reach the file", async () => {
    const { wb, ws } = writer();
    Cell.setValue(ws, "A1", "x");
    Cell.setValue(ws, { row: 1, col: 2 }, 42);
    Cell.setFont(ws, "A1", { bold: true });
    Cell.setNumFmt(ws, "B1", "0.00");
    Cell.setNote(ws, "A1", "hello");
    Workbook.addCellStyle(wb, "Totals", {
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } }
    });
    Cell.applyStyle(ws, "C1", "totals");
    expect(Cell.getValue(ws, "A1")).toBe("x");
    expect(Cell.getFont(ws, "A1")).toEqual({ bold: true });
    // A cell that does not exist answers as on a worksheet, and is not created.
    expect(Cell.getValue(ws, "D1")).toBeNull();
    expect(Cell.has(ws, "D1")).toBe(false);
    Stream.commitRow(ws.getRow(1));

    const back = sheetOf(await readBack(wb));
    expect(Cell.getValue(back, "A1")).toBe("x");
    expect(Cell.getValue(back, "B1")).toBe(42);
    expect(Cell.getFont(back, "A1")).toMatchObject({ bold: true });
    expect(Cell.getNumFmt(back, "B1")).toBe("0.00");
    expect(Cell.getNote(back, "A1")).toBeDefined();
    expect(Cell.findStyle(back, "C1")?.styleName).toBe("Totals");
  });

  it("an edit or read of a committed row is refused as Stream.setCell* refuses one", () => {
    const { ws } = writer();
    const row = ws.addRow(["a"]);
    Stream.commitRow(row);
    expect(() => Stream.setCellValue(Stream.getRowCell(row, 1), "b")).toThrow(
      ExcelStreamStateError
    );
    expect(() => Cell.setValue(ws, "A1", "b")).toThrow(ExcelStreamStateError);
    expect(() => Cell.setFont(ws, "B1", { bold: true })).toThrow(ExcelStreamStateError);
    expect(() => Cell.getStyle(ws, "A1")).toThrow(ExcelStreamStateError);
    expect(() => Cell.getValue(ws, "A1")).toThrow(ExcelStreamStateError);
    expect(() => Cell.getFont(ws, "A1")).toThrow(ExcelStreamStateError);
    // A later row is still open.
    Cell.setValue(ws, "A2", "c");
    expect(Cell.getValue(ws, "A2")).toBe("c");
  });

  it("a refused value leaves no cell behind, and a bad address is an address error first", () => {
    const { ws } = writer();
    expect(() => Cell.setFont(ws, "A1", { size: -1 } as never)).toThrow();
    expect(Cell.has(ws, "A1")).toBe(false);
    Stream.commitRow(ws.addRow([1]));
    expect(() => Cell.setValue(ws, "A0", 1)).toThrow(/A0/);
  });
});

describe("Row functions on a streaming worksheet writer", () => {
  it("add, set values, height, visibility, style and page breaks reach the file", async () => {
    const { wb, ws } = writer();
    expect(Row.add(ws, ["a", 1])).toBe(1);
    expect(Row.addMany(ws, [["b"], ["c"]])).toEqual([2, 3]);
    Row.setValues(ws, 2, ["B", 2]);
    Row.setHeight(ws, 2, 30);
    Row.setHidden(ws, 3, true);
    Row.setOutlineLevel(ws, 3, 1);
    Row.setFont(ws, 1, { italic: true });
    Row.addPageBreak(ws, 2);
    expect(Row.getValues(ws, 2)).toEqual(["B", 2]);
    expect(Row.getHeight(ws, 2)).toBe(30);

    const back = sheetOf(await readBack(wb));
    expect(Row.getValues(back, 1)).toEqual(["a", 1]);
    expect(Row.getValues(back, 2)).toEqual(["B", 2]);
    expect(Row.getHeight(back, 2)).toBe(30);
    expect(Row.isHidden(back, 3)).toBe(true);
    expect(Row.getOutlineLevel(back, 3)).toBe(1);
    expect(Cell.getFont(back, "A1")).toMatchObject({ italic: true });
    expect(back.rowBreaks.map(item => item.id)).toEqual([2]);
  });

  it("a committed row refuses edits and reads", () => {
    const { ws } = writer();
    Row.add(ws, ["a"]);
    Stream.commitRow(ws.getRow(1));
    expect(() => Row.setHeight(ws, 1, 20)).toThrow(ExcelStreamStateError);
    expect(() => Row.setValues(ws, 1, ["b"])).toThrow(ExcelStreamStateError);
    expect(() => Row.getValues(ws, 1)).toThrow(ExcelStreamStateError);
    expect(() => Row.setHeight(ws, 0, 20)).toThrow(/row|0/i);
  });

  it("what moves or walks the rows already written is refused by name", () => {
    const { ws } = writer();
    Row.add(ws, ["a"]);
    const refusals: [string, () => unknown][] = [
      ["Row.insert", () => Row.insert(ws as never, 1, ["x"])],
      ["Row.insertMany", () => Row.insertMany(ws as never, 1, [["x"]])],
      ["Row.duplicate", () => Row.duplicate(ws as never, 1, 1)],
      ["Row.splice", () => Row.splice(ws as never, 1, 1)],
      ["Row.getCount", () => Row.getCount(ws as never)],
      ["Row.getActualCount", () => Row.getActualCount(ws as never)],
      ["Row.getLast", () => Row.getLast(ws as never)],
      ["Row.each", () => Row.each(ws as never, () => undefined)],
      ["Row.autoFit", () => Row.autoFit(ws as never, 1)]
    ];
    for (const [operation, call] of refusals) {
      let error: unknown;
      try {
        call();
      } catch (thrown) {
        error = thrown;
      }
      expect(error, operation).toBeInstanceOf(ExcelUnsupportedFeatureError);
      expect((error as ExcelUnsupportedFeatureError).operation).toBe(operation);
    }
    expect(() => Row.add(ws, ["b"], { copyStyle: "above" })).toThrow(ExcelUnsupportedFeatureError);
  });
});

describe("Range functions on a streaming worksheet writer", () => {
  it("setValues and the style setters reach the file", async () => {
    const { wb, ws } = writer();
    Range.setValues(ws, "A1:B2", [
      ["a", 1],
      ["b", 2]
    ]);
    Range.setFont(ws, "A1:B1", { bold: true });
    Range.setStyle(ws, "A2:B2", { numFmt: "0.0" });
    expect(Range.getValues(ws, "A1:B2")).toEqual([
      ["a", 1],
      ["b", 2]
    ]);
    const back = sheetOf(await readBack(wb));
    expect(Range.getValues(back, "A1:B2")).toEqual([
      ["a", 1],
      ["b", 2]
    ]);
    expect(Cell.getFont(back, "B1")).toMatchObject({ bold: true });
    expect(Cell.getNumFmt(back, "B2")).toBe("0.0");
  });

  it("a range reaching a committed row is refused whole, and changes nothing", () => {
    const { ws } = writer();
    Range.setValues(ws, "A1:A3", [["a"], ["b"], ["c"]]);
    Stream.commitRow(ws.getRow(1));
    expect(() => Range.setValues(ws, "A1:B3", [[1], [2], [3]])).toThrow(ExcelStreamStateError);
    expect(() => Range.setFont(ws, "A1:A3", { bold: true })).toThrow(ExcelStreamStateError);
    expect(Range.getValues(ws, "A2:A3")).toEqual([["b"], ["c"]]);
    expect(Cell.getFont(ws, "A2")).toBeUndefined();
    // Reading a range with a written row in it is refused rather than answered as empty.
    expect(() => Range.getValues(ws, "A1:A3")).toThrow(ExcelStreamStateError);
  });
});
