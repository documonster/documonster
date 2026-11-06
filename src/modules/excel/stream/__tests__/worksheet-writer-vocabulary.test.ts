/**
 * The streaming writer's handles and methods answer in the document model's vocabulary where they share a job:
 * `removeWatermark` reports what it removed, as `Watermark.remove` does; a column handle has a public name; and the
 * deprecated `Row.commit` writes a streamed row exactly as `Stream.commitRow` does.
 */
import { Cell, Row, Stream, Workbook } from "@excel/index";
import { describe, expect, expectTypeOf, it } from "vitest";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("Stream.WorksheetWriter vocabulary", () => {
  it("removeWatermark returns whether the sheet had one", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    const imageId = wb.addImage({ buffer: PNG, extension: "png" });
    expect(ws.removeWatermark()).toBe(false);
    ws.addWatermark({ imageId });
    expect(ws.removeWatermark()).toBe(true);
    expect(ws.getWatermark()).toBeNull();
    expect(ws.removeWatermark()).toBe(false);
  });

  it("getColumn hands out a Stream.ColumnHandle, which the setColumn* operations take", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    const column = ws.getColumn(2);
    expectTypeOf(column).toEqualTypeOf<Stream.ColumnHandle>();
    Stream.setColumnWidth(column, 20);
    expect(column.width).toBe(20);
  });

  it("getRowCell creates the cell it names; findCell does not", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    const row = ws.addRow([1]);
    expect(ws.findCell(1, 3)).toBeUndefined();
    const cell = Stream.getRowCell(row, 3);
    expect(ws.findCell(1, 3)).toBe(cell);
    expect(Stream.getRowCell(row, "C")).toBe(cell);
  });

  it("the deprecated Row.commit writes the row as Stream.commitRow does", async () => {
    const write = async (commit: (ws: Stream.WorksheetWriter) => void) => {
      const wb = new Stream.WorkbookWriter({});
      const ws = wb.addWorksheet("S");
      ws.addRow(["a"]);
      commit(ws);
      expect(ws.isRowCommitted(1)).toBe(true);
      ws.commit();
      const back = await Workbook.read(await wb.toBuffer());
      return Cell.getValue(Workbook.getWorksheet(back, "S")!, "A1");
    };
    expect(await write(ws => Row.commit(ws as never, 1))).toBe("a");
    expect(await write(ws => Stream.commitRow(ws.getRow(1)))).toBe("a");
  });
});
