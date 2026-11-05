/**
 * Row and cell lookups by number on a streaming worksheet writer.
 *
 * A writer keeps only its uncommitted rows, behind an offset that moves as rows are written out and resets when it
 * compacts its buffer (after more than 1024 committed rows). The core lookups used to index the writer's row array as
 * `row - 1`, so once rows had been committed `Row.setHeight(ws, 2002, …)` styled a detached row and the file never
 * saw it.
 */

import { ExcelRowOutOfBoundsError } from "@excel/errors";
import { Cell, Row, Stream, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

function writerPast(committed: number) {
  const writer = new Stream.WorkbookWriter({});
  const ws = writer.addWorksheet("S");
  for (let i = 1; i <= committed; i++) {
    Stream.commitRow(ws.addRow([i]));
  }
  return { writer, ws };
}

describe("by-number lookups on a streaming writer", () => {
  it.each([3, 2000])("address the right row after %i committed rows", committed => {
    const { ws } = writerPast(committed);
    ws.addRow([committed + 1]);
    ws.addRow([committed + 2]);

    Row.setHeight(ws as never, committed + 2, 30);

    expect(ws.findRow(committed + 2)?.height).toBe(30);
    expect(ws.findRow(committed + 1)?.height).toBeUndefined();
    expect(Row.getValues(ws as never, committed + 2)).toEqual([committed + 2]);
    expect(Cell.getValue(ws as never, `A${committed + 1}`)).toBe(committed + 1);
  });

  it("refuses a row that has already been written out", () => {
    const { ws } = writerPast(2000);
    expect(() => Row.setHeight(ws as never, 5, 10)).toThrow(ExcelRowOutOfBoundsError);
  });

  it("writes what was set through the namespace into the file", async () => {
    const { writer, ws } = writerPast(2000);
    ws.addRow(["last"]);
    Row.setHeight(ws as never, 2001, 30);
    Stream.commitRow(ws.getRow(2001));
    const wb = await Workbook.read(await writer.toBuffer());
    const sheet = Workbook.getWorksheet(wb, "S")!;
    expect(Row.getHeight(sheet, 2001)).toBe(30);
    expect(Cell.getValue(sheet, "A2001")).toBe("last");
  });
});
