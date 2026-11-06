import { ArchiveFileTooLargeError } from "@archive/core/errors";
/**
 * The streaming reader holds the workbook, styles, shared strings and
 * relationship parts in memory, so it bounds each of them like the buffered
 * reader does — counting bytes as they inflate, since a streamed ZIP may not
 * declare sizes. Worksheets are streamed and not bounded by it.
 */
import { Cell, ExcelLimitExceededError, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { describe, it, expect } from "vitest";

async function workbookWithStrings(count: number): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (let i = 1; i <= count; i++) {
    Cell.setValue(ws, i, 1, `distinct string number ${i}`);
  }
  return Workbook.toBuffer(wb);
}

async function readAll(buffer: Uint8Array, maxEntrySize?: number): Promise<number> {
  const reader = new WorkbookReader(buffer, {
    worksheets: "emit",
    sharedStrings: "cache",
    styles: "cache",
    hyperlinks: "cache",
    ...(maxEntrySize === undefined ? {} : { maxEntrySize })
  });
  let rows = 0;
  for await (const ws of reader) {
    for await (const _row of ws) {
      rows++;
    }
  }
  return rows;
}

describe("streaming reader part-size limit", () => {
  it("rejects a cached part larger than maxEntrySize while it inflates", async () => {
    const buffer = await workbookWithStrings(2000);
    const err = await readAll(buffer, 4096).catch((e: unknown) => e);
    // An Excel error naming the limit and what it allows; the archive's refusal, naming the part, is the cause.
    expect(err).toBeInstanceOf(ExcelLimitExceededError);
    expect(err).toMatchObject({ limit: "maxEntrySize", allowed: 4096 });
    expect((err as Error).cause).toBeInstanceOf(ArchiveFileTooLargeError);
    expect((err as Error).cause).toMatchObject({ path: "xl/sharedStrings.xml" });
  });

  it("reads the same workbook under the default limit", async () => {
    expect(await readAll(await workbookWithStrings(2000))).toBe(2000);
  });

  it("does not bound streamed worksheets", async () => {
    // Inline values put the bulk in the sheet part, which is streamed.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 3000; i++) {
      Cell.setValue(ws, i, 1, i);
    }
    expect(await readAll(await Workbook.toBuffer(wb), 16 * 1024)).toBe(3000);
  });

  it.each([Number.NaN, -1])("rejects maxEntrySize %s as an ExcelArgumentError", bad => {
    expect(() => new WorkbookReader(new Uint8Array(0), { maxEntrySize: bad })).toThrow(
      expect.objectContaining({ name: "ExcelArgumentError", field: "options.maxEntrySize" })
    );
  });
});
