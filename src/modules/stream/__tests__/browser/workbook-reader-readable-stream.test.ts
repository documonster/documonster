import { cellSetValue } from "@excel/cell";
import { getCell } from "@excel/worksheet";
import { describe, it, expect } from "vitest";
describe("WorkbookReader (Browser) accepts ReadableStream input", () => {
  it("should read a workbook from ReadableStream<Uint8Array>", async () => {
    const excelModule = await import("../../../../index.browser");
    const { Workbook, WorkbookReader } = excelModule as any;

    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    cellSetValue(getCell(ws, "A1"), "hello");
    cellSetValue(getCell(ws, "A2"), 42);

    const data: Uint8Array = await Workbook.toXlsxBuffer(wb);

    const webStream = new ReadableStream<Uint8Array>({
      start(controller) {
        const chunkSize = 32 * 1024;
        for (let i = 0; i < data.length; i += chunkSize) {
          controller.enqueue(data.slice(i, i + chunkSize));
        }
        controller.close();
      }
    });

    const reader = new WorkbookReader(webStream, { worksheets: "emit" });
    let seen = false;

    for await (const worksheet of reader) {
      seen = true;
      expect(worksheet.name).toBe("Sheet1");

      let rowCount = 0;
      for await (const row of worksheet) {
        rowCount++;
        if (row.number === 1) {
          expect(row.getCell(1).value).toBe("hello");
        }
        if (row.number === 2) {
          expect(row.getCell(1).value).toBe(42);
        }
      }

      expect(rowCount).toBeGreaterThan(0);
    }

    expect(seen).toBe(true);
  });
});
