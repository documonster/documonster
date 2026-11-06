/**
 * A file path the streaming reader cannot open is an `ExcelFileError`, as it is for `Workbook.readFile` — the
 * stream opened for the path reports the refusal only once it is read, and it used to surface as the platform's
 * own `ENOENT` error, which `isExcelError` does not recognise.
 */
import { ExcelFileError, isExcelError, Stream } from "@excel/index";
import { describe, expect, it } from "vitest";

const missing = "/nonexistent/documonster/streamed.xlsx";

describe("Stream.WorkbookReader over a file path", () => {
  it("reports a missing file from parse() as ExcelFileError, the system error as its cause", async () => {
    const reader = new Stream.WorkbookReader(missing);
    const error = await (async () => {
      for await (const _event of reader.parse()) {
        // drain
      }
    })().catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ExcelFileError);
    expect(isExcelError(error)).toBe(true);
    expect((error as ExcelFileError).path).toBe(missing);
    expect((error as ExcelFileError).operation).toBe("read");
    expect(((error as Error).cause as { code?: string }).code).toBe("ENOENT");
  });

  it("reports it from read() the same way", async () => {
    const error = await new Stream.WorkbookReader(missing).read().catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ExcelFileError);
  });
});
