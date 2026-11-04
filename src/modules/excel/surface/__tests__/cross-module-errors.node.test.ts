/**
 * Cross-module consistency of the Excel entry's errors: the argument error names its `field`, the byte reader checks
 * its input like `readStream`, and `readFile` keeps the file system's error as the cause.
 */
import { tmpdir } from "node:os";
import path from "node:path";

import { Cell, ExcelArgumentError, ExcelFileError, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("ExcelArgumentError", () => {
  it("names the offending field as `field`, as the other modules' argument errors do", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const error = (() => {
      try {
        Cell.setFont(ws, "A1", { size: -3 });
      } catch (e) {
        return e;
      }
    })() as ExcelArgumentError;
    expect(error).toBeInstanceOf(ExcelArgumentError);
    // The path from the parameter name: `Cell.setFont(ws, addr, value)`.
    expect(error.field).toBe("value.size");
    expect("argument" in error).toBe(false);
  });
});

describe("Workbook.read input", () => {
  it.each([42, null, undefined, {}, "hello", true, [1, 2]])(
    "refuses %j, which is not bytes, as an argument error",
    async input => {
      const error = await Workbook.read(input as never).catch(e => e);
      expect(error).toBeInstanceOf(ExcelArgumentError);
      expect((error as ExcelArgumentError).field).toBe("data");
      const report = await Workbook.readWithDiagnostics(input as never).catch(e => e);
      expect(report).toBeInstanceOf(ExcelArgumentError);
      // The message describes the argument, not the ZIP loader it no longer reaches.
      for (const [where, e] of [
        ["Workbook.read", error],
        ["Workbook.readWithDiagnostics", report]
      ] as const) {
        expect((e as Error).message).toMatch(
          new RegExp(
            `^${where.replace(".", "\\.")}: data must be bytes \\(Uint8Array, ArrayBuffer, a typed-array view, a Blob\\) or a base64 string with \\{ base64: true \\}, got \\S`
          )
        );
        expect((e as Error).message).not.toContain("zip file");
      }
    }
  );

  it("names what it was given", async () => {
    const message = (input: unknown) =>
      Workbook.read(input as never).then(
        () => "",
        (e: Error) => e.message.slice(e.message.indexOf(", got ") + 6)
      );
    expect(await message(42)).toBe("number");
    expect(await message(null)).toBe("null");
    expect(await message(undefined)).toBe("undefined");
    expect(await message({})).toBe("object");
    expect(await message([1])).toBe("an array");
    expect(await message("hello")).toBe("a string without { base64: true }");
  });

  it("keeps ExcelFileError for real bytes that are not a workbook, and base64 text that is", async () => {
    await expect(Workbook.read(new Uint8Array([1, 2, 3]))).rejects.toThrow(ExcelFileError);
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    const base64 = Buffer.from(await Workbook.toBuffer(wb)).toString("base64");
    const back = await Workbook.read(base64, { base64: true });
    expect(Cell.getValue(Workbook.getWorksheet(back, 1)!, "A1")).toBe(1);
    // Text without `base64: true` is not bytes either.
    await expect(Workbook.read(base64)).rejects.toThrow(ExcelArgumentError);
  });

  it("accepts every ArrayBufferView", async () => {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    const bytes = await Workbook.toBuffer(wb);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const back = await Workbook.read(view);
    expect(Cell.getValue(Workbook.getWorksheet(back, 1)!, "A1")).toBe(1);
  });
});

describe("Workbook.readFile", () => {
  it("keeps the file system's error as the cause", async () => {
    const missing = path.join(tmpdir(), "documonster-no-such-file.xlsx");
    const error = await Workbook.readFile(missing).catch(e => e);
    expect(error).toBeInstanceOf(ExcelFileError);
    expect((error.cause as NodeJS.ErrnoException | undefined)?.code).toBe("ENOENT");
    const xlsb = await Workbook.readFile(`${missing}.xlsb`).catch(e => e);
    expect(xlsb).toBeInstanceOf(ExcelFileError);
    expect((xlsb.cause as NodeJS.ErrnoException | undefined)?.code).toBe("ENOENT");
  });

  it("refuses a path that is not a string", async () => {
    await expect(Workbook.readFile(42 as never)).rejects.toThrow(ExcelArgumentError);
  });
});
