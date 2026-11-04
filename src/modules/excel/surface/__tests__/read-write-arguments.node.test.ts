/**
 * What `Workbook.read*` and `Workbook.toBuffer` do with an input or an option they cannot use: a typed error from the
 * call that was given it — never a hang, a raw `TypeError` or a silently ignored option.
 */
import { Readable } from "node:stream";
import vm from "node:vm";

import { Cell, ExcelArgumentError, ExcelFileError, Workbook, isExcelError } from "@excel/index";
import { describe, expect, it } from "vitest";

async function bytes() {
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 42);
  return Workbook.toBuffer(wb);
}

describe("Workbook.readStream", () => {
  it.each([42, "a path", null, {}])("refuses %j, which is not a stream", async input => {
    await expect(Workbook.readStream(input as never)).rejects.toThrow(ExcelArgumentError);
  });

  it("rejects, rather than waiting forever, when the source fails before any data", async () => {
    const failing = new Readable({
      read() {
        this.destroy(new Error("source failed"));
      }
    });
    await expect(Workbook.readStream(failing)).rejects.toThrow("source failed");
  });

  it("still reads a stream", async () => {
    const wb = await Workbook.readStream(Readable.from([await bytes()]));
    expect(Cell.getValue(Workbook.getWorksheet(wb, 1)!, "A1")).toBe(42);
  });
});

describe("Workbook.read", () => {
  it("reports bytes that are not a package as an Excel error", async () => {
    const error = await Workbook.read(new Uint8Array([1, 2, 3])).catch(e => e);
    expect(error).toBeInstanceOf(ExcelFileError);
    expect(isExcelError(error)).toBe(true);
  });

  it("refuses options that are not an object, and a format it does not know", async () => {
    const buf = await bytes();
    await expect(Workbook.read(buf, 42 as never)).rejects.toThrow(ExcelArgumentError);
    await expect(Workbook.read(buf, { format: "bogus" as never })).rejects.toThrow(
      ExcelArgumentError
    );
  });

  it("reads an ArrayBuffer from another realm", async () => {
    const buf = await bytes();
    const foreign = vm.runInNewContext("new ArrayBuffer(n)", { n: buf.length }) as ArrayBuffer;
    new Uint8Array(foreign).set(buf);
    const wb = await Workbook.read(foreign);
    expect(Cell.getValue(Workbook.getWorksheet(wb, 1)!, "A1")).toBe(42);
  });
});

describe("Workbook.toBuffer", () => {
  it("refuses a format it cannot write, instead of writing XLSX", async () => {
    const wb = await Workbook.read(await bytes());
    await expect(Workbook.toBuffer(wb, { format: "csv" as never })).rejects.toThrow(
      ExcelArgumentError
    );
    await expect(Workbook.toBuffer(wb, 42 as never)).rejects.toThrow(ExcelArgumentError);
  });
});
