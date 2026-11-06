/**
 * `Workbook.read` / `readWithDiagnostics` / `readStream` / `readFile` return a new workbook from their input.
 * The forms that read *into* a workbook stay as deprecated overloads, and keep replacing what it held.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";

import { Cell, DefinedNames, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

async function sampleBytes(format?: "xlsb") {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Cell.setValue(ws, "A1", "hello");
  Cell.setValue(ws, "B2", 42);
  DefinedNames.add(Workbook.getDefinedNames(wb), "Data!$B$2", "Answer");
  DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Twice", "Data!$B$2*2");
  return Workbook.toBuffer(wb, format ? { format } : undefined);
}

function summary(wb: Workbook.Handle) {
  return Workbook.listWorksheets(wb).map(ws => ({
    name: Worksheet.getName(ws),
    a1: Cell.getValue(ws, "A1"),
    b2: Cell.getValue(ws, "B2")
  }));
}

describe("reads that return a new workbook", () => {
  it("read(bytes) matches the deprecated read(wb, bytes)", async () => {
    const bytes = await sampleBytes();
    const fresh = await Workbook.read(bytes);
    const into = Workbook.create();
    const returned = await Workbook.read(into, bytes);
    expect(returned).toBe(into);
    expect(summary(fresh)).toEqual(summary(into));
    expect(summary(fresh)).toEqual([{ name: "Data", a1: "hello", b2: 42 }]);
  });

  it("each call returns a distinct workbook", async () => {
    const bytes = await sampleBytes();
    const a = await Workbook.read(bytes);
    const b = await Workbook.read(bytes);
    expect(a).not.toBe(b);
    Cell.setValue(Workbook.getWorksheet(a, "Data")!, "A1", "changed");
    expect(Cell.getValue(Workbook.getWorksheet(b, "Data")!, "A1")).toBe("hello");
  });

  it("accepts every input the old form did: ArrayBuffer, a view, base64 text, XLSB", async () => {
    const bytes = await sampleBytes();
    const copy = Uint8Array.from(bytes);
    const expected = [{ name: "Data", a1: "hello", b2: 42 }];
    expect(summary(await Workbook.read(copy.buffer))).toEqual(expected);
    expect(summary(await Workbook.read(new DataView(copy.buffer)))).toEqual(expected);
    const base64 = Buffer.from(bytes).toString("base64");
    expect(summary(await Workbook.read(base64, { base64: true }))).toEqual(expected);
    expect(summary(await Workbook.read(await sampleBytes("xlsb")))).toEqual(expected);
  });

  it("passes options through, including format and the create-time syntax probe", async () => {
    const xlsb = await sampleBytes("xlsb");
    await expect(Workbook.read(xlsb, { format: "xlsx" })).rejects.toThrow();
    const seen: string[] = [];
    const wb = await Workbook.read(await sampleBytes(), {
      formulaSyntaxProbe: text => {
        seen.push(text);
        return false;
      }
    });
    expect(summary(wb)).toEqual([{ name: "Data", a1: "hello", b2: 42 }]);
    expect(DefinedNames.getNames(Workbook.getDefinedNames(wb), "Data!B2")).toEqual(["Answer"]);
    // The probe is the one the new workbook was created with, so it is what classified the formula name.
    expect(seen).toContain("Data!$B$2*2");
  });

  it("the deprecated form still replaces what the workbook held", async () => {
    const bytes = await sampleBytes();
    const target = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(target, "Leftover"), "A1", "old");
    await Workbook.read(target, bytes);
    expect(Workbook.listWorksheets(target).map(ws => Worksheet.getName(ws))).toEqual(["Data"]);
  });

  it("readWithDiagnostics(bytes) returns the new workbook on the report", async () => {
    const bytes = await sampleBytes("xlsb");
    const report = await Workbook.readWithDiagnostics(bytes);
    expect(summary(report.workbook)).toEqual([{ name: "Data", a1: "hello", b2: 42 }]);
    const into = Workbook.create();
    const legacy = await Workbook.readWithDiagnostics(into, bytes);
    expect(legacy.workbook).toBe(into);
    expect(report.lost).toEqual(legacy.lost);
    const xlsx = await Workbook.readWithDiagnostics(await sampleBytes());
    expect(summary(xlsx.workbook)).toEqual(summary(report.workbook));
  });

  it("readStream(stream) and readFile(path) return a new workbook", async () => {
    const bytes = await sampleBytes();
    const fromStream = await Workbook.readStream(Readable.from([Buffer.from(bytes)]));
    expect(summary(fromStream)).toEqual([{ name: "Data", a1: "hello", b2: 42 }]);

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "dm-read-"));
    try {
      const file = path.join(dir, "book.xlsx");
      fs.writeFileSync(file, bytes);
      const fromFile = await Workbook.readFile(file);
      expect(summary(fromFile)).toEqual(summary(fromStream));
      const into = Workbook.create();
      expect(await Workbook.readFile(into, file)).toBe(into);
      expect(summary(into)).toEqual(summary(fromFile));

      const xlsbFile = path.join(dir, "book.xlsb");
      fs.writeFileSync(xlsbFile, await sampleBytes("xlsb"));
      expect(summary(await Workbook.readFile(xlsbFile))).toEqual(summary(fromFile));
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});
