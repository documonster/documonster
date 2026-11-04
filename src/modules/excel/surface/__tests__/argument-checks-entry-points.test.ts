/**
 * Arguments a public entry point used to accept in silence, or fail on with an untyped `TypeError` from inside the
 * library. Each is now refused at the call, with the module's own error naming what was wrong.
 */
import { ExcelCsv } from "@excel/bridge/csv";
import type { CsvOptions } from "@excel/bridge/csv";
import {
  Cell,
  DataValidation,
  ExcelArgumentError,
  InvalidValueTypeError,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

const BAD = [42, null, [], true] as const;

describe("entry points refuse what is not an argument of theirs", () => {
  it("a cell refuses an invalid Date, which the writer could not write", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "S");
    expect(() => Cell.setValue(ws, "A1", new Date("x"))).toThrow(InvalidValueTypeError);
    expect(() => Cell.setValue(ws, "A2", { formula: "TODAY()", result: new Date("x") })).toThrow(
      InvalidValueTypeError
    );
    expect(Cell.getValue(ws, "A1")).toBeNull();
  });

  it.each(BAD)("DataValidation names a target that is neither a sheet nor rules (%j)", bad => {
    expect(() => DataValidation.find(bad as never, "A1")).toThrow(ExcelArgumentError);
    expect(() =>
      DataValidation.add(bad as never, "A1", {
        type: "whole",
        operator: "between",
        formulae: [1, 2]
      })
    ).toThrow(ExcelArgumentError);
  });

  it("Workbook.removeWorksheet removes nothing for a missing sheet, where it removed the first", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "A");
    Workbook.addWorksheet(wb, "B");
    expect(() => Workbook.removeWorksheet(wb, undefined as never)).toThrow(ExcelArgumentError);
    expect(() => Workbook.removeWorksheet(wb, true as never)).toThrow(ExcelArgumentError);
    expect(Workbook.removeWorksheet(wb, 99)).toBe(false);
    expect(Workbook.listWorksheets(wb).map(ws => Worksheet.getName(ws))).toEqual(["A", "B"]);
  });

  it.each(BAD)("ExcelCsv.read refuses %j instead of reading its text", async bad => {
    await expect(ExcelCsv.read(bad as never)).rejects.toThrow(ExcelArgumentError);
  });

  it("ExcelCsv.read still takes text and bytes", async () => {
    for (const input of [
      "a,b",
      new TextEncoder().encode("a,b"),
      new TextEncoder().encode("a,b").buffer
    ]) {
      const wb = await ExcelCsv.read(input);
      expect(Cell.getValue(Workbook.listWorksheets(wb)[0]!, "B1")).toBe("b");
    }
  });
});

describe("ExcelCsv options are the read half or the write half", () => {
  it("a read takes the parser's options, and an option only a write uses is a type error", async () => {
    const wb = await ExcelCsv.read("a;b\n1;2", { delimiter: ";", sheetName: "In" });
    expect(Cell.getValue(Workbook.getWorksheet(wb, "In")!, "B2")).toBe(2);
    // @ts-expect-error — `dateFormat` is how a write formats dates; a read ignored it.
    await ExcelCsv.read("a", { dateFormat: "yyyy" });
    // @ts-expect-error — `delimitersToGuess` guides a parser; a write ignored it.
    ExcelCsv.format(wb, { delimitersToGuess: [";"] });
  });

  it("one bag of both still serves both, and `info` / `raw` no longer break a read", async () => {
    const both: CsvOptions = { delimiter: ";", dateFormat: "YYYY", sheetName: "In" };
    const wb = await ExcelCsv.read("a;b\n1;2", both);
    expect(ExcelCsv.format(wb, both)).toBe("a;b\n1;2");
    for (const extra of [{ info: true }, { raw: true }, { info: true, headers: true }]) {
      const read = await ExcelCsv.read("a,b\n1,2", extra as CsvOptions);
      expect(Cell.getValue(Workbook.listWorksheets(read)[0]!, "B2")).toBe(2);
    }
  });
});
