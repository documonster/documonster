import * as CsvEntry from "@excel/bridge/csv";
import * as MarkdownEntry from "@excel/bridge/markdown";
import * as ChartEntry from "@excel/chart/index";
import * as ExcelEntry from "@excel/index";
/**
 * The unprefixed error names kept as deprecated forwards until 1.0: each is the
 * same class as its `Excel*` replacement, on every entry that published it.
 */
import { describe, expect, it } from "vitest";

const RENAMED = [
  "WorksheetNameError",
  "InvalidAddressError",
  "ColumnOutOfBoundsError",
  "RowOutOfBoundsError",
  "MergeConflictError",
  "InvalidValueTypeError",
  "PivotTableError",
  "ChartOptionsError",
  "TableError",
  "DefinedNameError",
  "ImageError"
] as const;

type ErrorClass = new (message: string) => Error;

describe("deprecated Excel error names", () => {
  const entry = ExcelEntry as unknown as Record<string, ErrorClass>;

  it.each(RENAMED)("%s forwards to its Excel-prefixed class", oldName => {
    const newName = `Excel${oldName}`;
    const NewClass = entry[newName];
    expect(NewClass).toBeTypeOf("function");
    expect(entry[oldName]).toBe(NewClass);
    const err = new NewClass("boom");
    expect(err).toBeInstanceOf(entry[oldName]);
    expect(err).toBeInstanceOf(ExcelEntry.ExcelError);
    expect(err.name).toBe(newName);
  });

  it.each([
    ["MaxItemsExceededError", "ExcelLimitExceededError"],
    ["ExcelNotSupportedError", "ExcelUnsupportedFeatureError"]
  ] as const)("%s forwards to %s, the name the other modules use", (oldName, newName) => {
    const NewClass = entry[newName];
    expect(NewClass).toBeTypeOf("function");
    expect(entry[oldName]).toBe(NewClass);
    const err = new (NewClass as unknown as new (a: string, b: number) => Error)("row", 1);
    expect(err).toBeInstanceOf(entry[oldName]);
    expect(err).toBeInstanceOf(ExcelEntry.ExcelError);
    expect(err.name).toBe(newName);
  });

  it("keeps the aliases on the bridge and chart entries", () => {
    expect(ChartEntry.ChartOptionsError).toBe(ChartEntry.ExcelChartOptionsError);
    for (const mod of [CsvEntry, MarkdownEntry]) {
      expect(mod.InvalidValueTypeError).toBe(mod.ExcelInvalidValueTypeError);
      expect(mod.WorksheetNameError).toBe(mod.ExcelWorksheetNameError);
    }
  });

  it("names the table error after its class", () => {
    expect(new ExcelEntry.ExcelTableError("x").name).toBe("ExcelTableError");
  });
});
