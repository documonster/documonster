/**
 * Round-4 review fixes in the CSV and Markdown bridges: `encoding` decodes every byte input, write options are checked
 * as the sibling bridge checks them, generated sheet names fit, and each entry exports the errors it throws.
 */
import { CsvError as CsvModuleError, isCsvError as isCsvModuleError } from "@csv";
import * as ExcelEntry from "@excel";
import { Row, Workbook, Worksheet, WorksheetNameError } from "@excel";
import * as CsvEntry from "@excel/bridge/csv";
import { ExcelCsv } from "@excel/bridge/csv";
import * as FormulaEntry from "@excel/bridge/formula";
import * as MarkdownEntry from "@excel/bridge/markdown";
import { ExcelMarkdown } from "@excel/bridge/markdown";
import { describe, expect, it } from "vitest";

/** `n\ncafé\n` in windows-1252: `é` is the single byte 0xE9, which UTF-8 cannot read. */
const WIN1252 = new Uint8Array([0x6e, 0x0a, 0x63, 0x61, 0x66, 0xe9, 0x0a]);

async function cell(workbook: Promise<ReturnType<typeof Workbook.create>>): Promise<unknown> {
  return Row.getValues(Workbook.listWorksheets(await workbook)[0]!, 2)[0];
}

describe("X1 / X2 ExcelCsv decodes every byte input as options.encoding says", () => {
  it.each<[string, () => unknown]>([
    ["Uint8Array", () => WIN1252],
    ["ArrayBuffer", () => WIN1252.slice().buffer],
    ["Blob", () => new Blob([WIN1252])],
    ["File", () => new File([WIN1252], "data.csv")],
    [
      "a byte stream",
      () =>
        (async function* () {
          yield WIN1252.subarray(0, 4);
          yield WIN1252.subarray(4);
        })()
    ]
  ])("%s", async (_kind, input) => {
    expect(await cell(ExcelCsv.read(input() as never, { encoding: "windows-1252" }))).toBe("café");
  });

  it("a File streams, and reports progress, without FileReader", async () => {
    const seen: [number, number][] = [];
    const file = new File([WIN1252], "data.csv");
    expect(
      await cell(
        ExcelCsv.read(file, { encoding: "windows-1252", onProgress: (a, b) => seen.push([a, b]) })
      )
    ).toBe("café");
    expect(seen.at(-1)).toEqual([WIN1252.length, WIN1252.length]);
    expect(await cell(ExcelCsv.read(file, { encoding: "windows-1252", stream: true }))).toBe(
      "café"
    );
  });

  it("defaults to UTF-8 and refuses a label TextDecoder does not know", async () => {
    expect(await cell(ExcelCsv.read(new TextEncoder().encode("n\ncafé\n")))).toBe("café");
    await expect(ExcelCsv.read(WIN1252, { encoding: "bogus" })).rejects.toThrow(
      expect.objectContaining({ field: "options.encoding" })
    );
    await expect(ExcelCsv.read("a", { onProgress: 42 } as never)).rejects.toThrow(
      expect.objectContaining({ field: "options.onProgress" })
    );
  });
});

describe("R4-C3 / X8 ExcelCsv write options are checked", () => {
  function book() {
    const wb = Workbook.create();
    Row.add(Workbook.addWorksheet(wb, "S"), ["h"]);
    return wb;
  }

  it.each<[string, Record<string, unknown>]>([
    ["options.includeEmptyRows", { includeEmptyRows: "no" }],
    ["options.append", { append: "false" }],
    ["options.highWaterMark", { highWaterMark: -1 }],
    ["options.dateUTC", { dateUTC: "no" }]
  ])("%s", (field, options) => {
    expect(() => ExcelCsv.format(book(), options as never)).toThrow(
      expect.objectContaining({ field })
    );
  });

  it("the writers that produce UTF-8 refuse another encoding rather than ignore it", async () => {
    expect(() => ExcelCsv.toBufferSync(book(), { encoding: "latin1" })).toThrow(
      expect.objectContaining({ field: "options.encoding" })
    );
    await expect(ExcelCsv.toBuffer(book(), { encoding: "utf16le" })).rejects.toThrow(
      /ExcelCsv\.writeFile only/
    );
    expect(new TextDecoder().decode(ExcelCsv.toBufferSync(book(), { encoding: "UTF-8" }))).toBe(
      "h"
    );
  });

  it("uses the Markdown bridge's wording for the options both share", () => {
    for (const run of [
      () => ExcelCsv.format(book(), { dateUTC: "no" } as never),
      () => ExcelMarkdown.format(book(), { dateUTC: "no" } as never)
    ]) {
      expect(run).toThrow(/options\.dateUTC must be a boolean, got "no"/);
    }
  });
});

describe("R4-C7 ExcelMarkdown makes up only names it can keep", () => {
  const twoTables = "| a |\n|---|\n| 1 |\n\n| b |\n|---|\n| 2 |\n";

  it("shortens the base so a suffixed name fits 31 characters", () => {
    const base = "m".repeat(31);
    const names = Workbook.listWorksheets(
      ExcelMarkdown.read(twoTables, { sheetName: base, tables: "all" })
    ).map(Worksheet.getName);
    expect(names).toEqual([base, `${"m".repeat(29)}_2`]);
  });

  it("refuses a clash before adding any sheet", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "T_2");
    expect(() => ExcelMarkdown.addWorksheets(wb, twoTables, { sheetName: "T" })).toThrow(
      WorksheetNameError
    );
    expect(Workbook.listWorksheets(wb).map(Worksheet.getName)).toEqual(["T_2"]);
  });
});

describe("cross: the bridge entries export the errors they throw", () => {
  it("as the same class objects documonster/excel exports", () => {
    for (const entry of [CsvEntry, MarkdownEntry]) {
      expect(entry.ExcelArgumentError).toBe(ExcelEntry.ExcelArgumentError);
      expect(entry.ExcelFileError).toBe(ExcelEntry.ExcelFileError);
      expect(entry.WorksheetNameError).toBe(ExcelEntry.WorksheetNameError);
      expect(entry.InvalidValueTypeError).toBe(ExcelEntry.InvalidValueTypeError);
      expect(entry.isExcelError).toBe(ExcelEntry.isExcelError);
    }
    expect(CsvEntry.ExcelDownloadError).toBe(ExcelEntry.ExcelDownloadError);
    expect(CsvEntry.CsvError).toBe(CsvModuleError);
    expect(CsvEntry.isCsvError).toBe(isCsvModuleError);
    expect(MarkdownEntry.MarkdownParseError).toBeDefined();
    expect(FormulaEntry.ExcelArgumentError).toBe(ExcelEntry.ExcelArgumentError);
    expect(FormulaEntry.isExcelError).toBe(ExcelEntry.isExcelError);
  });
});
