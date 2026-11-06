/**
 * Round-3 review fixes in the CSV and Markdown bridges: the sheet-name rule, typed errors for a non-workbook, a
 * non-string input and non-object options.
 */
import { ExcelArgumentError, Workbook, Worksheet, ExcelWorksheetNameError } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { ExcelMarkdown } from "@excel/bridge/markdown";
import { describe, expect, it } from "vitest";

const MD = "| a |\n| --- |\n| 1 |\n";

describe("R3-C2 the 31-character rule", () => {
  it("ExcelCsv reads", async () => {
    const wb = Workbook.create();
    await expect(ExcelCsv.addWorksheet(wb, "a", { sheetName: "c".repeat(40) })).rejects.toThrow(
      ExcelWorksheetNameError
    );
    await expect(ExcelCsv.read("a", { sheetName: "c".repeat(40) })).rejects.toThrow(
      ExcelWorksheetNameError
    );
    expect(Workbook.listWorksheets(wb)).toEqual([]);
  });

  it("ExcelMarkdown reads, before any sheet is added", () => {
    const wb = Workbook.create();
    expect(() => ExcelMarkdown.addWorksheet(wb, MD, { sheetName: "m".repeat(40) })).toThrow(
      ExcelWorksheetNameError
    );
    expect(Workbook.listWorksheets(wb)).toEqual([]);
    // A name this bridge makes up (`…_2`) is shortened to fit rather than refused: the caller never wrote it. See
    // `r4-bridges.test.ts`.
    expect(
      ExcelMarkdown.addWorksheets(wb, `${MD}\n${MD}`, { sheetName: "m".repeat(30) }).map(
        ws => Worksheet.getName(ws).length
      )
    ).toEqual([30, 31]);
  });
});

describe("R3-C5 / C6 typed errors", () => {
  it("format of something that is not a workbook", () => {
    expect(() => ExcelCsv.format(null as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.format(null as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.format({} as never)).toThrow(ExcelArgumentError);
  });

  it("ExcelMarkdown.read of something that is not text, as ExcelCsv.read refuses it", async () => {
    expect(() => ExcelMarkdown.read(42 as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.addWorksheet(Workbook.create(), 42 as never)).toThrow(
      ExcelArgumentError
    );
    await expect(ExcelCsv.read(42 as never)).rejects.toThrow(ExcelArgumentError);
  });

  it("options that are not an object", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    expect(() => ExcelCsv.format(wb, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.format(wb, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.read(MD, 42 as never)).toThrow(ExcelArgumentError);
    await expect(ExcelCsv.read("a", 42 as never)).rejects.toThrow(ExcelArgumentError);
  });
});

describe("X6 every argument error of the bridges is an ExcelArgumentError naming the call", () => {
  function caught(fn: () => unknown): ExcelArgumentError {
    try {
      fn();
    } catch (error) {
      expect(error).toBeInstanceOf(ExcelArgumentError);
      return error as ExcelArgumentError;
    }
    throw new Error("expected an error");
  }
  async function rejected(promise: Promise<unknown>): Promise<ExcelArgumentError> {
    const error = await promise.then(
      () => undefined,
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(ExcelArgumentError);
    return error as ExcelArgumentError;
  }
  function book() {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    return wb;
  }

  it("ExcelCsv parse and format options", async () => {
    let e = await rejected(ExcelCsv.read("a", { delimiter: 42 as never }));
    expect(e.field).toBe("options.delimiter");
    expect(e.message).toMatch(/^ExcelCsv\.read: /);
    e = await rejected(ExcelCsv.addWorksheet(book(), "a", { quote: 7 as never }));
    expect(e.message).toMatch(/^ExcelCsv\.addWorksheet: /);
    e = caught(() => ExcelCsv.format(book(), { delimiter: 42 as never }));
    expect(e.field).toBe("options.delimiter");
    expect(e.message).toMatch(/^ExcelCsv\.format: /);
    e = caught(() => ExcelCsv.toBufferSync(book(), { lineEnding: 1 as never }));
    expect(e.message).toMatch(/^ExcelCsv\.toBufferSync: /);
    e = caught(() => ExcelCsv.createImportStream(book(), { delimiter: 42 as never }));
    expect(e.message).toMatch(/^ExcelCsv\.createImportStream: /);
    e = caught(() => ExcelCsv.format(book(), { sheetName: "nope" }));
    expect(e.message).toMatch(/^ExcelCsv\.format: /);
    e = await rejected(ExcelCsv.read(42 as never));
    expect(e.message).toMatch(/^ExcelCsv\.read: /);
    e = await rejected(ExcelCsv.read("a", { dateFormats: ["nope"] as never }));
    expect(e.field).toBe("options.dateFormats");
  });

  it("ExcelMarkdown read and format options", () => {
    let e = caught(() => ExcelMarkdown.read(MD, { sheetName: 42 as never }));
    expect(e.field).toBe("options.sheetName");
    expect(e.message).toMatch(/^ExcelMarkdown\.read: /);
    e = caught(() => ExcelMarkdown.format(book(), { alignment: 42 as never }));
    expect(e.field).toBe("options.alignment");
    expect(e.message).toMatch(/^ExcelMarkdown\.format: /);
    e = caught(() => ExcelMarkdown.toBufferSync(book(), { padding: "yes" as never }));
    expect(e.message).toMatch(/^ExcelMarkdown\.toBufferSync: /);
    e = caught(() => ExcelMarkdown.read(42 as never));
    expect(e.message).toMatch(/^ExcelMarkdown\.read: /);
    e = caught(() => ExcelMarkdown.addWorksheets(book(), MD, 42 as never));
    expect(e.message).toMatch(/^ExcelMarkdown\.addWorksheets: /);
    e = caught(() => ExcelMarkdown.format(book(), { sheetName: "nope" }));
    expect(e.message).toMatch(/^ExcelMarkdown\.format: /);
  });

  it("content the parser cannot read is not an argument error", () => {
    // Malformed CSV text is the data's fault: it keeps the CSV module's error.
    expect(() => ExcelMarkdown.read("not a table")).not.toThrow(ExcelArgumentError);
  });
});
