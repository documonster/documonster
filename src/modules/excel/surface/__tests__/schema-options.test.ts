/**
 * The public options objects are checked against schemas (`@utils/schema`): every refusal names the option by its
 * path from the parameter and reads as one sentence in the schema's wording, and the options the checks used to skip
 * are checked too.
 */
import { ExcelCsv } from "@excel/bridge/csv";
import {
  Cell,
  Column,
  ExcelArgumentError,
  ExcelColumnOutOfBoundsError,
  Sparkline,
  Stream,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", "h");
  Cell.setValue(ws, "A2", 1);
  return { wb, ws };
}

function refusal(field: string, message?: RegExp) {
  return expect.objectContaining({
    name: "ExcelArgumentError",
    field,
    ...(message ? { message: expect.stringMatching(message) } : {})
  });
}

const CHART = { type: "bar", series: [{ values: "S!$A$1:$A$2" }] } as never;

describe("wording", () => {
  it("reads as the schema words it, without a trailing period", () => {
    const { ws } = sheet();
    expect(() => Worksheet.toJson(ws, { header: 2 } as never)).toThrow(
      'Worksheet.toJson: options.header must be 1, "A" or an array of names, got 2'
    );
    expect(() => Worksheet.freeze(ws, { columns: -1 })).toThrow(
      "Worksheet.freeze: options.columns must be an integer from 0 to 16383, got -1"
    );
    expect(() =>
      Workbook.addWorksheet(Workbook.create(), "T", { useSharedStrings: "yes" } as never)
    ).toThrow('Workbook.addWorksheet: options.useSharedStrings must be a boolean, got "yes"');
  });
});

describe("protect", () => {
  it("names a flag of the options object by its path from options", async () => {
    const { wb, ws } = sheet();
    await expect(Worksheet.protect(ws, { formatCells: "yes" } as never)).rejects.toEqual(
      refusal("options.formatCells", /options\.formatCells must be a boolean/)
    );
    await expect(Worksheet.protect(ws, { password: 42 } as never)).rejects.toEqual(
      refusal("options.password")
    );
    await expect(Workbook.protect(wb, { lockStructure: 1 } as never)).rejects.toEqual(
      refusal("options.lockStructure")
    );
    // The deprecated positional form: its options are the `options` parameter too.
    await expect(Worksheet.protect(ws, "pw", { sort: "no" } as never)).rejects.toEqual(
      refusal("options.sort")
    );
  });
});

describe("Workbook write options", () => {
  it("take null as no options in every write, as the check documents", async () => {
    const { wb } = sheet();
    const bytes = await Workbook.toBuffer(wb, null as never);
    expect(bytes.length).toBeGreaterThan(0);
  });

  it("check toStream's highWaterMark", () => {
    const { wb } = sheet();
    expect(() => Workbook.toStream(wb, { highWaterMark: 0 })).toThrow(
      refusal("options.highWaterMark")
    );
  });

  it("refuse a create option a read does not know the shape of", async () => {
    const bytes = await Workbook.toBuffer(sheet().wb);
    await expect(Workbook.read(bytes, { formulaSyntaxProbe: 5 } as never)).rejects.toEqual(
      refusal("options.formulaSyntaxProbe")
    );
    expect(() => Workbook.create({ formulaSyntaxProbe: null } as never)).not.toThrow();
  });
});

describe("streaming writer and reader options", () => {
  it.each<[string, Record<string, unknown>]>([
    ["options.useStyles", { useStyles: "yes" }],
    ["options.created", { created: "2024-01-01" }],
    ["options.format", { format: "csv" }],
    ["options.zip.zlib.level", { zip: { zlib: { level: 99 } } }],
    ["options.stream", { stream: "out.xlsx" }],
    ["options.filename", { filename: 42 }]
  ])("WorkbookWriter refuses %s", (field, options) => {
    expect(() => new Stream.WorkbookWriter(options as never)).toThrow(refusal(field));
  });

  it("WorkbookReader refuses an array and a negative buffer limit", () => {
    expect(() => new Stream.WorkbookReader(new Uint8Array(0), [] as never)).toThrow(
      refusal("options")
    );
    expect(
      () => new Stream.WorkbookReader(new Uint8Array(0), { maxBufferedWorksheetBytes: -1 })
    ).toThrow(refusal("options.maxBufferedWorksheetBytes"));
  });
});

describe("worksheet options", () => {
  it("checks toJson's flags", () => {
    const { ws } = sheet();
    expect(() => Worksheet.toJson(ws, { raw: "yes" } as never)).toThrow(refusal("options.raw"));
    expect(() => Worksheet.toJson(ws, { dateFormat: 1 } as never)).toThrow(
      refusal("options.dateFormat")
    );
  });

  it("refuses an auto-fit span past the last column instead of counting up to it", () => {
    const { ws } = sheet();
    expect(() => Column.autoFit(ws, { start: 1, end: 2 ** 40 })).toThrow(
      ExcelColumnOutOfBoundsError
    );
    expect(() => Column.autoFit(ws, { includeHiddenRows: "yes" } as never)).toThrow(
      refusal("options.includeHiddenRows")
    );
    expect(() => Column.autoFit(ws, { start: "A", end: "B" })).not.toThrow();
  });

  it("refuses a sparkline without the cell it is drawn in", () => {
    const { ws } = sheet();
    expect(() =>
      Sparkline.add(ws, { type: "line", sparklines: [{ dataRef: "S!A1:A2" }] } as never)
    ).toThrow(refusal("options.sparklines[0].cellRef"));
  });
});

describe("chartsheet options", () => {
  it("are checked as the Chartsheet setters check them", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    expect(() => Workbook.addChartsheet(wb, "C", 42 as never)).toThrow(refusal("options"));
    expect(() => Workbook.addChartsheet(wb, "C", { chart: CHART, state: "gone" } as never)).toThrow(
      refusal("options.state")
    );
    expect(() => Workbook.addChartsheet(wb, "C", { chart: CHART, zoomScale: -1 })).toThrow(
      refusal("options.zoomScale")
    );
    expect(() =>
      Workbook.addChartsheet(wb, "C", { chart: CHART, pageMargins: { side: 1 } } as never)
    ).toThrow(refusal('options.pageMargins["side"]'));
    expect(() =>
      Workbook.addChartsheet(wb, "C", { chart: CHART, pageSetup: { orientation: "up" } } as never)
    ).toThrow(refusal("options.pageSetup.orientation"));
    expect(Workbook.listChartsheets(wb)).toHaveLength(0);
  });
});

describe("ExcelCsv read options", () => {
  it("checks the options the bridge used to pass over", async () => {
    await expect(ExcelCsv.read("a\n1", { decimalSeparator: ";" } as never)).rejects.toBeInstanceOf(
      ExcelArgumentError
    );
    await expect(ExcelCsv.read("a\n1", { requestHeaders: { a: 1 } } as never)).rejects.toEqual(
      refusal('options.requestHeaders["a"]')
    );
    await expect(ExcelCsv.read("a\n1", { signal: "stop" } as never)).rejects.toEqual(
      refusal("options.signal")
    );
  });
});
