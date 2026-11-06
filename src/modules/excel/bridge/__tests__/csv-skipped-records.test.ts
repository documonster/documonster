import { CsvError } from "@csv/errors";
import { Workbook, Worksheet } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { describe, expect, it } from "vitest";

// With a header row, a record with the wrong number of fields used to vanish from the sheet without a word.
const RAGGED = "a,b\n1\n2,3\n4,5,6\n";

function chunks(text: string): AsyncIterable<string> {
  return (async function* () {
    yield text;
  })();
}

describe("ExcelCsv reads: a record the parser leaves out", () => {
  it("fails the read when nothing receives it, and leaves the workbook as it was", async () => {
    const wb = Workbook.create();
    await expect(ExcelCsv.addWorksheet(wb, RAGGED, { headers: true })).rejects.toThrow(CsvError);
    await expect(ExcelCsv.addWorksheet(wb, RAGGED, { headers: true })).rejects.toThrow(
      /line 2 \[TooFewFields\] \(and 1 more\); pass onSkip/
    );
    expect(Workbook.listWorksheets(wb)).toHaveLength(0);
  });

  it("fails a streamed read too", async () => {
    await expect(ExcelCsv.read(chunks(RAGGED), { headers: true })).rejects.toThrow(/TooFewFields/);
  });

  it("fails an import stream as it is met", async () => {
    const wb = Workbook.create();
    const sink = ExcelCsv.createImportStream(wb, { headers: true });
    const failed = new Promise<unknown>(resolve => sink.on("error", resolve));
    sink.end(RAGGED);
    expect(await failed).toBeInstanceOf(CsvError);
  });

  it("is handed to onSkip, and the rest is imported", async () => {
    const skipped: number[] = [];
    const wb = await ExcelCsv.read(RAGGED, { headers: true, onSkip: e => skipped.push(e.line) });
    expect(skipped).toEqual([2, 4]);
    expect(Worksheet.getValues(Workbook.listWorksheets(wb)[0]!, { base: 0 })).toEqual([
      ["a", "b"],
      [2, 3]
    ]);
  });

  it("is imported when columnMismatch keeps it", async () => {
    const wb = await ExcelCsv.read(RAGGED, {
      headers: true,
      columnMismatch: { less: "pad", more: "truncate" }
    });
    expect(Worksheet.getValues(Workbook.listWorksheets(wb)[0]!, { base: 0 })).toEqual([
      ["a", "b"],
      [1],
      [2, 3],
      [4, 5]
    ]);
  });
});
