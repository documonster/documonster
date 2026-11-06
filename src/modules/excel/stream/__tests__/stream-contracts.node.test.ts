/**
 * The streaming writer and reader check what they are given as the document model does, so a streamed workbook can be
 * read back and nothing a caller wrote is dropped in silence.
 */
import { Readable, Writable } from "node:stream";

import {
  Cell,
  ExcelArgumentError,
  ExcelStreamStateError,
  ExcelImageError,
  Stream,
  Workbook,
  Worksheet,
  ExcelWorksheetNameError
} from "@excel";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

function writer() {
  const chunks: Uint8Array[] = [];
  const stream = new Writable({
    write(chunk, _encoding, callback) {
      chunks.push(chunk);
      callback();
    }
  });
  const wb = new Stream.WorkbookWriter({ stream });
  return { wb, bytes: () => Buffer.concat(chunks) };
}

describe("WorkbookWriter.addWorksheet", () => {
  it.each(["S", "s", "a/b", "x".repeat(32), "", "History", "'quoted'"])(
    "refuses %j beside a sheet called S",
    name => {
      const { wb } = writer();
      wb.addWorksheet("S");
      expect(() => wb.addWorksheet(name)).toThrow(ExcelWorksheetNameError);
    }
  );

  it("writes a workbook that reads back", async () => {
    const { wb, bytes } = writer();
    Stream.commitRow(wb.addWorksheet("One").addRow([1]));
    Stream.commitRow(wb.addWorksheet("x".repeat(31)).addRow([2]));
    await wb.commit();
    const back = await Workbook.read(bytes());
    expect(Workbook.listWorksheets(back).map(ws => Worksheet.getName(ws))).toEqual([
      "One",
      "x".repeat(31)
    ]);
  });
});

describe("streamed cells and rows", () => {
  it("setCellNote and setCellComment check the note before the row is committed", () => {
    const { wb } = writer();
    const ws = wb.addWorksheet("S");
    const cell = ws.getCell("A1");
    expect(() => Stream.setCellNote(cell, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Stream.setCellComment(cell, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Stream.setCellComment(cell, { texts: [{ text: 1 }] } as never)).toThrow(
      ExcelArgumentError
    );
    Stream.setCellNote(cell, "fine");
    expect(() => ws.commit()).not.toThrow();
  });

  it("a committed sheet refuses more rows, merges, formats and pictures", () => {
    const { wb } = writer();
    const id = wb.addImage({ buffer: PNG, extension: "png" });
    const ws = wb.addWorksheet("S");
    ws.addRow([1]);
    ws.commit();
    expect(() => ws.addRow([2])).toThrow(ExcelStreamStateError);
    expect(() => ws.getCell("A5")).toThrow(ExcelStreamStateError);
    expect(() => ws.mergeCells("A10:B10")).toThrow(ExcelStreamStateError);
    expect(() => ws.addImage(id, "B2:C3")).toThrow(ExcelStreamStateError);
    expect(() =>
      ws.addConditionalFormatting({
        ref: "A1",
        rules: [{ type: "expression", formulae: ["TRUE"] }]
      })
    ).toThrow(ExcelStreamStateError);
    expect(ws.findRow(1)).toBeUndefined();
    const seen: number[] = [];
    ws.eachRow((_row, n) => seen.push(n));
    expect(seen).toEqual([]);
  });

  it("pictures must name an image of the workbook", () => {
    const { wb } = writer();
    const ws = wb.addWorksheet("S");
    expect(() => ws.addImage(999, "B2:C3")).toThrow(ExcelImageError);
    expect(() => ws.addBackgroundImage(999)).toThrow(ExcelImageError);
    expect(() => wb.addImage({} as never)).toThrow(ExcelImageError);
    expect(ws.getImages()).toEqual([]);
  });

  it("conditional formatting is checked as ConditionalFormatting.add checks it, and copied", async () => {
    const { wb, bytes } = writer();
    const ws = wb.addWorksheet("S");
    expect(() =>
      ws.addConditionalFormatting({ ref: "A1", rules: [{ type: "bogus" }] } as never)
    ).toThrow(ExcelArgumentError);
    const cf = { ref: "A1:A3", rules: [{ type: "expression" as const, formulae: ["A1>1"] }] };
    ws.addConditionalFormatting(cf);
    cf.rules[0].formulae[0] = "FALSE";
    ws.addRow([1]);
    ws.commit();
    await wb.commit();
    const back = await Workbook.read(bytes());
    const sheet = Workbook.getWorksheet(back, "S")!;
    expect(Cell.getValue(sheet, "A1")).toBe(1);
    const { ConditionalFormatting } = await import("@excel");
    expect(ConditionalFormatting.list(sheet)[0].rules[0]).toMatchObject({ formulae: ["A1>1"] });
  });
});

describe("WorkbookReader options", () => {
  async function book(): Promise<Uint8Array> {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 7);
    return Workbook.toBuffer(wb);
  }

  it.each([{ worksheets: "emitt" }, { styles: "bogus" }, { sharedStrings: 1 }, { entries: true }])(
    "refuses %j",
    options => {
      expect(() => new Stream.WorkbookReader(Readable.from([]), options as never)).toThrow(
        ExcelArgumentError
      );
    }
  );

  it("parse(input, options) keeps the defaults it does not override", async () => {
    const bytes = await book();
    const reader = new Stream.WorkbookReader(Readable.from([bytes]));
    const values: unknown[] = [];
    for await (const event of reader.parse(undefined, { styles: "cache" })) {
      if (event.eventType === "worksheet") {
        for await (const rows of event.value.parse()) {
          for (const row of rows) {
            if (row.eventType === "row") {
              values.push(Stream.getRowValues(row.value)[0]);
            }
          }
        }
      }
    }
    expect(values).toEqual([7]);
  });
});
