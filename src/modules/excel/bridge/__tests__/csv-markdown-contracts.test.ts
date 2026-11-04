/**
 * A write names a sheet the workbook has, the value options are checked, and a Markdown table formatted from a
 * workbook reads back with its values.
 */
import { Cell, ExcelArgumentError, Workbook, Worksheet } from "@excel";
import { ExcelCsv } from "@excel/bridge/csv";
import { ExcelMarkdown, readMarkdown } from "@excel/bridge/markdown";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Worksheet.addAoa(ws, [
    ["text", "num", "date", "bool", "err"],
    ["a\nb", 1.5, new Date(Date.UTC(2020, 0, 15, 10, 30)), true, { error: "#N/A" }]
  ]);
  return { wb, ws };
}

describe("ExcelCsv", () => {
  it("refuses a sheet the workbook does not have", async () => {
    const { wb } = book();
    expect(() => ExcelCsv.format(wb, { sheetName: "nope" })).toThrow(ExcelArgumentError);
    expect(() => ExcelCsv.format(wb, { sheetId: 99 })).toThrow(ExcelArgumentError);
    await expect(ExcelCsv.toBuffer(wb, { sheetName: "nope" })).rejects.toThrow(ExcelArgumentError);
    expect(ExcelCsv.format(Workbook.create())).toBe("");
    expect(ExcelCsv.format(wb, { sheetName: "S" })).toContain("text,num");
  });

  it("refuses date options of the wrong kind", async () => {
    const { wb } = book();
    expect(() => ExcelCsv.format(wb, { dateFormat: 42 as never })).toThrow(ExcelArgumentError);
    await expect(ExcelCsv.read("2020-01-01", { dateFormats: ["bogus" as never] })).rejects.toThrow(
      ExcelArgumentError
    );
    await expect(ExcelCsv.read("a", { map: 42 as never })).rejects.toThrow(ExcelArgumentError);
  });

  it("a refused read adds no sheet", async () => {
    const wb = Workbook.create();
    await expect(
      ExcelCsv.addWorksheet(wb, "a", { dateFormats: ["bogus" as never] })
    ).rejects.toThrow(ExcelArgumentError);
    expect(Workbook.listWorksheets(wb)).toEqual([]);
  });
});

describe("ExcelMarkdown", () => {
  it("refuses a sheet the workbook does not have", () => {
    const { wb } = book();
    expect(() => ExcelMarkdown.format(wb, { sheetName: "nope" })).toThrow(ExcelArgumentError);
    expect(() => ExcelMarkdown.format(wb, { sheetId: 99 })).toThrow(ExcelArgumentError);
    expect(ExcelMarkdown.format(Workbook.create())).toBe("");
  });

  it("reads back what it formats", () => {
    const { wb } = book();
    const md = ExcelMarkdown.format(wb, { dateUTC: true });
    const back = ExcelMarkdown.read(md);
    const ws = Workbook.listWorksheets(back)[0];
    expect(["A1", "B1"].map(a => Cell.getValue(ws, a))).toEqual(["text", "num"]);
    expect(Cell.getValue(ws, "A2")).toBe("a\nb");
    expect(Cell.getValue(ws, "B2")).toBe(1.5);
    expect(Cell.getValue(ws, "C2")).toEqual(new Date(Date.UTC(2020, 0, 15, 10, 30)));
    expect(Cell.getValue(ws, "D2")).toBe(true);
    expect(Cell.getValue(ws, "E2")).toEqual({ error: "#N/A" });
  });

  it('addWorksheet and read({ tables: "all" }) read the same way; map and convertBr still decide', () => {
    const md = "| a | b |\n| --- | --- |\n| 1 | x<br>y |\n";
    const wb = Workbook.create();
    const ws = ExcelMarkdown.addWorksheet(wb, md);
    expect([Cell.getValue(ws, "A2"), Cell.getValue(ws, "B2")]).toEqual([1, "x\ny"]);
    const all = Workbook.listWorksheets(ExcelMarkdown.read(`${md}\n${md}`, { tables: "all" }));
    expect(all.map(sheet => Cell.getValue(sheet, "A2"))).toEqual([1, 1]);
    const raw = ExcelMarkdown.addWorksheet(wb, md, { map: value => value, convertBr: false });
    expect([Cell.getValue(raw, "A2"), Cell.getValue(raw, "B2")]).toEqual(["1", "x<br>y"]);
  });

  it("the deprecated flat read keeps reading text", () => {
    const wb = Workbook.create();
    const ws = readMarkdown(wb, "| a |\n| --- |\n| 1 |\n");
    expect(Cell.getValue(ws, "A2")).toBe("1");
  });
});
