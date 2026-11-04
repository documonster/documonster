/**
 * A legacy array formula (`{=A1:A3}`, `shareType: "array"`) computes every cell of its range. Excel stores the master
 * with the formula and every other cell of the range as a plain cached value (`<c r="B27"><v>0</v></c>` in Excel 16's
 * save of `tmp/office-tests/r4-formula/r4-blank.xlsx`), and Excel computes 1, 0, 3 for `{=E!A1:A3}` over 1, empty, 3.
 * The writeback used to write only cells already holding the formula, so the rest of the range kept nothing, or kept
 * a stale value. A dynamic array's spill is checked alongside.
 */
import { Archive } from "@archive";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "A3", 3);
  return { wb, ws };
}

const column = (ws: Parameters<typeof Cell.getResult>[0], col: string, rows = [1, 2, 3]) =>
  rows.map(row => {
    const value = Cell.getValue(ws, `${col}${row}`);
    return typeof value === "object" && value !== null && "formula" in value ? value.result : value;
  });

async function sheetXml(bytes: Uint8Array): Promise<string> {
  const files = await Archive.unpackZip(bytes);
  return new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
}

describe("legacy array formula writeback", () => {
  it("writes every cell of the range", () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "B1", { formula: "A1:A3", shareType: "array", ref: "B1:B3" });
    calculateFormulas(wb);
    expect(column(ws, "B")).toEqual([1, 0, 3]);
    // The master keeps the formula; the others are cached values of it, as Excel stores them.
    expect(Cell.getValue(ws, "B1")).toMatchObject({
      formula: "A1:A3",
      shareType: "array",
      ref: "B1:B3"
    });
    expect(Cell.getValue(ws, "B2")).toBe(0);
  });

  it("pads past the result with #N/A, and fills a two-dimensional range with a single value", () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "D1", { formula: "A1:A2*10", shareType: "array", ref: "D1:D3" });
    Cell.setValue(ws, "G1", { formula: "1+1", shareType: "array", ref: "G1:H2" });
    calculateFormulas(wb);
    expect(column(ws, "D")).toEqual([10, 0, { error: "#N/A" }]);
    expect([...column(ws, "G", [1, 2]), ...column(ws, "H", [1, 2])]).toEqual([2, 2, 2, 2]);
  });

  it("replaces the stale cached values a file carried", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "B1", { formula: "A1:A3", shareType: "array", ref: "B1:B3", result: 7 });
    Cell.setValue(ws, "B2", 8);
    Cell.setValue(ws, "B3", 9);
    const read = await Workbook.read(await Workbook.toBuffer(wb));
    const sheet = Workbook.getWorksheet(read, "S")!;
    calculateFormulas(read);
    expect(column(sheet, "B")).toEqual([1, 0, 3]);
  });

  it.each(["xlsx", "xlsb"] as const)("writes and re-reads the range through %s", async format => {
    const { wb, ws } = book();
    Cell.setValue(ws, "B1", { formula: "A1:A3", shareType: "array", ref: "B1:B3" });
    calculateFormulas(wb);
    const bytes = await Workbook.toBuffer(wb, { format });
    if (format === "xlsx") {
      const xml = await sheetXml(bytes);
      expect(xml).toContain('<c r="B1"><f t="array" ref="B1:B3">A1:A3</f><v>1</v></c>');
      expect(xml).toContain('<c r="B2"><v>0</v></c>');
      expect(xml).toContain('<c r="B3"><v>3</v></c>');
    }
    const read = Workbook.getWorksheet(await Workbook.read(bytes), "S")!;
    expect(Cell.getValue(read, "B1")).toMatchObject({
      formula: "A1:A3",
      shareType: "array",
      ref: "B1:B3",
      result: 1
    });
    expect(column(read, "B")).toEqual([1, 0, 3]);
  });

  it("recomputes the same range after a read-back", async () => {
    const { wb, ws } = book();
    Cell.setValue(ws, "B1", { formula: "A1:A3", shareType: "array", ref: "B1:B3" });
    calculateFormulas(wb);
    const read = await Workbook.read(await Workbook.toBuffer(wb));
    const sheet = Workbook.getWorksheet(read, "S")!;
    Cell.setValue(sheet, "A2", 5);
    calculateFormulas(read);
    expect(column(sheet, "B")).toEqual([1, 5, 3]);
  });
});

/**
 * Excel stores a spilled dynamic array as its master, with `ref` covering the spill, and every other cell of the spill
 * as a plain cached value. Those values are the formula's own spill, not data in its way: they used to block it, and
 * the master read back from Excel's own file computed `#SPILL!`.
 */
async function spilledFile(format: "xlsx" | "xlsb", formula = "A1:A3", ref = "B1:B3") {
  const { wb, ws } = book();
  Cell.setValue(ws, "B1", { formula, shareType: "array", ref, isDynamicArray: true, result: 1 });
  Cell.setValue(ws, "B2", 0);
  Cell.setValue(ws, "B3", 3);
  const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
  return { read, sheet: Workbook.getWorksheet(read, "S")! };
}

describe.each(["xlsx", "xlsb"] as const)("a dynamic array's cached spill read from %s", format => {
  it("is the formula's own spill, not a blockage", async () => {
    const { read, sheet } = await spilledFile(format);
    Cell.setValue(sheet, "A2", 5);
    calculateFormulas(read);
    expect(column(sheet, "B")).toEqual([1, 5, 3]);
  });

  it("is cleared where the spill no longer reaches", async () => {
    const { read, sheet } = await spilledFile(format, "SEQUENCE(A5)");
    Cell.setValue(sheet, "A5", 2);
    calculateFormulas(read);
    expect(column(sheet, "B")).toEqual([1, 2, null]);
  });

  it("blocks the spill once a value is set over it", async () => {
    const { read, sheet } = await spilledFile(format);
    Cell.setValue(sheet, "B2", "mine");
    calculateFormulas(read);
    expect(Cell.getResult(sheet, "B1")).toEqual({ error: "#SPILL!" });
    expect(Cell.getValue(sheet, "B2")).toBe("mine");
  });

  it("is written back as it came when nothing is calculated", async () => {
    const { read } = await spilledFile(format);
    const again = Workbook.getWorksheet(
      await Workbook.read(await Workbook.toBuffer(read, { format })),
      "S"
    )!;
    expect(column(again, "B")).toEqual([1, 0, 3]);
  });
});
