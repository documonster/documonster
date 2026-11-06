/**
 * A reader never creates, materialises or mutates anything (AGENTS.md "Public API
 * Conventions" → Accessors). `Cell.getStyle` is the one sanctioned exception; the facet
 * getters and the model reader must answer for an absent cell exactly as they would
 * once it existed, while leaving the sheet — and what it saves — untouched.
 */
import { Cell, Chart, Column, Image, Row, Workbook, Worksheet } from "@excel/index";
import { WorkbookWriter } from "@excel/stream/workbook-writer";
import { describe, expect, it } from "vitest";

function styledSheet(): Worksheet.Handle {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Column.setStyle(ws, "B", {
    numFmt: "0.00",
    font: { bold: true },
    alignment: { horizontal: "center" },
    border: { top: { style: "thin" } },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } },
    protection: { locked: false }
  });
  Row.setStyle(ws, 7, { font: { italic: true } });
  return ws;
}

function shape(ws: Worksheet.Handle): string {
  const m = Worksheet.getModel(ws);
  return JSON.stringify({
    rows: m.rows?.length,
    cols: m.cols,
    dim: m.dimensions,
    rowCount: Row.getCount(ws)
  });
}

const FACET_READERS = [
  "getFont",
  "getAlignment",
  "getBorder",
  "getFill",
  "getProtection",
  "getModel",
  "getMergeMaster"
] as const;

describe("Cell readers leave the sheet untouched", () => {
  for (const name of FACET_READERS) {
    it(`Cell.${name} creates neither the row nor the cell`, () => {
      const ws = styledSheet();
      const before = shape(ws);
      (Cell[name] as (ws: Worksheet.Handle, addr: string) => unknown)(ws, "Q50");
      (Cell[name] as (ws: Worksheet.Handle, r: number, c: number) => unknown)(ws, 60, 2);
      expect(shape(ws)).toBe(before);
      expect(Cell.has(ws, "Q50")).toBe(false);
      expect(Cell.has(ws, "B60")).toBe(false);
    });
  }

  it("answers for an absent cell what the created cell would report", () => {
    const ws = styledSheet();
    const before = shape(ws);
    expect(Cell.getFont(ws, "B50")).toEqual({ bold: true });
    expect(Cell.getAlignment(ws, "B50")).toEqual({ horizontal: "center" });
    expect(Cell.getBorder(ws, "B50")).toEqual({ top: { style: "thin" } });
    expect(Cell.getFill(ws, "B50")).toEqual({
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF0000" }
    });
    expect(Cell.getProtection(ws, "B50")).toEqual({ locked: false });
    expect(Cell.getFont(ws, "Q7")).toEqual({ italic: true });
    expect(Cell.getFont(ws, "B7")).toEqual({ italic: true });
    expect(Cell.getFont(ws, "Q50")).toBeUndefined();
    expect(Cell.getModel(ws, "Q50")).toEqual({ address: "Q50", type: 0, style: {} });
    expect(Cell.getModel(ws, "B50")).toEqual({
      address: "B50",
      type: 0,
      style: {
        numFmt: "0.00",
        font: { bold: true },
        alignment: { horizontal: "center" },
        border: { top: { style: "thin" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } },
        protection: { locked: false }
      }
    });
    expect(Cell.view(Cell.getMergeMaster(ws, "C9")).value).toBeNull();
    expect(Cell.getMergeMaster(ws, "C9").address).toBe("C9");
    expect(shape(ws)).toBe(before);
  });

  it("editing what an absent cell's facet reader returned changes nothing", () => {
    const ws = styledSheet();
    Cell.getFont(ws, "B50")!.bold = false;
    expect(Column.findStyle(ws, "B")?.font).toEqual({ bold: true });
    expect(Cell.getFont(ws, "B51")).toEqual({ bold: true });
  });

  it("editing what an existing cell's facet reader returned changes nothing", () => {
    const ws = styledSheet();
    Cell.setValue(ws, "B3", 1);
    Cell.getFont(ws, "B3")!.italic = true;
    Cell.getAlignment(ws, "B3")!.horizontal = "right";
    Cell.getBorder(ws, "B3")!.top = { style: "thick" };
    Cell.getFill(ws, "B3")!.type = "gradient";
    Cell.getProtection(ws, "B3")!.locked = true;
    expect(Cell.getStyle(ws, "B3")).toEqual(Column.findStyle(ws, "B"));
    expect(Column.findStyle(ws, "B")?.font).toEqual({ bold: true });
  });

  it("a facet is changed through its setter or the live getStyle record", () => {
    const ws = styledSheet();
    Cell.setFont(ws, "B3", { ...Cell.getFont(ws, "B3"), italic: true });
    expect(Cell.getFont(ws, "B3")).toEqual({ bold: true, italic: true });
    Cell.getStyle(ws, "B4").font = { underline: true };
    expect(Cell.getFont(ws, "B4")).toEqual({ underline: true });
    expect(Column.findStyle(ws, "B")?.font).toEqual({ bold: true });
  });

  it("Cell, Row and Column facet readers all return copies", () => {
    const ws = styledSheet();
    Row.setFont(ws, 7, { size: 20 });
    Cell.setValue(ws, "B3", 1);
    const reads = [
      () => Cell.getFont(ws, "B3"),
      () => Cell.getFont(ws, "B60"),
      () => Row.getFont(ws, 7),
      () => Column.getFont(ws, "B")
    ];
    for (const read of reads) {
      const first = read()!;
      expect(read()).not.toBe(first);
      first.strike = true;
      expect(read()!.strike).toBeUndefined();
    }
  });

  it("Row.commit on a document-model sheet creates no row", () => {
    const ws = styledSheet();
    const before = shape(ws);
    Row.commit(ws, 80);
    expect(Row.has(ws, 80)).toBe(false);
    expect(shape(ws)).toBe(before);
    expect(() => Row.commit(ws, 0)).toThrow();
  });

  it("Row.commit still commits a streaming writer's row", () => {
    const wb = new WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    Cell.setValue(ws as never, "A1", "x");
    expect(() => ws.getRow(1)).not.toThrow();
    Row.commit(ws as never, 1);
    expect(() => ws.getRow(1)).toThrow();
  });

  it("Row.eachCell over an absent row visits nothing and creates nothing", () => {
    const ws = styledSheet();
    const before = shape(ws);
    const seen: number[] = [];
    Row.eachCell(ws, 50, (_c, n) => seen.push(n));
    Row.eachCell(ws, 50, { includeEmpty: true }, (_c, n) => seen.push(n));
    expect(seen).toEqual([]);
    expect(shape(ws)).toBe(before);
  });
});

describe("model readers hand out copies all the way down", () => {
  const PNG =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  it("Image.getModel: editing the placement's extent changes nothing", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    const image = Image.place(ws, id, { tl: { col: 1, row: 1 }, ext: { width: 10, height: 20 } });
    const model = Image.getModel(image);
    if (model.type !== "image") {
      throw new Error("expected a picture");
    }
    model.range.ext!.width = 999;
    expect(Image.getModel(image)).toMatchObject({ range: { ext: { width: 10, height: 20 } } });
  });

  it("Chart.getAnchorModel: editing the extent changes nothing", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["a", 1]);
    const chart = Chart.place(
      ws,
      { tl: { col: 3, row: 0 }, ext: { cx: 3000000, cy: 2000000 } },
      {
        type: "bar",
        series: [{ values: "S!$B$1:$B$1", categories: "S!$A$1:$A$1" }]
      }
    );
    Chart.getAnchorModel(chart).range.ext!.cx = 1;
    expect(Chart.getAnchorModel(chart).range.ext).toEqual({ cx: 3000000, cy: 2000000 });
  });

  it("Workbook.getModel: editing it changes nothing, and a second workbook given it shares nothing", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    const model = Workbook.getModel(wb);
    model.properties.date1904 = true;
    model.views.push({ activeTab: 3 } as never);
    expect(Workbook.getModel(wb).properties.date1904).not.toBe(true);
    expect(Workbook.getModel(wb).views).toHaveLength(
      Workbook.getModel(Workbook.create()).views.length
    );

    const other = Workbook.create();
    Workbook.setModel(other, Workbook.getModel(wb));
    Workbook.getModel(other).properties.date1904 = true;
    expect(Workbook.getModel(wb).properties.date1904).not.toBe(true);
    expect("definedNamesInstance" in Workbook.getModel(wb)).toBe(false);
  });
});
