/**
 * A reader never creates, materialises or mutates anything (AGENTS.md "Public API
 * Conventions" → Accessors). `Cell.getStyle` is the one sanctioned exception; the facet
 * getters and the model reader must answer for an absent cell exactly as they would
 * once it existed, while leaving the sheet — and what it saves — untouched.
 */
import { Cell, Column, Row, Workbook, Worksheet } from "@excel/index";
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

  it("an existing cell's facet is still live", () => {
    const ws = styledSheet();
    Cell.setValue(ws, "B3", 1);
    Cell.getFont(ws, "B3")!.italic = true;
    expect(Cell.getFont(ws, "B3")).toEqual({ bold: true, italic: true });
    // The column snapshot the cell inherited from is not reached through it.
    expect(Column.findStyle(ws, "B")?.font).toEqual({ bold: true });
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
