import { Column, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

// excelts 10.2.0's `col.collapsed`: an outline level at or above the sheet's `outlineLevelCol`.
describe("Column.isCollapsed", () => {
  it("answers as excelts's col.collapsed did", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(Column.isCollapsed(ws, "B")).toBe(false);
    Column.setOutlineLevel(ws, "B", 1);
    expect(Column.isCollapsed(ws, "B")).toBe(true);
    ws.properties.outlineLevelCol = 2;
    expect(Column.isCollapsed(ws, 2)).toBe(false);
    Column.setOutlineLevel(ws, 2, 2);
    expect(Column.isCollapsed(ws, 2)).toBe(true);
    expect(Column.has(ws, "Z")).toBe(false);
    expect(Column.isCollapsed(ws, "Z")).toBe(false);
    expect(Column.has(ws, "Z")).toBe(false);
  });
});
