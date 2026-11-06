/**
 * Readers that hand back a snapshot hand back a copy: editing what they return changes nothing, and nothing a caller
 * edits can reach the workbook around the setters' checks.
 */
import { Cell, Column, ExcelArgumentError, Row, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "B1", { richText: [{ text: "rich" }] });
  Cell.setFont(ws, "A1", { bold: true });
  Row.setHeight(ws, 1, 20);
  return { wb, ws };
}

describe("Row.getModel", () => {
  it("is a copy", async () => {
    const { wb, ws } = sheet();
    const model = Row.getModel(ws, 1)!;
    model.cells[0].value = NaN;
    model.cells[0].style!.font = { bold: false };
    (model.cells[1].value as { richText: { text: string }[] }).richText[0].text = "changed";
    model.height = 99;
    expect(Cell.getValue(ws, "A1")).toBe(1);
    expect(Cell.getFont(ws, "A1")).toEqual({ bold: true });
    expect(Cell.getValue(ws, "B1")).toEqual({ richText: [{ text: "rich" }] });
    expect(Row.getHeight(ws, 1)).toBe(20);
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Cell.getValue(back, "A1")).toBe(1);
  });
});

describe("Column.getDefinition / getDefinitions", () => {
  it("are copies, style included", () => {
    const { ws } = sheet();
    Column.setWidth(ws, "A", 10);
    Column.setStyle(ws, "A", { font: { bold: true } });
    const [first] = Column.getDefinitions(ws);
    first.width = 99;
    first.style!.font!.bold = false;
    const one = Column.getDefinition(ws, "A");
    one.style!.font!.italic = true;
    expect(Column.getWidth(ws, "A")).toBe(10);
    expect(Column.findStyle(ws, "A")).toEqual({ font: { bold: true } });
  });
});

describe("Worksheet.getPanes / getAutoFilter", () => {
  it("are copies", () => {
    const { ws } = sheet();
    Worksheet.freeze(ws, { rows: 1 });
    (Worksheet.getPanes(ws) as { rows: number }).rows = 7;
    expect(Worksheet.getPanes(ws)).toEqual({ state: "frozen", columns: 0, rows: 1 });
    expect(ws.views[0]).toMatchObject({ ySplit: 1 });
    Worksheet.setAutoFilter(ws, { from: "A1", to: "B3" });
    (Worksheet.getAutoFilter(ws) as { to: string }).to = "Z99";
    expect(Worksheet.getAutoFilter(ws)).toEqual({ from: "A1", to: "B3" });
  });
});

describe("named cell styles", () => {
  it("are handed out as copies, so an edit cannot skip the style checks", async () => {
    const { wb } = sheet();
    const added = Workbook.addCellStyle(wb, "Live", { font: { bold: true } });
    added.font = { size: -1 };
    Workbook.getCellStyle(wb, "Live")!.font!.bold = false;
    Workbook.listCellStyles(wb)[0].font = { name: 42 as never };
    expect(Workbook.getCellStyle(wb, "Live")).toEqual({ name: "Live", font: { bold: true } });
    await expect(Workbook.toBuffer(wb)).resolves.toBeDefined();
  });
});

describe("external links", () => {
  it("are live, and an edit that breaks one is refused when the workbook is written", async () => {
    const { wb } = sheet();
    const link = Workbook.addExternalLink(wb, { target: "ref.xlsx" });
    link.cachedValues = { Sheet1: { A1: 5 } } as never;
    expect(Workbook.getExternalLink(wb, 1)?.cachedValues).toEqual({ Sheet1: { A1: 5 } });
    link.target = "";
    await expect(Workbook.toBuffer(wb)).rejects.toThrow(ExcelArgumentError);
  });
});
