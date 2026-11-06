/**
 * Optional settings travel in one trailing options object; the positional forms these members had stay
 * as deprecated overloads with their old behaviour. Also: the read-only `findStyle` readers, and the
 * names that replaced `eachSheet` / `importSheet` / `defineCellStyle` / `useBuiltinCellStyle` and the
 * chart / format / cell aliases.
 */
import {
  Anchor,
  Cell,
  Chart,
  Chartsheet,
  Column,
  Format,
  Note,
  Row,
  Table,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

describe("options objects", () => {
  it("freeze / split take their counts as options, and the positional forms still work", () => {
    const { ws } = sheet();
    Worksheet.freeze(ws, { columns: 1, rows: 2 });
    const byOptions = Worksheet.getModel(ws).views;
    Worksheet.freeze(ws, 1, 2);
    expect(Worksheet.getModel(ws).views).toEqual(byOptions);
    expect(byOptions[0]).toMatchObject({
      state: "frozen",
      xSplit: 1,
      ySplit: 2,
      topLeftCell: "B3"
    });

    Worksheet.split(ws, { x: 2000, y: 1000 });
    const split = Worksheet.getModel(ws).views;
    Worksheet.split(ws, 2000, 1000);
    expect(Worksheet.getModel(ws).views).toEqual(split);
    expect(split[0]).toMatchObject({ state: "split", xSplit: 2000, ySplit: 1000 });

    Worksheet.freeze(ws, {});
    expect(Worksheet.getModel(ws).views).toEqual([]);
  });

  it("protect takes the password in its options and does not store it as a flag", async () => {
    const { wb, ws } = sheet();
    await Worksheet.protect(ws, { password: "pw", formatCells: true });
    expect(await Worksheet.verifyPassword(ws, "pw")).toBe(true);
    const stored = Worksheet.getModel(ws).sheetProtection as Record<string, unknown>;
    expect(stored.formatCells).toBe(true);
    expect(stored).not.toHaveProperty("password");

    await Worksheet.protect(ws, "old", { formatCells: true });
    expect(await Worksheet.verifyPassword(ws, "old")).toBe(true);

    await Workbook.protect(wb, { password: "book", lockStructure: true });
    expect(Workbook.getModel(wb).protection).toMatchObject({ lockStructure: true });
    expect(Workbook.getModel(wb).protection).not.toHaveProperty("password");
    await Workbook.protect(wb, "book", { lockStructure: true });
    expect(Workbook.getModel(wb).protection).toMatchObject({ lockStructure: true });
  });

  it("registerPerson takes the identity as options, and the positional form registers the same person", () => {
    const { wb } = sheet();
    const id = Workbook.registerPerson(wb, "Ada", { userId: "ada@x", providerId: "AD" });
    expect(Workbook.registerPerson(wb, "Ada", "ada@x", "AD")).toBe(id);
    expect(Workbook.registerPerson(wb, "Ada")).not.toBe(id);
  });

  it("Note.create and Anchor.create take their extras as options", () => {
    expect(Note.create("hi", { author: "Al" })).toEqual({ note: "hi", author: "Al" });
    expect(Note.create("hi", "Al")).toEqual(Note.create("hi", { author: "Al" }));

    const { ws } = sheet();
    const a = Anchor.create(ws, "C4", { offset: -1 });
    expect([Anchor.getCol(a), Anchor.getRow(a)]).toEqual([2, 3]);
    expect(Anchor.getModel(Anchor.create(ws, "C4", -1))).toEqual(Anchor.getModel(a));
    expect(Anchor.getModel(Anchor.create(ws, "C4"))).toMatchObject({ nativeCol: 3, nativeRow: 4 });
  });
});

describe("fillFormula takes its results and share type as options", () => {
  function filled(fill: (ws: Worksheet.Handle) => void) {
    const { ws } = sheet();
    fill(ws);
    return ["A1", "A2", "B1", "B2"].map(a => Cell.getModel(ws, a));
  }

  it("matches the positional form for a grid, a list, a function and an array formula", () => {
    const grid = [
      [2, 3],
      [3, 4]
    ];
    expect(
      filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()+COLUMN()", { results: grid }))
    ).toEqual(filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()+COLUMN()", grid)));
    expect(
      filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", { results: [1, 1, 2, 2] }))
    ).toEqual(filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", [1, 1, 2, 2])));
    const fn = (row: number, col: number) => row * 10 + col;
    expect(filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", { results: fn }))).toEqual(
      filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", fn))
    );
    const byOptions = filled(ws =>
      Worksheet.fillFormula(ws, "A1:B2", "ROW()", { results: [1, 1, 2, 2], shareType: "array" })
    );
    expect(byOptions).toEqual(
      filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", [1, 1, 2, 2], "array"))
    );
    expect(byOptions[0]).toMatchObject({ shareType: "array", formula: "ROW()", ref: "A1:B2" });
    expect(filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()", {}))).toEqual(
      filled(ws => Worksheet.fillFormula(ws, "A1:B2", "ROW()"))
    );
  });
});

describe("findStyle reads without creating", () => {
  it("Row / Column / Cell", () => {
    const { ws } = sheet();
    expect(Row.findStyle(ws, 5)).toBeUndefined();
    expect(Column.findStyle(ws, "C")).toBeUndefined();
    expect(Cell.findStyle(ws, "B2")).toBeUndefined();
    expect(Row.has(ws, 5)).toBe(false);
    expect(Cell.has(ws, "B2")).toBe(false);
    expect(Column.getDefinitions(ws)).toEqual([]);

    Row.setFont(ws, 5, { bold: true });
    Column.setFont(ws, "C", { italic: true });
    Cell.setStyle(ws, "B2", { font: { size: 14 } });
    const rowStyle = Row.findStyle(ws, 5)!;
    expect(rowStyle.font).toEqual({ bold: true });
    rowStyle.font!.bold = false;
    expect(Row.getFont(ws, 5)).toEqual({ bold: true });
    expect(Column.findStyle(ws, "C")!.font).toEqual({ italic: true });
    const cellStyle = Cell.findStyle(ws, "B2")!;
    expect(cellStyle.font).toEqual({ size: 14 });
    cellStyle.font!.size = 1;
    expect(Cell.findStyle(ws, "B2")!.font).toEqual({ size: 14 });
  });
});

describe("renamed members keep deprecated forwards", () => {
  it("Workbook", () => {
    const { wb, ws } = sheet();
    expect(Workbook.eachSheet).toBe(Workbook.eachWorksheet);
    expect(Workbook.importSheet).toBe(Workbook.importWorksheet);
    expect(Workbook.defineCellStyle).toBe(Workbook.addCellStyle);
    expect(Workbook.useBuiltinCellStyle).toBe(Workbook.useBuiltInCellStyle);
    const entry = Workbook.addCellStyle(wb, "Accent", { font: { bold: true } });
    expect(entry).toEqual(Workbook.getCellStyle(wb, "Accent"));
    Cell.applyStyle(ws, "A1", "Accent");
    expect(Cell.applyCellStyle).toBe(Cell.applyStyle);
    expect(Cell.getStyle(ws, "A1").styleName).toBe("Accent");
  });

  it("Chart / Format", () => {
    expect(Chart.isChartEx).toBe(Chart.isEx);
    expect(Chart.mutateChartEx).toBe(Chart.mutateEx);
    expect(Chart.seriesFromColumns).toBe(Chart.createSeriesFromColumns);
    expect(Chart.chartModel).toBe(Chart.getModel);
    expect(Chart.chartExModel).toBe(Chart.getExModel);
    expect(Format.parseValueByFormat).toBe(Format.parseValue);
  });

  it("Table / Chartsheet boolean readers", () => {
    expect(Table.headerRow).toBe(Table.hasHeaderRow);
    expect(Table.totalsRow).toBe(Table.hasTotalsRow);
    expect(Table.showFirstColumn).toBe(Table.hasFirstColumnStyle);
    expect(Table.showLastColumn).toBe(Table.hasLastColumnStyle);
    expect(Table.showRowStripes).toBe(Table.hasRowStripes);
    expect(Table.showColumnStripes).toBe(Table.hasColumnStripes);
    expect(Table.columnFilterButton).toBe(Table.hasColumnFilterButton);
    expect(Chartsheet.tabSelected).toBe(Chartsheet.isTabSelected);
    expect(Chartsheet.zoomToFit).toBe(Chartsheet.isZoomToFit);
  });
});
