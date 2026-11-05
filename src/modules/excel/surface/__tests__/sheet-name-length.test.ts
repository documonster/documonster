/**
 * A sheet name a caller chose is never silently shortened: past Excel's 31 characters it is refused. A file read from
 * another program still has its long names shortened, as before.
 */
import { Cell, Workbook, Worksheet, ExcelWorksheetNameError } from "@excel/index";
import { describe, expect, it } from "vitest";

const LONG = "x".repeat(32);

describe("sheet names past 31 characters", () => {
  it("are refused by Workbook.addWorksheet", () => {
    const wb = Workbook.create();
    expect(() => Workbook.addWorksheet(wb, LONG)).toThrow(ExcelWorksheetNameError);
    expect(Workbook.listWorksheets(wb)).toEqual([]);
    expect(Worksheet.getName(Workbook.addWorksheet(wb, "x".repeat(31)))).toBe("x".repeat(31));
  });

  it("are refused by Worksheet.setName and Worksheet.setModel, which leave the name as it was", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => Worksheet.setName(ws, LONG)).toThrow(ExcelWorksheetNameError);
    expect(() => Worksheet.setModel(ws, { ...Worksheet.getModel(ws), name: LONG })).toThrow(
      ExcelWorksheetNameError
    );
    expect(Worksheet.getName(ws)).toBe("S");
  });

  it("are refused by Workbook.importWorksheet", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => Workbook.importWorksheet(wb, ws, LONG)).toThrow(ExcelWorksheetNameError);
    expect(Workbook.listWorksheets(wb)).toEqual([ws]);
  });

  it("are refused by Workbook.addChartsheet", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    expect(() =>
      Workbook.addChartsheet(wb, LONG, {
        chart: { type: "bar", series: [{ values: "S!$A$1:$A$1" }] }
      } as never)
    ).toThrow(ExcelWorksheetNameError);
    expect(Workbook.listChartsheets(wb)).toEqual([]);
  });
});
