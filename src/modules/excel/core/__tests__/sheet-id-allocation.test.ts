/**
 * A sheet id is never handed out twice in a session, as Excel does it: Excel 16.114, given sheets 1, 2, 3, gave a
 * sheet added after deleting sheet 3 the id 4, and one added after deleting sheet 1 the id 5
 * (`tmp/office-tests/r6-sheetid/README.md`). Worksheets and chartsheets share the one pool. The lowest free id used
 * to be reused, so a numeric id held across a remove named a different sheet afterwards.
 */
import { Chartsheet, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const CHART = { chart: { type: "bar", series: [{ values: "A!$A$1:$A$2" }] } } as never;

describe("sheet ids", () => {
  it("are not reused after the highest is removed, nor after a lower one", () => {
    const wb = Workbook.create();
    for (const name of ["A", "B", "C"]) {
      Workbook.addWorksheet(wb, name);
    }
    Workbook.removeWorksheet(wb, "C");
    expect(Workbook.addWorksheet(wb, "D").id).toBe(4);
    Workbook.removeWorksheet(wb, "A");
    expect(Workbook.addWorksheet(wb, "E").id).toBe(5);
    expect(Workbook.getWorksheet(wb, 1)).toBeUndefined();
    expect(Workbook.getWorksheet(wb, 3)).toBeUndefined();
  });

  it("are shared with chartsheets, and a worksheet does not inherit a removed chartsheet's", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "A");
    Workbook.addChartsheet(wb, "CS1", CHART);
    Workbook.addChartsheet(wb, "CS2", CHART);
    expect(Workbook.listChartsheets(wb).map(cs => Chartsheet.getId(cs))).toEqual([2, 3]);
    Workbook.removeChartsheet(wb, "CS2");
    expect(Workbook.addWorksheet(wb, "B").id).toBe(4);
    expect(Chartsheet.getId(Workbook.addChartsheet(wb, "CS3", CHART))).toBe(5);
  });

  it("continue past the highest id of a workbook that was read", async () => {
    const wb = Workbook.create();
    for (const name of ["A", "B", "C"]) {
      Workbook.addWorksheet(wb, name);
    }
    Workbook.removeWorksheet(wb, "A");
    const read = await Workbook.read(await Workbook.toBuffer(wb));
    expect(Workbook.listWorksheets(read).map(ws => ws.id)).toEqual([2, 3]);
    expect(Workbook.addWorksheet(read, "D").id).toBe(4);
  });
});
