/** `Worksheet.merge` keeps only merges the file keeps, and refuses a range on another sheet. */
import { Cell, InvalidAddressError, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheets() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Workbook.addWorksheet(wb, "Other");
  return { wb, ws };
}

describe("Worksheet.merge", () => {
  it("leaves a single cell as it is, as the written file does", async () => {
    const { wb, ws } = sheets();
    Worksheet.merge(ws, "A1");
    Worksheet.mergeWithoutStyle(ws, 3, 3, 3, 3);
    expect(Worksheet.getMergedRegions(ws)).toEqual([]);
    expect(Cell.isMerged(ws, "A1")).toBe(false);
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Worksheet.getMergedRegions(back)).toEqual([]);
  });

  it("refuses a range qualified by another sheet, and accepts one qualified by this sheet", () => {
    const { ws } = sheets();
    expect(() => Worksheet.merge(ws, "Other!C1:D2")).toThrow(InvalidAddressError);
    expect(() => Worksheet.unmerge(ws, "Other!C1:D2")).toThrow(InvalidAddressError);
    expect(Worksheet.getMergedRegions(ws)).toEqual([]);
    Worksheet.merge(ws, "S!C1:D2");
    expect(Worksheet.getMergedRegions(ws)).toEqual([{ top: 1, left: 3, bottom: 2, right: 4 }]);
  });

  it("refuses a range that names no cells", () => {
    const { ws } = sheets();
    expect(() => Worksheet.merge(ws, "garbage")).toThrow(InvalidAddressError);
    expect(() => Worksheet.merge(ws, 0, 0, 0, 0)).toThrow(InvalidAddressError);
  });
});
