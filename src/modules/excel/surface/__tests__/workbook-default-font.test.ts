import { Cell, ExcelArgumentError, Stream, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("Workbook.getDefaultFont / setDefaultFont", () => {
  it("is written as the workbook's default font and read back", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", "x");
    expect(Workbook.getDefaultFont(wb)).toBeUndefined();
    Workbook.setDefaultFont(wb, { name: "Arial", size: 12 });
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(Workbook.getDefaultFont(back)).toMatchObject({ name: "Arial", size: 12 });
  });

  it("copies in and out", () => {
    const wb = Workbook.create();
    const font = { name: "Arial", size: 12 };
    Workbook.setDefaultFont(wb, font);
    font.size = 20;
    const read = Workbook.getDefaultFont(wb)!;
    read.size = 30;
    expect(Workbook.getDefaultFont(wb)).toEqual({ name: "Arial", size: 12 });
  });

  it("resets with undefined", () => {
    const wb = Workbook.create();
    Workbook.setDefaultFont(wb, { name: "Arial" });
    Workbook.setDefaultFont(wb, undefined);
    expect(Workbook.getDefaultFont(wb)).toBeUndefined();
  });

  it("refuses a font that is not one", () => {
    const wb = Workbook.create();
    expect(() => Workbook.setDefaultFont(wb, { size: "big" } as never)).toThrow(ExcelArgumentError);
    expect(Workbook.getDefaultFont(wb)).toBeUndefined();
  });

  it("sets a streaming writer's default font", async () => {
    const writer = new Stream.WorkbookWriter({ useStyles: true });
    Workbook.setDefaultFont(writer, { name: "Arial", size: 12 });
    expect(Workbook.getDefaultFont(writer)).toMatchObject({ name: "Arial", size: 12 });
    const ws = Workbook.addWorksheet(writer, "S");
    Stream.commitRow(ws.addRow(["x"]));
    const back = await Workbook.read(await writer.toBuffer());
    expect(Workbook.getDefaultFont(back)).toMatchObject({ name: "Arial", size: 12 });
  });
});
