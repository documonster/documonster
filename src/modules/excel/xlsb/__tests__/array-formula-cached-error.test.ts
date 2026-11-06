/**
 * An array or dynamic-array formula whose cached result is an error is still an array formula after an XLSB read: the
 * reader's error branch used to drop the `BrtArrFmla`'s range and the dynamic-array flag, so the master read back as
 * an ordinary formula — which Excel intersects.
 */
import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

describe("XLSB array formula with a cached error", () => {
  it("keeps its kind, its range and its dynamic-array flag", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", {
      formula: "SEQUENCE(3)",
      shareType: "array",
      ref: "A1",
      isDynamicArray: true,
      result: { error: "#REF!" }
    });
    Cell.setValue(ws, "B1", {
      formula: "1/0",
      shareType: "array",
      ref: "B1:B2",
      result: { error: "#DIV/0!" }
    });
    const read = Workbook.getWorksheet(
      await Workbook.read(await Workbook.toBuffer(wb, { format: "xlsb" })),
      "S"
    )!;
    expect(Cell.getValue(read, "A1")).toMatchObject({
      shareType: "array",
      ref: "A1",
      isDynamicArray: true,
      result: { error: "#REF!" }
    });
    expect(Cell.getValue(read, "B1")).toMatchObject({
      shareType: "array",
      ref: "B1:B2",
      result: { error: "#DIV/0!" }
    });
  });
});
