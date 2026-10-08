import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("XLSX literal number formats", () => {
  it.each([
    ["\\$0.0,,\\M", 5000000000, "$5000.0M"],
    ["\\$0.0,,\\M", 1000000, "$1.0M"],
    ["\\$0.0,,\\M", 0, "$0.0M"],
    ["\\$0.0,,\\M", -1000000, "$-1.0M"],
    ['$0.0,,"M"', 5000000000, "$5000.0M"],
    ["[Green]\\$0.0,,\\M;[Red](\\$0.0,,\\M)", -1000000, "($1.0M)"],
    ["0.0\\m", 12, "12.0m"],
    ["0.0\\s", 12, "12.0s"],
    ['0" days"', 12, "12 days"],
    ["0_m", 12, "12 "],
    ["0*m", 12, "12"]
  ])(
    "preserves numeric values and formula results with %s (%s)",
    async (numFmt, value, display) => {
      let workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "Repro");
      Cell.setValue(sheet, "A1", value);
      Cell.setValue(sheet, "A2", { formula: String(value), result: value });
      Cell.setNumFmt(sheet, "A1", numFmt);
      Cell.setNumFmt(sheet, "A2", numFmt);

      // A second save/load proves the imported model remains serializable without losing the format.
      for (let pass = 0; pass < 2; pass++) {
        const loaded = Workbook.create();
        await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
        const loadedSheet = Workbook.getWorksheet(loaded, "Repro")!;
        expect(Cell.getValue(loadedSheet, "A1")).toBe(value);
        expect(Cell.getValue(loadedSheet, "A2")).toEqual({ formula: String(value), result: value });
        for (const address of ["A1", "A2"]) {
          expect(Cell.getNumFmt(loadedSheet, address)).toBe(numFmt);
          expect(Cell.getDisplayText(loadedSheet, address)).toBe(display);
        }
        workbook = loaded;
      }
    }
  );

  it.each([
    ["yyyy-mm-dd", 45952, "2025-10-22"],
    ["yyyy\\-mm\\-dd", 45952, "2025-10-22"],
    ["hh:mm:ss", 0.5, "12:00:00"],
    ["[h]:mm:ss", 1.5, "36:00:00"]
  ])("still imports genuine date/time formats: %s", async (numFmt, value, display) => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "Dates");
    Cell.setValue(sheet, "A1", value);
    Cell.setValue(sheet, "A2", { formula: String(value), result: value });
    Cell.setNumFmt(sheet, "A1", numFmt);
    Cell.setNumFmt(sheet, "A2", numFmt);
    const loaded = Workbook.create();
    await Workbook.read(loaded, new Uint8Array(await Workbook.toBuffer(workbook)));
    const loadedSheet = Workbook.getWorksheet(loaded, "Dates")!;
    expect(Cell.getValue(loadedSheet, "A1")).toBeInstanceOf(Date);
    expect(Cell.getResult(loadedSheet, "A2")).toBeInstanceOf(Date);
    for (const address of ["A1", "A2"]) {
      expect(Cell.getDisplayText(loadedSheet, address)).toBe(display);
    }
  });
});
