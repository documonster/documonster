import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * Browser-safe end-to-end check for named cell styles.
 *
 * Runs under BOTH the Node and the browser (chromium) test runners — it uses
 * only in-memory `toBuffer` / `read`, no Node `fs`. This is the runtime the
 * feature was actually broken in (#185 follow-up: `Workbook.defineCellStyle`
 * was `undefined` in browser/bundler builds), so exercising the real call here
 * guards the browser code path directly.
 */
describe("named cell styles (cross-platform e2e)", () => {
  it("defines, applies, and round-trips a custom named style in-memory", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    Workbook.defineCellStyle(wb, "Heading 1", {
      font: { name: "Arial", size: 20, bold: true }
    });
    Cell.setValue(ws, "A1", "Title");
    Cell.applyCellStyle(ws, "A1", "Heading 1");

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

    expect(Cell.getStyle(ws2, "A1").styleName).toBe("Heading 1");
    expect(Workbook.getCellStyle(wb2, "Heading 1")?.font).toMatchObject({
      name: "Arial",
      size: 20,
      bold: true
    });
  });

  it("applies a built-in preset with the correct builtinId in-memory", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    const name = Workbook.useBuiltinCellStyle(wb, "Heading1");
    expect(name).toBe("Heading 1");
    Cell.setValue(ws, "A1", "Title");
    Cell.applyCellStyle(ws, "A1", name);

    const buffer = new Uint8Array(await Workbook.toBuffer(wb));
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);
    expect(Workbook.getCellStyle(wb2, "Heading 1")?.builtinId).toBe(16);
  });
});
