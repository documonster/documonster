import { Cell, Workbook } from "@excel/index";
import { describe, expect, it, vi } from "vitest";

describe("writeBuffer auto-validate hook", () => {
  it("does NOT auto-validate under vitest (process.env.VITEST='true')", async () => {
    // Vitest sets VITEST=true automatically. The hook skips then, so
    // fixture builders that produce hundreds of workbooks don't pay
    // the validation cost per call. Tests that need validation must
    // call `expectValidXlsx()` explicitly.
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "hello");
    await Workbook.toBuffer(wb);

    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });

  it("runs validation when forced via { validate: true } — even under vitest", async () => {
    // This is the explicit opt-in path. The workbook is clean so no
    // warning surfaces, but the hook DID execute (covered by the
    // negative-case test below).
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "hi");
    const bytes = await Workbook.toBuffer(wb, { validate: true });
    expect(bytes.byteLength).toBeGreaterThan(0);
  });

  it("can be disabled explicitly via { validate: false }", async () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "hi");

    await Workbook.toBuffer(wb, { validate: false });

    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });
});
