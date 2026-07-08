import * as WorkbookNode from "@excel/surface/workbook";
import * as WorkbookBrowser from "@excel/surface/workbook.browser";
import { describe, expect, it } from "vitest";

/**
 * Guards against the `Workbook` surface drifting between the Node
 * (`surface/workbook.ts`) and browser (`surface/workbook.browser.ts`) entries.
 *
 * Regression: the named-cell-style functions (`defineCellStyle`, …) were added
 * to the Node surface only, so `Workbook.defineCellStyle` was `undefined` at
 * runtime in browser/bundler builds (issue #185 follow-up).
 */
describe("Workbook surface node/browser parity", () => {
  const runtimeFns = (mod: Record<string, unknown>) =>
    Object.keys(mod)
      .filter(k => typeof mod[k] === "function")
      .sort();

  it("exposes the named-cell-style API on both entries", () => {
    const api = [
      "defineCellStyle",
      "getCellStyle",
      "listCellStyles",
      "removeCellStyle",
      "useBuiltinCellStyle"
    ] as const;
    for (const fn of api) {
      expect(typeof (WorkbookNode as Record<string, unknown>)[fn]).toBe("function");
      expect(typeof (WorkbookBrowser as Record<string, unknown>)[fn]).toBe("function");
    }
  });

  it("differs only by Node file-path IO (readFile/writeFile)", () => {
    const node = runtimeFns(WorkbookNode as Record<string, unknown>);
    const browser = runtimeFns(WorkbookBrowser as Record<string, unknown>);
    // Every browser function must also exist on Node.
    expect(browser.filter(k => !node.includes(k))).toEqual([]);
    // Node adds only the file-path IO helpers on top of the shared surface.
    expect(node.filter(k => !browser.includes(k)).sort()).toEqual(["readFile", "writeFile"]);
  });
});
