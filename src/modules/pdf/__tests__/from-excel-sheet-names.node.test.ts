/**
 * `Pdf.fromExcel({ sheets })`: names are the one way; a numeric position still works, through
 * a `@deprecated` overload, so an editor strikes it through.
 *
 * Deprecation is a suggestion, which `tsc` never reports, and the TypeScript in this
 * repository (7, native) has no JS language-service API — so it is read from the type-aware
 * lint rule `typescript/no-deprecated` (tsgolint) over a probe file (`@test/deprecation-probe`):
 * a call naming sheets is not reported, a call with a number is.
 */
import { Cell, Workbook } from "@excel/index";
import { deprecatedLines } from "@test/deprecation-probe";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";

describe("Pdf.fromExcel sheets: names, numbers deprecated", () => {
  it("reports a numeric sheet as deprecated, and nothing else", () => {
    const lines = [
      'import type { Workbook } from "@excel/index";',
      'import { Pdf } from "@pdf/index";',
      'import type { PdfFromExcelOptions } from "@pdf/index";',
      "declare const wb: Workbook.Handle;",
      'const names: readonly string[] = ["A"];',
      'const published: PdfFromExcelOptions = { sheets: ["Sales"] };',
      "const inferred = { sheets: [1, 2] };",
      // 8: names, a names variable, `{}`, no options — not deprecated.
      'void Pdf.fromExcel(wb, { sheets: ["Summary"] });',
      "void Pdf.fromExcel(wb, { sheets: names, scale: 1 });",
      "void Pdf.fromExcel(wb, {});",
      "void Pdf.fromExcel(wb);",
      // 12: a value typed with the published options interface — not deprecated either.
      "void Pdf.fromExcel(wb, published);",
      // 13: numbers — deprecated, in a literal or in an inferred variable.
      "void Pdf.fromExcel(wb, { sheets: [2] });",
      'void Pdf.fromExcel(wb, { sheets: ["Summary", 2] });',
      "void Pdf.fromExcel(wb, inferred);"
    ];
    expect(deprecatedLines(lines)).toEqual([13, 14, 15]);
  }, 60_000);

  it("a number still selects the same sheet at runtime", async () => {
    const wb = Workbook.create();
    for (const name of ["First", "Second"]) {
      Cell.setValue(Workbook.addWorksheet(wb, name), "A1", `cell of ${name}`);
    }
    const bytes = await Pdf.fromExcel(wb, { sheets: [2] });
    expect((await Pdf.extract(bytes)).pages[0].text).toContain("cell of Second");
  });
});
