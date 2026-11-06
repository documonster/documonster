/**
 * `WordExcel.toDocx({ sheets })`: names are the one way; a numeric position still works,
 * through a `@deprecated` overload — the same three overloads as `Pdf.fromExcel` (see
 * `pdf/__tests__/from-excel-sheet-names.node.test.ts`, which this mirrors):
 *
 * 1. names only — not deprecated;
 * 2. `sheets` required and able to hold numbers — `@deprecated`;
 * 3. the published `DocxFromExcelOptions` — not deprecated, so a variable typed with it is not
 *    struck through whatever it holds.
 *
 * Deprecation is read from the type-aware lint rule `typescript/no-deprecated` (TypeScript 7
 * has no JS language-service API) over a probe (`@test/deprecation-probe`).
 */
import { Row, Workbook } from "@excel/index";
import { deprecatedLines } from "@test/deprecation-probe";
import { WordExcel } from "@word/excel";
import { describe, expect, it } from "vitest";

describe("WordExcel.toDocx sheets: names, numbers deprecated", () => {
  it("reports a numeric sheet as deprecated, and nothing else", () => {
    const lines = [
      'import type { Workbook } from "@excel/index";',
      'import { WordExcel } from "@word/excel";',
      'import type { DocxFromExcelOptions } from "@word/excel";',
      "declare const wb: Workbook.Handle;",
      'const names: readonly string[] = ["A"];',
      'const published: DocxFromExcelOptions = { sheets: ["Sales"] };',
      "const inferred = { sheets: [1, 2] };",
      // 8: names, a names variable, `{}`, no options — not deprecated.
      'void WordExcel.toDocx(wb, { sheets: ["Sales"] });',
      "void WordExcel.toDocx(wb, { sheets: names, maxRows: 10 });",
      "void WordExcel.toDocx(wb, {});",
      "void WordExcel.toDocx(wb);",
      // 12: a value typed with the published options interface — not deprecated either.
      "void WordExcel.toDocx(wb, published);",
      // 13: numbers — deprecated, in a literal or in an inferred variable.
      "void WordExcel.toDocx(wb, { sheets: [0] });",
      'void WordExcel.toDocx(wb, { sheets: ["Sales", 0] });',
      "void WordExcel.toDocx(wb, inferred);"
    ];
    expect(deprecatedLines(lines)).toEqual([13, 14, 15]);
  }, 60_000);

  it("a number still selects the same sheet at runtime", () => {
    const wb = Workbook.create();
    Row.add(Workbook.addWorksheet(wb, "First"), ["first"]);
    Row.add(Workbook.addWorksheet(wb, "Second"), ["second"]);
    const doc = WordExcel.toDocx(wb, { sheets: [1], includeSheetHeadings: false });
    expect(JSON.stringify(doc.body)).toContain("second");
    expect(JSON.stringify(doc.body)).not.toContain("first");
  });
});
