/**
 * `WordExcel.toDocx({ sheets })`: names are the one way; a numeric position still works,
 * through a `@deprecated` overload — the same three overloads as `Pdf.fromExcel` (see
 * `pdf/__tests__/from-excel-sheet-names.node.test.ts`, which this mirrors):
 *
 * 1. names only — not deprecated;
 * 2. `sheets` required and able to hold numbers — `@deprecated`;
 * 3. the published `ExcelToDocxOptions` — not deprecated, so a variable typed with it is not
 *    struck through whatever it holds.
 *
 * Deprecation is read from the type-aware lint rule `typescript/no-deprecated` (TypeScript 7
 * has no JS language-service API) over a probe written beside this test and removed after.
 */
import { execFileSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import path from "node:path";

import { Row, Workbook } from "@excel/index";
import { WordExcel } from "@word/excel";
import { describe, expect, it } from "vitest";

const ROOT = path.resolve(__dirname, "../../../..");

function deprecatedLines(lines: readonly string[]): number[] {
  const probe = path.join(__dirname, `__deprecation-probe-${process.pid}__.ts`);
  writeFileSync(probe, lines.join("\n") + "\n");
  try {
    let output: string;
    try {
      output = execFileSync(
        path.join(ROOT, "node_modules/.bin/oxlint"),
        ["--type-aware", "-A", "all", "-D", "typescript/no-deprecated", probe],
        { cwd: ROOT, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }
      );
    } catch (err) {
      // oxlint exits 1 when it reports something; the report is on stdout.
      output = (err as { stdout?: string }).stdout ?? "";
    }
    return [...output.matchAll(/__\.ts:(\d+):\d+: error typescript\(no-deprecated\)/g)].map(m =>
      Number(m[1])
    );
  } finally {
    rmSync(probe, { force: true });
  }
}

describe("WordExcel.toDocx sheets: names, numbers deprecated", () => {
  it("reports a numeric sheet as deprecated, and nothing else", () => {
    const lines = [
      'import type { Workbook } from "@excel/index";',
      'import { WordExcel } from "@word/excel";',
      'import type { ExcelToDocxOptions } from "@word/excel";',
      "declare const wb: Workbook.Handle;",
      'const names: readonly string[] = ["A"];',
      'const published: ExcelToDocxOptions = { sheets: ["Sales"] };',
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
