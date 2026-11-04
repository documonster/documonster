/**
 * `Pdf.fromExcel({ sheets })`: names are the one way; a numeric position still works, through
 * a `@deprecated` overload, so an editor strikes it through.
 *
 * Deprecation is a suggestion, which `tsc` never reports, and the TypeScript in this
 * repository (7, native) has no JS language-service API — so it is read from the type-aware
 * lint rule `typescript/no-deprecated` (tsgolint) over a probe file: a call naming sheets is
 * not reported, a call with a number is. The probe sits beside this test so the project's
 * `tsconfig.json` (paths, strictness) covers it, and is removed afterwards.
 */
import { execFileSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import path from "node:path";

import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";

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

describe("Pdf.fromExcel sheets: names, numbers deprecated", () => {
  it("reports a numeric sheet as deprecated, and nothing else", () => {
    const lines = [
      'import type { Workbook } from "@excel/index";',
      'import { Pdf } from "@pdf/index";',
      'import type { ExcelToPdfOptions } from "@pdf/index";',
      "declare const wb: Workbook.Handle;",
      'const names: readonly string[] = ["A"];',
      'const published: ExcelToPdfOptions = { sheets: ["Sales"] };',
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
