/**
 * The accounting formats, end to end through a template a Chinese Excel produced.
 *
 * **Split out of `numfmt-builtins.test.ts` so that file keeps its browser coverage.** This case reads a template off
 * disk, which is Node-only — and the browser configuration excludes `*.node.test.ts` by glob.
 */
import { readFileSync } from "node:fs";

import { Cell, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const TEMPLATE = "src/modules/excel/examples/data/comments.xlsx";

/** The four accounting codes the template declares, with `<numFmt numFmtId="41…44">`. */
const DECLARED = [
  '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ',
  '_ "￥"* #,##0_ ;_ "￥"* \\-#,##0_ ;_ "￥"* "-"_ ;_ @_ ',
  '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  '_ "￥"* #,##0.00_ ;_ "￥"* \\-#,##0.00_ ;_ "￥"* "-"??_ ;_ @_ '
];

describe("a format code a file declares for an accounting id", () => {
  // The codes are the template's, so they are written as codes and a reader in any locale shows them as its author
  // did. Excel opens the result keeping all 49 of the template's named styles (checked in Excel for both containers):
  // custom format ids were once blamed for Excel discarding them, which was the style XFs' `ixfParent`.
  it.each(["xlsx", "xlsb"] as const)(
    "is read and written back as that code in %s",
    async format => {
      const source = await Workbook.read(readFileSync(TEMPLATE));
      const ws = Workbook.addWorksheet(source, "Probe");
      DECLARED.forEach((code, i) => {
        Cell.setValue(ws, { row: i + 1, col: 1 }, -1234.5);
        Cell.setNumFmt(ws, { row: i + 1, col: 1 }, code);
      });
      const back = await Workbook.read(
        await Workbook.toBuffer(source, { format, unsupported: "ignore" }),
        { format }
      );
      const probe = Workbook.getWorksheet(back, "Probe")!;
      expect(DECLARED.map((_, i) => Cell.getNumFmt(probe, { row: i + 1, col: 1 }))).toEqual(
        DECLARED
      );
      expect(Cell.getDisplayText(probe, "A2")).toBe(" ￥-1,235 ");
    }
  );
});
