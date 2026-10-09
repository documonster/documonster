/**
 * Number-format rendering checked against LibreOffice, an independent implementation of the same
 * language.
 *
 * Opt-in like the other LibreOffice tests (`DOCUMONSTER_LIBREOFFICE_SMOKE=1`); skipped, visibly, without it,
 * and failing rather than passing when it is set but no LibreOffice can be found. Two paths are compared:
 *
 * - **Display.** Each value is written to a cell with its format; LibreOffice exports what it displays as
 *   CSV, and that has to equal what this library displays.
 * - **`TEXT()`.** Each formula is written with no cached result, so LibreOffice has to compute it; the
 *   exported result has to equal what this library's formula engine computes.
 *
 * LibreOffice pads `?` with U+2007 FIGURE SPACE where Excel uses an ordinary space, so that one character is
 * normalised.
 *
 * Only cases where LibreOffice and Excel agree belong here. Where they differ this library follows Excel —
 * measured, in `src/utils/__tests__/number-format-excel.test.ts` — and the cases are pinned there and in
 * `number-format-render.test.ts` instead. Among them:
 *
 * - Fractions are the last continued-fraction convergent (`3  1/7 `), not the closest fraction (`3 14/99`).
 * - Conditional sections drop or keep the sign by Excel's rules (`[<=-5]0;0` shows -7 as `7`).
 * - `General` fits eleven characters (`1.23457E+11`); `"x"` shows -5 as `-x`; `hh dd mm` reads minutes.
 * - A negative time shows hashes; a date rounds to the second even when no time is shown.
 *
 * - `#.##` keeps its decimal point (`1.`); LibreOffice drops it.
 * - Each `%` multiplies by 100 (`0%%` shows `100%%` for 0.01); LibreOffice applies only one.
 * - A clock reading rounds to the shown precision (23:59:59.99999 is `00:00:00` the next day); LibreOffice
 *   truncates.
 * - `y` and `yyy`, Buddhist `b`, and the case of `am/pm` follow Excel's documentation.
 * - Serial 60 is Excel's 1900-02-29, and a negative date shows hashes.
 * - A comma ending the integer placeholders scales by 1,000 even before a decimal point (`0,.0` shows 1234
 *   as `1.2`), as Excel's documentation states for any comma that follows a digit placeholder; LibreOffice
 *   shows `1234`.
 * - The character after `_` is always its argument: `0_";0` pads with the width of `"` (`12 `, as SheetJS
 *   also reads it), where LibreOffice opens a string there.
 */
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";

import { libreOfficeAvailable } from "@excel/__tests__/helpers/libreoffice-smoke";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { renderNumberFormat } from "@utils/number-format-render";
import { describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);

/** 2024-01-15 under the 1900 epoch. */
const JAN_15_2024 = 45306;

const CASES: readonly [string, number][] = [
  ["$#,##0", -1234],
  ["####.##", 12.5],
  ["[$€-407]#,##0.00", 1234.56],
  ["0.00", 1.005],
  ["0.00", 2.675],
  ["0", 0.5],
  ["0", -0.5],
  ["#", 0],
  ["#,###", 0.4],
  ["0.00", -0.001],
  ["0.00;(0.00)", -0.001],
  ["#,##0;-#,##0", -0.4],
  ["(###) ###-####", 5551234],
  ["0000-000", 1234567],
  ["??0.0?", 1.5],
  ["#,##0.0", 1234567.89],
  ["0,0", 1234],
  ["#,", 1500],
  ["#,##0.00,,", 1234567890],
  ["0.0,,\\M", 999999999],
  ["\\$0.0,,\\M", -5e9],
  ['$0.0,,"M"', 5e9],
  ["0.0\\s", 12],
  ["*-0", 5],
  ['0%" off"', 0.25],
  ["0.0%", -0.0004],
  ["0.00E+00", 0],
  ["0.00E+00", 9.999],
  ["0.00E-00", 1234],
  ["0.0e+0", 1234],
  ["##0.0E+0", 12345],
  ["##0.0E+0", 0.00012345],
  ["# ?/?", 5],
  ["# ?/?", 0],
  ["# ?/?", 0.5],
  ["# ?/?", 0.97],
  ["# ?/?", -3.25],
  ["# ??/??", 0.333],
  ["?/?", 1.5],
  ["# ?/8", 1.5],
  ['[<0]"neg";[>=1000]0.0,"K";0', -5],
  ['[>=1000]0.0,"K";0', 12345],
  ["[<0]0;0", -5],
  ["[<0]0;0", 5],
  ["[>100]0.00;0", 50],
  ['[>100]0;[<0]"m"0;0', -5],
  ["[>=0]0;[<0]0", -3],
  ['[=0]"z";0', 0],
  ["0;-0;;@", 0],
  ["0.00;@", -3],
  ['"n="@', 5],
  ["[Red]General", -1.5],
  ['General" kg"', 3],
  ["General", 1e-12],
  ['_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 1234.5],
  ['_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 0],
  ["yyyy-mm-dd hh:mm:ss", JAN_15_2024 + 0.5],
  ['h "at" mm', 0.75],
  ["ddd, mmm d", JAN_15_2024],
  ["mm/dd", JAN_15_2024],
  ["m", JAN_15_2024],
  ["h:mm AM/PM", 0.6],
  ["[h]:mm:ss", 1.5],
  ["[mm]:ss", 0.05],
  ["[m]", 0.5],
  ["0_;0", 12],
  ["?/00", 0.5],
  ["?/?0", 0.5],
  ["?/#0", 0.5],
  ["00/??", 0.5],
  ["# ??/##", 0.5],
  ["# ?/16", 0.3],
  ["[ss].00", 3735.8 / 86400],
  ["h:mm:ss.00", (16 * 3600 + 36 * 60 + 3.75) / 86400],
  ["0.00E+00", 1e-300],
  ["yyyy-mm-dd hh:mm:ss.000", 200000.5]
];

/** `TEXT()` calls whose results LibreOffice and this library's formula engine must agree on. */
const TEXT_FORMULAS: readonly string[] = [
  'TEXT(1234.5,"#,##0.00")',
  'TEXT(-5,"$#,##0")',
  'TEXT(5000000000,"$0.0,,""M""")',
  'TEXT(3,"0"" hours""")',
  'TEXT(0.25,"0%"" off""")',
  'TEXT(1.5,"General")',
  'TEXT(0.5,"# ?/?")',
  'TEXT(-5,"[<0]0;0")',
  'TEXT(1.5,"[h]:mm")',
  'TEXT(45306,"dddd, mmmm d, yyyy")',
  'TEXT("12","0.00")',
  'TEXT("12","0.00;@")',
  'TEXT("12","0;0;0;@")',
  'TEXT("hi","0")',
  'TEXT("hi","0;@")',
  'TEXT("","0")',
  'TEXT(5,"@")',
  'TEXT(5,"""n=""@")'
];

/** One CSV field, unquoted. Every row here is a single field. */
function unquote(field: string): string {
  return field.startsWith('"') ? field.slice(1, -1).replace(/""/g, '"') : field;
}

const enabled = process.env.DOCUMONSTER_LIBREOFFICE_SMOKE === "1";

describe("number formats agree with LibreOffice", () => {
  it.skipIf(!enabled)(
    "renders every case as LibreOffice displays it",
    async () => {
      const binary = await libreOfficeAvailable();
      if (!binary) {
        throw new Error(
          "DOCUMONSTER_LIBREOFFICE_SMOKE=1 is set but no LibreOffice executable was found."
        );
      }
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "Formats");
      CASES.forEach(([fmt, value], i) => {
        Cell.setValue(sheet, `A${i + 1}`, value);
        Cell.setNumFmt(sheet, `A${i + 1}`, fmt);
      });
      // TEXT() results are computed here first, then the formulas are rewritten without them so that
      // LibreOffice cannot read a cached answer and has to compute its own.
      const textRow = (i: number): string => `A${CASES.length + i + 1}`;
      TEXT_FORMULAS.forEach((formula, i) => Cell.setValue(sheet, textRow(i), { formula }));
      calculateFormulas(workbook);
      const ourText = TEXT_FORMULAS.map((_, i) => String(Cell.getResult(sheet, textRow(i))));
      TEXT_FORMULAS.forEach((formula, i) => Cell.setValue(sheet, textRow(i), { formula }));
      const dir = await mkdtemp(join(tmpdir(), "documonster-numfmt-"));
      try {
        await writeFile(
          join(dir, "formats.xlsx"),
          new Uint8Array(await Workbook.toBuffer(workbook))
        );
        // Token 9 ("save cell contents as shown") is what makes the export the displayed text.
        await execFileAsync(
          binary,
          [
            "--headless",
            "--convert-to",
            "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true",
            "--outdir",
            dir,
            join(dir, "formats.xlsx")
          ],
          {
            timeout: 120_000,
            // A profile of its own: LibreOffice serialises processes sharing one, and the second exits zero
            // having converted nothing.
            env: { ...process.env, UserInstallation: `file://${join(dir, "profile")}` }
          }
        );
        const lines = (await readFile(join(dir, "formats.csv"), "utf8")).split(/\r?\n/);
        const shown = lines.map(line => unquote(line).replace(/\u2007/g, " "));
        const ours = CASES.map(([fmt, value]) => renderNumberFormat(fmt, value));
        expect(ours).toEqual(shown.slice(0, CASES.length));
        expect(ourText).toEqual(shown.slice(CASES.length, CASES.length + TEXT_FORMULAS.length));
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    },
    180_000
  );
});
