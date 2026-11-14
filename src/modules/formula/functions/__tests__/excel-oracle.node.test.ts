/**
 * Every formula in `data/excel-oracle/*.xlsx` against the result Excel for Mac 16.x calculated and
 * saved for it. The workbooks were written by this library, opened, fully recalculated and saved
 * by Excel; the cached values in them are Excel's answers.
 *
 * They settle what no documentation does: that weekend and `return_type` codes are truncated, the
 * text length `FIXED` and `DOLLAR` stop at, that the securities functions divide by `YEARFRAC`'s
 * year fraction on every basis, the month-end rule of the coupon schedule, the day counts `PRICE`,
 * `YIELD` and `DURATION` use, Excel's own `COUPDAYS` on basis 1 (which depends on where settlement
 * falls in its period) and `ACCRINT`'s quasi-coupon accrual with both calc_methods.
 */

import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const DIR = join(__dirname, "data", "excel-oracle");

interface Case {
  file: string;
  formula: string;
  excel: unknown;
  ours: unknown;
}

async function loadCases(): Promise<Case[]> {
  const cases: Case[] = [];
  for (const file of readdirSync(DIR).filter(name => name.endsWith(".xlsx"))) {
    const saved = await Workbook.read(readFileSync(join(DIR, file)));
    const sheet = Workbook.getWorksheet(saved, "oracle")!;
    const fresh = Workbook.create();
    const freshSheet = Workbook.addWorksheet(fresh, "oracle");
    const rows: { row: number; formula: string; excel: unknown }[] = [];
    for (let row = 2; Cell.getValue(sheet, `A${row}`) != null; row++) {
      const formula = Cell.getFormula(sheet, `D${row}`)!;
      rows.push({ row, formula, excel: Cell.getResult(sheet, `D${row}`) });
      Cell.setValue(freshSheet, `D${row}`, { formula, result: 0 });
    }
    calculateFormulas(fresh);
    for (const { row, formula, excel } of rows) {
      cases.push({ file, formula, excel, ours: Cell.getResult(freshSheet, `D${row}`) });
    }
  }
  return cases;
}

function agrees(excel: unknown, ours: unknown): boolean {
  if (typeof excel === "number" && typeof ours === "number") {
    // YIELD is a root found by iteration, in Excel as here: they agree to about 1e-12.
    return Math.abs(excel - ours) <= 1e-11 * Math.max(1, Math.abs(excel));
  }
  return JSON.stringify(excel) === JSON.stringify(ours);
}

describe("formulas agree with the results Excel saved", async () => {
  const cases = await loadCases();

  it("reads every oracle workbook", () => {
    expect(cases.length).toBeGreaterThan(16000);
  });

  it("agrees on every formula", () => {
    const differing = cases.filter(c => !agrees(c.excel, c.ours));
    expect(
      differing.map(
        c =>
          `${c.file}: ${c.formula} Excel=${JSON.stringify(c.excel)} ours=${JSON.stringify(c.ours)}`
      )
    ).toEqual([]);
  });
});
