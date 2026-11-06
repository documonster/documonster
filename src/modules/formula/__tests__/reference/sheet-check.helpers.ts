/**
 * Calculate formulas on a fresh workbook and compare each result with Excel's. Each formula
 * sits in its own block of sheet `F` (so no spill collides) in column A; `seed` fills the data
 * sheets, and may create `F` itself to give a formula data beside it.
 */
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { expect } from "vitest";

export type Expected = number | string | boolean | { error: string } | Expected[][];

type Sheet = ReturnType<typeof Workbook.addWorksheet>;

/** The data most tests use: S!A1:C1 = 5, "x", 7; S!A2 = 3; S!A3 = 9. */
export function seedDefault(wb: ReturnType<typeof Workbook.create>): void {
  const data = Workbook.addWorksheet(wb, "S");
  Cell.setValue(data, "A1", 5);
  Cell.setValue(data, "B1", "x");
  Cell.setValue(data, "C1", 7);
  Cell.setValue(data, "A2", 3);
  Cell.setValue(data, "A3", 9);
}

/**
 * A formula and Excel's result. `"dynamic"` enters it as a dynamic-array formula (it spills, as
 * one typed in Excel 365 does); otherwise it is an ordinary formula.
 */
export type Case = readonly [formula: string, expected: Expected, kind?: "dynamic"];

export function check(
  cases: ReadonlyArray<Case>,
  seed: (wb: ReturnType<typeof Workbook.create>) => void = seedDefault
): void {
  const wb = Workbook.create();
  seed(wb);
  const ws: Sheet = Workbook.getWorksheet(wb, "F") ?? Workbook.addWorksheet(wb, "F");
  const row = (i: number) => i * 10 + 1;
  cases.forEach(([formula, , kind], i) =>
    Cell.setValue(
      ws,
      `A${row(i)}`,
      kind === "dynamic" ? { formula, result: 0, isDynamicArray: true } : { formula, result: 0 }
    )
  );
  calculateFormulas(wb);
  const cols = "ABCDEFGHIJ";
  cases.forEach(([formula, expected], i) => {
    if (Array.isArray(expected)) {
      const got = expected.map((line, r) =>
        line.map((_, c) => {
          const address = `${cols[c]}${row(i) + r}`;
          return r === 0 && c === 0 ? Cell.getResult(ws, address) : Cell.getValue(ws, address);
        })
      );
      expect(got, formula).toEqual(expected);
      return;
    }
    expect(Cell.getResult(ws, `A${row(i)}`), formula).toEqual(expected);
  });
}
