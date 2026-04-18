/**
 * Excel reference-value conformance tests.
 *
 * Each entry in `excel-reference.json` pairs a formula with the value
 * Excel 365 / Microsoft Office documentation says it should produce.
 * We run every formula through a small throw-away Workbook and assert
 * the computed `cell.result` matches.
 *
 * Scope: function-level correctness. The reference JSON focuses on
 * canonical inputs (documented examples, textbook values, round
 * numbers) rather than exhaustive edge cases — those live in the
 * per-function saturation tests alongside their source.
 *
 * To grow: append new `{formula, expected}` entries to the JSON.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "@excel/workbook";
import { describe, expect, it } from "vitest";

const __dirname = dirname(fileURLToPath(import.meta.url));
const refPath = join(__dirname, "excel-reference.json");

interface ReferenceNumber {
  num: number;
  eps: number;
}
interface ReferenceError {
  error: string;
}
interface ReferenceArray {
  array: (number | string | boolean | ReferenceError)[][];
}
interface ReferenceBlank {
  blank: true;
}
type Expected =
  | number
  | string
  | boolean
  | ReferenceNumber
  | ReferenceError
  | ReferenceArray
  | ReferenceBlank
  | null;

interface ReferenceCase {
  formula: string;
  expected: Expected;
}

interface ReferenceFile {
  cases: ReferenceCase[];
}

const raw = readFileSync(refPath, "utf8");
const data: ReferenceFile = JSON.parse(raw);

/**
 * Evaluate a formula through a one-shot workbook and return the cell
 * result. Uses A1 as the formula cell; any references in the formula
 * resolve against other cells we leave empty, which matches our
 * reference examples (they use either literal arrays or no refs).
 */
function evaluate(formula: string): unknown {
  const wb = new Workbook();
  const ws = wb.addWorksheet("S");
  // Strip leading '=' to match the { formula: ... } API convention.
  const body = formula.startsWith("=") ? formula.slice(1) : formula;
  ws.getCell("A1").value = { formula: body, result: 0 };
  wb.calculateFormulas();
  return ws.getCell("A1").result;
}

/** Distinguish expected-value shapes. */
function isRefNumber(v: Expected): v is ReferenceNumber {
  return typeof v === "object" && v !== null && "num" in v && "eps" in v;
}
function isRefError(v: Expected): v is ReferenceError {
  return typeof v === "object" && v !== null && "error" in v;
}
function isRefArray(v: Expected): v is ReferenceArray {
  return typeof v === "object" && v !== null && "array" in v;
}
function isRefBlank(v: Expected): v is ReferenceBlank {
  return typeof v === "object" && v !== null && "blank" in v;
}

describe("Excel reference values", () => {
  for (const tc of data.cases) {
    it(`${tc.formula}`, () => {
      const actual = evaluate(tc.formula);
      const expected = tc.expected;

      if (isRefNumber(expected)) {
        expect(typeof actual).toBe("number");
        expect(Math.abs((actual as number) - expected.num)).toBeLessThanOrEqual(expected.eps);
        return;
      }
      if (isRefError(expected)) {
        expect(actual).toEqual({ error: expected.error });
        return;
      }
      if (isRefArray(expected)) {
        // Dynamic-array results land in A1 (scalar result = top-left) and
        // spill into siblings. Don't try to verify each cell here — the
        // per-function tests cover spill detail. We just verify A1
        // matches the top-left of the expected grid.
        expect(actual).toEqual(expected.array[0][0]);
        return;
      }
      if (isRefBlank(expected)) {
        expect(actual === null || actual === undefined || actual === 0).toBe(true);
        return;
      }
      expect(actual).toEqual(expected);
    });
  }
});
