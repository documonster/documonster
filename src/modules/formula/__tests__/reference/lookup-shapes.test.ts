import { Cell, Workbook } from "@excel/index";
/**
 * Lookup functions refuse arrays of the wrong shape as Excel does, rather than reading cells
 * that are not there as blank. XLOOKUP: a vertical lookup_array (one column) needs a
 * return_array with as many rows (any number of columns, which come back as a row); a
 * horizontal one (one row) needs as many columns (rows come back as a column); otherwise, or
 * for a 2-D lookup_array, #VALUE!. MATCH needs a one-row or one-column lookup_array (#N/A).
 */
import { check, type Expected } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

// S!A1:A3 = 1, 2, 3; S!B1:C3 = 10..15; S!E1:G1 = 1, 2, 3; S!E2:G3 = 20..25
function seed(wb: ReturnType<typeof Workbook.create>): void {
  const s = Workbook.addWorksheet(wb, "S");
  [1, 2, 3].forEach((v, i) => Cell.setValue(s, `A${i + 1}`, v));
  [10, 11, 12, 13, 14, 15].forEach((v, i) =>
    Cell.setValue(s, `${i % 2 ? "C" : "B"}${(i >> 1) + 1}`, v)
  );
  ["E", "F", "G"].forEach((col, i) => {
    Cell.setValue(s, `${col}1`, i + 1);
    Cell.setValue(s, `${col}2`, 20 + i);
    Cell.setValue(s, `${col}3`, 23 + i);
  });
}

const run = (cases: ReadonlyArray<readonly [string, Expected]>) => check(cases, seed);
const VALUE = { error: "#VALUE!" };

describe("XLOOKUP return_array size", () => {
  it("vertical lookup: as many rows; columns come back as a row", () => {
    run([
      ["XLOOKUP(2,S!A1:A3,S!B1:B3)", 12],
      ["XLOOKUP(2,S!A1:A3,S!B1:C3)", [[12, 13]]],
      ["XLOOKUP(2,S!A1:A3,S!B1:B2)", VALUE],
      ["XLOOKUP(1,S!A1:A3,S!B1)", VALUE],
      ["XLOOKUP(2,S!A1:A3,S!E1:G1)", VALUE]
    ]);
  });

  it("horizontal lookup: as many columns; rows come back as a column", () => {
    run([
      ["XLOOKUP(2,S!E1:G1,S!E2:G2)", 21],
      ["XLOOKUP(2,S!E1:G1,S!E2:G3)", [[21], [24]]],
      ["XLOOKUP(2,S!E1:G1,S!E2:F2)", VALUE],
      ["XLOOKUP(2,S!E1:G1,S!A1:A3)", VALUE]
    ]);
  });

  it("a 2-D lookup_array is #VALUE!; not found still returns if_not_found", () => {
    run([
      ["XLOOKUP(12,S!B1:C3,S!B1:C3)", VALUE],
      ['XLOOKUP(9,S!A1:A3,S!B1:B3,"none")', "none"]
    ]);
  });
});

describe("MATCH / XMATCH lookup_array shape", () => {
  it("a 2-D lookup_array: MATCH #N/A, XMATCH #VALUE!", () => {
    run([
      ["MATCH(13,S!B1:C3,0)", { error: "#N/A" }],
      ["XMATCH(13,S!B1:C3)", VALUE],
      ["MATCH(3,S!A1:A3,0)", 3],
      ["XMATCH(3,S!E1:G1)", 3]
    ]);
  });
});

describe("VLOOKUP / HLOOKUP index past the table", () => {
  it("is #REF!", () => {
    run([
      ["VLOOKUP(2,S!A1:C3,4,FALSE)", { error: "#REF!" }],
      ["HLOOKUP(2,S!E1:G3,4,FALSE)", { error: "#REF!" }],
      ["VLOOKUP(2,S!A1:C3,3,FALSE)", 13]
    ]);
  });
});
