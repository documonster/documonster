import { sanitizeTableName, tableNameProblem } from "@excel/core/table-name";
import { describe, expect, it } from "vitest";

describe("tableNameProblem and sanitizeTableName agree", () => {
  const NAMES = [
    "Sales",
    "_T1",
    "T1",
    "t1",
    "XFD1048576",
    "XFE1",
    "AAAA1",
    "R1C1",
    "RC",
    "R",
    "c",
    "r2c",
    "RC3",
    "My Table",
    "a\tb",
    "Sales-2024",
    "Maß",
    "销售表",
    "1st",
    "_",
    "\\",
    "\\Name",
    "a\\b",
    "a.b",
    "a?b",
    "𝔘nicode",
    "x".repeat(255),
    "x".repeat(256),
    "é",
    "e\u0301"
  ];

  it.each(NAMES)(
    "%j: refused exactly when it would be changed, and the suggestion is accepted",
    name => {
      const problem = tableNameProblem(name);
      const fixed = sanitizeTableName(name);
      expect(problem === undefined).toBe(fixed === name);
      expect(tableNameProblem(fixed)).toBeUndefined();
      if (problem !== undefined) {
        expect(problem).toContain(JSON.stringify(fixed));
      }
    }
  );

  it("reads letters and a number, or R1C1 with both numbers, as a cell reference", () => {
    expect(tableNameProblem("XFD1048576")).toMatch(/cell reference/);
    expect(tableNameProblem("R2C3")).toMatch(/cell reference/);
    expect(tableNameProblem("RC")).toBeUndefined();
  });
});
