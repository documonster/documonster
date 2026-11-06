/**
 * A constant or a computed value where a function takes only a reference — `RANK(2,{2,2,1},0)`. Excel refuses such a
 * formula when it is typed and refuses to open a workbook holding one, so both writers report it through
 * `unsupported`: refused by default, written with its cached value under `"ignore"`.
 */
import { extractAll } from "@archive/unzip/extract";
import { Cell, ConditionalFormatting, DataValidation, DefinedNames, Workbook } from "@excel";
import { ExcelUnsupportedFeatureError } from "@excel/errors";
import { unwritableFormulaReason } from "@excel/utils/array-evaluation";
import {
  FUNCTION_ID_BY_NAME,
  functionClasses,
  futureFunctionClasses,
  paramClassAt
} from "@formula/syntax/function-table";
import { describe, expect, it } from "vitest";

const REFUSED = [
  ["RANK(2,{2,2,1},0)", "RANK argument 2 must be a reference"],
  ["_xlfn.RANK.EQ(2,{2,2,1})", "RANK.EQ argument 2 must be a reference"],
  ["RANK.AVG(2,1+1)", "RANK.AVG argument 2 must be a reference"],
  ["COUNTIF(5,1)", "COUNTIF argument 1 must be a reference"],
  ["COUNTIFS(A1:A3,1,{1,2,3},2)", "COUNTIFS argument 3 must be a reference"],
  ['SUMIF(A1:A3,">0","x")', "SUMIF argument 3 must be a reference"],
  ["OFFSET(1,1,1)", "OFFSET argument 1 must be a reference"],
  ["SUBTOTAL(9,A1:A3,{1,2})", "SUBTOTAL argument 3 must be a reference"],
  ["1+ABS(AREAS(-A1))", "AREAS argument 1 must be a reference"]
] as const;

const WRITTEN = [
  "IF(1,2,3)",
  "VLOOKUP(1,{1,2;3,4},2,0)",
  "IFERROR(1/0,{1,2})",
  "RANK(2,A1:A3,0)",
  "COUNTIF(OFFSET(A1,0,0,3),1)",
  "SUMIF(A1:A3,1)",
  "ROW()",
  "AREAS((A1:A2,B3))",
  'CELL("type")',
  "RANK(1,Data)",
  'COUNTIF(INDIRECT("A1:A3"),1)'
];

function workbookWith(formula: string) {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValue(sheet, "A1", 1);
  Cell.setValue(sheet, "B1", { formula, result: 7 } as never);
  Cell.setValue(sheet, "B2", { formula: "1+1", result: 2 } as never);
  return workbook;
}

async function reopen(bytes: Uint8Array) {
  const workbook = Workbook.create();
  await Workbook.read(workbook, bytes);
  return Workbook.getWorksheet(workbook, "S")!;
}

describe("the function table's reference-only parameters", () => {
  const at = (name: string, index: number) => {
    const id = FUNCTION_ID_BY_NAME.get(name);
    return paramClassAt(id === undefined ? futureFunctionClasses(name) : functionClasses(id), index)
      .referenceOnly;
  };

  it("marks the parameters Excel documents as references", () => {
    expect(at("RANK", 1)).toBe(true);
    expect(at("RANK.EQ", 1)).toBe(true);
    expect(at("RANK.AVG", 1)).toBe(true);
    expect(at("OFFSET", 0)).toBe(true);
    expect(at("AREAS", 0)).toBe(true);
    expect(at("ROW", 0)).toBe(true);
    expect(at("COLUMN", 0)).toBe(true);
    expect(at("CELL", 1)).toBe(true);
    expect(at("SUBTOTAL", 1)).toBe(true);
    expect(at("SUBTOTAL", 5)).toBe(true);
    expect(at("COUNTBLANK", 0)).toBe(true);
    expect(at("FORMULATEXT", 0)).toBe(true);
    expect(at("COUNTIF", 0)).toBe(true);
    expect(at("SUMIF", 0)).toBe(true);
    expect(at("SUMIF", 2)).toBe(true);
    expect(at("AVERAGEIF", 2)).toBe(true);
    // Pairs repeat: criteria ranges are references, criteria are values.
    expect(at("COUNTIFS", 2)).toBe(true);
    expect(at("COUNTIFS", 3)).toBeUndefined();
    expect(at("SUMIFS", 0)).toBe(true);
    expect(at("SUMIFS", 3)).toBe(true);
    expect(at("SUMIFS", 4)).toBeUndefined();
  });

  it("leaves parameters that take any value unmarked, although their class is the same", () => {
    expect(at("RANK", 0)).toBeUndefined();
    expect(at("IF", 1)).toBeUndefined();
    expect(at("CHOOSE", 1)).toBeUndefined();
    expect(at("VLOOKUP", 1)).toBeUndefined();
    expect(at("IFERROR", 1)).toBeUndefined();
    expect(at("ROWS", 0)).toBeUndefined();
    expect(at("INDIRECT", 0)).toBeUndefined();
    expect(at("AGGREGATE", 2)).toBeUndefined();
  });
});

describe("a non-reference in a reference-only parameter", () => {
  it.each(REFUSED)("is named for %s", (formula, reason) => {
    expect(unwritableFormulaReason(formula)).toBe(`formula (${reason})`);
  });

  it.each(WRITTEN)("is not found in %s", formula => {
    expect(unwritableFormulaReason(formula)).toBeUndefined();
  });

  it.each(["xlsx", "xlsb"] as const)("refuses the %s write by default", async format => {
    const error = await Workbook.toBuffer(workbookWith("RANK(2,{2,2,1},0)"), { format }).catch(
      (cause: unknown) => cause
    );
    expect(error).toBeInstanceOf(ExcelUnsupportedFeatureError);
    expect((error as ExcelUnsupportedFeatureError).items).toEqual([
      "S!B1: formula (RANK argument 2 must be a reference)"
    ]);
  });

  it.each([
    ["xlsx", "RANK(2,{2,2,1},0)"],
    ["xlsx", "COUNTIF(5,1)"],
    ["xlsx", "OFFSET(1,1,1)"],
    ["xlsb", "RANK(2,{2,2,1},0)"],
    ["xlsb", "COUNTIF(5,1)"],
    ["xlsb", "OFFSET(1,1,1)"]
  ] as const)("writes the cached value under ignore (%s, %s)", async (format, formula) => {
    const bytes = await Workbook.toBuffer(workbookWith(formula), { format, unsupported: "ignore" });
    const sheet = await reopen(bytes);
    expect(Cell.getValue(sheet, "B1")).toBe(7);
    expect(Cell.getFormula(sheet, "B2")).toBe("1+1");
  });

  it.each(["xlsx", "xlsb"] as const)("still writes valid formulas (%s)", async format => {
    for (const formula of ["IF(1,2,3)", "VLOOKUP(1,{1,2;3,4},2,0)", "IFERROR(1/0,{1,2})"]) {
      const bytes = await Workbook.toBuffer(workbookWith(formula), { format });
      const sheet = await reopen(bytes);
      expect(Cell.getFormula(sheet, "B1")).toBe(formula);
    }
  });

  it.each(["xlsx", "xlsb"] as const)(
    "takes a shared formula's dependents with its master (%s)",
    async format => {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      Cell.setValue(sheet, "A1", {
        formula: "RANK(1,{1,2})",
        result: 1,
        shareType: "shared",
        ref: "A1:A2"
      } as never);
      Cell.setValue(sheet, "A2", { sharedFormula: "A1", result: 1 } as never);
      const error = await Workbook.toBuffer(workbook, { format }).catch((cause: unknown) => cause);
      expect((error as ExcelUnsupportedFeatureError).items).toContain(
        "S!A1: formula (RANK argument 2 must be a reference)"
      );
      const bytes = await Workbook.toBuffer(workbook, { format, unsupported: "ignore" });
      const back = await reopen(bytes);
      expect(Cell.getFormula(back, "A1")).toBeUndefined();
      expect(Cell.getFormula(back, "A2")).toBeUndefined();
      expect(Cell.getValue(back, "A2")).toBe(1);
    }
  );

  it.each(["xlsx", "xlsb"] as const)(
    "drops an array formula to its cached value (%s)",
    async format => {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      Cell.setValue(sheet, "A1", {
        formula: "ROW(1)",
        result: 3,
        shareType: "array",
        ref: "A1"
      } as never);
      const error = await Workbook.toBuffer(workbook, { format }).catch((cause: unknown) => cause);
      expect((error as ExcelUnsupportedFeatureError).items).toEqual([
        "S!A1: formula (ROW argument 1 must be a reference)"
      ]);
      const back = await reopen(
        await Workbook.toBuffer(workbook, { format, unsupported: "ignore" })
      );
      expect(Cell.getFormula(back, "A1")).toBeUndefined();
      expect(Cell.getValue(back, "A1")).toBe(3);
    }
  );

  it.each(["xlsx", "xlsb"] as const)(
    "drops a conditional format, a validation and a defined name that carry one (%s)",
    async format => {
      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "S");
      Cell.setValue(sheet, "A1", 1);
      ConditionalFormatting.add(sheet, {
        ref: "A1:A3",
        rules: [
          { type: "expression", formulae: ["RANK(A1,{1,2,3})=1"], style: { font: { bold: true } } },
          { type: "expression", formulae: ["A1>0"], style: { font: { italic: true } } }
        ]
      });
      DataValidation.add(sheet.dataValidations, "B1", {
        type: "custom",
        formulae: ["COUNTIF({1,2},B1)=0"]
      });
      DataValidation.add(sheet.dataValidations, "B2", { type: "custom", formulae: ["B2>0"] });
      DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Bad", "OFFSET(1,0,0)");
      DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Good", "OFFSET(S!$A$1,0,0)");
      const error = await Workbook.toBuffer(workbook, { format }).catch((cause: unknown) => cause);
      expect(error).toBeInstanceOf(ExcelUnsupportedFeatureError);
      // Each container words these losses as it words any formula of theirs it cannot write.
      expect((error as ExcelUnsupportedFeatureError).items).toEqual(
        format === "xlsx"
          ? [
              "S!A1:A3: conditional formatting formula (RANK argument 2 must be a reference)",
              "S!B1: data validation formula (COUNTIF argument 1 must be a reference)",
              "Bad: defined name formula (OFFSET argument 1 must be a reference)"
            ]
          : [
              'S!conditional formatting formula "RANK(A1,{1,2,3})=1"',
              "S!B1: data validation",
              "Bad: defined name definition"
            ]
      );

      const bytes = await Workbook.toBuffer(workbook, { format, unsupported: "ignore" });
      const parts = await extractAll(bytes);
      const text = [...parts.values()].map(part => new TextDecoder().decode(part.data)).join("\n");
      if (format === "xlsx") {
        expect(text).not.toContain("RANK(");
        expect(text).not.toContain("COUNTIF(");
        expect(text).not.toContain("OFFSET(1");
        expect(text).toContain("A1&gt;0");
        expect(text).toContain("B2&gt;0");
        expect(text).toContain("OFFSET(S!$A$1,0,0)");
      }
      const reopened = Workbook.create();
      await Workbook.read(reopened, bytes);
      const names = Workbook.getModel(reopened).definedNames.map(entry => entry.name);
      expect(names).toContain("Good");
      expect(names).not.toContain("Bad");
    }
  );
});
