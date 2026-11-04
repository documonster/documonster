/**
 * Rule collections keep their own copies of what they are given, and refuse a rule Excel cannot open.
 */
import {
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Row,
  Sparkline,
  Workbook
} from "@excel";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  return { wb, s1: Workbook.addWorksheet(wb, "S1"), s2: Workbook.addWorksheet(wb, "S2") };
}

describe("ConditionalFormatting.add keeps a copy", () => {
  it("one options object added to two sheets makes two blocks", () => {
    const { s1, s2 } = book();
    const options = {
      ref: "A2:A10",
      rules: [{ type: "expression" as const, formulae: ["A2>5"], style: { font: { bold: true } } }]
    };
    const a = ConditionalFormatting.add(s1, options);
    const b = ConditionalFormatting.add(s2, options);
    expect(a).not.toBe(options);
    expect(a).not.toBe(b);
    Row.insert(s1, 1, [1]);
    expect(ConditionalFormatting.list(s1)[0].ref).toBe("A3:A11");
    expect(ConditionalFormatting.list(s2)[0].ref).toBe("A2:A10");
    expect(ConditionalFormatting.list(s2)[0].rules[0]).toMatchObject({ formulae: ["A2>5"] });
    expect(options.ref).toBe("A2:A10");
  });

  it("editing the options afterwards does not reach the sheet", () => {
    const { s1 } = book();
    const options = { ref: "A1", rules: [{ type: "expression" as const, formulae: ["TRUE"] }] };
    const added = ConditionalFormatting.add(s1, options);
    options.rules[0].formulae[0] = "FALSE";
    expect(ConditionalFormatting.list(s1)).toEqual([added]);
    expect(added.rules[0]).toMatchObject({ formulae: ["TRUE"] });
    expect(ConditionalFormatting.remove(s1, added)).toBe(true);
  });
});

describe("ConditionalFormatting.add refuses what Excel cannot open", () => {
  it("a cellIs rule without its operator", () => {
    // Excel 16 (macOS) does not open such a workbook without a repair prompt.
    const { s1 } = book();
    expect(() =>
      ConditionalFormatting.add(s1, {
        ref: "A1",
        rules: [{ type: "cellIs", formulae: ["1"] } as never]
      })
    ).toThrow(ExcelArgumentError);
    expect(ConditionalFormatting.list(s1)).toEqual([]);
  });

  it("colours that are not colours", () => {
    const { s1 } = book();
    const cfvo = [{ type: "min" as const }, { type: "max" as const }];
    expect(() =>
      ConditionalFormatting.add(s1, {
        ref: "A1",
        rules: [{ type: "dataBar", cfvo, color: 42 as never }]
      })
    ).toThrow(ExcelArgumentError);
    expect(() =>
      ConditionalFormatting.add(s1, {
        ref: "A1",
        rules: [{ type: "colorScale", cfvo, color: [1, 2] as never }]
      })
    ).toThrow(ExcelArgumentError);
    expect(() =>
      ConditionalFormatting.add(s1, {
        ref: "A1",
        rules: [{ type: "colorScale", cfvo, color: [{ argb: "FFFF0000" }, { argb: "nope" }] }]
      })
    ).toThrow(ExcelArgumentError);
    ConditionalFormatting.add(s1, {
      ref: "A1",
      rules: [{ type: "colorScale", cfvo, color: [{ argb: "FFFF0000" }, { theme: 4 }] }]
    });
    expect(ConditionalFormatting.list(s1)).toHaveLength(1);
  });
});

describe("Sparkline.add keeps a copy of the sparklines", () => {
  it("reusing the options does not share them", () => {
    const { s1, s2 } = book();
    const options = { type: "line" as const, sparklines: [{ dataRef: "S1!A2:E2", cellRef: "F2" }] };
    const a = Sparkline.add(s1, options);
    const b = Sparkline.add(s2, options);
    expect(a.sparklines).not.toBe(options.sparklines);
    expect(a.sparklines).not.toBe(b.sparklines);
    options.sparklines[0].cellRef = "Z9";
    expect(Sparkline.list(s1)[0].sparklines[0].cellRef).toBe("F2");
  });
});

describe("DataValidation.add", () => {
  it("requires a rule; DataValidation.remove is what removes one", () => {
    const { s1 } = book();
    DataValidation.add(s1, "A1", { type: "whole", operator: "greaterThan", formulae: [1] });
    expect(() => DataValidation.add(s1, "A1", undefined as never)).toThrow(ExcelArgumentError);
    expect(DataValidation.find(s1, "A1")).toMatchObject({ type: "whole" });
    expect(DataValidation.remove(s1, "A1")).toBe(true);
  });
});
