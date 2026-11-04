/**
 * A cell value of a known shape — rich text, a formula, a shared formula, a hyperlink, an error, a checkbox — is checked
 * as that shape when it is set, so a malformed one is refused at the call instead of failing at write time or being
 * written as its JSON text. An object or array of no known shape keeps being stored as JSON text, as it always was.
 */
import { Cell, ExcelArgumentError, Row, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", { formula: "1+1", result: 2 });
  return { wb, ws };
}

describe("malformed values of a known shape", () => {
  it.each<[string, unknown]>([
    ["an unknown error literal", { error: "#BOGUS" }],
    ["an error that is not text", { error: 42 }],
    ["rich text that is not a list", { richText: 42 }],
    ["a rich-text run whose text is not text", { richText: [{ text: 42 }] }],
    ["a rich-text run that is not an object", { richText: ["x"] }],
    ["a rich-text run whose font is not an object", { richText: [{ text: "x", font: 5 }] }],
    ["an empty formula", { formula: "" }],
    ["a formula that is not text", { formula: 42 }],
    ["a formula result that is an object", { formula: "1", result: {} }],
    ["a formula result that is an unknown error", { formula: "1", result: { error: "#NOPE" } }],
    ["a shared formula whose master does not exist", { sharedFormula: "Z99" }],
    ["a shared formula whose master is not a formula", { sharedFormula: "B1" }],
    ["a shared formula whose master comes after it", { sharedFormula: "C9" }],
    ["an empty shared formula", { sharedFormula: "" }],
    ["a hyperlink that is not text", { text: "x", hyperlink: 42 }],
    ["a hyperlink text that is not text", { text: 5, hyperlink: "https://example.com" }],
    ["a hyperlink tooltip that is not text", { text: "x", hyperlink: "https://e.com", tooltip: 1 }],
    ["a checkbox that is not a boolean", { checkbox: "yes" }]
  ])("%s is refused at the call", async (name, value) => {
    const { wb, ws } = sheet();
    Cell.setValue(ws, "B1", 5);
    Cell.setValue(ws, "C9", { formula: "B1*2" });
    expect(() => Cell.setValue(ws, "B2", value as never)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "B2")).toBe(false);
    // Where a shared formula's master is is a question of position, which only `Cell.setValue` checks; the shape
    // checks apply to every value setter.
    if (!name.startsWith("a shared formula whose master")) {
      expect(() => Row.setValues(ws, 5, [value as never])).toThrow(ExcelArgumentError);
      expect(Row.has(ws, 5)).toBe(false);
      expect(() => Worksheet.addAoa(ws, [[value as never]], { origin: "H1" })).toThrow(
        ExcelArgumentError
      );
    }
    await expect(Workbook.toBuffer(wb)).resolves.toBeDefined();
  });
});

describe("well-formed values keep working", () => {
  it.each<[string, unknown]>([
    ["an error", { error: "#N/A" }],
    ["a newer error", { error: "#SPILL!" }],
    ["rich text", { richText: [{ text: "a", font: { bold: true } }, { text: "b" }] }],
    ["a formula with a result", { formula: "1+1", result: 2 }],
    ["a formula with an error result", { formula: "1/0", result: { error: "#DIV/0!" } }],
    ["a shared formula", { sharedFormula: "A1" }],
    ["a hyperlink", { text: "x", hyperlink: "https://example.com", tooltip: "go" }],
    ["a checkbox", { checkbox: true }]
  ])("%s", async (_name, value) => {
    const { wb, ws } = sheet();
    Cell.setValue(ws, "A2", value as never);
    expect(Cell.has(ws, "A2")).toBe(true);
    await expect(Workbook.toBuffer(wb)).resolves.toBeDefined();
  });

  it("an object or array of no known shape is stored as JSON text, as before", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "B2", { a: 1 } as never);
    Cell.setValue(ws, "B3", [1, 2] as never);
    expect(Cell.getText(ws, "B2")).toBe('{"a":1}');
    expect(Cell.getText(ws, "B3")).toBe("[1,2]");
  });
});
