/**
 * Read and write options are checked where they are passed. (A compression failure inside the ZIP writer rejecting
 * the write is pinned in the `.node` sibling.)
 */
import { Cell, ExcelArgumentError, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function workbook() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (let i = 1; i <= 10; i++) {
    Cell.setValue(ws, `A${i}`, `v${i}`);
  }
  return wb;
}

describe("write options", () => {
  it.each<[string, Record<string, unknown>]>([
    ["zip.level 42", { zip: { level: 42 } }],
    ["zip.level -5", { zip: { level: -5 } }],
    ["zip.level '9'", { zip: { level: "9" } }],
    ["zip.level 1.5", { zip: { level: 1.5 } }],
    ["zip 42", { zip: 42 }],
    ["zip.modTime invalid", { zip: { modTime: new Date(NaN) } }],
    ["zip.modTime a string", { zip: { modTime: "2024-01-01" } }],
    ["zip.timestamps bogus", { zip: { timestamps: "bogus" } }],
    ["useSharedStrings 'no'", { useSharedStrings: "no" }],
    ["useStyles 1", { useStyles: 1 }],
    ["templateMode bogus", { templateMode: "bogus" }],
    ["strictTemplateMode 'yes'", { strictTemplateMode: "yes" }],
    ["validate 'yes'", { validate: "yes" }]
  ])("%s is refused by toBuffer, writeStream and toStream", async (_name, options) => {
    const wb = workbook();
    await expect(Workbook.toBuffer(wb, options as never)).rejects.toThrow(ExcelArgumentError);
    expect(() => Workbook.toStream(wb, options as never)).toThrow(ExcelArgumentError);
  });

  it("accepts every level from 0 to 9", async () => {
    const wb = workbook();
    for (let level = 0; level <= 9; level++) {
      await expect(Workbook.toBuffer(wb, { zip: { level } })).resolves.toBeDefined();
    }
  });
});

describe("read options", () => {
  it.each<[string, Record<string, unknown>]>([
    ["maxRows NaN", { maxRows: NaN }],
    ["maxRows 0", { maxRows: 0 }],
    ["maxRows '5'", { maxRows: "5" }],
    ["maxRows 1.5", { maxRows: 1.5 }],
    ["maxCols -1", { maxCols: -1 }],
    ["ignoreNodes a string", { ignoreNodes: "sheetData" }],
    ["ignoreNodes [42]", { ignoreNodes: [42] }],
    ["zip 42", { zip: 42 }],
    ["zip.maxEntries 'x'", { zip: { maxEntries: "x" } }],
    ["zip.maxEntrySize -1", { zip: { maxEntrySize: -1 } }],
    ["zip.checkCrc32 'no'", { zip: { checkCrc32: "no" } }],
    ["blankCells bogus", { blankCells: "bogus" }],
    ["formulas bogus", { formulas: "bogus" }],
    ["formulaSyntaxProbe 42", { formulaSyntaxProbe: 42 }]
  ])("%s is refused", async (_name, options) => {
    const buf = await Workbook.toBuffer(workbook());
    await expect(Workbook.read(buf, options as never)).rejects.toThrow(ExcelArgumentError);
    await expect(Workbook.readWithDiagnostics(buf, options as never)).rejects.toThrow(
      ExcelArgumentError
    );
  });

  it("accepts Infinity to lift a limit, and an array of node names", async () => {
    const buf = await Workbook.toBuffer(workbook());
    const wb = await Workbook.read(buf, {
      maxRows: Infinity,
      maxCols: 5,
      ignoreNodes: ["dataValidations"],
      zip: { maxEntries: Infinity, maxEntrySize: 1e9, checkCrc32: false }
    });
    expect(Cell.getValue(Workbook.getWorksheet(wb, "S")!, "A10")).toBe("v10");
  });
});
