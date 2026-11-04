/**
 * A row or column index past the sheet's grid is refused where it is decoded, not modelled.
 *
 * A fuzzed workbook carried a defined name whose area started at row `0x00EE0000` — 15.6 million rows past the last
 * one. The reader decoded it into an ordinary range and handed it to the defined-name registry, which walked the area
 * cell by cell and allocated until the process aborted: an out-of-memory abort, which no caller can catch. The same
 * four numbers arrive in a merge, a hyperlink, a validation, a conditional format and a formula token, so each is
 * checked here.
 */
import { extractAll } from "@archive/unzip/extract";
import { packZipSync } from "@archive/zip-one-shot";
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Workbook,
  Worksheet
} from "@excel";
import { XlsbParseError } from "@excel/errors";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { readWorksheetPart } from "@excel/xlsb/read/parts";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

const OFF_GRID = 0x00ee0000;

/** A sheet whose every feature covers `A5:B6` — rows 4 and 5, which no other field of these records holds. */
async function packageWith(feature: string): Promise<Map<string, Uint8Array>> {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValue(sheet, "A1", 1);
  switch (feature) {
    case "BrtName":
      DefinedNames.add(Workbook.getDefinedNames(workbook), "S!$A$5:$B$6", "Big");
      break;
    case "BrtMergeCell":
      Worksheet.merge(sheet, "A5:B6");
      break;
    case "BrtHLink":
      Cell.setValue(sheet, "A5", { text: "x", hyperlink: "https://example.com" });
      break;
    case "BrtDVal":
      DataValidation.add(sheet.dataValidations, "A5:B6", {
        type: "whole",
        operator: "between",
        formulae: [1, 9]
      });
      break;
    case "BrtBeginConditionalFormatting":
      ConditionalFormatting.add(sheet, {
        ref: "A5:B6",
        rules: [{ type: "expression", formulae: ["A5>1"], style: { font: { bold: true } } }]
      });
      break;
    case "BrtFmlaNum":
      Cell.setValue(sheet, "C1", { formula: "SUM(S!$A$5:$B$6)", result: 0 });
      break;
  }
  const archive = await extractAll(await Workbook.toBuffer(workbook, { format: "xlsb" }));
  return new Map([...archive].map(([path, entry]) => [path, entry.data.slice()]));
}

/**
 * Move the first row of every `A5:B6` the named records hold off the grid.
 *
 * An area stores `rowFirst`, `rowLast` as two consecutive `u32`s, both in an `RfX` and in a `PtgArea`; `A5:B6` is
 * `4, 5` (and the single cell `A5` is `4, 4`), which nothing else in these small records spells.
 */
function moveOffGrid(part: Uint8Array, record: string): number {
  let patched = 0;
  for (const entry of iterateInterpretableRecords(part, "p")) {
    if (recordSpec(entry.id)?.name !== record) {
      continue;
    }
    const view = new DataView(entry.payload.buffer, entry.payload.byteOffset);
    for (let at = 0; at + 8 <= entry.payload.length; at++) {
      const last = view.getUint32(at + 4, true);
      if (view.getUint32(at, true) === 4 && (last === 4 || last === 5)) {
        view.setUint32(at, OFF_GRID, true);
        patched++;
      }
    }
  }
  return patched;
}

describe("indexes past the grid", () => {
  it.each(["BrtMergeCell", "BrtHLink", "BrtBeginConditionalFormatting"])(
    "a %s range is a parse error",
    async record => {
      const sheet = (await packageWith(record)).get("xl/worksheets/sheet1.bin")!;
      expect(moveOffGrid(sheet, record)).toBeGreaterThan(0);
      expect(() => readWorksheetPart(sheet, "xl/worksheets/sheet1.bin", [])).toThrow(
        XlsbParseError
      );
    }
  );

  // A `BrtDVal` that does not decode is skipped as a whole — the policy for every malformed validation — so the
  // range is never modelled.
  it("a BrtDVal range drops the validation", async () => {
    const sheet = (await packageWith("BrtDVal")).get("xl/worksheets/sheet1.bin")!;
    expect(moveOffGrid(sheet, "BrtDVal")).toBe(1);
    expect(readWorksheetPart(sheet, "xl/worksheets/sheet1.bin", []).validations).toEqual([]);
  });

  it("a defined name's area is reported as undecodable and the workbook still reads", async () => {
    const files = await packageWith("BrtName");
    expect(moveOffGrid(files.get("xl/workbook.bin")!, "BrtName")).toBe(1);
    const read = await Workbook.read(Workbook.create(), packZipSync(files));
    expect(DefinedNames.list(Workbook.getDefinedNames(read)).map(n => n.name)).not.toContain("Big");
  });

  it("a formula's area leaves the cell's cached value and reports the formula", async () => {
    const files = await packageWith("BrtFmlaNum");
    expect(moveOffGrid(files.get("xl/worksheets/sheet1.bin")!, "BrtFmlaNum")).toBe(1);
    const read = await Workbook.read(Workbook.create(), packZipSync(files), {
      unsupported: "ignore"
    });
    expect(Cell.getFormula(Workbook.getWorksheet(read, "S")!, "C1")).toBeUndefined();
  });
});
