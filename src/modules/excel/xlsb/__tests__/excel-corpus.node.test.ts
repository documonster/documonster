/**
 * Formulas Excel wrote, read from both of its containers.
 *
 * Each pair under `__tests__/data/*.excel.{xlsb,xlsx}` is one workbook saved by real Excel twice — once as XLSB, once
 * as XLSX — from a single source (`formula-matrix` is this library's own acceptance matrix written as XLSX;
 * `formula-edges` is hand-written XML of the forms that are hard to get right: spacing, line breaks, display
 * parentheses, reference operators, deleted references, sheet-local names, multi-range anchors, every conditional
 * format kind and every defined-name shape). XLSX stores a formula as text and XLSB as tokens, so the two reads
 * agreeing is the statement "the token decoder produces what Excel shows" — checked against Excel, not against
 * this codec's own encoder.
 *
 * What it found the first time: `SUM(F5-I5)` read as `F5-I5` (`PtgAttrSum` ignored), `_xlfn.CONCAT` read without its
 * prefix, every space and redundant parenthesis dropped, `{1,2;3,4}` refused, `Other!LocalO` refused, constant and
 * formula names dropped, local names made global, conditional formats and validations anchored at the first range
 * instead of the bounding box, `EXACT` and `LENB` unreadable, `Tbl[#All]` printed `Tbl[[#All]]`.
 *
 * `.node` because it reads fixtures off disk, which the browser config excludes by glob.
 */
import { readFile } from "node:fs/promises";

import { Cell, Workbook, Worksheet } from "@excel";
import { describe, expect, it } from "vitest";

const DIRECTORY = "src/modules/excel/xlsb/__tests__/data";
const FIXTURES = ["formula-matrix", "formula-edges"];

/** Every formula-bearing thing in a workbook, keyed so the two reads line up. */
async function formulas(path: string, format?: "xlsb"): Promise<Record<string, unknown>> {
  const workbook = Workbook.create();
  await Workbook.read(
    workbook,
    await readFile(path),
    format === undefined ? undefined : { format }
  );
  const out: Record<string, unknown> = {};
  for (const sheet of Workbook.listWorksheets(workbook)) {
    const model = Worksheet.getModel(sheet);
    for (const row of model.rows ?? []) {
      for (const cell of row.cells ?? []) {
        const c = cell as {
          formula?: string;
          sharedFormula?: string;
          shareType?: string;
          ref?: string;
        };
        if (c.formula === undefined && c.sharedFormula === undefined && c.shareType === undefined) {
          continue;
        }
        // The text a caller sees — a shared formula's member resolved against its master.
        out[`${model.name}!${cell.address}`] = Cell.getFormula(sheet, cell.address);
        // An array formula's range. *Not* a shared formula's: whether a fill is stored shared is the container's
        // choice — Excel keeps `E20:E23` shared in XLSX and writes it out cell by cell in XLSB — and the text above
        // already says what each cell computes.
        if (c.shareType === "array") {
          out[`${model.name}!${cell.address} array`] = c.ref;
        }
      }
    }
    for (const formatting of model.conditionalFormattings ?? []) {
      formatting.rules.forEach((rule, index) => {
        // Everything but the differential format, whose numbering is the container's.
        const {
          dxfId: _dxfId,
          style: _style,
          ...shape
        } = rule as unknown as Record<string, unknown>;
        out[`${model.name} CF ${formatting.ref.split(" ").sort().join(" ")} #${index}`] = shape;
      });
    }
    for (const [address, validation] of Object.entries(model.dataValidations ?? {})) {
      out[`${model.name} DV ${address}`] = validation;
    }
  }
  for (const name of Workbook.getModel(workbook).definedNames ?? []) {
    out[`name ${name.name}@${name.localSheetId ?? "workbook"}`] = name;
  }
  return out;
}

describe.each(FIXTURES)("%s: XLSB reads what XLSX reads", fixture => {
  it("for every cell, array, conditional format, validation and defined name", async () => {
    const xlsx = await formulas(`${DIRECTORY}/${fixture}.excel.xlsx`);
    const xlsb = await formulas(`${DIRECTORY}/${fixture}.excel.xlsb`, "xlsb");
    // Not vacuous: the fixtures are formula-heavy by construction.
    expect(Object.keys(xlsx).length).toBeGreaterThan(fixture === "formula-matrix" ? 190 : 100);
    expect(xlsb).toEqual(xlsx);
  });
});
