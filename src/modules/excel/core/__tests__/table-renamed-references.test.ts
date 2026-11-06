/**
 * A table a file names with something Excel cannot hold — a name that reads as a cell address — is stored under a
 * sanitised one (`T2` → `_T2`), and every formula the file wrote with the file's name has to say the stored one — or
 * Excel refuses the workbook. `Table.add` refuses such a name, so only a file can carry one.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Cell, ConditionalFormatting, DefinedNames, Table, Workbook } from "@excel";
import { restateFormula } from "@excel/core/table-references";
import { encodeTablePart } from "@excel/xlsb/tables";
import { describe, expect, it } from "vitest";

function build() {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValue(sheet, "F1", { formula: "SUM(_T2[x])" });
  Table.add(sheet, {
    name: "_T2",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "x" }, { name: "y" }],
    rows: [
      [1, 2],
      [3, 4]
    ]
  });
  Cell.setValue(sheet, "F2", { formula: 'COUNTA(_T2[[#All],[y]])&"T2[x]"' });
  ConditionalFormatting.add(sheet, {
    ref: "F1",
    rules: [{ type: "expression", formulae: ["F1>MAX(_T2[y])"], style: { font: { bold: true } } }]
  });
  DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Total", "SUM(_T2[x])");
  return workbook;
}

describe("references to a table renamed by sanitisation", () => {
  it("restates only a name that is not a table and sanitises to one", () => {
    const tables = new Set(["_t2", "tb"]);
    expect(restateFormula("SUM(T2[x])+tb[y]", tables)).toBe("SUM(_T2[x])+tb[y]");
    expect(restateFormula('"T2[x]"&_T2[[#Headers],[T2[]]', tables)).toBe(
      '"T2[x]"&_T2[[#Headers],[T2[]]'
    );
    expect(restateFormula("Other[x]", tables)).toBe("Other[x]");
  });

  it("is a name Table.add refuses, so a workbook built here never needs it", () => {
    const sheet = Workbook.addWorksheet(Workbook.create(), "S");
    expect(() =>
      Table.add(sheet, { name: "T2", ref: "A1", columns: [{ name: "x" }], rows: [[1]] })
    ).toThrow(/"_T2" is one Excel accepts/);
  });

  /**
   * A file whose table is literally named `T2` — another producer's, since this library never writes one. The workbook
   * stores the table as `_T2`, and the formulas read with it used to keep `T2[x]`, naming no table the workbook holds.
   * They are read as `_T2[x]` now, in both containers, so the model is consistent and a write reproduces it.
   */
  it.each(["xlsx", "xlsb"] as const)(
    "reads formulas naming a table the file calls T2 under its stored name (%s)",
    async format => {
      const entries = await extractAll(await Workbook.toBuffer(build(), { format }));
      const patched = [...entries.entries()].map(([name, file]) => {
        if (format === "xlsx" && /^xl\/(tables\/table1|worksheets\/sheet1)\.xml$/.test(name)) {
          const text = new TextDecoder().decode(file.data).replaceAll("_T2", "T2");
          return { name, data: new TextEncoder().encode(text) };
        }
        if (format === "xlsb" && name === "xl/tables/table1.bin") {
          // The table part again, under the name a file may give it; `PtgList` names the table by id.
          return {
            name,
            data: encodeTablePart({
              ref: "A1:B3",
              name: "T2",
              id: 1,
              columns: [{ name: "x" }, { name: "y" }]
            })!
          };
        }
        return { name, data: file.data };
      });
      const reread = Workbook.create();
      await Workbook.read(reread, await createZip(patched));
      const sheet = Workbook.getWorksheet(reread, "S")!;
      expect(Table.list(sheet).map(table => Table.getName(table))).toEqual(["_T2"]);
      expect(Cell.getFormula(sheet, "F1")).toBe("SUM(_T2[x])");
      const again = Workbook.create();
      await Workbook.read(again, await Workbook.toBuffer(reread, { format }));
      expect(Cell.getFormula(Workbook.getWorksheet(again, "S")!, "F1")).toBe("SUM(_T2[x])");
    }
  );
});
