import { readFileSync } from "node:fs";
import { join } from "node:path";

import { Cell, Table, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * A table column's data format in XLSB: `BrtBeginListCol.nDxfInsertRow`, the XLSX `dataDxfId`. It was written as none
 * and not read, so a column's format survived one container and not the other.
 */

const DATA = join(import.meta.dirname, "data");

function columnFormats(wb: Workbook.Handle, sheet: string, table: string): unknown[] {
  const handle = Table.get(Workbook.getWorksheet(wb, sheet)!, table)!;
  return Table.getModel(handle).columns.map(column => column.style);
}

describe("an XLSB table column's data format", () => {
  it("reads the format Excel stored, as the same workbook's XLSX gives it", async () => {
    // Excel saved this workbook both ways: `dataDxfId="1"` and `"0"` in the XLSX are `nDxfInsertRow` 1 and 0.
    const xlsx = await Workbook.read(readFileSync(join(DATA, "table-this-row.excel.xlsx")));
    const xlsb = await Workbook.read(readFileSync(join(DATA, "table-this-row.excel.xlsb")));
    const [sheet] = Workbook.listWorksheets(xlsx);
    const [table] = Table.list(sheet!);
    const name = Table.getModel(table!).name;
    const sheetName = Worksheet.getName(sheet!);
    const fromXlsx = columnFormats(xlsx, sheetName, name);
    expect(fromXlsx.filter(style => style !== undefined)).toHaveLength(2);
    expect(columnFormats(xlsb, sheetName, name)).toEqual(fromXlsx);
  });

  it("writes it, so it reads back as the XLSX does", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", "x");
    Table.add(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      columns: [
        { name: "Name" },
        { name: "Amount", style: { numFmt: "0.00", font: { bold: true } } }
      ],
      rows: [
        ["a", 1],
        ["b", 2]
      ]
    });
    const viaXlsb = await Workbook.read(await Workbook.toBuffer(wb, { format: "xlsb" }));
    const viaXlsx = await Workbook.read(await Workbook.toBuffer(wb));
    const formats = columnFormats(viaXlsb, "S", "T");
    expect(formats[0]).toBeUndefined();
    expect(formats[1]).toMatchObject({ numFmt: "0.00", font: { bold: true } });
    expect(formats).toEqual(columnFormats(viaXlsx, "S", "T"));
  });
});
