import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

import { getWorksheets } from "@excel/core/workbook";
import { addTable, getTable } from "@excel/core/worksheet";
import { Cell, Workbook } from "@excel/index";
/**
 * End-to-end test for OOXML _xHHHH_ escape decoding in table column headers.
 *
 * Issue: https://github.com/documonster/documonster/issues/94
 *
 * Root cause: Excel stores newlines in `<tableColumn name="..."/>` attributes
 * as `_x000a_` (OOXML ST_Xstring escaping). On load, `Table.store()` copies
 * `column.name` into the header row cells, so an undecoded `_x000a_` leaks
 * into both `table.columns[].name` and the cell value visible to the user.
 */
import { describe, it, expect } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

describe("OOXML _xHHHH_ escape in table column headers (issue #94)", () => {
  it("reads Excel-generated file with _x000a_ in table column header", async () => {
    // This is the exact file attached to the bug report.
    // Inside the XLSX:
    //   xl/sharedStrings.xml  → <t>Col3⏎new line</t>    (literal newline)
    //   xl/tables/table1.xml  → name="Col3_x000a_new line" (OOXML escape)
    //
    // On load, Table.store() overwrites cell C1 with column.name,
    // so if _x000a_ isn't decoded the user sees the literal string.
    const filePath = path.join(__dirname, "../../__tests__/data/ooxml-escape-table-header.xlsx");
    const buffer = fs.readFileSync(filePath);

    const wb = Workbook.create();
    await Workbook.read(wb, buffer);
    const ws = getWorksheets(wb)[0]!;

    // Table column name must have the decoded newline
    const table = getTable(ws, "Table1_1");
    expect(table).toBeDefined();
    expect(table.table.columns[2].name).toBe("Col3\nnew line");

    // Cell value (overwritten by store()) must also have the decoded newline
    expect(Cell.getValue(ws, "C1")).toBe("Col3\nnew line");
  });

  it("table column header with newline survives write → read roundtrip", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");

    addTable(ws, {
      name: "TestTable",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Col1", filterButton: true },
        { name: "Col2", filterButton: true },
        { name: "Col3\nnew line", filterButton: true }
      ],
      rows: [
        ["a", "b", "c"],
        ["d", "e", "f"]
      ]
    });

    const buffer = await Workbook.toBuffer(wb);

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer as Buffer);
    const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

    const table = getTable(ws2, "TestTable");
    expect(table.table.columns[2].name).toBe("Col3\nnew line");
    expect(Cell.getValue(ws2, "C1")).toBe("Col3\nnew line");
  });
});
