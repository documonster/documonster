import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

/**
 * End-to-end test for OOXML _xHHHH_ escape decoding in table column headers.
 *
 * Issue: https://github.com/cjnoname/excelts/issues/94
 *
 * Root cause: Excel stores newlines in `<tableColumn name="..."/>` attributes
 * as `_x000a_` (OOXML ST_Xstring escaping). On load, `Table.store()` copies
 * `column.name` into the header row cells, so an undecoded `_x000a_` leaks
 * into both `table.columns[].name` and the cell value visible to the user.
 */
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

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
    const filePath = path.join(__dirname, "data/ooxml-escape-table-header.xlsx");
    const buffer = fs.readFileSync(filePath);

    const wb = new Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0]!;

    // Table column name must have the decoded newline
    const table = ws.getTable("Table1_1");
    expect(table).toBeDefined();
    expect(table.table.columns[2].name).toBe("Col3\nnew line");

    // Cell value (overwritten by store()) must also have the decoded newline
    expect(ws.getCell("C1").value).toBe("Col3\nnew line");
  });

  it("table column header with newline survives write → read roundtrip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.addTable({
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

    const buffer = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer as Buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    const table = ws2.getTable("TestTable");
    expect(table.table.columns[2].name).toBe("Col3\nnew line");
    expect(ws2.getCell("C1").value).toBe("Col3\nnew line");
  });
});
