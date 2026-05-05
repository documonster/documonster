import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";
import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

describe("Table", () => {
  it("supports getTable().addRow() after loading a workbook", async () => {
    const wb1 = new Workbook();
    const ws1: any = wb1.addWorksheet("Data");

    ws1.addTable({
      name: "TestTable",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [
        { name: "Name", filterButton: true },
        { name: "Age", filterButton: true }
      ],
      rows: [
        ["Alice", 30],
        ["Bob", 25]
      ]
    });

    const template = await wb1.xlsx.writeBuffer();
    await expectValidXlsx(template, { label: "table-addrow-loaded template" });

    const wb2 = new Workbook();
    await wb2.xlsx.load(template);

    const ws2: any = wb2.getWorksheet("Data");
    const table2 = ws2.getTable("TestTable");

    // The key workflow: addRow should update the worksheet and table refs without requiring
    // the caller to manually call table.commit().
    table2.addRow(["Carol", 41]);

    expect(ws2.getCell("A4").value).toBe("Carol");
    expect(ws2.getCell("B4").value).toBe(41);

    // Table reference should expand to include the new row
    expect(table2.model.tableRef).toBe("A1:B4");

    // AutoFilter reference must cover the entire filterable range —
    // header row + all data rows (excluding any totals row). Real
    // Excel emits the filter ref at the full table extent; a
    // header-only ref like `A1:B1` makes Excel reject the table on
    // open with "Removed Records: Table from /xl/tables/tableN.xml".
    expect(table2.model.autoFilterRef).toBe("A1:B4");

    const out = await wb2.xlsx.writeBuffer();
    await expectValidXlsx(out, { label: "table-addrow-loaded after addRow" });

    const wb3 = new Workbook();
    await wb3.xlsx.load(out);
    const ws3: any = wb3.getWorksheet("Data");
    const table3 = ws3.getTable("TestTable");

    expect(table3.model.tableRef).toBe("A1:B4");
    expect(table3.model.autoFilterRef).toBe("A1:B4");
    expect(ws3.getCell("A4").value).toBe("Carol");
    expect(ws3.getCell("B4").value).toBe(41);
  });

  it("throws on duplicate table names across worksheets", () => {
    const wb = new Workbook();
    const ws1: any = wb.addWorksheet("Sheet1");
    const ws2: any = wb.addWorksheet("Sheet2");

    const tableConfig = {
      name: "MyTable",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    };

    ws1.addTable(tableConfig);
    expect(() => ws2.addTable(tableConfig)).toThrow(/already exists/i);
  });

  it("throws on duplicate table names (case-insensitive)", () => {
    const wb = new Workbook();
    const ws1: any = wb.addWorksheet("Sheet1");
    const ws2: any = wb.addWorksheet("Sheet2");

    ws1.addTable({
      name: "Sales",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      ws2.addTable({
        name: "SALES",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("allows distinct table names across worksheets", async () => {
    const wb = new Workbook();
    const ws1: any = wb.addWorksheet("Sheet1");
    const ws2: any = wb.addWorksheet("Sheet2");

    ws1.addTable({
      name: "Table1",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });
    ws2.addTable({
      name: "Table2",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    // Should not throw
    const buffer = await wb.xlsx.writeBuffer();
    await expectValidXlsx(buffer, { label: "distinct table names across sheets" });
    expect(buffer).toBeTruthy();
  });

  it("throws on duplicate table names within the same worksheet", () => {
    const wb = new Workbook();
    const ws: any = wb.addWorksheet("Sheet1");

    ws.addTable({
      name: "Orders",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      ws.addTable({
        name: "Orders",
        ref: "D1",
        columns: [{ name: "Col2", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("throws when sanitized table names collide on the same worksheet", () => {
    const wb = new Workbook();
    const ws: any = wb.addWorksheet("Sheet1");

    // "My Table" sanitizes to "My_Table" (spaces become underscores)
    ws.addTable({
      name: "My Table",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      ws.addTable({
        name: "My_Table",
        ref: "D1",
        columns: [{ name: "Col2", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("allows reusing a table name after removeTable", () => {
    const wb = new Workbook();
    const ws: any = wb.addWorksheet("Sheet1");

    ws.addTable({
      name: "Reuse",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    ws.removeTable("Reuse");

    // Should succeed — name is freed
    expect(() =>
      ws.addTable({
        name: "Reuse",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["new data"]]
      })
    ).not.toThrow();
  });

  it("allows reusing a table name on another worksheet after removeTable", () => {
    const wb = new Workbook();
    const ws1: any = wb.addWorksheet("Sheet1");
    const ws2: any = wb.addWorksheet("Sheet2");

    ws1.addTable({
      name: "Shared",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    ws1.removeTable("Shared");

    // Name freed globally — another worksheet can now use it
    expect(() =>
      ws2.addTable({
        name: "Shared",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["other"]]
      })
    ).not.toThrow();
  });
});
