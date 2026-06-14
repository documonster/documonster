import { cellGetValue } from "@excel/cell";
import { Table, Workbook } from "@excel/index";
import { tableAddRow, tableModel } from "@excel/table";
import { getCell } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

describe("Table", () => {
  it("supports getTable().addRow() after loading a workbook", async () => {
    const wb1 = Workbook.create();
    const ws1: any = Workbook.addWorksheet(wb1, "Data");

    Table.add(ws1, {
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

    const template = await Workbook.toXlsxBuffer(wb1);
    await expectValidXlsx(template, { label: "table-addrow-loaded template" });

    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, template);

    const ws2: any = Workbook.getWorksheet(wb2, "Data")!;
    const table2 = Table.get(ws2, "TestTable");

    // The key workflow: addRow should update the worksheet and table refs without requiring
    // the caller to manually call tableCommit(table).
    tableAddRow(table2, ["Carol", 41]);

    expect(cellGetValue(getCell(ws2, "A4"))).toBe("Carol");
    expect(cellGetValue(getCell(ws2, "B4"))).toBe(41);

    // Table reference should expand to include the new row
    expect(tableModel(table2).tableRef).toBe("A1:B4");

    // AutoFilter reference must cover the entire filterable range —
    // header row + all data rows (excluding any totals row). Real
    // Excel emits the filter ref at the full table extent; a
    // header-only ref like `A1:B1` makes Excel reject the table on
    // open with "Removed Records: Table from /xl/tables/tableN.xml".
    expect(tableModel(table2).autoFilterRef).toBe("A1:B4");

    const out = await Workbook.toXlsxBuffer(wb2);
    await expectValidXlsx(out, { label: "table-addrow-loaded after addRow" });

    const wb3 = Workbook.create();
    await Workbook.loadXlsx(wb3, out);
    const ws3: any = Workbook.getWorksheet(wb3, "Data")!;
    const table3 = Table.get(ws3, "TestTable");

    expect(tableModel(table3).tableRef).toBe("A1:B4");
    expect(tableModel(table3).autoFilterRef).toBe("A1:B4");
    expect(cellGetValue(getCell(ws3, "A4"))).toBe("Carol");
    expect(cellGetValue(getCell(ws3, "B4"))).toBe(41);
  });

  it("throws on duplicate table names across worksheets", () => {
    const wb = Workbook.create();
    const ws1: any = Workbook.addWorksheet(wb, "Sheet1");
    const ws2: any = Workbook.addWorksheet(wb, "Sheet2");

    const tableConfig = {
      name: "MyTable",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    };

    Table.add(ws1, tableConfig);
    expect(() => Table.add(ws2, tableConfig)).toThrow(/already exists/i);
  });

  it("throws on duplicate table names (case-insensitive)", () => {
    const wb = Workbook.create();
    const ws1: any = Workbook.addWorksheet(wb, "Sheet1");
    const ws2: any = Workbook.addWorksheet(wb, "Sheet2");

    Table.add(ws1, {
      name: "Sales",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      Table.add(ws2, {
        name: "SALES",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("allows distinct table names across worksheets", async () => {
    const wb = Workbook.create();
    const ws1: any = Workbook.addWorksheet(wb, "Sheet1");
    const ws2: any = Workbook.addWorksheet(wb, "Sheet2");

    Table.add(ws1, {
      name: "Table1",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });
    Table.add(ws2, {
      name: "Table2",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    // Should not throw
    const buffer = await Workbook.toXlsxBuffer(wb);
    await expectValidXlsx(buffer, { label: "distinct table names across sheets" });
    expect(buffer).toBeTruthy();
  });

  it("throws on duplicate table names within the same worksheet", () => {
    const wb = Workbook.create();
    const ws: any = Workbook.addWorksheet(wb, "Sheet1");

    Table.add(ws, {
      name: "Orders",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      Table.add(ws, {
        name: "Orders",
        ref: "D1",
        columns: [{ name: "Col2", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("throws when sanitized table names collide on the same worksheet", () => {
    const wb = Workbook.create();
    const ws: any = Workbook.addWorksheet(wb, "Sheet1");

    // "My Table" sanitizes to "My_Table" (spaces become underscores)
    Table.add(ws, {
      name: "My Table",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    expect(() =>
      Table.add(ws, {
        name: "My_Table",
        ref: "D1",
        columns: [{ name: "Col2", filterButton: true }],
        rows: [["other"]]
      })
    ).toThrow(/already exists/i);
  });

  it("allows reusing a table name after removeTable", () => {
    const wb = Workbook.create();
    const ws: any = Workbook.addWorksheet(wb, "Sheet1");

    Table.add(ws, {
      name: "Reuse",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    Table.remove(ws, "Reuse");

    // Should succeed — name is freed
    expect(() =>
      Table.add(ws, {
        name: "Reuse",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["new data"]]
      })
    ).not.toThrow();
  });

  it("allows reusing a table name on another worksheet after removeTable", () => {
    const wb = Workbook.create();
    const ws1: any = Workbook.addWorksheet(wb, "Sheet1");
    const ws2: any = Workbook.addWorksheet(wb, "Sheet2");

    Table.add(ws1, {
      name: "Shared",
      ref: "A1",
      columns: [{ name: "Col1", filterButton: true }],
      rows: [["data"]]
    });

    Table.remove(ws1, "Shared");

    // Name freed globally — another worksheet can now use it
    expect(() =>
      Table.add(ws2, {
        name: "Shared",
        ref: "A1",
        columns: [{ name: "Col1", filterButton: true }],
        rows: [["other"]]
      })
    ).not.toThrow();
  });
});
