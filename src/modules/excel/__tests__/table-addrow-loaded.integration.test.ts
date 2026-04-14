import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

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

    // AutoFilter reference should remain header-row-only
    expect(table2.model.autoFilterRef).toBe("A1:B1");

    const out = await wb2.xlsx.writeBuffer();

    const wb3 = new Workbook();
    await wb3.xlsx.load(out);
    const ws3: any = wb3.getWorksheet("Data");
    const table3 = ws3.getTable("TestTable");

    expect(table3.model.tableRef).toBe("A1:B4");
    expect(table3.model.autoFilterRef).toBe("A1:B1");
    expect(ws3.getCell("A4").value).toBe("Carol");
    expect(ws3.getCell("B4").value).toBe(41);
  });
});
