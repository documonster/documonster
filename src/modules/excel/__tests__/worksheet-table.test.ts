import { extractAll } from "@archive/unzip/extract";
import { sanitizeTableName } from "@excel/table";
import { colCache } from "@excel/utils/col-cache";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

const spliceArray = (a: any[], index: number, count: number, ...rest: any[]) => {
  const clone = [...a];
  clone.splice(index, count, ...rest);
  return clone;
};

const values = [
  ["Date", "Id", "Word"],
  [new Date("2019-08-01"), 1, "Bird"],
  [new Date("2019-08-02"), 2, "is"],
  [new Date("2019-08-03"), 3, "the"],
  [new Date("2019-08-04"), 4, "Word"],
  ["Totals", { formula: "SUBTOTAL(104,TestTable[Id])", result: 4 }, null]
];

function addTable(ref: string, ws: any) {
  return ws.addTable({
    name: "TestTable",
    ref,
    headerRow: true,
    totalsRow: true,
    style: {
      theme: "TableStyleDark3",
      showRowStripes: true
    },
    columns: [
      { name: "Date", totalsRowLabel: "Totals", filterButton: true },
      {
        name: "Id",
        totalsRowFunction: "max",
        filterButton: true,
        totalsRowResult: 4
      },
      {
        name: "Word",
        filterButton: false,
        style: { font: { bold: true, name: "Comic Sans MS" } }
      }
    ],
    rows: [
      [new Date("2019-08-01"), 1, "Bird"],
      [new Date("2019-08-02"), 2, "is"],
      [new Date("2019-08-03"), 3, "the"],
      [new Date("2019-08-04"), 4, "Word"]
    ]
  });
}

function checkTable(ref: string, ws: any, testValues: any[]) {
  const a = colCache.decodeAddress(ref);

  for (let i = -1; i <= testValues.length + 1; i++) {
    const vRow = testValues[i];
    const nRow = i + a.row;
    const row = nRow >= 1 && ws.getRow(nRow);
    if (!row) {
      continue;
    }
    for (let j = -1; j <= testValues[0].length + 1; j++) {
      const value = (vRow && vRow[j]) || null;
      const nCol = j + a.col;
      const cellValue = nCol >= 1 && row.getCell(nCol).value;
      if (!cellValue) {
        continue;
      }

      if (value instanceof Date) {
        expect(cellValue).toEqual(value);
      } else if (value === null) {
        expect(cellValue).toBeNull();
      } else if (typeof value === "object") {
        expect(cellValue).toEqual(value);
      } else {
        expect(cellValue).toBe(value);
      }
    }
  }
}

describe("Worksheet", () => {
  describe("Table", () => {
    it("creates a table", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      addTable("A1", ws);

      checkTable("A1", ws, values);
    });

    it("removes header", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.headerRow = false;
      table.commit();

      const newValues = spliceArray(values, 0, 1);
      checkTable("A1", ws, newValues);
    });

    it("removes totals", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.totalsRow = false;
      table.commit();

      const newValues = spliceArray(values, 5, 1);
      checkTable("A1", ws, newValues);
    });

    it("moves the table", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.ref = "C2";
      table.commit();

      checkTable("C2", ws, values);
    });

    it("removes a row", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.removeRows(1);
      table.commit();

      const newValues = spliceArray(values, 2, 1);
      checkTable("A1", ws, newValues);
    });

    it("adds a row", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.addRow([new Date("2019-08-05"), 5, "Bird"]);
      table.commit();

      const newValues = spliceArray(values, 5, 0, [new Date("2019-08-05"), 5, "Bird"]);
      checkTable("A1", ws, newValues);
    });

    it("removes a column", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.removeColumns(1);
      table.commit();

      const newValues = values.map(rVals => spliceArray(rVals, 1, 1));
      checkTable("A1", ws, newValues);
    });

    it("adds a column", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      table.addColumn(
        {
          name: "Letter",
          totalsRowFunction: "custom",
          totalsRowFormula: "ROW()",
          totalsRowResult: 6,
          filterButton: true
        },
        ["a", "b", "c", "d"],
        2
      );
      table.commit();

      const colValues = ["Letter", "a", "b", "c", "d", { formula: "ROW()", result: 6 }];
      const newValues = values.map((rVals, i) => spliceArray(rVals, 2, 0, colValues[i]));
      checkTable("A1", ws, newValues);
    });

    it("renames a column", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const table = addTable("A1", ws);

      const column = table.getColumn(1);
      column.name = "Code";
      table.commit();

      const newValues = [...values];
      newValues.splice(0, 1, ["Date", "Code", "Word"]);
      newValues.splice(5, 1, [
        "Totals",
        { formula: "SUBTOTAL(104,TestTable[Code])", result: 4 },
        null
      ]);

      checkTable("A1", ws, newValues);
    });

    it("keeps implicit structured references by default", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.addTable({
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "[@A]" }]]
      });

      const cellValue = ws.getRow(2).getCell(2).value;
      expect(cellValue).toEqual({ formula: "[@A]" });
    });

    it("qualifies implicit structured references when enabled", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.addTable({
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        qualifyImplicitStructuredReferences: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "[@A]" }]]
      });

      const cellValue = ws.getRow(2).getCell(2).value;
      expect(cellValue).toEqual({ formula: "TestTable[[#This Row],[A]]" });
    });

    it("writes CONCAT([@A]) without leading @", async () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();

      worksheet.addTable({
        name: "table",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "CONCAT([@A])" }]]
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const zipData = await extractAll(new Uint8Array(buffer));

      const sheet1 = zipData.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();

      const xml = new TextDecoder().decode(sheet1!.data);
      const match = xml.match(/<c[^>]*\br="B2"[^>]*>[\s\S]*?<f[^>]*>([^<]*)<\/f>/);
      expect(match).toBeTruthy();

      const formula = match![1];
      expect(formula).toBe("CONCAT([@A])");
      expect(formula.startsWith("@")).toBe(false);
    });

    it("loads a table with calculatedColumnFormula without crashing (issue #76)", async () => {
      // Create a workbook with a 3-column table
      const wb1 = new Workbook();
      const ws1: any = wb1.addWorksheet("Data");

      ws1.addTable({
        name: "CalcTable",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [
          { name: "Value", filterButton: true },
          { name: "Double", filterButton: true },
          { name: "Label", filterButton: true }
        ],
        rows: [
          [10, 20, "a"],
          [30, 60, "b"]
        ]
      });

      const buffer = await wb1.xlsx.writeBuffer();

      // Manually inject a <calculatedColumnFormula> child element into the table XML
      // to simulate what Excel produces for calculated columns.
      const zipData = await extractAll(new Uint8Array(buffer));
      const tablePath = [...zipData.keys()].find(k => k.startsWith("xl/tables/"));
      expect(tablePath).toBeDefined();

      const tableXml = new TextDecoder().decode(zipData.get(tablePath!)!.data);
      // Replace the self-closing <tableColumn> for "Double" with one that has a child element
      const modifiedXml = tableXml.replace(
        /(<tableColumn id="2" name="Double"[^/]*)\/>/,
        "$1><calculatedColumnFormula>[Value]*2</calculatedColumnFormula></tableColumn>"
      );
      expect(modifiedXml).toContain("<calculatedColumnFormula>");

      zipData.get(tablePath!)!.data = new TextEncoder().encode(modifiedXml);

      // Re-pack into a ZIP buffer
      const { createZip } = await import("@archive/zip/zip-bytes");
      const entries = [...zipData.entries()].map(([name, file]) => ({
        name,
        data: file.data
      }));
      const modifiedBuffer = await createZip(entries);

      // This should NOT throw: "Cannot read properties of undefined (reading 'style')"
      const wb2 = new Workbook();
      await wb2.xlsx.load(modifiedBuffer);

      const ws2: any = wb2.getWorksheet("Data");
      expect(ws2).toBeDefined();

      const table2 = ws2.getTable("CalcTable");
      expect(table2).toBeDefined();
      expect(table2.table.columns).toHaveLength(3);
      expect(table2.table.columns[0].name).toBe("Value");
      expect(table2.table.columns[1].name).toBe("Double");
      expect(table2.table.columns[1].calculatedColumnFormula).toBe("[Value]*2");
      expect(table2.table.columns[2].name).toBe("Label");

      // Round-trip: write and reload — calculatedColumnFormula should survive
      const buffer2 = await wb2.xlsx.writeBuffer();
      const wb3 = new Workbook();
      await wb3.xlsx.load(buffer2);
      const ws3: any = wb3.getWorksheet("Data");
      const table3 = ws3.getTable("CalcTable");
      expect(table3.table.columns[1].calculatedColumnFormula).toBe("[Value]*2");
    });
  });

  // ========================================================================
  // sanitizeTableName unit tests
  // ========================================================================
  describe("sanitizeTableName", () => {
    it("returns valid names unchanged", () => {
      expect(sanitizeTableName("TestTable")).toBe("TestTable");
      expect(sanitizeTableName("_private")).toBe("_private");
      expect(sanitizeTableName("Table1")).toBe("Table1");
      expect(sanitizeTableName("my.table")).toBe("my.table");
    });

    it("replaces spaces with underscores", () => {
      expect(sanitizeTableName("test table")).toBe("test_table");
      expect(sanitizeTableName("my  table  name")).toBe("my__table__name");
    });

    it("replaces all whitespace characters (tab, newline) with underscores", () => {
      expect(sanitizeTableName("test\ttable")).toBe("test_table");
      expect(sanitizeTableName("test\ntable")).toBe("test_table");
      expect(sanitizeTableName("test\r\ntable")).toBe("test__table");
    });

    it("strips invalid characters", () => {
      expect(sanitizeTableName("test@table!")).toBe("testtable");
      expect(sanitizeTableName("table#1$2%3")).toBe("table123");
      expect(sanitizeTableName("hello-world")).toBe("helloworld");
    });

    it("prefixes with underscore when first char is digit", () => {
      expect(sanitizeTableName("1Table")).toBe("_1Table");
      expect(sanitizeTableName("123")).toBe("_123");
    });

    it("prefixes with underscore when first char is period", () => {
      expect(sanitizeTableName(".table")).toBe("_.table");
    });

    it("returns _Table for empty string or all-invalid characters", () => {
      expect(sanitizeTableName("")).toBe("_Table");
      expect(sanitizeTableName("@#$%^&")).toBe("_Table");
    });

    it("converts all-spaces to underscores (not empty)", () => {
      expect(sanitizeTableName("   ")).toBe("___");
    });

    it("avoids names that look like A1-style cell references", () => {
      expect(sanitizeTableName("A1")).toBe("_A1");
      expect(sanitizeTableName("XFD1048576")).toBe("_XFD1048576");
      expect(sanitizeTableName("Z99")).toBe("_Z99");
    });

    it("avoids names that look like R1C1-style cell references", () => {
      expect(sanitizeTableName("R1C1")).toBe("_R1C1");
      expect(sanitizeTableName("R100C200")).toBe("_R100C200");
    });

    it("does not reject bare RC as a cell reference", () => {
      // "RC" is not a valid R1C1 reference (needs digits), so it's a valid name
      expect(sanitizeTableName("RC")).toBe("RC");
      expect(sanitizeTableName("Rc")).toBe("Rc");
    });

    it("prefixes reserved single-character names (C, c, R, r)", () => {
      expect(sanitizeTableName("C")).toBe("_C");
      expect(sanitizeTableName("c")).toBe("_c");
      expect(sanitizeTableName("R")).toBe("_R");
      expect(sanitizeTableName("r")).toBe("_r");
    });

    it("allows other single-character names", () => {
      expect(sanitizeTableName("A")).toBe("A");
      expect(sanitizeTableName("Z")).toBe("Z");
      expect(sanitizeTableName("_")).toBe("_");
    });

    it("preserves leading backslash (valid only as first char)", () => {
      expect(sanitizeTableName("\\name")).toBe("\\name");
      expect(sanitizeTableName("\\")).toBe("\\");
    });

    it("strips backslash in non-first positions", () => {
      expect(sanitizeTableName("name\\value")).toBe("namevalue");
      expect(sanitizeTableName("a\\b\\c")).toBe("abc");
    });

    it("preserves Unicode letters (CJK, etc.)", () => {
      expect(sanitizeTableName("销售数据")).toBe("销售数据");
      expect(sanitizeTableName("テーブル1")).toBe("テーブル1");
      expect(sanitizeTableName("Données")).toBe("Données");
      expect(sanitizeTableName("表格 测试")).toBe("表格_测试");
    });

    it("truncates to 255 characters", () => {
      const long = "A".repeat(300);
      expect(sanitizeTableName(long)).toHaveLength(255);
    });

    it("handles the exact reproduction case from issue #91", () => {
      expect(sanitizeTableName("test table")).toBe("test_table");
    });
  });

  // ========================================================================
  // Issue #91: table names with spaces should be auto-sanitized
  // ========================================================================
  describe("table name sanitization integration", () => {
    it("sanitizes table name with spaces on addTable", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      const table = ws.addTable({
        name: "test table",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [
          { name: "A", filterButton: true },
          { name: "B", filterButton: true },
          { name: "C", filterButton: true }
        ],
        rows: [
          ["test", 2, "a4f"],
          ["test 2", 1, "a4f"],
          ["test 3", 6, "a4f"]
        ]
      });

      // Name should have been sanitized
      expect(table.name).toBe("test_table");
      expect(table.displayName).toBe("test_table");

      // Should be retrievable by sanitized name
      expect(ws.getTable("test_table")).toBe(table);
    });

    it("issue #91 reproduction: writeBuffer produces valid OOXML", async () => {
      const columns = ["A", "B", "C"];
      const data = [
        ["test", 2, "a4f"],
        ["test 2", 1, "a4f"],
        ["test 3", 6, "a4f"]
      ];

      const workbook = new Workbook();
      const sheet = workbook.addWorksheet("test");
      sheet.addTable({
        columns: columns.map(i => ({ name: i, filterButton: true })),
        headerRow: true,
        name: "test table",
        ref: "A1",
        rows: data,
        totalsRow: false
      });

      const buffer = await workbook.xlsx.writeBuffer();

      // Verify the table XML has sanitized name
      const entries = await extractAll(new Uint8Array(buffer));
      const tableEntry = entries.get("xl/tables/table1.xml");
      expect(tableEntry).toBeDefined();
      const tableXml = new TextDecoder().decode(tableEntry!.data);

      // Must NOT contain spaces in name/displayName
      expect(tableXml).toContain('name="test_table"');
      expect(tableXml).toContain('displayName="test_table"');
      expect(tableXml).not.toContain('name="test table"');
      expect(tableXml).not.toContain('displayName="test table"');
    });

    it("sanitizes displayName independently when provided", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      const table = ws.addTable({
        name: "my table",
        displayName: "My Display Name",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Col1" }],
        rows: [["val"]]
      });

      expect(table.name).toBe("my_table");
      expect(table.displayName).toBe("My_Display_Name");
    });

    it("sanitizes name when set via setter", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      const table = ws.addTable({
        name: "ValidName",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Col1" }],
        rows: [["val"]]
      });

      table.name = "new name with spaces";
      expect(table.name).toBe("new_name_with_spaces");

      table.displayName = "another display name";
      expect(table.displayName).toBe("another_display_name");
    });

    it("sanitized name survives round-trip (write + load)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      ws.addTable({
        name: "my table",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "X", filterButton: true }],
        rows: [["val1"], ["val2"]]
      });

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("test")!;
      const table2 = (ws2 as any).getTable("my_table");
      expect(table2).toBeDefined();
      expect(table2.name).toBe("my_table");
    });

    it("totalsRow formulas use sanitized table name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");
      ws.addTable({
        name: "sales data",
        ref: "A1",
        headerRow: true,
        totalsRow: true,
        columns: [
          { name: "Category", totalsRowLabel: "Total" },
          { name: "Amount", totalsRowFunction: "sum" }
        ],
        rows: [
          ["A", 10],
          ["B", 20]
        ]
      });

      // The totals row formula cell should reference the sanitized name
      // Row 1 = header, rows 2-3 = data, row 4 = totals
      const totalsCell = ws.getCell("B4");
      expect(totalsCell.value).toBeDefined();
      const formula = (totalsCell.value as any).formula;
      expect(formula).toContain("sales_data");
      expect(formula).not.toContain("sales data");
    });
  });

  // ===========================================================================
  // Table name lifecycle: rename and delete must keep worksheet.tables and
  // workbook._tableNames in sync, otherwise getTable / duplicate-name checks
  // and cross-sheet name reuse silently break.
  // ===========================================================================

  describe("table name lifecycle", () => {
    it("renaming a table updates worksheet.tables and releases the old name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const table = ws.addTable({
        name: "Original",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      table.name = "Renamed";

      expect(table.name).toBe("Renamed");
      expect(ws.getTable("Renamed")).toBe(table);
      expect(ws.getTable("Original")).toBeUndefined();

      // The old name must be free for reuse on another sheet.
      const ws2 = wb.addWorksheet("Sheet2");
      expect(() =>
        ws2.addTable({
          name: "Original",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("renaming to an existing workbook-wide name throws", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addTable({
        name: "Existing",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });
      const t2 = ws.addTable({
        name: "Other",
        ref: "C1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      expect(() => {
        t2.name = "Existing";
      }).toThrow(/already exists/i);
      // After the failed rename the table must still be reachable by its old name.
      expect(t2.name).toBe("Other");
      expect(ws.getTable("Other")).toBe(t2);
    });

    it("removeTable releases the workbook-wide name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addTable({
        name: "Temp",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });
      ws.removeTable("Temp");

      const ws2 = wb.addWorksheet("Sheet2");
      expect(() =>
        ws2.addTable({
          name: "Temp",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("removing a worksheet releases all of its table names", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      ws1.addTable({
        name: "Releasable",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      wb.removeWorksheet(ws1.id);

      // The name must be reusable on a brand-new sheet.
      const ws2 = wb.addWorksheet("Sheet2");
      expect(() =>
        ws2.addTable({
          name: "Releasable",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("setting column.style triggers commit() to propagate to cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const table = ws.addTable({
        name: "Styled",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Col1" }, { name: "Col2" }],
        rows: [
          ["a", "b"],
          ["c", "d"]
        ]
      });

      const col = table.getColumn(1);
      col.style = { font: { bold: true } };

      // Without cacheState() being called, commit() returns early and the
      // newly-applied style never reaches the data cells.
      table.commit();

      // Data cells in the second column (B2..B3) should now carry the bold font.
      expect((ws.getCell("B2").style as any).font?.bold).toBe(true);
      expect((ws.getCell("B3").style as any).font?.bold).toBe(true);
    });
  });
});
