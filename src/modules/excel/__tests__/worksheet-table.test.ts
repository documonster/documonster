import { extractAll } from "@archive/unzip/extract";
import { cellGetValue } from "@excel/core/cell";
import {
  sanitizeTableName,
  tableAddColumn,
  tableAddRow,
  tableColumnSetName,
  tableColumnSetStyle,
  tableCommit,
  tableDisplayName,
  tableGetColumn,
  tableName,
  tableRemoveColumns,
  tableRemoveRows,
  tableSetDisplayName,
  tableSetHeaderRow,
  tableSetName,
  tableSetRef,
  tableSetTotalsRow
} from "@excel/core/table";
import { getCell, getTable, removeTable, rowGetCell } from "@excel/core/worksheet";
import { Cell, Table, Workbook, Worksheet } from "@excel/index";
import { colCache } from "@excel/utils/col-cache";
import { describe, it, expect } from "vitest";

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

function buildTable(ref: string, ws: any) {
  return Table.add(ws, {
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
    const row = nRow >= 1 && Worksheet.getRow(ws, nRow);
    if (!row) {
      continue;
    }
    for (let j = -1; j <= testValues[0].length + 1; j++) {
      const value = (vRow && vRow[j]) || null;
      const nCol = j + a.col;
      const cellValue = nCol >= 1 && cellGetValue(rowGetCell(row, nCol));
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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      buildTable("A1", ws);

      checkTable("A1", ws, values);
    });

    it("removes header", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableSetHeaderRow(table, false);
      tableCommit(table);

      const newValues = spliceArray(values, 0, 1);
      checkTable("A1", ws, newValues);
    });

    it("removes totals", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableSetTotalsRow(table, false);
      tableCommit(table);

      const newValues = spliceArray(values, 5, 1);
      checkTable("A1", ws, newValues);
    });

    it("moves the table", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableSetRef(table, "C2");
      tableCommit(table);

      checkTable("C2", ws, values);
    });

    it("removes a row", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableRemoveRows(table, 1);
      tableCommit(table);

      const newValues = spliceArray(values, 2, 1);
      checkTable("A1", ws, newValues);
    });

    it("adds a row", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableAddRow(table, [new Date("2019-08-05"), 5, "Bird"]);
      tableCommit(table);

      const newValues = spliceArray(values, 5, 0, [new Date("2019-08-05"), 5, "Bird"]);
      checkTable("A1", ws, newValues);
    });

    it("removes a column", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableRemoveColumns(table, 1);
      tableCommit(table);

      const newValues = values.map(rVals => spliceArray(rVals, 1, 1));
      checkTable("A1", ws, newValues);
    });

    it("adds a column", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      tableAddColumn(
        table,
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
      tableCommit(table);

      const colValues = ["Letter", "a", "b", "c", "d", { formula: "ROW()", result: 6 }];
      const newValues = values.map((rVals, i) => spliceArray(rVals, 2, 0, colValues[i]));
      checkTable("A1", ws, newValues);
    });

    it("renames a column", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");
      const table = buildTable("A1", ws);

      const column = tableGetColumn(table, 1);
      tableColumnSetName(column, "Code");
      tableCommit(table);

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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Table.add(ws, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "[@A]" }]]
      });

      const cellValue = cellGetValue(rowGetCell(Worksheet.getRow(ws, 2), 2));
      expect(cellValue).toEqual({ formula: "[@A]" });
    });

    it("qualifies implicit structured references when enabled", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Table.add(ws, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        qualifyImplicitStructuredReferences: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "[@A]" }]]
      });

      const cellValue = cellGetValue(rowGetCell(Worksheet.getRow(ws, 2), 2));
      expect(cellValue).toEqual({ formula: "TestTable[[#This Row],[A]]" });
    });

    it("writes CONCAT([@A]) without leading @", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      Table.add(worksheet, {
        name: "table",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }],
        rows: [["a1", { formula: "CONCAT([@A])" }]]
      });

      const buffer = await Workbook.toBuffer(workbook);
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
      const wb1 = Workbook.create();
      const ws1: any = Workbook.addWorksheet(wb1, "Data");

      Table.add(ws1, {
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

      const buffer = await Workbook.toBuffer(wb1);

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
      const wb2 = Workbook.create();
      await Workbook.read(wb2, modifiedBuffer);

      const ws2: any = Workbook.getWorksheet(wb2, "Data")!;
      expect(ws2).toBeDefined();

      const table2 = Table.get(ws2, "CalcTable");
      expect(table2).toBeDefined();
      expect(table2.table.columns).toHaveLength(3);
      expect(table2.table.columns[0].name).toBe("Value");
      expect(table2.table.columns[1].name).toBe("Double");
      expect(table2.table.columns[1].calculatedColumnFormula).toBe("[Value]*2");
      expect(table2.table.columns[2].name).toBe("Label");

      // Round-trip: write and reload — calculatedColumnFormula should survive
      const buffer2 = await Workbook.toBuffer(wb2);
      const wb3 = Workbook.create();
      await Workbook.read(wb3, buffer2);
      const ws3: any = Workbook.getWorksheet(wb3, "Data")!;
      const table3 = Table.get(ws3, "CalcTable");
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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");
      const table = Table.add(ws, {
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
      expect(tableName(table)).toBe("test_table");
      expect(tableDisplayName(table)).toBe("test_table");

      // Should be retrievable by sanitized name
      expect(getTable(ws, "test_table")).toBe(table);
    });

    it("issue #91 reproduction: writeBuffer produces valid OOXML", async () => {
      const columns = ["A", "B", "C"];
      const data = [
        ["test", 2, "a4f"],
        ["test 2", 1, "a4f"],
        ["test 3", 6, "a4f"]
      ];

      const workbook = Workbook.create();
      const sheet = Workbook.addWorksheet(workbook, "test");
      Table.add(sheet, {
        columns: columns.map(i => ({ name: i, filterButton: true })),
        headerRow: true,
        name: "test table",
        ref: "A1",
        rows: data,
        totalsRow: false
      });

      const buffer = await Workbook.toBuffer(workbook);

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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");
      const table = Table.add(ws, {
        name: "my table",
        displayName: "My Display Name",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Col1" }],
        rows: [["val"]]
      });

      expect(tableName(table)).toBe("my_table");
      expect(tableDisplayName(table)).toBe("My_Display_Name");
    });

    it("sanitizes name when set via setter", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");
      const table = Table.add(ws, {
        name: "ValidName",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Col1" }],
        rows: [["val"]]
      });

      tableSetName(table, "new name with spaces");
      expect(tableName(table)).toBe("new_name_with_spaces");

      tableSetDisplayName(table, "another display name");
      expect(tableDisplayName(table)).toBe("another_display_name");
    });

    it("sanitized name survives round-trip (write + load)", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");
      Table.add(ws, {
        name: "my table",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "X", filterButton: true }],
        rows: [["val1"], ["val2"]]
      });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "test")!;
      const table2 = Table.get(ws2, "my_table");
      expect(table2).toBeDefined();
      expect(tableName(table2)).toBe("my_table");
    });

    it("totalsRow formulas use sanitized table name", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");
      Table.add(ws, {
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
      const totalsCell = getCell(ws, "B4");
      expect(cellGetValue(totalsCell)).toBeDefined();
      const formula = (cellGetValue(totalsCell) as any).formula;
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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      const table = Table.add(ws, {
        name: "Original",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      tableSetName(table, "Renamed");

      expect(tableName(table)).toBe("Renamed");
      expect(getTable(ws, "Renamed")).toBe(table);
      expect(getTable(ws, "Original")).toBeUndefined();

      // The old name must be free for reuse on another sheet.
      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      expect(() =>
        Table.add(ws2, {
          name: "Original",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("renaming to an existing workbook-wide name throws", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      Table.add(ws, {
        name: "Existing",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });
      const t2 = Table.add(ws, {
        name: "Other",
        ref: "C1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      expect(() => {
        tableSetName(t2, "Existing");
      }).toThrow(/already exists/i);
      // After the failed rename the table must still be reachable by its old name.
      expect(tableName(t2)).toBe("Other");
      expect(getTable(ws, "Other")).toBe(t2);
    });

    it("removeTable releases the workbook-wide name", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      Table.add(ws, {
        name: "Temp",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });
      removeTable(ws, "Temp");

      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      expect(() =>
        Table.add(ws2, {
          name: "Temp",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("removing a worksheet releases all of its table names", () => {
      const wb = Workbook.create();
      const ws1 = Workbook.addWorksheet(wb, "Sheet1");
      Table.add(ws1, {
        name: "Releasable",
        ref: "A1",
        columns: [{ name: "Col1" }],
        rows: [["v1"]]
      });

      Workbook.removeWorksheet(wb, ws1.id);

      // The name must be reusable on a brand-new sheet.
      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      expect(() =>
        Table.add(ws2, {
          name: "Releasable",
          ref: "A1",
          columns: [{ name: "Col1" }],
          rows: [["v1"]]
        })
      ).not.toThrow();
    });

    it("setting column.style triggers commit() to propagate to cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      const table = Table.add(ws, {
        name: "Styled",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Col1" }, { name: "Col2" }],
        rows: [
          ["a", "b"],
          ["c", "d"]
        ]
      });

      const col = tableGetColumn(table, 1);
      tableColumnSetStyle(col, { font: { bold: true } });

      // Without cacheState() being called, commit() returns early and the
      // newly-applied style never reaches the data cells.
      tableCommit(table);

      // Data cells in the second column (B2..B3) should now carry the bold font.
      expect((Cell.getStyle(ws, "B2") as any).font?.bold).toBe(true);
      expect((Cell.getStyle(ws, "B3") as any).font?.bold).toBe(true);
    });
  });
});
