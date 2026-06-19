import fs from "fs";
import { promisify } from "util";

import type { WorkbookData } from "@excel/core/workbook-core";
import { Column, Workbook, Worksheet } from "@excel/index";
import { describe, it, expect } from "vitest";

const fsReadFileAsync = promisify(fs.readFile);
import { ZipParser } from "@archive/unzip/zip-parser";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";

const PIVOT_TABLE_FILEPATHS = [
  "xl/pivotCache/pivotCacheRecords1.xml",
  "xl/pivotCache/pivotCacheDefinition1.xml",
  "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
  "xl/pivotTables/pivotTable1.xml",
  "xl/pivotTables/_rels/pivotTable1.xml.rels"
];

import { addPivotTable, addTable } from "@excel/core/worksheet";
import { testFilePath } from "@test/utils";

const TEST_XLSX_FILEPATH = testFilePath("workbook-pivot.test");
const TEST_XLSX_TABLE_FILEPATH = testFilePath("workbook-pivot-table.test");

// ---------------------------------------------------------------------------
// Helper: write workbook to buffer, parse zip, return decoded entries
// ---------------------------------------------------------------------------
type ZipEntries = Record<string, Uint8Array>;

async function writeThenParseZip(workbook: WorkbookData): Promise<ZipEntries>;
async function writeThenParseZip(workbook: WorkbookData, filePath: string): Promise<ZipEntries>;
async function writeThenParseZip(workbook: WorkbookData, filePath?: string): Promise<ZipEntries> {
  if (filePath) {
    await Workbook.writeFile(workbook, filePath);
    const buffer = await fsReadFileAsync(filePath);
    await expectValidXlsx(buffer, { label: `writeThenParseZip ${filePath}` });
    return new ZipParser(buffer).extractAllSync();
  }
  const buffer = await Workbook.toBuffer(workbook);
  await expectValidXlsx(buffer, { label: "writeThenParseZip buffer" });
  return new ZipParser(buffer as Buffer).extractAllSync();
}

function decodeXml(zipData: ZipEntries, path: string): string {
  return new TextDecoder().decode(zipData[path]);
}

const TEST_DATA = [
  ["A", "B", "C", "D", "E"],
  ["a1", "b1", "c1", 4, 5],
  ["a1", "b2", "c1", 4, 5],
  ["a2", "b1", "c2", 14, 24],
  ["a2", "b2", "c2", 24, 35],
  ["a3", "b1", "c3", 34, 45],
  ["a3", "b2", "c3", 44, 45]
];

// =============================================================================
// Tests

describe("Workbook", () => {
  describe("Pivot Tables", () => {
    it("if pivot table added with sourceSheet, then certain xml and rels files are added", async () => {
      const workbook = Workbook.create();

      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["E"],
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeDefined();
      }
    });

    it("if pivot table added with sourceTable, then certain xml and rels files are added", async () => {
      const workbook = Workbook.create();

      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      // Create a table with the same data structure as TEST_DATA
      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
        rows: [
          ["a1", "b1", "c1", 4, 5],
          ["a1", "b2", "c1", 4, 5],
          ["a2", "b1", "c2", 14, 24],
          ["a2", "b2", "c2", 24, 35],
          ["a3", "b1", "c3", 34, 45],
          ["a3", "b2", "c3", 44, 45]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["E"],
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_TABLE_FILEPATH);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeDefined();
      }
    });

    it("if pivot table NOT added, then certain xml and rels files are not added", async () => {
      const workbook = Workbook.create();

      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      Workbook.addWorksheet(workbook, "Sheet2");

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeUndefined();
      }
    });

    it("throws error if neither sourceSheet nor sourceTable is provided", () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      expect(() => {
        addPivotTable(worksheet, {
          rows: ["A"],
          columns: ["B"],
          values: ["C"],
          metric: "sum"
        } as any);
      }).toThrow("Either sourceSheet or sourceTable must be provided.");
    });

    it("throws error if both sourceSheet and sourceTable are provided", () => {
      const workbook = Workbook.create();

      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const table = addTable(worksheet1, {
        name: "TestTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [["a1", "b1", "c1"]]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          sourceTable: table,
          rows: ["A"],
          columns: ["B"],
          values: ["C"],
          metric: "sum"
        });
      }).toThrow("Cannot specify both sourceSheet and sourceTable. Choose one.");
    });

    it("throws error if header name not found in sourceTable", () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [["a1", "b1", "c1"]]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: ["NonExistent"],
          values: ["C"],
          metric: "sum"
        });
      }).toThrow('The header name "NonExistent" was not found in Sheet1.');
    });

    it("throws error if sourceTable has no data rows", () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      const table = addTable(worksheet, {
        name: "EmptyTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [] // empty rows
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: ["B"],
          values: ["C"],
          metric: "sum"
        });
      }).toThrow("Cannot create pivot table from an empty table. Add data rows to the table.");
    });

    it("throws error if sourceTable has duplicate column names", () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      const table = addTable(worksheet, {
        name: "DuplicateColumnsTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "A" }], // duplicate 'A'
        rows: [["a1", "b1", "a2"]]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: ["B"],
          values: ["A"],
          metric: "sum"
        });
      }).toThrow(
        'Duplicate column name "A" found in table. Pivot tables require unique column names.'
      );
    });

    it("works with sourceTable not starting at A1", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      // Table starting at C5 instead of A1
      const table = addTable(worksheet, {
        name: "OffsetTable",
        ref: "C5",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
        rows: [
          ["a1", "b1", "c1", 4, 5],
          ["a1", "b2", "c1", 4, 5],
          ["a2", "b1", "c2", 14, 24]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["A"],
        columns: ["B"],
        values: ["E"],
        metric: "sum"
      });

      const offsetFilePath = testFilePath("workbook-pivot-offset.test");
      const zipData = await writeThenParseZip(workbook, offsetFilePath);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeDefined();
      }
    });

    it("supports multiple values when columns is empty", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

      const table = addTable(worksheet, {
        name: "MultiValuesTable",
        ref: "A1",
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
        rows: [
          ["a1", "b1", "c1", 4, 5],
          ["a1", "b2", "c1", 4, 5],
          ["a2", "b1", "c2", 14, 24],
          ["a2", "b2", "c2", 24, 35]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["A", "B"],
        columns: [], // Empty columns - allows multiple values
        values: ["D", "E"], // Multiple values
        metric: "sum"
      });

      const multiValuesFilePath = testFilePath("workbook-pivot-multi-values.test");
      const zipData = await writeThenParseZip(workbook, multiValuesFilePath);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeDefined();
      }
    });

    it("supports empty columns with single value", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: [], // Empty columns
        values: ["E"],
        metric: "sum"
      });

      const emptyColsFilePath = testFilePath("workbook-pivot-empty-cols.test");
      const zipData = await writeThenParseZip(workbook, emptyColsFilePath);
      for (const filepath of PIVOT_TABLE_FILEPATHS) {
        expect(zipData[filepath]).toBeDefined();
      }
    });

    it("supports multiple values with non-empty columns (field x=-2 appended)", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A"],
        columns: ["B"], // Non-empty columns
        values: ["D", "E"], // Multiple values
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook);
      const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      // colFields should have B's field index + the -2 sentinel
      expect(pivotXml).toContain('colFields count="2"');
      expect(pivotXml).toContain('field x="-2"');
      // dataFields should have both values
      expect(pivotXml).toContain('dataFields count="2"');
      expect(pivotXml).toContain("Sum of D");
      expect(pivotXml).toContain("Sum of E");
    });

    it("throws error if no values specified", () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: [], // No values
          metric: "sum"
        });
      }).toThrow("Must have at least one value.");
    });

    // R8-T2: Additional validate() branch coverage

    it("throws error if no rows specified", () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: [],
          columns: ["B"],
          values: ["D"],
          metric: "sum"
        });
      }).toThrow("No pivot table rows specified.");
    });

    it("throws error for empty header name in source sheet", () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      // Header row with an empty string header
      Worksheet.addRows(worksheet1, [
        ["A", "", "C"],
        ["a1", "b1", "c1"]
      ]);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["C"],
          metric: "sum"
        });
      }).toThrow(/Empty or missing header name at column 2/);
    });

    it("throws error for whitespace-only header name in source sheet", () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, [
        ["A", "   ", "C"],
        ["a1", "b1", "c1"]
      ]);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["C"],
          metric: "sum"
        });
      }).toThrow(/Empty or missing header name at column 2/);
    });

    it("throws error for duplicate value field names", () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      expect(() => {
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["D", "D"],
          metric: "sum"
        });
      }).toThrow('Duplicate value field "D". Each value field name must be unique.');
    });

    it("supports applyWidthHeightFormats option to preserve column widths", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      // Set custom column widths before creating pivot table
      Column.setWidth(worksheet2, 1, 30);
      Column.setWidth(worksheet2, 2, 15);

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["D"],
        metric: "sum",
        applyWidthHeightFormats: "0" // Preserve worksheet column widths
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      expect(pivotTableXml).toContain('applyWidthHeightFormats="0"');
    });

    it("defaults applyWidthHeightFormats to 1 when not specified", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["D"],
        metric: "sum"
        // applyWidthHeightFormats not specified, should default to "1"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      expect(pivotTableXml).toContain('applyWidthHeightFormats="1"');
    });

    it("supports omitting columns (Excel uses 'Values' as column field)", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      // Create pivot table without specifying columns
      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        // columns is omitted - should default to []
        values: ["D"],
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);

      // Verify pivot table XML exists
      expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
    });

    it("handles XML special characters in pivot table data", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");

      // Data with XML special characters: &, <, >, ", '
      // Use special characters in both row fields AND value field names
      Worksheet.addRows(worksheet1, [
        ["Company", "Product", "Sales & Revenue"],
        ["Johnson & Johnson", "Drug A", 1000],
        ["BioTech <Special>", "Drug B", 1500],
        ['PharmaCorp "Elite"', "Drug C", 1200],
        ["Gene's Labs", "Drug D", 1800]
      ]);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["Company"],
        columns: ["Product"],
        values: ["Sales & Revenue"], // Value field name contains &
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);

      // Verify pivot cache definition contains properly escaped XML
      const cacheDefinition = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");

      // Check that XML special characters are escaped in sharedItems
      expect(cacheDefinition).toContain("Johnson &amp; Johnson");
      expect(cacheDefinition).toContain("BioTech &lt;Special&gt;");
      expect(cacheDefinition).toContain("PharmaCorp &quot;Elite&quot;");
      expect(cacheDefinition).toContain("Gene&apos;s Labs");

      // Verify the XML is valid (no unescaped special chars)
      expect(cacheDefinition).not.toContain('v="Johnson & Johnson"');
      expect(cacheDefinition).not.toContain('v="BioTech <Special>"');

      // Verify pivot table definition has escaped dataField name
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");
      expect(pivotTableXml).toContain("Sum of Sales &amp; Revenue");
      expect(pivotTableXml).not.toContain("Sum of Sales & Revenue");
    });

    it("handles null and undefined values in pivot table data", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");

      // Data with null/undefined values
      Worksheet.addRows(worksheet1, [
        ["Region", "Territory", "Amount"],
        ["North", "NE", 1000],
        ["South", null, 1500], // null territory
        ["East", undefined, 2000], // undefined territory
        ["West", "NW", 2200]
      ]);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["Region", "Territory"], // Territory has null/undefined values
        values: ["Amount"],
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);

      // Null values in row fields are added to sharedItems as <m /> and referenced via <x v="..."/>
      const cacheDef = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");
      expect(cacheDef).toContain("<m />");
      expect(cacheDef).toContain('containsBlank="1"');

      // Cache records use index references for row fields (no inline <m /> in records)
      const cacheRecords = decodeXml(zipData, "xl/pivotCache/pivotCacheRecords1.xml");
      expect(cacheRecords).toContain('<x v="');

      // Verify pivot table XML exists
      expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
    });

    it("supports multiple pivot tables from same source data", async () => {
      const workbook = Workbook.create();

      // Create source data with multiple dimensions
      const sourceSheet = Workbook.addWorksheet(workbook, "Sales Data");
      Worksheet.addRows(sourceSheet, [
        ["Region", "Product", "Salesperson", "Quarter", "Revenue", "Units"],
        ["North", "Widget A", "Alice", "Q1", 10000, 100],
        ["South", "Widget B", "Bob", "Q1", 15000, 150],
        ["North", "Widget A", "Alice", "Q2", 12000, 120],
        ["South", "Widget B", "Bob", "Q2", 18000, 180],
        ["East", "Widget C", "Charlie", "Q1", 20000, 200],
        ["West", "Widget C", "Diana", "Q2", 22000, 220]
      ]);

      // First pivot table: Revenue by Region and Product
      const pivot1Sheet = Workbook.addWorksheet(workbook, "Pivot 1 - Region x Product");
      addPivotTable(pivot1Sheet, {
        sourceSheet,
        rows: ["Region", "Product"],
        columns: ["Quarter"],
        values: ["Revenue"],
        metric: "sum"
      });

      // Second pivot table: Units by Salesperson (completely different fields)
      const pivot2Sheet = Workbook.addWorksheet(workbook, "Pivot 2 - Salesperson");
      addPivotTable(pivot2Sheet, {
        sourceSheet,
        rows: ["Salesperson"],
        columns: ["Quarter"],
        values: ["Units"],
        metric: "sum"
      });

      // Third pivot table: Another different configuration
      const pivot3Sheet = Workbook.addWorksheet(workbook, "Pivot 3 - Product x Region");
      addPivotTable(pivot3Sheet, {
        sourceSheet,
        rows: ["Product"],
        columns: ["Region"],
        values: ["Revenue"],
        metric: "sum"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);

      // Verify all three pivot tables exist
      expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
      expect(zipData["xl/pivotTables/pivotTable2.xml"]).toBeDefined();
      expect(zipData["xl/pivotTables/pivotTable3.xml"]).toBeDefined();

      // Verify all three pivot cache definitions exist
      expect(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]).toBeDefined();
      expect(zipData["xl/pivotCache/pivotCacheDefinition2.xml"]).toBeDefined();
      expect(zipData["xl/pivotCache/pivotCacheDefinition3.xml"]).toBeDefined();

      // Verify each pivot table has unique cacheId
      const pivotTable1Xml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");
      const pivotTable2Xml = decodeXml(zipData, "xl/pivotTables/pivotTable2.xml");
      const pivotTable3Xml = decodeXml(zipData, "xl/pivotTables/pivotTable3.xml");

      expect(pivotTable1Xml).toContain('cacheId="10"');
      expect(pivotTable2Xml).toContain('cacheId="11"');
      expect(pivotTable3Xml).toContain('cacheId="12"');
    });

    it("supports 'count' metric for pivot tables", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["D"],
        metric: "count"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      // dataField should have name="Count of D" and subtotal="count"
      expect(pivotTableXml).toContain("Count of D");
      expect(pivotTableXml).toContain('subtotal="count"');
      expect(pivotTableXml).not.toContain("Sum of");
    });

    it("defaults to 'sum' metric when not specified", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["D"]
        // metric not specified - should default to 'sum'
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      // dataField should have name="Sum of D" and no subtotal attribute
      expect(pivotTableXml).toContain("Sum of D");
      expect(pivotTableXml).not.toContain('subtotal="count"');
    });

    it("supports 'average' metric for pivot tables", async () => {
      const workbook = Workbook.create();
      const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
      Worksheet.addRows(worksheet1, TEST_DATA);

      const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

      addPivotTable(worksheet2, {
        sourceSheet: worksheet1,
        rows: ["A", "B"],
        columns: ["C"],
        values: ["D"],
        metric: "average"
      });

      const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
      const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

      expect(pivotTableXml).toContain("Average of D");
      expect(pivotTableXml).toContain('subtotal="average"');
      expect(pivotTableXml).not.toContain("Sum of");
    });

    // ==========================================================================
    // Pivot Table Read and Preserve Tests
    // ==========================================================================

    describe("Pivot Table Preservation (Load/Save)", () => {
      const ROUNDTRIP_FILEPATH = testFilePath("workbook-pivot-roundtrip.test");

      it("preserves pivot table through load/save cycle", async () => {
        // Step 1: Create workbook with pivot table
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A", "B"],
          columns: ["C"],
          values: ["E"],
          metric: "sum"
        });

        // Step 2: Save to file
        await Workbook.writeFile(workbook, ROUNDTRIP_FILEPATH);

        // Step 3: Read the file back
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, ROUNDTRIP_FILEPATH);

        // Step 4: Check that loaded workbook has pivot tables
        expect(loadedWorkbook.pivotTables.length).toBe(1);
        const loadedPivot = loadedWorkbook.pivotTables[0];
        expect(loadedPivot).toBeDefined();
        expect(loadedPivot.isLoaded).toBe(true);
        expect(loadedPivot.tableNumber).toBe(1);

        // Step 5: Save again and verify
        const ROUNDTRIP_FILEPATH2 = testFilePath("workbook-pivot-roundtrip2.test");
        const zipData = await writeThenParseZip(loadedWorkbook, ROUNDTRIP_FILEPATH2);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zipData[filepath]).toBeDefined();
        }
      });

      it("preserves multiple pivot tables through load/save cycle", async () => {
        // Step 1: Create workbook with multiple pivot tables
        const workbook = Workbook.create();
        const sourceSheet = Workbook.addWorksheet(workbook, "Source");
        Worksheet.addRows(sourceSheet, TEST_DATA);

        const pivotSheet = Workbook.addWorksheet(workbook, "Pivots");

        // Add two pivot tables
        addPivotTable(pivotSheet, {
          sourceSheet: sourceSheet,
          rows: ["A"],
          columns: ["C"],
          values: ["D"],
          metric: "sum"
        });

        addPivotTable(pivotSheet, {
          sourceSheet: sourceSheet,
          rows: ["B"],
          columns: ["C"],
          values: ["E"],
          metric: "count"
        });

        expect(workbook.pivotTables.length).toBe(2);

        // Step 2: Save
        const MULTI_PIVOT_PATH = testFilePath("workbook-multi-pivot-roundtrip.test");
        await Workbook.writeFile(workbook, MULTI_PIVOT_PATH);

        // Step 3: Load
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, MULTI_PIVOT_PATH);

        // Step 4: Verify both pivot tables are loaded
        expect(loadedWorkbook.pivotTables.length).toBe(2);

        // Step 5: Save again and verify
        const MULTI_PIVOT_PATH2 = testFilePath("workbook-multi-pivot-roundtrip2.test");
        const zipData = await writeThenParseZip(loadedWorkbook, MULTI_PIVOT_PATH2);

        // Both pivot tables should have their files
        expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
        expect(zipData["xl/pivotTables/pivotTable2.xml"]).toBeDefined();
        expect(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]).toBeDefined();
        expect(zipData["xl/pivotCache/pivotCacheDefinition2.xml"]).toBeDefined();
      });

      it("preserves pivot table cache fields correctly", async () => {
        // Create workbook with specific data
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, [
          ["Category", "Value"],
          ["Alpha", 100],
          ["Beta", 200],
          ["Alpha", 150]
        ]);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["Category"],
          columns: [],
          values: ["Value"],
          metric: "sum"
        });

        const CACHE_FILEPATH = testFilePath("workbook-pivot-cache.test");
        await Workbook.writeFile(workbook, CACHE_FILEPATH);

        // Load and verify cache fields
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, CACHE_FILEPATH);

        expect(loadedWorkbook.pivotTables.length).toBe(1);
        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.cacheFields).toBeDefined();
        expect(pivot.cacheFields.length).toBe(2);
        expect(pivot.cacheFields[0].name).toBe("Category");
        expect(pivot.cacheFields[1].name).toBe("Value");
      });

      it("preserves pivot table data fields correctly", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: ["D", "E"], // Multiple values
          metric: "sum"
        });

        const DATAFIELD_FILEPATH = testFilePath("workbook-pivot-datafields.test");
        await Workbook.writeFile(workbook, DATAFIELD_FILEPATH);

        // Load and verify data fields
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, DATAFIELD_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.dataFields).toBeDefined();
        expect(pivot.dataFields!.length).toBe(2);
        expect(pivot.dataFields![0].name).toContain("D");
        expect(pivot.dataFields![1].name).toContain("E");
      });

      it("preserves count metric through load/save", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: ["D"],
          metric: "count"
        });

        const COUNT_FILEPATH = testFilePath("workbook-pivot-count.test");
        await Workbook.writeFile(workbook, COUNT_FILEPATH);

        // Load and verify metric
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, COUNT_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.metric).toBe("count");
      });

      it("preserves applyWidthHeightFormats option", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["C"],
          values: ["D"],
          applyWidthHeightFormats: "0" // Preserve column widths
        });

        const FORMATS_FILEPATH = testFilePath("workbook-pivot-formats.test");
        await Workbook.writeFile(workbook, FORMATS_FILEPATH);

        // Load and verify
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, FORMATS_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.applyWidthHeightFormats).toBe("0");
      });

      it("preserves XML special characters in cache field names", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, [
          ["Name<>", "Value&"],
          ["Test'1", 100],
          ['Test"2', 200]
        ]);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["Name<>"],
          columns: [],
          values: ["Value&"],
          metric: "sum"
        });

        const SPECIAL_CHARS_FILEPATH = testFilePath("workbook-pivot-special-chars.test");
        await Workbook.writeFile(workbook, SPECIAL_CHARS_FILEPATH);

        // Load and verify special characters are preserved
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, SPECIAL_CHARS_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.cacheFields[0].name).toBe("Name<>");
        expect(pivot.cacheFields[1].name).toBe("Value&");
      });

      it("handles pivot table with shared items", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A", "B"], // These columns will have shared items
          columns: ["C"],
          values: ["E"],
          metric: "sum"
        });

        const SHARED_ITEMS_FILEPATH = testFilePath("workbook-pivot-shared-items.test");
        await Workbook.writeFile(workbook, SHARED_ITEMS_FILEPATH);

        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, SHARED_ITEMS_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        // Check that row fields have shared items
        expect(pivot.cacheFields[0].sharedItems).toBeDefined();
        expect(pivot.cacheFields[0].sharedItems).toContain("a1");
        expect(pivot.cacheFields[1].sharedItems).toBeDefined();
        expect(pivot.cacheFields[1].sharedItems).toContain("b1");
      });
    });

    describe("Pivot tables with non-sequential cacheId", () => {
      // This tests the fix where pivot tables with
      // non-sequential cache IDs (e.g., cacheId=23 instead of 10)
      // were not being properly linked to their cache data
      const NON_SEQ_CACHE_FILEPATH = testFilePath("workbook-pivot-non-seq-cache.test");

      it("should correctly load and save pivot tables with high cacheId values", async () => {
        // Create a workbook with pivot table
        const workbook = Workbook.create();
        const sheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(sheet1, TEST_DATA);

        const sheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(sheet2, {
          sourceSheet: sheet1,
          rows: ["A"],
          columns: ["C"],
          values: ["D"],
          metric: "sum"
        });

        // Save it
        await Workbook.writeFile(workbook, NON_SEQ_CACHE_FILEPATH);

        // Load and resave
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, NON_SEQ_CACHE_FILEPATH);

        expect(loadedWorkbook.pivotTables.length).toBe(1);
        const pivot = loadedWorkbook.pivotTables[0];

        // Verify cache data is properly linked
        expect(pivot.isLoaded).toBe(true);
        expect(pivot.cacheDefinition).toBeDefined();
        expect(pivot.cacheRecords).toBeDefined();

        // Save again
        const RESAVED_FILEPATH = testFilePath("workbook-pivot-non-seq-cache-resaved.test");
        await Workbook.writeFile(loadedWorkbook, RESAVED_FILEPATH);

        // Load the resaved file
        const finalWorkbook = Workbook.create();
        await Workbook.readFile(finalWorkbook, RESAVED_FILEPATH);

        // Verify pivot table is still intact
        expect(finalWorkbook.pivotTables.length).toBe(1);
        const finalPivot = finalWorkbook.pivotTables[0];
        expect(finalPivot.isLoaded).toBe(true);
        expect(finalPivot.cacheDefinition).toBeDefined();
        expect(finalPivot.cacheRecords).toBeDefined();

        // Verify the data is preserved
        expect(finalPivot.cacheFields.length).toBeGreaterThan(0);
        expect(finalPivot.cacheRecords!.records.length).toBe(6); // 6 data rows

        // Clean up
        await promisify(fs.unlink)(NON_SEQ_CACHE_FILEPATH);
        await promisify(fs.unlink)(RESAVED_FILEPATH);
      });
    });

    describe("minimal rowItems/colItems (required by Excel)", () => {
      it("emits minimal rowItems and colItems — refreshOnLoad rebuilds full expansion", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook, "table");

        const table = addTable(worksheet, {
          name: "table",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a1", "b2", 5],
            ["a2", "b1", 24],
            ["a2", "b2", 35],
            ["a3", "b1", 45],
            ["a3", "b2", 45]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: ["B"],
          values: ["C"],
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-rowitems-and-recordcount.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);

        // Check pivotTable1.xml
        expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
        const pivotTableStr = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Minimal rowItems (grand total) and colItems are required by Excel
        expect(pivotTableStr).toContain('rowItems count="1"');
        expect(pivotTableStr).toContain('<i t="grand">');
        expect(pivotTableStr).toContain('colItems count="1"');
        // colFields present because columns=["B"] is non-empty
        expect(pivotTableStr).toContain("colFields");

        // Check pivotCacheDefinition1.xml for correct recordCount
        const cacheDefStr = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");
        expect(cacheDefStr).toMatch(/recordCount="6"/);

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });

      it("emits minimal colItems when columns is empty with single value", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook, "table");

        const table = addTable(worksheet, {
          name: "table",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a2", "b2", 10]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: [], // empty columns
          values: ["C"],
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-no-cols.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);
        const pivotTableStr = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Minimal rowItems and colItems present
        expect(pivotTableStr).toContain('rowItems count="1"');
        expect(pivotTableStr).toContain('colItems count="1"');
        // No colFields when columns is empty
        expect(pivotTableStr).not.toContain("colFields");

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });

      it("emits colFields with field x=-2 and multi-value colItems when columns is empty with multiple values", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook, "table");

        const table = addTable(worksheet, {
          name: "table",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a2", "b2", 10]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: [], // empty columns
          values: ["B", "C"], // multiple values
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-multi-vals.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);
        const pivotTableStr = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Minimal rowItems present
        expect(pivotTableStr).toContain('rowItems count="1"');
        // Multi-value requires colFields with the synthetic "Values" pseudo-field
        expect(pivotTableStr).toContain('colFields count="1"');
        expect(pivotTableStr).toContain('field x="-2"');
        // colItems: one per value + grand total = 3
        expect(pivotTableStr).toContain('colItems count="3"');

        // dataFields should still be correct
        expect(pivotTableStr).toContain('dataFields count="2"');

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });
    });

    describe("worksheetSource uses table name attribute", () => {
      it("uses table name (not sheet+ref) in pivotCacheDefinition worksheetSource", async () => {
        const workbook = Workbook.create();
        // Create a worksheet named "DataSheet" with a table named "MyTable"
        const worksheet = Workbook.addWorksheet(workbook, "DataSheet");

        const table = addTable(worksheet, {
          name: "MyTable", // Table name is different from worksheet name
          ref: "A1",
          headerRow: true,
          columns: [{ name: "Col1" }, { name: "Col2" }, { name: "Value" }],
          rows: [
            ["a", "x", 10],
            ["b", "y", 20]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "PivotSheet");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["Col1"],
          columns: ["Col2"],
          values: ["Value"],
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-worksheet-name.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);
        const cacheDefStr = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");

        // The worksheetSource should use the table name attribute
        expect(cacheDefStr).toContain('name="MyTable"');
        // Should NOT have sheet+ref format when using sourceTable
        expect(cacheDefStr).not.toContain("sheet=");
        expect(cacheDefStr).not.toContain("ref=");

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });
    });

    describe("same field in rows and values", () => {
      it("handles numeric field used as both row and value field correctly", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook);

        const table = addTable(worksheet, {
          name: "table",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a1", "b2", 5],
            ["a2", "b1", 24],
            ["a2", "b2", 35],
            ["a3", "b1", 45],
            ["a3", "b2", 45]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        // Use same field "C" for both rows and values
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["C"],
          columns: ["B"],
          values: ["C"],
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-shared-field-row-and-value.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);

        // Check pivotCacheDefinition1.xml
        const cacheDefStr = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");

        // Field C should have numeric shared items (not string)
        expect(cacheDefStr).toContain('name="C"');
        expect(cacheDefStr).toContain('containsNumber="1"');
        expect(cacheDefStr).toContain('<n v="5"');
        expect(cacheDefStr).toContain('<n v="24"');
        expect(cacheDefStr).toContain('<n v="35"');
        expect(cacheDefStr).toContain('<n v="45"');
        expect(cacheDefStr).not.toContain('<s v="5"');
        expect(cacheDefStr).not.toContain('<s v="24"');

        // Check pivotCacheRecords1.xml - records should use index references
        const cacheRecStr = decodeXml(zipData, "xl/pivotCache/pivotCacheRecords1.xml");
        expect(cacheRecStr).toContain('<x v="');

        // Check pivotTable1.xml
        const pivotStr = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");
        expect(pivotStr).toContain("Sum of C");

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });

      it("handles same numeric field used as both row and column", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook);

        const table = addTable(worksheet, {
          name: "table",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", 1, 100],
            ["a1", 2, 200],
            ["a2", 1, 300],
            ["a2", 2, 400]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["B"],
          columns: ["A"],
          values: ["C"],
          metric: "sum"
        });

        const pivotFilePath = testFilePath("workbook-pivot-numeric-rows.test");
        const zipData = await writeThenParseZip(workbook, pivotFilePath);
        const cacheDefStr = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");

        // Field B should have numeric shared items
        expect(cacheDefStr).toContain('name="B"');
        expect(cacheDefStr).toContain('<n v="1"');
        expect(cacheDefStr).toContain('<n v="2"');

        // Clean up
        await promisify(fs.unlink)(pivotFilePath);
      });
    });

    // ==========================================================================
    // Page Fields (Report Filters) Tests
    // ==========================================================================

    describe("Page Fields (Report Filters)", () => {
      it("creates pivot table with a single page field", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: ["C"],
          metric: "sum"
        });

        const filePath = testFilePath("workbook-pivot-page-single.test");
        const zipData = await writeThenParseZip(workbook, filePath);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Should have pageFields element
        expect(pivotXml).toContain("<pageFields");
        expect(pivotXml).toMatch(/<pageFields count="1">/);
        expect(pivotXml).toMatch(/<pageField fld="2" hier="-1"/);
        expect(pivotXml).toContain('axis="axisPage"');

        // Should also have sharedItems for "C" in cache definition
        const cacheDefXml = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");
        expect(cacheDefXml).toContain('name="C"');

        // Clean up
        await promisify(fs.unlink)(filePath);
      });

      it("creates pivot table with multiple page fields", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["E"],
          pages: ["B", "C"],
          metric: "sum"
        });

        const filePath = testFilePath("workbook-pivot-page-multi.test");
        const zipData = await writeThenParseZip(workbook, filePath);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Should have 2 page fields
        expect(pivotXml).toMatch(/<pageFields count="2">/);
        expect(pivotXml).toMatch(/<pageField fld="1" hier="-1"/);
        expect(pivotXml).toMatch(/<pageField fld="2" hier="-1"/);

        // Both pivotFields should have axis="axisPage"
        const axisPageMatches = pivotXml.match(/axis="axisPage"/g);
        expect(axisPageMatches).toHaveLength(2);

        // Clean up
        await promisify(fs.unlink)(filePath);
      });

      it("creates pivot table with pages, rows, columns, and values combined", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: ["C"],
          metric: "sum"
        });

        const filePath = testFilePath("workbook-pivot-page-combined.test");
        const zipData = await writeThenParseZip(workbook, filePath);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Verify all areas present (minimal rowItems/colItems/colFields required by Excel)
        expect(pivotXml).toContain("<rowFields");
        expect(pivotXml).toContain("<pageFields");
        expect(pivotXml).toContain("<dataFields");
        // colFields present because columns=["B"] is non-empty
        expect(pivotXml).toContain("colFields");
        // Minimal rowItems and colItems required by Excel
        expect(pivotXml).toContain('rowItems count="1"');
        expect(pivotXml).toContain('colItems count="1"');

        // Verify correct axis types
        expect(pivotXml).toContain('axis="axisRow"');
        expect(pivotXml).toContain('axis="axisCol"');
        expect(pivotXml).toContain('axis="axisPage"');

        // Clean up
        await promisify(fs.unlink)(filePath);
      });

      it("adjusts location ref when page fields are present", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        // Without page fields: location starts at A3
        const ws2 = Workbook.addWorksheet(workbook, "NoPagesSheet");
        addPivotTable(ws2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          metric: "sum"
        });

        // With 1 page field: location starts at A5 (3 + 1 page + 1 separator)
        const ws3 = Workbook.addWorksheet(workbook, "OnePageSheet");
        addPivotTable(ws3, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: ["C"],
          metric: "sum"
        });

        // With 2 page fields: location starts at A6 (3 + 2 pages + 1 separator)
        const ws4 = Workbook.addWorksheet(workbook, "TwoPagesSheet");
        addPivotTable(ws4, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: ["C", "D"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);

        // No pages: starts at A3
        const xml1 = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");
        expect(xml1).toMatch(/ref="A3:/);

        // 1 page field: starts at A5
        const xml2 = decodeXml(zipData, "xl/pivotTables/pivotTable2.xml");
        expect(xml2).toMatch(/ref="A5:/);

        // 2 page fields: starts at A6
        const xml3 = decodeXml(zipData, "xl/pivotTables/pivotTable3.xml");
        expect(xml3).toMatch(/ref="A6:/);
      });

      it("page fields have sharedItems in cache definition", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["E"],
          pages: ["C"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);

        const cacheDefXml = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");

        // Field "C" (page field) should have shared items with c1, c2, c3
        expect(cacheDefXml).toContain('name="C"');
        expect(cacheDefXml).toContain('<s v="c1"');
        expect(cacheDefXml).toContain('<s v="c2"');
        expect(cacheDefXml).toContain('<s v="c3"');
      });

      it("throws error for invalid page field name", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            values: ["E"],
            pages: ["NonExistent"],
            metric: "sum"
          });
        }).toThrow('The header name "NonExistent" was not found in Sheet1.');
      });

      it("throws error when same field appears in rows and columns", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: ["A"],
            values: ["E"],
            metric: "sum"
          });
        }).toThrow(
          'Field "A" cannot appear in both rows and columns. Each field can only be assigned to one axis area.'
        );
      });

      it("throws error when same field appears in rows and pages", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            values: ["E"],
            pages: ["A"],
            metric: "sum"
          });
        }).toThrow(
          'Field "A" cannot appear in both rows and pages. Each field can only be assigned to one axis area.'
        );
      });

      it("throws error when same field appears in columns and pages", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: ["B"],
            values: ["E"],
            pages: ["B"],
            metric: "sum"
          });
        }).toThrow(
          'Field "B" cannot appear in both columns and pages. Each field can only be assigned to one axis area.'
        );
      });

      it("allows same field in values and another axis area", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        // This should NOT throw - values can overlap with axis areas (dataField="1")
        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            values: ["A"],
            metric: "count"
          });
        }).not.toThrow();
      });

      it("no pageFields element when pages is empty", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: [], // explicitly empty
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);

        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Should NOT have pageFields element
        expect(pivotXml).not.toContain("<pageFields");
        expect(pivotXml).not.toContain("axisPage");

        // Location should still start at A3
        expect(pivotXml).toMatch(/ref="A3:/);
      });

      it("preserves page fields through load/save roundtrip", async () => {
        // Step 1: Create workbook with page fields
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["E"],
          pages: ["C"],
          metric: "sum"
        });

        // Step 2: Save
        const filePath = testFilePath("workbook-pivot-page-roundtrip1.test");
        await Workbook.writeFile(workbook, filePath);

        // Step 3: Load
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, filePath);

        expect(loadedWorkbook.pivotTables.length).toBe(1);
        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.isLoaded).toBe(true);

        // Step 4: Save again and verify
        const zipData = await writeThenParseZip(loadedWorkbook);

        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // pageFields should still be present
        expect(pivotXml).toContain("<pageFields");
        expect(pivotXml).toMatch(/<pageFields count="1">/);
        expect(pivotXml).toMatch(/<pageField fld="2"/);

        // pivotField with axisPage should still be present
        expect(pivotXml).toContain('axis="axisPage"');

        // Clean up
        await promisify(fs.unlink)(filePath);
      });

      it("supports page fields with sourceTable", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

        const table = addTable(worksheet, {
          name: "TestTable",
          ref: "A1",
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
          rows: [
            ["a1", "b1", "c1", 4, 5],
            ["a1", "b2", "c1", 4, 5],
            ["a2", "b1", "c2", 14, 24],
            ["a2", "b2", "c2", 24, 35]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          values: ["E"],
          pages: ["B", "C"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);

        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(pivotXml).toMatch(/<pageFields count="2">/);
        expect(pivotXml).toContain('axis="axisPage"');
      });

      it("supports 3 page fields with correct location offset", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          values: ["E"],
          pages: ["B", "C", "D"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // 3 page fields
        expect(pivotXml).toMatch(/<pageFields count="3">/);
        // Location: 3 base + 3 pages + 1 separator = A7
        expect(pivotXml).toMatch(/ref="A7:/);
        expect(pivotXml).toContain('rowPageCount="3"');
        expect(pivotXml).toContain('colPageCount="1"');
        // 3 axisPage pivotFields
        const axisPageMatches = pivotXml.match(/axis="axisPage"/g);
        expect(axisPageMatches).toHaveLength(3);
      });
    });

    describe("sourceSheet with multiple values", () => {
      it("supports sourceSheet with multiple values and no columns", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: ["D", "E"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Multi-value with no columns
        expect(pivotXml).toContain('colFields count="1"');
        expect(pivotXml).toContain('field x="-2"');
        expect(pivotXml).toContain('colItems count="3"'); // 2 values + grand total
        expect(pivotXml).toContain('dataFields count="2"');
        expect(pivotXml).toContain("Sum of D");
        expect(pivotXml).toContain("Sum of E");
      });

      it("supports sourceSheet with multiple values + pages", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: ["D", "E"],
          pages: ["B", "C"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // Multi-value + pages
        expect(pivotXml).toContain('colFields count="1"');
        expect(pivotXml).toContain('field x="-2"');
        expect(pivotXml).toContain('pageFields count="2"');
        expect(pivotXml).toContain('dataFields count="2"');
        // Location offset for 2 pages: A6
        expect(pivotXml).toMatch(/ref="A6:/);
      });
    });

    describe("deep row nesting", () => {
      it("supports 4-level row nesting", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, [
          ["L1", "L2", "L3", "L4", "Value"],
          ["a", "b", "c", "d", 10],
          ["a", "b", "c", "e", 20],
          ["a", "b", "f", "g", 30],
          ["h", "i", "j", "k", 40]
        ]);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["L1", "L2", "L3", "L4"],
          columns: [],
          values: ["Value"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // 4 row fields
        expect(pivotXml).toContain('rowFields count="4"');
        // firstDataCol should be 4 (4 row fields)
        expect(pivotXml).toContain('firstDataCol="4"');
        // 4 axisRow pivotFields
        const axisRowMatches = pivotXml.match(/axis="axisRow"/g);
        expect(axisRowMatches).toHaveLength(4);
      });
    });

    describe("multiple values with columns (columns>0 && values>1)", () => {
      it("supports multiple values with non-empty columns using sourceTable", async () => {
        const workbook = Workbook.create();
        const worksheet = Workbook.addWorksheet(workbook, "Sheet1");

        const table = addTable(worksheet, {
          name: "TestTable",
          ref: "A1",
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
          rows: [
            ["a1", "b1", "c1", 4, 5],
            ["a1", "b2", "c1", 4, 5],
            ["a2", "b1", "c2", 14, 24],
            ["a2", "b2", "c2", 24, 35]
          ]
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: ["B"],
          values: ["D", "E"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // colFields should have B + -2 sentinel = 2
        expect(pivotXml).toContain('colFields count="2"');
        expect(pivotXml).toContain('field x="-2"');
        // dataFields should have both values
        expect(pivotXml).toContain('dataFields count="2"');
        expect(pivotXml).toContain("Sum of D");
        expect(pivotXml).toContain("Sum of E");
      });

      it("supports multiple values with columns and pages combined", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: ["D", "E"],
          pages: ["C"],
          metric: "sum"
        });

        const zipData = await writeThenParseZip(workbook);
        const pivotXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // colFields: B + -2 = 2
        expect(pivotXml).toContain('colFields count="2"');
        expect(pivotXml).toContain('field x="-2"');
        // pageFields
        expect(pivotXml).toContain('pageFields count="1"');
        // dataFields
        expect(pivotXml).toContain('dataFields count="2"');
        // Location offset for 1 page: A5
        expect(pivotXml).toMatch(/ref="A5:/);
      });
    });

    describe("all metric types", () => {
      const ALL_METRICS = [
        { metric: "average" as const, display: "Average", hasSubtotal: true },
        { metric: "max" as const, display: "Max", hasSubtotal: true },
        { metric: "min" as const, display: "Min", hasSubtotal: true },
        { metric: "product" as const, display: "Product", hasSubtotal: true },
        { metric: "countNums" as const, display: "Count Numbers", hasSubtotal: true },
        { metric: "stdDev" as const, display: "StdDev", hasSubtotal: true },
        { metric: "stdDevP" as const, display: "StdDevP", hasSubtotal: true },
        { metric: "var" as const, display: "Var", hasSubtotal: true },
        { metric: "varP" as const, display: "VarP", hasSubtotal: true }
      ];

      for (const { metric, display, hasSubtotal } of ALL_METRICS) {
        it(`supports '${metric}' metric with display name "${display} of D"`, async () => {
          const workbook = Workbook.create();
          const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
          Worksheet.addRows(worksheet1, TEST_DATA);

          const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: ["C"],
            values: ["D"],
            metric
          });

          const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
          const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

          expect(pivotTableXml).toContain(`${display} of D`);
          if (hasSubtotal) {
            expect(pivotTableXml).toContain(`subtotal="${metric}"`);
          }
        });
      }
    });

    describe("per-value metric overrides", () => {
      it("supports per-value metrics with PivotTableValue objects", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: [
            { name: "D", metric: "sum" },
            { name: "E", metric: "average" }
          ]
        });

        const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
        const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(pivotTableXml).toContain("Sum of D");
        expect(pivotTableXml).toContain("Average of E");
        expect(pivotTableXml).toContain('subtotal="average"');
        expect(pivotTableXml).toContain('dataFields count="2"');
      });

      it("supports mixed string and PivotTableValue in values array", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: ["D", { name: "E", metric: "max" }],
          metric: "count"
        });

        const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
        const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        // "D" inherits table-wide "count"
        expect(pivotTableXml).toContain("Count of D");
        // "E" overrides with "max"
        expect(pivotTableXml).toContain("Max of E");
        expect(pivotTableXml).toContain('subtotal="count"');
        expect(pivotTableXml).toContain('subtotal="max"');
      });

      it("supports per-value metrics with columns and pages", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: ["B"],
          values: [
            { name: "D", metric: "min" },
            { name: "E", metric: "stdDev" }
          ],
          pages: ["C"]
        });

        const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
        const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(pivotTableXml).toContain("Min of D");
        expect(pivotTableXml).toContain("StdDev of E");
        expect(pivotTableXml).toContain('subtotal="min"');
        expect(pivotTableXml).toContain('subtotal="stdDev"');
        // Multi-value with columns → field x="-2"
        expect(pivotTableXml).toContain('field x="-2"');
        expect(pivotTableXml).toContain('pageFields count="1"');
      });

      it("preserves per-value metrics through load/save", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(worksheet2, {
          sourceSheet: worksheet1,
          rows: ["A"],
          columns: [],
          values: [
            { name: "D", metric: "sum" },
            { name: "E", metric: "average" }
          ]
        });

        const PERVALUE_FILEPATH = testFilePath("workbook-pivot-pervalue.test");
        await Workbook.writeFile(workbook, PERVALUE_FILEPATH);

        // Load and verify
        const loadedWorkbook = Workbook.create();
        await Workbook.readFile(loadedWorkbook, PERVALUE_FILEPATH);

        const pivot = loadedWorkbook.pivotTables[0];
        expect(pivot.valueMetrics).toBeDefined();
        expect(pivot.valueMetrics).toHaveLength(2);
        expect(pivot.valueMetrics[0]).toBe("sum");
        expect(pivot.valueMetrics[1]).toBe("average");
      });

      it("supports sourceTable with per-value metrics", async () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const table = addTable(worksheet1, {
          name: "MetricTable",
          ref: "A1",
          headerRow: true,
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "E" }],
          rows: TEST_DATA.slice(1)
        });

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");
        addPivotTable(worksheet2, {
          sourceTable: table,
          rows: ["A"],
          columns: [],
          values: [
            { name: "D", metric: "product" },
            { name: "E", metric: "varP" }
          ]
        });

        const zipData = await writeThenParseZip(workbook, TEST_XLSX_FILEPATH);
        const pivotTableXml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(pivotTableXml).toContain("Product of D");
        expect(pivotTableXml).toContain("VarP of E");
        expect(pivotTableXml).toContain('subtotal="product"');
        expect(pivotTableXml).toContain('subtotal="varP"');
      });
    });

    describe("intra-axis duplicate field validation (bug #3)", () => {
      it("throws error when same field appears twice in rows", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A", "A"],
            columns: ["B"],
            values: ["E"],
            metric: "sum"
          });
        }).toThrow('Duplicate field "A" in rows');
      });

      it("throws error when same field appears twice in columns", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: ["B", "B"],
            values: ["E"],
            metric: "sum"
          });
        }).toThrow('Duplicate field "B" in columns');
      });

      it("throws error when same field appears twice in pages", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            values: ["E"],
            pages: ["B", "B"],
            metric: "sum"
          });
        }).toThrow('Duplicate field "B" in pages');
      });
    });

    describe("invalid metric validation (bug #4)", () => {
      it("throws error for invalid table-wide metric", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: ["B"],
            values: ["E"],
            metric: "foo" as any
          });
        }).toThrow('Invalid metric "foo"');
      });

      it("throws error for invalid per-value metric", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const worksheet2 = Workbook.addWorksheet(workbook, "Sheet2");

        expect(() => {
          addPivotTable(worksheet2, {
            sourceSheet: worksheet1,
            rows: ["A"],
            columns: [],
            values: [
              { name: "D", metric: "sum" },
              { name: "E", metric: "banana" as any }
            ]
          });
        }).toThrow('Invalid metric "banana" on value field "E"');
      });

      it("accepts all valid metric strings without throwing", () => {
        const workbook = Workbook.create();
        const worksheet1 = Workbook.addWorksheet(workbook, "Sheet1");
        Worksheet.addRows(worksheet1, TEST_DATA);

        const validMetrics = [
          "sum",
          "count",
          "average",
          "max",
          "min",
          "product",
          "countNums",
          "stdDev",
          "stdDevP",
          "var",
          "varP"
        ] as const;

        for (const metric of validMetrics) {
          const ws = Workbook.addWorksheet(workbook, `M_${metric}`);
          expect(() => {
            addPivotTable(ws, {
              sourceSheet: worksheet1,
              rows: ["A"],
              columns: [],
              values: ["E"],
              metric
            });
          }).not.toThrow();
        }
      });
    });

    // =========================================================================
    // R9 Bug Fixes — Tests
    // =========================================================================

    describe("R9-B1: tableNumber collision avoidance", () => {
      it("assigns non-colliding tableNumber when loaded tables have non-contiguous numbering", async () => {
        // Step 1: Create 3 pivot tables → tableNumbers 1, 2, 3
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws1 = Workbook.addWorksheet(workbook, "P1");
        addPivotTable(ws1, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });

        const ws2 = Workbook.addWorksheet(workbook, "P2");
        addPivotTable(ws2, { sourceSheet: source, rows: ["B"], values: ["D"], metric: "sum" });

        const ws3 = Workbook.addWorksheet(workbook, "P3");
        addPivotTable(ws3, { sourceSheet: source, rows: ["C"], values: ["D"], metric: "sum" });

        expect(workbook.pivotTables.map(pt => pt.tableNumber)).toEqual([1, 2, 3]);

        // Step 2: Save → Load (loaded tables keep their tableNumbers)
        const filepath = testFilePath("r9-b1-table-number.test");
        await Workbook.writeFile(workbook, filepath);

        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        expect(loaded.pivotTables.length).toBe(3);
        expect(loaded.pivotTables.map(pt => pt.tableNumber).sort((a, b) => a - b)).toEqual([
          1, 2, 3
        ]);

        // Step 3: Add a new pivot table to the loaded workbook
        const newSource = Workbook.getWorksheet(loaded, "Data")!;
        const newWs = Workbook.addWorksheet(loaded, "P4");
        addPivotTable(newWs, { sourceSheet: newSource, rows: ["A"], values: ["E"], metric: "sum" });

        // The new table should get tableNumber 4 (max(1,2,3)+1)
        const newPivot = loaded.pivotTables[3];
        expect(newPivot.tableNumber).toBe(4);

        // No duplicates
        const tableNumbers = loaded.pivotTables.map(pt => pt.tableNumber);
        expect(new Set(tableNumbers).size).toBe(tableNumbers.length);

        // Step 4: Save again and verify all 4 pivot table files exist
        const zipData = await writeThenParseZip(loaded);
        for (let i = 1; i <= 4; i++) {
          expect(zipData[`xl/pivotTables/pivotTable${i}.xml`]).toBeDefined();
        }
      });

      it("handles first-ever pivot table getting tableNumber 1", () => {
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivot");

        addPivotTable(ws, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });
        expect(workbook.pivotTables[0].tableNumber).toBe(1);
      });
    });

    describe("R9-B4+B5: loaded pivot table without cacheRecords", () => {
      it("writes correctly when loaded pivot table has no cacheRecords", async () => {
        // Step 1: Create a normal pivot table and save
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(ws, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });

        const filepath = testFilePath("r9-b4-no-records.test");
        await Workbook.writeFile(workbook, filepath);

        // Step 2: Load and artificially remove cacheRecords (simulating OLAP scenario)
        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        const pivot = loaded.pivotTables[0];
        expect(pivot.cacheRecords).toBeDefined();

        // Simulate missing cache records
        delete (pivot as any).cacheRecords;

        // Step 3: Save again — should not throw
        const zipData = await writeThenParseZip(loaded);

        // Pivot table file should exist
        expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
        // Cache definition should exist
        expect(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]).toBeDefined();
        // Cache records should NOT exist
        expect(zipData["xl/pivotCache/pivotCacheRecords1.xml"]).toBeUndefined();
        // Cache definition rels should NOT exist (no records to point to)
        expect(zipData["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"]).toBeUndefined();

        // Content types should NOT mention pivotCacheRecords
        const contentTypesXml = decodeXml(zipData, "[Content_Types].xml");
        expect(contentTypesXml).not.toContain("pivotCacheRecords");
        // But should mention pivotCacheDefinition and pivotTable
        expect(contentTypesXml).toContain("pivotCacheDefinition");
        expect(contentTypesXml).toContain("pivotTable");
      });
    });

    describe("R9-B3: cache definition rels rId consistency", () => {
      it("writes rels Id matching cache definition r:id", async () => {
        // Create and save a pivot table
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(ws, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });

        const filepath = testFilePath("r9-b3-rid-consistency.test");
        await Workbook.writeFile(workbook, filepath);

        // Load, then save again
        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        const zipData = await writeThenParseZip(loaded);

        // Check that cache definition r:id and rels Id are consistent
        const cacheDefXml = decodeXml(zipData, "xl/pivotCache/pivotCacheDefinition1.xml");
        const relsXml = decodeXml(zipData, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");

        // Extract r:id from cache definition
        const rIdMatch = cacheDefXml.match(/r:id="([^"]+)"/);
        expect(rIdMatch).not.toBeNull();
        const cacheDefRId = rIdMatch![1];

        // Extract Id from rels
        const relIdMatch = relsXml.match(/Id="([^"]+)"/);
        expect(relIdMatch).not.toBeNull();
        const relsId = relIdMatch![1];

        // They must match
        expect(cacheDefRId).toBe(relsId);
      });
    });

    describe("R9-B6: shared cache deduplication", () => {
      it("does not duplicate cache files when pivot tables share the same cacheId", async () => {
        // Create a workbook with a pivot table
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws1 = Workbook.addWorksheet(workbook, "Pivot1");
        addPivotTable(ws1, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });

        const ws2 = Workbook.addWorksheet(workbook, "Pivot2");
        addPivotTable(ws2, { sourceSheet: source, rows: ["B"], values: ["E"], metric: "count" });

        // Save and load (so pivot tables get loaded with their own cacheIds)
        const filepath = testFilePath("r9-b6-shared-cache.test");
        await Workbook.writeFile(workbook, filepath);

        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        // Force both pivot tables to share the same cacheId (simulate shared cache)
        const pivot1 = loaded.pivotTables[0];
        const pivot2 = loaded.pivotTables[1];
        const sharedCacheId = pivot1.cacheId;
        pivot2.cacheId = sharedCacheId;
        // Share the cache data
        pivot2.cacheDefinition = pivot1.cacheDefinition;
        pivot2.cacheRecords = pivot1.cacheRecords;

        // Save again
        const zipData = await writeThenParseZip(loaded);

        // Both pivot tables should exist
        expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
        expect(zipData["xl/pivotTables/pivotTable2.xml"]).toBeDefined();

        // Only one cache definition should exist (the first table's)
        expect(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]).toBeDefined();
        // The second table's cache definition should NOT be written
        expect(zipData["xl/pivotCache/pivotCacheDefinition2.xml"]).toBeUndefined();

        // Content types: only one pivotCacheDefinition entry
        const contentTypesXml = decodeXml(zipData, "[Content_Types].xml");
        const cacheDefMatches = contentTypesXml.match(/pivotCacheDefinition/g) ?? [];
        // One in content type attribute value, one in PartName → 2 occurrences per entry
        // With dedup we expect exactly 1 Override element = 2 text occurrences
        expect(cacheDefMatches.length).toBe(2);

        // Workbook.xml: only one pivotCache element for the shared cacheId
        const workbookXml = decodeXml(zipData, "xl/workbook.xml");
        const pivotCacheMatches = workbookXml.match(/<pivotCache /g) ?? [];
        expect(pivotCacheMatches.length).toBe(1);

        // Workbook rels: only one pivotCacheDefinition relationship
        const workbookRelsXml = decodeXml(zipData, "xl/_rels/workbook.xml.rels");
        const cacheRelMatches = workbookRelsXml.match(/pivotCacheDefinition/g) ?? [];
        // One in Type attribute, one in Target → 2 occurrences per relationship
        expect(cacheRelMatches.length).toBe(2);

        // Pivot table 2's rels should point to pivotTable1's cache definition
        const pt2RelsXml = decodeXml(zipData, "xl/pivotTables/_rels/pivotTable2.xml.rels");
        expect(pt2RelsXml).toContain("pivotCacheDefinition1.xml");
      });
    });

    describe("R9-T5: mixed loaded + new pivot tables", () => {
      it("correctly writes both loaded and new pivot tables in the same workbook", async () => {
        // Step 1: Create and save a workbook with one pivot table
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivot1");
        addPivotTable(ws, { sourceSheet: source, rows: ["A"], values: ["D"], metric: "sum" });

        const filepath = testFilePath("r9-t5-mixed.test");
        await Workbook.writeFile(workbook, filepath);

        // Step 2: Load and add another new pivot table
        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        expect(loaded.pivotTables.length).toBe(1);
        expect(loaded.pivotTables[0].isLoaded).toBe(true);

        const loadedSource = Workbook.getWorksheet(loaded, "Data")!;
        const newWs = Workbook.addWorksheet(loaded, "Pivot2");
        addPivotTable(newWs, {
          sourceSheet: loadedSource,
          rows: ["B"],
          values: ["E"],
          metric: "count"
        });

        expect(loaded.pivotTables.length).toBe(2);
        expect(loaded.pivotTables[1].isLoaded).toBeUndefined();

        // Step 3: Save and verify both pivot tables
        const zipData = await writeThenParseZip(loaded);

        // Both pivot table files should exist
        expect(zipData["xl/pivotTables/pivotTable1.xml"]).toBeDefined();
        expect(zipData["xl/pivotTables/pivotTable2.xml"]).toBeDefined();

        // Both cache definitions should exist
        expect(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]).toBeDefined();
        expect(zipData["xl/pivotCache/pivotCacheDefinition2.xml"]).toBeDefined();

        // Both cache records should exist
        expect(zipData["xl/pivotCache/pivotCacheRecords1.xml"]).toBeDefined();
        expect(zipData["xl/pivotCache/pivotCacheRecords2.xml"]).toBeDefined();

        // Content types should have entries for both
        const contentTypesXml = decodeXml(zipData, "[Content_Types].xml");
        expect(contentTypesXml).toContain("pivotTable1.xml");
        expect(contentTypesXml).toContain("pivotTable2.xml");

        // Workbook rels should have two cache definition relationships
        const workbookRelsXml = decodeXml(zipData, "xl/_rels/workbook.xml.rels");
        const cacheRelMatches = workbookRelsXml.match(/pivotCacheDefinition\d+\.xml/g) ?? [];
        expect(cacheRelMatches.length).toBe(2);

        // Step 4: Load the result again and verify integrity
        const buffer = await Workbook.toBuffer(loaded);
        await expectValidXlsx(buffer, { label: "R9-T5 final load" });
        const final = Workbook.create();
        await Workbook.read(final, buffer as Buffer);

        expect(final.pivotTables.length).toBe(2);
        final.pivotTables.forEach(pt => {
          expect(pt.cacheDefinition).toBeDefined();
          expect(pt.cacheRecords).toBeDefined();
          expect(pt.cacheFields.length).toBeGreaterThan(0);
        });
      });
    });

    describe("R9-B8: pivotCacheDefinitionRels parsing optimization", () => {
      it("roundtrips correctly without parsing cache definition rels", async () => {
        // Verify that skipping cache def rels parsing doesn't break roundtrip
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivot");
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["A"],
          columns: ["B"],
          values: ["D"],
          metric: "sum"
        });

        const filepath = testFilePath("r9-b8-rels-skip.test");
        await Workbook.writeFile(workbook, filepath);

        // Load (this will skip parsing cache def rels per R9-B8)
        const loaded = Workbook.create();
        await Workbook.readFile(loaded, filepath);

        // Verify pivot table is fully functional
        const pivot = loaded.pivotTables[0];
        expect(pivot.cacheDefinition).toBeDefined();
        expect(pivot.cacheRecords).toBeDefined();
        expect(pivot.cacheFields.length).toBe(5);

        // Save again
        const zipData = await writeThenParseZip(loaded);

        // The cache def rels should still be written correctly
        const relsXml = decodeXml(zipData, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");
        expect(relsXml).toContain("pivotCacheRecords1.xml");
      });
    });

    describe("explicit ref (anchor) for multiple pivots on one sheet", () => {
      it("honours model.ref when writing the <location> element", async () => {
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws = Workbook.addWorksheet(workbook, "Pivots");
        // Pivot anchored at C7 — rows=["A"] (1 row field) + values=["E"] (1 data col)
        // => body spans C7:D8 (rows.length=1 so endCol = C+1 = D).
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["A"],
          values: ["E"],
          metric: "sum",
          ref: "C7"
        });

        const zipData = await writeThenParseZip(workbook);
        const xml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(xml).toContain('ref="C7:D8"');
        // firstDataCol is the offset of data columns *within the pivot*, so it
        // must stay at 1 (not absolute col 3) regardless of the anchor column.
        expect(xml).toContain('firstDataCol="1"');
      });

      it("shifts the body by pageOffset when page filters are present", async () => {
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws = Workbook.addWorksheet(workbook, "Pivots");
        // 1 page filter → pageOffset = 2 (page row + blank). Anchor at A10
        // puts the page field at row 10 and the pivot body at A12.
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["A"],
          values: ["E"],
          pages: ["C"],
          metric: "sum",
          ref: "A10"
        });

        const zipData = await writeThenParseZip(workbook);
        const xml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(xml).toMatch(/ref="A12:B13"/);
        expect(xml).toContain('rowPageCount="1"');
      });

      it("keeps multiple pivots on one sheet from sharing the same location", async () => {
        // Regression: previously every new pivot defaulted to column A row
        // 3 + pageOffset, so stacking pivots on a single sheet caused Excel
        // to report "there's already a PivotTable there" on refresh.
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws = Workbook.addWorksheet(workbook, "Stacked");
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["A"],
          values: ["D"],
          metric: "sum",
          ref: "A3"
        });
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["B"],
          values: ["D"],
          metric: "sum",
          ref: "A20"
        });
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["C"],
          values: ["D"],
          pages: ["A"],
          metric: "sum",
          ref: "A40"
        });

        const zipData = await writeThenParseZip(workbook);

        const xml1 = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");
        const xml2 = decodeXml(zipData, "xl/pivotTables/pivotTable2.xml");
        const xml3 = decodeXml(zipData, "xl/pivotTables/pivotTable3.xml");

        const refMatch = (xml: string): string => {
          const m = xml.match(/<location[^/]*ref="([^"]+)"/);
          if (!m) {
            throw new Error("no <location> element");
          }
          return m[1];
        };

        expect(refMatch(xml1)).toBe("A3:B4");
        expect(refMatch(xml2)).toBe("A20:B21");
        // 1 page filter shifts the body 2 rows below the anchor.
        expect(refMatch(xml3)).toBe("A42:B43");
      });

      it("accepts a range and reduces it to the top-left cell", async () => {
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);

        const ws = Workbook.addWorksheet(workbook, "Pivots");
        addPivotTable(ws, {
          sourceSheet: source,
          rows: ["A"],
          values: ["E"],
          metric: "sum",
          ref: "E5:Z99"
        });

        const zipData = await writeThenParseZip(workbook);
        const xml = decodeXml(zipData, "xl/pivotTables/pivotTable1.xml");

        expect(xml).toContain('ref="E5:F6"');
      });

      it("rejects malformed ref values with a PivotTableError", async () => {
        const workbook = Workbook.create();
        const source = Workbook.addWorksheet(workbook, "Data");
        Worksheet.addRows(source, TEST_DATA);
        const ws = Workbook.addWorksheet(workbook, "Pivots");

        expect(() =>
          addPivotTable(ws, {
            sourceSheet: source,
            rows: ["A"],
            values: ["E"],
            metric: "sum",
            ref: "not-a-cell"
          })
        ).toThrow(/Invalid pivot table ref/);

        // Column-only refs are not a cell address either.
        expect(() =>
          addPivotTable(ws, {
            sourceSheet: source,
            rows: ["A"],
            values: ["E"],
            metric: "sum",
            ref: "A"
          })
        ).toThrow(/Invalid pivot table ref/);
      });
    });
  });
});
