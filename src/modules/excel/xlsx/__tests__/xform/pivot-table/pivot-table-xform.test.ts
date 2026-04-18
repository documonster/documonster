import { ZipParser } from "@archive/unzip/zip-parser";
import type { CellFormulaValue, CellValue } from "@excel/types";
import { Workbook } from "@excel/workbook";
import { PivotTableXform } from "@excel/xlsx/xform/pivot-table/pivot-table-xform";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

// ---------------------------------------------------------------------------
// Helper: build a pivot table from a table source and return decoded XML strings
// ---------------------------------------------------------------------------
interface TableDef {
  columns: { name: string }[];
  rows: Array<Array<CellValue | CellFormulaValue>>;
}

interface PivotXml {
  pivotTableXml: string;
  cacheDefXml: string;
  cacheRecXml: string;
}

async function buildPivotXml(
  tableDef: TableDef,
  pivotOptions: Omit<
    Parameters<InstanceType<typeof Workbook>["worksheets"][0]["addPivotTable"]>[0],
    "sourceTable" | "sourceSheet"
  >
): Promise<PivotXml> {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();

  const table = worksheet.addTable({
    name: "TestTable",
    ref: "A1",
    headerRow: true,
    columns: tableDef.columns,
    rows: tableDef.rows
  });

  const worksheet2 = workbook.addWorksheet("Pivot");
  worksheet2.addPivotTable({ sourceTable: table, ...pivotOptions });

  const buffer = await workbook.xlsx.writeBuffer();
  const zipData = new ZipParser(buffer as Buffer).extractAllSync();

  return {
    pivotTableXml: new TextDecoder().decode(zipData["xl/pivotTables/pivotTable1.xml"]),
    cacheDefXml: new TextDecoder().decode(zipData["xl/pivotCache/pivotCacheDefinition1.xml"]),
    cacheRecXml: new TextDecoder().decode(zipData["xl/pivotCache/pivotCacheRecords1.xml"])
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("PivotTableXform - renderPivotFields", () => {
  describe("dataField attribute", () => {
    it("should add dataField=1 when field is used as both row and value", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a2", "b2", 10]
          ]
        },
        { rows: ["C"], columns: ["B"], values: ["C"], metric: "sum" }
      );

      expect(pivotTableXml).toMatch(/axis="axisRow"[^>]*dataField="1"/);
    });

    it("should add dataField=1 when field is used as both column and value", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a2", "b2", 10]
          ]
        },
        { rows: ["A"], columns: ["C"], values: ["C"], metric: "sum" }
      );

      expect(pivotTableXml).toMatch(/axis="axisCol"[^>]*dataField="1"/);
    });

    it("should NOT add dataField=1 for row-only fields", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["A", 10],
            ["B", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('axis="axisRow"');
      expect(pivotTableXml).not.toMatch(/axis="axisRow"[^>]*dataField="1"/);
    });

    it("should add dataField=1 for value-only fields", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["A", 10],
            ["B", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('dataField="1"');
    });
  });

  describe("sharedItems in pivotCacheDefinition", () => {
    it("should use <n> for numeric sharedItems", async () => {
      const { cacheDefXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
          rows: [
            ["a1", "b1", 5],
            ["a1", "b2", 5],
            ["a2", "b1", 24],
            ["a2", "b2", 35]
          ]
        },
        { rows: ["C"], columns: ["B"], values: ["C"], metric: "sum" }
      );

      expect(cacheDefXml).toContain('containsNumber="1"');
      expect(cacheDefXml).toContain('<n v="5"');
      expect(cacheDefXml).toContain('<n v="24"');
      expect(cacheDefXml).toContain('<n v="35"');
      expect(cacheDefXml).not.toContain('<s v="5"');
      expect(cacheDefXml).not.toContain('<s v="24"');
    });

    it("should use empty <sharedItems/> for unused fields", async () => {
      const { cacheDefXml } = await buildPivotXml(
        {
          columns: [{ name: "Unused" }, { name: "Row" }, { name: "Value" }],
          rows: [
            ["x", "A", 10],
            ["y", "B", 20]
          ]
        },
        { rows: ["Row"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(cacheDefXml).toMatch(/name="Unused"[^>]*>[\s\S]*?<sharedItems\s*\/>/);
    });

    it("should use <s> for string sharedItems", async () => {
      const { cacheDefXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["Apple", 10],
            ["Banana", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(cacheDefXml).toContain('<s v="Apple"');
      expect(cacheDefXml).toContain('<s v="Banana"');
    });
  });

  describe("minimal rowItems and colItems (required by Excel)", () => {
    it("should emit minimal rowItems (grand total) and colItems", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["Same", 10],
            ["Same", 20],
            ["Same", 30]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      // Minimal rowItems with grand total required by Excel
      expect(pivotTableXml).toContain('rowItems count="1"');
      expect(pivotTableXml).toContain('<i t="grand">');
      // Minimal colItems required by Excel
      expect(pivotTableXml).toContain('colItems count="1"');
      // No colFields when columns is empty
      expect(pivotTableXml).not.toContain("colFields");
    });

    it("should emit minimal rowItems regardless of unique value count", async () => {
      const rows: (string | number)[][] = [];
      for (let i = 1; i <= 10; i++) {
        rows.push([`Cat${i}`, i * 10]);
      }

      const { pivotTableXml } = await buildPivotXml(
        { columns: [{ name: "Category" }, { name: "Value" }], rows },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      // Only 1 rowItem (grand total), NOT expanded per unique value
      expect(pivotTableXml).toContain('rowItems count="1"');
      expect(pivotTableXml).toContain('<i t="grand">');
      expect(pivotTableXml).toContain('colItems count="1"');
      expect(pivotTableXml).not.toContain("colFields");
    });

    it("should emit colFields with field x=-2 and multi-value colItems when multiple values and no column fields", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Val1" }, { name: "Val2" }, { name: "Val3" }],
          rows: [
            ["A", 10, 20, 30],
            ["B", 40, 50, 60]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Val1", "Val2", "Val3"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('rowItems count="1"');
      // Multi-value colFields: synthetic "Values" pseudo-field
      expect(pivotTableXml).toContain('colFields count="1"');
      expect(pivotTableXml).toContain('field x="-2"');
      // colItems: one per value + grand total = 4
      expect(pivotTableXml).toContain('colItems count="4"');
      expect(pivotTableXml).toContain('dataFields count="3"');
    });
  });

  describe("pivotField edge cases", () => {
    it("should handle field used in both rows and columns", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "Value" }],
          rows: [
            ["a1", "b1", 10],
            ["a2", "b2", 20]
          ]
        },
        { rows: ["A"], columns: ["B"], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('axis="axisRow"');
      expect(pivotTableXml).toContain('axis="axisCol"');
    });

    it("should handle Unicode field names in pivot table", async () => {
      const { cacheDefXml } = await buildPivotXml(
        {
          columns: [{ name: "カテゴリ" }, { name: "数値" }],
          rows: [
            ["東京", 100],
            ["大阪", 200]
          ]
        },
        { rows: ["カテゴリ"], columns: [], values: ["数値"], metric: "sum" }
      );

      expect(cacheDefXml).toContain('name="カテゴリ"');
      expect(cacheDefXml).toContain('name="数値"');
      expect(cacheDefXml).toContain('<s v="東京"');
      expect(cacheDefXml).toContain('<s v="大阪"');
    });

    it("should handle XML special characters in dataField names", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value <A&B>" }],
          rows: [
            ["X", 10],
            ["Y", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value <A&B>"], metric: "sum" }
      );

      expect(pivotTableXml).toContain("Sum of Value &lt;A&amp;B&gt;");
    });
  });

  describe("page fields (report filters)", () => {
    it("should render pivotField with axis='axisPage' for page fields", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Region" }, { name: "Category" }, { name: "Value" }],
          rows: [
            ["East", "A", 10],
            ["West", "B", 20]
          ]
        },
        { rows: ["Category"], columns: [], pages: ["Region"], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('axis="axisPage"');
      expect(pivotTableXml).toMatch(/axis="axisPage"[^>]*>[\s\S]*?<items count="3"/);
    });

    it("should render <pageFields> element with correct field indices", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "Value" }],
          rows: [
            ["a1", "b1", "c1", 10],
            ["a2", "b2", "c2", 20]
          ]
        },
        { rows: ["A"], columns: [], pages: ["B", "C"], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).toContain('pageFields count="2"');
      expect(pivotTableXml).toContain('fld="1"');
      expect(pivotTableXml).toContain('fld="2"');
      expect(pivotTableXml).toContain('hier="-1"');
    });

    it("should not render <pageFields> when pages is empty", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["A", 10],
            ["B", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(pivotTableXml).not.toContain("pageFields");
      expect(pivotTableXml).not.toContain("axisPage");
    });

    it("should place pageFields between colItems and dataFields in XML order", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Region" }, { name: "Category" }, { name: "Value" }],
          rows: [
            ["East", "A", 10],
            ["West", "B", 20]
          ]
        },
        { rows: ["Category"], columns: [], pages: ["Region"], values: ["Value"], metric: "sum" }
      );

      // OOXML order: rowFields → rowItems → colItems → pageFields → dataFields
      const rowItemsPos = pivotTableXml.indexOf("</rowItems>");
      const colItemsPos = pivotTableXml.indexOf("</colItems>");
      const pageFieldsPos = pivotTableXml.indexOf("<pageFields");
      const dataFieldsPos = pivotTableXml.indexOf("<dataFields");

      expect(rowItemsPos).toBeGreaterThan(-1);
      expect(colItemsPos).toBeGreaterThan(-1);
      expect(pageFieldsPos).toBeGreaterThan(-1);
      expect(dataFieldsPos).toBeGreaterThan(-1);
      expect(rowItemsPos).toBeLessThan(colItemsPos);
      expect(colItemsPos).toBeLessThan(pageFieldsPos);
      expect(pageFieldsPos).toBeLessThan(dataFieldsPos);
    });

    it("should render page field with items including sharedItems and default", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Region" }, { name: "Category" }, { name: "Value" }],
          rows: [
            ["East", "A", 10],
            ["West", "B", 20],
            ["North", "C", 30]
          ]
        },
        { rows: ["Category"], columns: [], pages: ["Region"], values: ["Value"], metric: "sum" }
      );

      const pageFieldMatch = pivotTableXml.match(
        /axis="axisPage"[\s\S]*?<items count="(\d+)">([\s\S]*?)<\/items>/
      );
      expect(pageFieldMatch).not.toBeNull();
      expect(pageFieldMatch![1]).toBe("4"); // 3 regions + 1 default
      expect(pageFieldMatch![2]).toContain('<item x="0"');
      expect(pageFieldMatch![2]).toContain('<item x="1"');
      expect(pageFieldMatch![2]).toContain('<item x="2"');
      expect(pageFieldMatch![2]).toContain('<item t="default"');
    });

    it("should render 3 page fields with correct location offset", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "Value" }],
          rows: [
            ["a1", "b1", "c1", "d1", 10],
            ["a2", "b2", "c2", "d2", 20]
          ]
        },
        {
          rows: ["A"],
          columns: [],
          pages: ["B", "C", "D"],
          values: ["Value"],
          metric: "sum"
        }
      );

      // 3 page fields
      expect(pivotTableXml).toContain('pageFields count="3"');
      // Location: 3 base + 3 pages + 1 separator = A7
      expect(pivotTableXml).toMatch(/ref="A7:/);
      expect(pivotTableXml).toContain('rowPageCount="3"');
      expect(pivotTableXml).toContain('colPageCount="1"');
      // All three should have axisPage
      const axisPageMatches = pivotTableXml.match(/axis="axisPage"/g);
      expect(axisPageMatches).toHaveLength(3);
    });

    it("should render isPage && isValue with dataField=1 on axisPage pivotField", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Region" }, { name: "Category" }, { name: "Value" }],
          rows: [
            ["East", "A", 10],
            ["West", "B", 20]
          ]
        },
        {
          rows: ["Category"],
          columns: [],
          pages: ["Value"],
          values: ["Value"],
          metric: "count"
        }
      );

      // The "Value" field is both page and value, so it should have
      // axis="axisPage" AND dataField="1"
      expect(pivotTableXml).toMatch(/axis="axisPage"[^>]*dataField="1"/);
    });
  });

  describe("multi-value with columns (field x=-2 appended)", () => {
    it("should emit colFields with explicit columns + field x=-2 when values>1 and columns>0", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "Val1" }, { name: "Val2" }],
          rows: [
            ["a1", "b1", 10, 20],
            ["a2", "b2", 30, 40]
          ]
        },
        { rows: ["A"], columns: ["B"], values: ["Val1", "Val2"], metric: "sum" }
      );

      // colFields should have B's field + -2 sentinel = 2 fields
      expect(pivotTableXml).toContain('colFields count="2"');
      expect(pivotTableXml).toContain('field x="1"'); // B is index 1
      expect(pivotTableXml).toContain('field x="-2"');
      // colItems should be minimal (single <i/>), Excel rebuilds on refresh
      expect(pivotTableXml).toContain('colItems count="1"');
      // Both dataFields present
      expect(pivotTableXml).toContain('dataFields count="2"');
      expect(pivotTableXml).toContain("Sum of Val1");
      expect(pivotTableXml).toContain("Sum of Val2");
    });
  });

  describe("deep row nesting", () => {
    it("should handle 4-level row nesting with correct firstDataCol", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }, { name: "Value" }],
          rows: [
            ["a1", "b1", "c1", "d1", 10],
            ["a2", "b2", "c2", "d2", 20]
          ]
        },
        { rows: ["A", "B", "C", "D"], columns: [], values: ["Value"], metric: "sum" }
      );

      // 4 row fields
      expect(pivotTableXml).toContain('rowFields count="4"');
      // firstDataCol should equal number of row fields (4)
      expect(pivotTableXml).toContain('firstDataCol="4"');
      // All 4 row fields should have axis="axisRow"
      const axisRowMatches = pivotTableXml.match(/axis="axisRow"/g);
      expect(axisRowMatches).toHaveLength(4);
    });
  });

  describe("sourceSheet multi-value", () => {
    it("should support sourceSheet with multiple values and no columns", async () => {
      // This test uses buildPivotXml which uses sourceTable, but we verify
      // the XML structure is correct for multi-value no-column cases
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Dept" }, { name: "Score" }, { name: "Projects" }, { name: "Bonus" }],
          rows: [
            ["Eng", 85, 3, 5000],
            ["Sales", 90, 5, 8000],
            ["HR", 75, 2, 3000]
          ]
        },
        {
          rows: ["Dept"],
          columns: [],
          values: ["Score", "Projects", "Bonus"],
          metric: "sum"
        }
      );

      // 3 values, no columns → colFields with -2, colItems with 4 entries
      expect(pivotTableXml).toContain('colFields count="1"');
      expect(pivotTableXml).toContain('field x="-2"');
      expect(pivotTableXml).toContain('colItems count="4"'); // 3 values + grand total
      expect(pivotTableXml).toContain('dataFields count="3"');
      expect(pivotTableXml).toContain("Sum of Score");
      expect(pivotTableXml).toContain("Sum of Projects");
      expect(pivotTableXml).toContain("Sum of Bonus");
    });
  });

  describe("numeric column fields", () => {
    it("should handle numeric values as column fields with <n> sharedItems", async () => {
      const { pivotTableXml, cacheDefXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Year" }, { name: "Value" }],
          rows: [
            ["A", 2023, 100],
            ["A", 2024, 200],
            ["B", 2023, 300],
            ["B", 2024, 400]
          ]
        },
        { rows: ["Category"], columns: ["Year"], values: ["Value"], metric: "sum" }
      );

      // Year should be a column field with numeric shared items
      expect(pivotTableXml).toContain('axis="axisCol"');
      expect(cacheDefXml).toContain('containsNumber="1"');
      expect(cacheDefXml).toContain('containsInteger="1"');
      expect(cacheDefXml).toContain('<n v="2023"');
      expect(cacheDefXml).toContain('<n v="2024"');
    });
  });

  describe("all metric types", () => {
    const TABLE_DEF: TableDef = {
      columns: [{ name: "Category" }, { name: "Value" }],
      rows: [
        ["A", 100],
        ["B", 200]
      ]
    };

    it("should support 'average' metric with correct display name and subtotal", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "average"
      });
      expect(pivotTableXml).toContain("Average of Value");
      expect(pivotTableXml).toContain('subtotal="average"');
    });

    it("should support 'max' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "max"
      });
      expect(pivotTableXml).toContain("Max of Value");
      expect(pivotTableXml).toContain('subtotal="max"');
    });

    it("should support 'min' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "min"
      });
      expect(pivotTableXml).toContain("Min of Value");
      expect(pivotTableXml).toContain('subtotal="min"');
    });

    it("should support 'product' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "product"
      });
      expect(pivotTableXml).toContain("Product of Value");
      expect(pivotTableXml).toContain('subtotal="product"');
    });

    it("should support 'countNums' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "countNums"
      });
      expect(pivotTableXml).toContain("Count Numbers of Value");
      expect(pivotTableXml).toContain('subtotal="countNums"');
    });

    it("should support 'stdDev' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "stdDev"
      });
      expect(pivotTableXml).toContain("StdDev of Value");
      expect(pivotTableXml).toContain('subtotal="stdDev"');
    });

    it("should support 'stdDevP' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "stdDevP"
      });
      expect(pivotTableXml).toContain("StdDevP of Value");
      expect(pivotTableXml).toContain('subtotal="stdDevP"');
    });

    it("should support 'var' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "var"
      });
      expect(pivotTableXml).toContain("Var of Value");
      expect(pivotTableXml).toContain('subtotal="var"');
    });

    it("should support 'varP' metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "varP"
      });
      expect(pivotTableXml).toContain("VarP of Value");
      expect(pivotTableXml).toContain('subtotal="varP"');
    });

    it("should omit subtotal attribute for 'sum' (OOXML default)", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });
      expect(pivotTableXml).toContain("Sum of Value");
      expect(pivotTableXml).not.toMatch(/subtotal=/);
    });

    it("should emit subtotal='count' for count metric", async () => {
      const { pivotTableXml } = await buildPivotXml(TABLE_DEF, {
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "count"
      });
      expect(pivotTableXml).toContain("Count of Value");
      expect(pivotTableXml).toContain('subtotal="count"');
    });
  });

  describe("per-value metric overrides", () => {
    it("should support mixed metrics across value fields", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Sales" }, { name: "Qty" }, { name: "Price" }],
          rows: [
            ["A", 1000, 50, 20],
            ["B", 2000, 30, 65]
          ]
        },
        {
          rows: ["Category"],
          columns: [],
          values: [
            { name: "Sales", metric: "sum" },
            { name: "Qty", metric: "count" },
            { name: "Price", metric: "average" }
          ]
        }
      );

      expect(pivotTableXml).toContain("Sum of Sales");
      expect(pivotTableXml).toContain("Count of Qty");
      expect(pivotTableXml).toContain("Average of Price");
      expect(pivotTableXml).toContain('subtotal="count"');
      expect(pivotTableXml).toContain('subtotal="average"');
      // 3 data fields
      expect(pivotTableXml).toContain('dataFields count="3"');
    });

    it("should inherit table-wide metric for plain string values", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Sales" }, { name: "Qty" }],
          rows: [
            ["A", 1000, 50],
            ["B", 2000, 30]
          ]
        },
        {
          rows: ["Category"],
          columns: [],
          values: ["Sales", { name: "Qty", metric: "max" }],
          metric: "average"
        }
      );

      // "Sales" inherits table-wide metric "average"
      expect(pivotTableXml).toContain("Average of Sales");
      expect(pivotTableXml).toContain('subtotal="average"');
      // "Qty" overrides with "max"
      expect(pivotTableXml).toContain("Max of Qty");
      expect(pivotTableXml).toContain('subtotal="max"');
    });

    it("should default per-value metric to sum when no table-wide metric", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Sales" }, { name: "Qty" }],
          rows: [
            ["A", 1000, 50],
            ["B", 2000, 30]
          ]
        },
        {
          rows: ["Category"],
          columns: [],
          values: ["Sales", { name: "Qty", metric: "min" }]
          // no metric specified → defaults to "sum"
        }
      );

      expect(pivotTableXml).toContain("Sum of Sales");
      expect(pivotTableXml).toContain("Min of Qty");
      expect(pivotTableXml).toContain('subtotal="min"');
    });

    it("should support per-value metrics with columns present", async () => {
      const { pivotTableXml } = await buildPivotXml(
        {
          columns: [
            { name: "Region" },
            { name: "Product" },
            { name: "Revenue" },
            { name: "Units" }
          ],
          rows: [
            ["East", "Widget", 5000, 100],
            ["West", "Gadget", 8000, 200]
          ]
        },
        {
          rows: ["Region"],
          columns: ["Product"],
          values: [
            { name: "Revenue", metric: "sum" },
            { name: "Units", metric: "average" }
          ]
        }
      );

      expect(pivotTableXml).toContain("Sum of Revenue");
      expect(pivotTableXml).toContain("Average of Units");
      expect(pivotTableXml).toContain('subtotal="average"');
      // Should have -2 sentinel for multi-value with columns
      expect(pivotTableXml).toContain('field x="-2"');
      expect(pivotTableXml).toContain('colFields count="2"');
    });
  });

  describe("null values in row fields (bug #1)", () => {
    it("should render null row values as <m /> in cache records without crashing", async () => {
      const { cacheDefXml, cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Region" }, { name: "Value" }],
          rows: [
            ["East", 10],
            [null, 20],
            ["West", 30]
          ]
        },
        { rows: ["Region"], columns: [], values: ["Value"], metric: "sum" }
      );

      // Cache definition should have containsBlank="1" for the Region field
      expect(cacheDefXml).toContain('containsBlank="1"');
      expect(cacheDefXml).toContain("<m />");

      // Cache records should use <x v="..."/> index references (not crash)
      expect(cacheRecXml).toContain('<x v="');
      // Should not contain raw null rendering
      expect(cacheRecXml).not.toContain("null");
    });

    it("should handle all-null row field", async () => {
      const { cacheDefXml, cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            [null, 10],
            [null, 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      expect(cacheDefXml).toContain('containsBlank="1"');
      expect(cacheDefXml).toContain('containsSemiMixedTypes="0"');
      expect(cacheDefXml).toContain('containsString="0"');
      // Records should not crash
      expect(cacheRecXml).toContain('<x v="');
    });
  });

  describe("Date values in row fields (bug #2)", () => {
    it("should render Date row values with <d> elements in cache definition", async () => {
      const d1 = new Date("2024-01-15T00:00:00.000Z");
      const d2 = new Date("2024-06-30T00:00:00.000Z");
      const { cacheDefXml, cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Date" }, { name: "Value" }],
          rows: [
            [d1, 100],
            [d2, 200],
            [d1, 150] // duplicate date to test indexOf
          ]
        },
        { rows: ["Date"], columns: [], values: ["Value"], metric: "sum" }
      );

      // Cache definition should have <d> elements
      expect(cacheDefXml).toContain('containsDate="1"');
      expect(cacheDefXml).toContain(`<d v="${d1.toISOString().replace(/\.\d{3}Z$/, "")}"`);
      expect(cacheDefXml).toContain(`<d v="${d2.toISOString().replace(/\.\d{3}Z$/, "")}"`);

      // Cache records should use <x v="..."/> index references (not crash on Date indexOf)
      expect(cacheRecXml).toContain('<x v="');
    });
  });

  describe("Boolean values in value-only fields (bug #9)", () => {
    it("should render boolean values as <b> in cache records for inline (no sharedItems) fields", async () => {
      const { cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Active" }],
          rows: [
            ["A", true],
            ["B", false],
            ["C", true]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Active"], metric: "count" }
      );

      // Value-only fields render inline in cache records
      // Boolean values should be <b v="1"/> or <b v="0"/>, not <s>
      expect(cacheRecXml).toContain('<b v="1"');
      expect(cacheRecXml).toContain('<b v="0"');
      expect(cacheRecXml).not.toContain('<s v="true"');
      expect(cacheRecXml).not.toContain('<s v="false"');
    });
  });
});

// ---------------------------------------------------------------------------
// Round 4 Bug Fix Tests
// ---------------------------------------------------------------------------

describe("PivotTableXform - Round 4 fixes", () => {
  // Helper: render a loaded model to XML string
  function renderLoaded(model: Record<string, any>): string {
    const xform = new PivotTableXform();
    const xmlStream = new XmlWriter();
    xform.render(xmlStream, {
      isLoaded: true,
      name: "PivotTable1",
      cacheId: 0,
      pivotFields: [],
      rowFields: [],
      colFields: [],
      pageFields: [],
      dataFields: [],
      rowItems: [],
      colItems: [],
      chartFormats: [],
      ...model
    } as any);
    return xmlStream.xml;
  }

  describe("Fix 5a: conditionalFormats/chartFormats render order", () => {
    it("should render conditionalFormats before chartFormats (OOXML schema order)", () => {
      const xml = renderLoaded({
        formatsXml: '<formats count="1"><format><pivotArea/></format></formats>',
        conditionalFormatsXml:
          '<conditionalFormats count="1"><conditionalFormat><pivotAreas/></conditionalFormat></conditionalFormats>',
        chartFormats: [{ chart: 0, format: 0, pivotAreaXml: "<pivotArea/>" }]
      });

      const formatsPos = xml.indexOf("<formats ");
      const condFormatsPos = xml.indexOf("<conditionalFormats ");
      const chartFormatsPos = xml.indexOf("<chartFormats ");

      expect(formatsPos).toBeGreaterThan(-1);
      expect(condFormatsPos).toBeGreaterThan(-1);
      expect(chartFormatsPos).toBeGreaterThan(-1);
      // OOXML order: formats → conditionalFormats → chartFormats
      expect(formatsPos).toBeLessThan(condFormatsPos);
      expect(condFormatsPos).toBeLessThan(chartFormatsPos);
    });
  });

  describe("Fix 5: XML attribute ordering", () => {
    it("should render useAutoFormatting before createdVersion in pivotTableDefinition", () => {
      const xml = renderLoaded({
        useAutoFormatting: "1",
        itemPrintTitles: "1",
        multipleFieldFilters: "0"
      });

      const useAutoPos = xml.indexOf("useAutoFormatting=");
      const createdVersionPos = xml.indexOf("createdVersion=");
      expect(useAutoPos).toBeGreaterThan(-1);
      expect(createdVersionPos).toBeGreaterThan(-1);
      expect(useAutoPos).toBeLessThan(createdVersionPos);
    });

    it("should render itemPrintTitles before createdVersion in pivotTableDefinition", () => {
      const xml = renderLoaded({
        itemPrintTitles: "1"
      });

      const itemPrintPos = xml.indexOf("itemPrintTitles=");
      const createdVersionPos = xml.indexOf("createdVersion=");
      expect(itemPrintPos).toBeGreaterThan(-1);
      expect(createdVersionPos).toBeGreaterThan(-1);
      expect(itemPrintPos).toBeLessThan(createdVersionPos);
    });

    it("should render showAll before defaultSubtotal in pivotField", () => {
      const xml = renderLoaded({
        pivotFields: [{ showAll: false, defaultSubtotal: false, compact: false, outline: false }]
      });

      const showAllPos = xml.indexOf("showAll=");
      const defaultSubtotalPos = xml.indexOf("defaultSubtotal=");
      expect(showAllPos).toBeGreaterThan(-1);
      expect(defaultSubtotalPos).toBeGreaterThan(-1);
      expect(showAllPos).toBeLessThan(defaultSubtotalPos);
    });
  });

  describe("Fix 6: pivotField extraAttrs bag", () => {
    it("should parse and preserve unknown pivotField attributes on roundtrip", () => {
      const xform = new PivotTableXform();
      // Start the root element
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "pivotFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "pivotField",
        attributes: {
          showAll: "0",
          dragToRow: "0",
          dragToCol: "0",
          dragToPage: "0",
          dragToData: "0",
          showPropCell: "1",
          serverField: "1"
        }
      });
      xform.parseClose("pivotField");
      xform.parseClose("pivotFields");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.pivotFields).toHaveLength(1);
      const pf = model.pivotFields[0];
      expect(pf.extraAttrs).toBeDefined();
      expect(pf.extraAttrs!["dragToRow"]).toBe("0");
      expect(pf.extraAttrs!["dragToCol"]).toBe("0");
      expect(pf.extraAttrs!["dragToPage"]).toBe("0");
      expect(pf.extraAttrs!["dragToData"]).toBe("0");
      expect(pf.extraAttrs!["showPropCell"]).toBe("1");
      expect(pf.extraAttrs!["serverField"]).toBe("1");

      // Now render back and verify they appear
      const xml = renderLoaded({
        pivotFields: [pf]
      });

      expect(xml).toContain('dragToRow="0"');
      expect(xml).toContain('dragToCol="0"');
      expect(xml).toContain('dragToPage="0"');
      expect(xml).toContain('dragToData="0"');
      expect(xml).toContain('showPropCell="1"');
      expect(xml).toContain('serverField="1"');
    });

    it("should not create extraAttrs when all attributes are known", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "pivotFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "pivotField",
        attributes: {
          axis: "axisRow",
          showAll: "0",
          compact: "0",
          outline: "0"
        }
      });
      xform.parseClose("pivotField");
      xform.parseClose("pivotFields");
      xform.parseClose("pivotTableDefinition");

      const pf = xform.model!.pivotFields[0];
      expect(pf.extraAttrs).toBeUndefined();
    });
  });

  describe("Fix 7: dataField subtotal validation", () => {
    it("should accept valid subtotal values", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "dataFields", attributes: { count: "2" } });
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Sum of Sales", fld: "1", subtotal: "sum" }
      });
      xform.parseClose("dataField");
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Count of Qty", fld: "2", subtotal: "count" }
      });
      xform.parseClose("dataField");
      xform.parseClose("dataFields");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.dataFields[0].subtotal).toBe("sum");
      expect(model.dataFields[1].subtotal).toBe("count");
    });

    it("should fall back to undefined for invalid subtotal values", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "dataFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Bad Field", fld: "1", subtotal: "invalid_garbage" }
      });
      xform.parseClose("dataField");
      xform.parseClose("dataFields");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.dataFields[0].subtotal).toBeUndefined();
    });

    it("should handle missing subtotal attribute (defaults to undefined → sum in OOXML)", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "dataFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Sum of Sales", fld: "1" }
      });
      xform.parseClose("dataField");
      xform.parseClose("dataFields");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      // undefined subtotal → OOXML default is "sum", correctly not validated
      expect(model.dataFields[0].subtotal).toBeUndefined();
    });

    it("should validate all 11 OOXML subtotal values", () => {
      const validSubtotals = [
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
      ];
      for (const subtotal of validSubtotals) {
        const xform = new PivotTableXform();
        xform.parseOpen({
          name: "pivotTableDefinition",
          attributes: { name: "PT1", cacheId: "0" }
        });
        xform.parseOpen({ name: "dataFields", attributes: { count: "1" } });
        xform.parseOpen({
          name: "dataField",
          attributes: { name: `Field`, fld: "0", subtotal }
        });
        xform.parseClose("dataField");
        xform.parseClose("dataFields");
        xform.parseClose("pivotTableDefinition");

        expect(xform.model!.dataFields[0].subtotal).toBe(subtotal);
      }
    });
  });

  describe("Fix 2: dataField numFmtId roundtrip", () => {
    it("should parse numFmtId from dataField attributes", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "dataFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Sum of Price", fld: "3", numFmtId: "44" }
      });
      xform.parseClose("dataField");
      xform.parseClose("dataFields");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.dataFields[0].numFmtId).toBe(44);
    });

    it("should render numFmtId in dataField when present", () => {
      const xml = renderLoaded({
        dataFields: [{ name: "Sum of Price", fld: 3, baseField: 0, baseItem: 0, numFmtId: 44 }]
      });

      expect(xml).toContain('numFmtId="44"');
    });

    it("should not render numFmtId in dataField when absent", () => {
      const xml = renderLoaded({
        dataFields: [{ name: "Sum of Sales", fld: 1, baseField: 0, baseItem: 0 }]
      });

      // Should not have numFmtId in the dataField element
      // (note: other elements may have numFmtId, so check specifically in dataField context)
      const dataFieldMatch = xml.match(/<dataField[^/]*\/>/);
      expect(dataFieldMatch).not.toBeNull();
      expect(dataFieldMatch![0]).not.toContain("numFmtId");
    });
  });

  describe("Fix 3: filters raw XML preservation", () => {
    it("should parse and preserve filters XML on roundtrip", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      // Simulate filters element with nested content
      xform.parseOpen({
        name: "filters",
        attributes: { count: "1" }
      });
      xform.parseOpen({
        name: "filter",
        attributes: { fld: "2", type: "dateBetween", id: "1" }
      });
      xform.parseOpen({
        name: "autoFilter",
        attributes: { ref: "A1:E10" }
      });
      xform.parseClose("autoFilter");
      xform.parseClose("filter");
      xform.parseClose("filters");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.filtersXml).toBeDefined();
      expect(model.filtersXml).toContain("<filters");
      expect(model.filtersXml).toContain("</filters>");
      expect(model.filtersXml).toContain('fld="2"');

      // Render and verify filters appears between pivotTableStyleInfo and extLst
      const xml = renderLoaded({
        filtersXml: model.filtersXml
      });

      const styleInfoPos = xml.indexOf("pivotTableStyleInfo");
      const filtersPos = xml.indexOf("<filters");
      const closePos = xml.indexOf("</pivotTableDefinition");
      expect(filtersPos).toBeGreaterThan(styleInfoPos);
      expect(filtersPos).toBeLessThan(closePos);
    });
  });

  describe("Fix 1: reset() clears all buffers", () => {
    it("should clear all XML buffers on reset for clean reparse", () => {
      const xform = new PivotTableXform();

      // Parse a pivot table with filters
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "filters", attributes: { count: "1" } });
      xform.parseOpen({ name: "filter", attributes: { fld: "1" } });
      xform.parseClose("filter");
      xform.parseClose("filters");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.filtersXml).toBeDefined();

      // Now parse a second pivot table WITHOUT filters — should not leak
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT2", cacheId: "1" }
      });
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.filtersXml).toBeUndefined();
      expect(xform.model!.formatsXml).toBeUndefined();
      expect(xform.model!.conditionalFormatsXml).toBeUndefined();
    });
  });

  // ===========================================================================
  // Round 5 audit fixes
  // ===========================================================================

  describe("R5-Fix1: NaN values in source data treated as null", () => {
    it("should treat NaN as null in sharedItems", async () => {
      const { cacheDefXml, cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Category" }, { name: "Value" }],
          rows: [
            ["A", 10],
            ["B", NaN],
            ["C", 20]
          ]
        },
        { rows: ["Category"], columns: [], values: ["Value"], metric: "sum" }
      );

      // NaN should not appear as a numeric shared item
      expect(cacheDefXml).not.toContain('v="NaN"');
      // The records should generate without errors (NaN treated as null/missing)
      expect(cacheRecXml).toBeDefined();
      expect(cacheRecXml.length).toBeGreaterThan(0);
    });

    it("should not crash when NaN is in a row field", async () => {
      const { cacheDefXml, cacheRecXml } = await buildPivotXml(
        {
          columns: [{ name: "Group" }, { name: "Value" }],
          rows: [
            [1, 10],
            [NaN, 20],
            [2, 30]
          ]
        },
        { rows: ["Group"], columns: [], values: ["Value"], metric: "sum" }
      );

      // Should have null sentinel in sharedItems (for the NaN row)
      expect(cacheDefXml).toContain("<m />");
      // Records should exist
      expect(cacheRecXml).toBeDefined();
    });
  });

  describe("R5-Fix2: no double xmlDecode on SAX-parsed attributes", () => {
    it("should preserve entity-like sequences in field names on roundtrip", () => {
      const xform = new PivotTableXform();

      // Simulate SAX parser output — SAX already decoded &amp;lt; to &lt;
      // (original XML was: name="Revenue &amp;lt; Target")
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "dataFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "dataField",
        attributes: { name: "Revenue &lt; Target", fld: "0" }
      });
      xform.parseClose("dataField");
      xform.parseClose("dataFields");
      xform.parseClose("pivotTableDefinition");

      // The model should preserve the value as-is from SAX (no double decode)
      expect(xform.model!.dataFields[0].name).toBe("Revenue &lt; Target");
    });

    it("should preserve entity-like sequences in page field names", () => {
      const xform = new PivotTableXform();

      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "pageFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "pageField",
        attributes: { fld: "0", name: "A &amp; B" }
      });
      xform.parseClose("pageField");
      xform.parseClose("pageFields");
      xform.parseClose("pivotTableDefinition");

      // SAX would have decoded &amp; to &, so the attribute value is "A & B"
      // But if SAX passes "A &amp; B" (meaning the literal text),
      // the old code would double-decode it to "A & B" — now it should stay as-is
      expect(xform.model!.pageFields[0].name).toBe("A &amp; B");
    });
  });

  describe("R5-Fix3: catch-all collector for unknown top-level elements", () => {
    it("should preserve pivotHierarchies on roundtrip", () => {
      const xform = new PivotTableXform();

      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      // Simulate a pivotHierarchies element (OLAP)
      xform.parseOpen({
        name: "pivotHierarchies",
        attributes: { count: "2" }
      });
      xform.parseOpen({
        name: "pivotHierarchy",
        attributes: { dragToRow: "0", dragToCol: "0" }
      });
      xform.parseClose("pivotHierarchy");
      xform.parseOpen({
        name: "pivotHierarchy",
        attributes: { dragToRow: "1" }
      });
      xform.parseClose("pivotHierarchy");
      xform.parseClose("pivotHierarchies");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.unknownElementsXml).toBeDefined();
      expect(xform.model!.unknownElementsXml).toContain("<pivotHierarchies");
      expect(xform.model!.unknownElementsXml).toContain("</pivotHierarchies>");
      expect(xform.model!.unknownElementsXml).toContain('dragToRow="0"');
      expect(xform.model!.unknownElementsXml).toContain('dragToCol="0"');
    });

    it("should preserve multiple unknown elements", () => {
      const xform = new PivotTableXform();

      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      // First unknown element
      xform.parseOpen({
        name: "rowHierarchiesUsage",
        attributes: { count: "1" }
      });
      xform.parseOpen({
        name: "rowHierarchyUsage",
        attributes: { hierarchyUsage: "0" }
      });
      xform.parseClose("rowHierarchyUsage");
      xform.parseClose("rowHierarchiesUsage");
      // Second unknown element
      xform.parseOpen({
        name: "colHierarchiesUsage",
        attributes: { count: "1" }
      });
      xform.parseOpen({
        name: "colHierarchyUsage",
        attributes: { hierarchyUsage: "1" }
      });
      xform.parseClose("colHierarchyUsage");
      xform.parseClose("colHierarchiesUsage");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.unknownElementsXml).toContain("<rowHierarchiesUsage");
      expect(xform.model!.unknownElementsXml).toContain("</rowHierarchiesUsage>");
      expect(xform.model!.unknownElementsXml).toContain("<colHierarchiesUsage");
      expect(xform.model!.unknownElementsXml).toContain("</colHierarchiesUsage>");
    });

    it("should render unknown elements between filters and extLst", () => {
      const xml = renderLoaded({
        filtersXml: '<filters count="0"></filters>',
        unknownElementsXml:
          '<pivotHierarchies count="1"><pivotHierarchy dragToRow="0"/></pivotHierarchies>'
      });

      const filtersPos = xml.indexOf("<filters");
      const unknownPos = xml.indexOf("<pivotHierarchies");
      const closePos = xml.indexOf("</pivotTableDefinition");

      expect(unknownPos).toBeGreaterThan(filtersPos);
      expect(unknownPos).toBeLessThan(closePos);
    });

    it("should not collect known elements as unknown", () => {
      const xform = new PivotTableXform();

      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      // These are all known elements — should NOT end up in unknownElementsXml
      xform.parseOpen({ name: "pivotFields", attributes: { count: "0" } });
      xform.parseClose("pivotFields");
      xform.parseOpen({ name: "rowFields", attributes: { count: "0" } });
      xform.parseClose("rowFields");
      xform.parseOpen({ name: "dataFields", attributes: { count: "0" } });
      xform.parseClose("dataFields");
      xform.parseOpen({
        name: "pivotTableStyleInfo",
        attributes: { name: "PivotStyleLight16" }
      });
      xform.parseClose("pivotTableStyleInfo");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.unknownElementsXml).toBeUndefined();
    });

    it("should not leak unknown elements between parses", () => {
      const xform = new PivotTableXform();

      // First parse with unknown element
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({
        name: "pivotHierarchies",
        attributes: { count: "1" }
      });
      xform.parseClose("pivotHierarchies");
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.unknownElementsXml).toBeDefined();

      // Second parse without unknown elements — should not leak
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT2", cacheId: "1" }
      });
      xform.parseClose("pivotTableDefinition");

      expect(xform.model!.unknownElementsXml).toBeUndefined();
    });
  });

  // ===========================================================================
  // Round 6 Bug C: hasRowItems/hasColItems flags
  // ===========================================================================

  describe("R6-BugC: hasRowItems/hasColItems flags", () => {
    it("should not inject rowItems/colItems when original had none", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      // No rowItems or colItems elements parsed
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.hasRowItems).toBeUndefined();
      expect(model.hasColItems).toBeUndefined();

      // Render — should NOT contain rowItems or colItems
      const xml = renderLoaded(model);
      expect(xml).not.toContain("<rowItems");
      expect(xml).not.toContain("<colItems");
    });

    it("should emit fallback when rowItems/colItems elements were present but empty", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "rowItems", attributes: { count: "0" } });
      xform.parseClose("rowItems");
      xform.parseOpen({ name: "colItems", attributes: { count: "0" } });
      xform.parseClose("colItems");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.hasRowItems).toBe(true);
      expect(model.hasColItems).toBe(true);

      // Render — should contain fallback rowItems/colItems
      const xml = renderLoaded(model);
      expect(xml).toContain("<rowItems");
      expect(xml).toContain("<colItems");
    });

    it("should preserve actual rowItems/colItems when present", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "rowItems", attributes: { count: "1" } });
      xform.parseOpen({ name: "i", attributes: { t: "grand" } });
      xform.parseOpen({ name: "x", attributes: {} });
      xform.parseClose("x");
      xform.parseClose("i");
      xform.parseClose("rowItems");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.hasRowItems).toBe(true);
      expect(model.rowItems).toHaveLength(1);

      const xml = renderLoaded(model);
      expect(xml).toContain('<i t="grand">');
    });
  });

  // ===========================================================================
  // Round 7: Additional roundtrip tests for coverage gaps
  // ===========================================================================

  describe("R7: extLstXml roundtrip", () => {
    it("should parse and preserve extLst XML on roundtrip", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "extLst", attributes: {} });
      xform.parseOpen({
        name: "ext",
        attributes: { uri: "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" }
      });
      xform.parseOpen({
        name: "x14:pivotTableDefinition",
        attributes: { fillDownLabelsDefault: "1" }
      });
      xform.parseClose("x14:pivotTableDefinition");
      xform.parseClose("ext");
      xform.parseClose("extLst");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.extLstXml).toBeDefined();
      expect(model.extLstXml).toContain("<extLst");
      expect(model.extLstXml).toContain("</extLst>");
      expect(model.extLstXml).toContain('fillDownLabelsDefault="1"');

      // Render and verify extLst is emitted
      const xml = renderLoaded(model);
      expect(xml).toContain("<extLst");
      expect(xml).toContain("</extLst>");
      expect(xml).toContain('fillDownLabelsDefault="1"');
    });

    it("should emit empty string when loaded model has no extLst", () => {
      const xml = renderLoaded({
        // isLoaded: true (from renderLoaded defaults), no extLstXml
      });

      // Loaded model without extLstXml → no extLst emitted
      expect(xml).not.toContain("<extLst");
    });
  });

  describe("R7: autoSortScope roundtrip", () => {
    it("should parse and preserve autoSortScope XML in pivotField", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "pivotFields", attributes: { count: "1" } });
      xform.parseOpen({
        name: "pivotField",
        attributes: { axis: "axisRow", showAll: "0", sortType: "descending" }
      });
      xform.parseOpen({ name: "autoSortScope", attributes: {} });
      xform.parseOpen({ name: "pivotArea", attributes: { dataOnly: "0", outline: "0" } });
      xform.parseOpen({ name: "references", attributes: { count: "1" } });
      xform.parseOpen({
        name: "reference",
        attributes: { field: "4294967294", count: "1", selected: "0" }
      });
      xform.parseOpen({ name: "x", attributes: { v: "0" } });
      xform.parseClose("x");
      xform.parseClose("reference");
      xform.parseClose("references");
      xform.parseClose("pivotArea");
      xform.parseClose("autoSortScope");
      xform.parseClose("pivotField");
      xform.parseClose("pivotFields");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.pivotFields).toHaveLength(1);
      expect(model.pivotFields[0].autoSortScopeXml).toBeDefined();
      expect(model.pivotFields[0].autoSortScopeXml).toContain("<autoSortScope>");
      expect(model.pivotFields[0].autoSortScopeXml).toContain("</autoSortScope>");
      expect(model.pivotFields[0].autoSortScopeXml).toContain('dataOnly="0"');
      expect(model.pivotFields[0].sortType).toBe("descending");

      // Render and verify autoSortScope is emitted inside pivotField
      const xml = renderLoaded({ pivotFields: model.pivotFields });
      expect(xml).toContain("<autoSortScope>");
      expect(xml).toContain("</autoSortScope>");
      expect(xml).toContain('sortType="descending"');
    });
  });

  describe("R7: location attributes roundtrip", () => {
    it("should parse and render all location attributes", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({
        name: "location",
        attributes: {
          ref: "A4:E20",
          firstHeaderRow: "1",
          firstDataRow: "2",
          firstDataCol: "3",
          rowPageCount: "2",
          colPageCount: "1"
        }
      });
      xform.parseClose("location");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.location).toBeDefined();
      expect(model.location!.ref).toBe("A4:E20");
      expect(model.location!.firstHeaderRow).toBe(1);
      expect(model.location!.firstDataRow).toBe(2);
      expect(model.location!.firstDataCol).toBe(3);
      expect(model.location!.rowPageCount).toBe(2);
      expect(model.location!.colPageCount).toBe(1);

      // Render and verify
      const xml = renderLoaded({ location: model.location });
      expect(xml).toContain('ref="A4:E20"');
      expect(xml).toContain('firstHeaderRow="1"');
      expect(xml).toContain('firstDataRow="2"');
      expect(xml).toContain('firstDataCol="3"');
      expect(xml).toContain('rowPageCount="2"');
      expect(xml).toContain('colPageCount="1"');
    });

    it("should handle location with only required attributes (no page counts)", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({
        name: "location",
        attributes: {
          ref: "A1:C10",
          firstHeaderRow: "1",
          firstDataRow: "2",
          firstDataCol: "1"
        }
      });
      xform.parseClose("location");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.location!.rowPageCount).toBeUndefined();
      expect(model.location!.colPageCount).toBeUndefined();

      const xml = renderLoaded({ location: model.location });
      expect(xml).not.toContain("rowPageCount");
      expect(xml).not.toContain("colPageCount");
    });
  });

  describe("R7: formatsXml / conditionalFormatsXml / filtersXml roundtrip", () => {
    it("should preserve formatsXml through parse → render", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "formats", attributes: { count: "1" } });
      xform.parseOpen({ name: "format", attributes: { dxfId: "0" } });
      xform.parseOpen({ name: "pivotArea", attributes: { type: "data" } });
      xform.parseClose("pivotArea");
      xform.parseClose("format");
      xform.parseClose("formats");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.formatsXml).toBeDefined();
      expect(model.formatsXml).toContain("<formats");
      expect(model.formatsXml).toContain('dxfId="0"');

      const xml = renderLoaded(model);
      expect(xml).toContain("<formats");
      expect(xml).toContain('dxfId="0"');
    });

    it("should preserve conditionalFormatsXml through parse → render", () => {
      const xform = new PivotTableXform();
      xform.parseOpen({
        name: "pivotTableDefinition",
        attributes: { name: "PT1", cacheId: "0" }
      });
      xform.parseOpen({ name: "conditionalFormats", attributes: { count: "1" } });
      xform.parseOpen({ name: "conditionalFormat", attributes: { priority: "1" } });
      xform.parseOpen({ name: "pivotAreas", attributes: { count: "1" } });
      xform.parseOpen({ name: "pivotArea", attributes: { type: "data" } });
      xform.parseClose("pivotArea");
      xform.parseClose("pivotAreas");
      xform.parseClose("conditionalFormat");
      xform.parseClose("conditionalFormats");
      xform.parseClose("pivotTableDefinition");

      const model = xform.model!;
      expect(model.conditionalFormatsXml).toBeDefined();
      expect(model.conditionalFormatsXml).toContain("<conditionalFormats");

      const xml = renderLoaded(model);
      expect(xml).toContain("<conditionalFormats");
      expect(xml).toContain('priority="1"');
    });
  });

  describe("R7: unknownElementsXml roundtrip via renderLoaded", () => {
    it("should emit unknownElementsXml in rendered output", () => {
      const unknownXml =
        '<pivotHierarchies count="1"><pivotHierarchy dragToRow="0" dragToCol="0"/></pivotHierarchies>';
      const xml = renderLoaded({ unknownElementsXml: unknownXml });

      expect(xml).toContain("<pivotHierarchies");
      expect(xml).toContain('dragToRow="0"');
      expect(xml).toContain("</pivotHierarchies>");
    });
  });
});
