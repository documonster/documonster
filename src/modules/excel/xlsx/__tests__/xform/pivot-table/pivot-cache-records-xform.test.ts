import { pivotError, type ParsedCacheRecords, type CacheField } from "@excel/pivot-table";
import { PivotCacheRecordsXform } from "@excel/xlsx/xform/pivot-table/pivot-cache-records-xform";
import { describe, it, expect } from "vitest";

// ---------------------------------------------------------------------------
// Parse tests
// ---------------------------------------------------------------------------

describe("PivotCacheRecordsXform", () => {
  describe("parse", () => {
    it("should parse empty pivotCacheRecords", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "0" } });
      xform.parseClose("pivotCacheRecords");

      expect(xform.model).not.toBeNull();
      expect(xform.model!.records).toEqual([]);
      expect(xform.model!.count).toBe(0);
      expect(xform.model!.isLoaded).toBe(true);
    });

    it("should parse records with shared item index (<x>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "x", attributes: { v: "2" } });
      xform.parseClose("x");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records).toHaveLength(1);
      expect(xform.model!.records[0]).toEqual([{ type: "x", value: 2 }]);
    });

    it("should parse records with numeric value (<n>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "n", attributes: { v: "42.5" } });
      xform.parseClose("n");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "n", value: 42.5 }]);
    });

    it("should parse missing numeric v attribute as missing value", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "n", attributes: {} });
      xform.parseClose("n");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "m" }]);
    });

    it("should parse empty numeric v attribute as missing value", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "n", attributes: { v: "" } });
      xform.parseClose("n");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "m" }]);
    });

    it("should parse string value (<s>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "s", attributes: { v: "hello" } });
      xform.parseClose("s");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "s", value: "hello" }]);
    });

    it("should parse boolean value (<b>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "2" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "b", attributes: { v: "1" } });
      xform.parseClose("b");
      xform.parseOpen({ name: "b", attributes: { v: "0" } });
      xform.parseClose("b");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([
        { type: "b", value: true },
        { type: "b", value: false }
      ]);
    });

    it("should parse missing value (<m>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "m", attributes: {} });
      xform.parseClose("m");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "m" }]);
    });

    it("should parse date value (<d>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "d", attributes: { v: "2024-01-15T00:00:00" } });
      xform.parseClose("d");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      const record = xform.model!.records[0][0];
      expect(record.type).toBe("d");
      if (record.type === "d") {
        expect(record.value).toBeInstanceOf(Date);
        expect(record.value.getUTCFullYear()).toBe(2024);
        expect(record.value.getUTCMonth()).toBe(0); // January
        expect(record.value.getUTCDate()).toBe(15);
      }
    });

    it("should parse missing date v attribute as missing value", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "d", attributes: {} });
      xform.parseClose("d");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "m" }]);
    });

    it("should parse error value (<e>)", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "e", attributes: { v: "REF!" } });
      xform.parseClose("e");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "e", value: "REF!" }]);
    });

    it("should parse multiple records with mixed types", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "2" } });
      // Record 1
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "x", attributes: { v: "0" } });
      xform.parseClose("x");
      xform.parseOpen({ name: "n", attributes: { v: "100" } });
      xform.parseClose("n");
      xform.parseClose("r");
      // Record 2
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "x", attributes: { v: "1" } });
      xform.parseClose("x");
      xform.parseOpen({ name: "m", attributes: {} });
      xform.parseClose("m");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records).toHaveLength(2);
      expect(xform.model!.records[0]).toEqual([
        { type: "x", value: 0 },
        { type: "n", value: 100 }
      ]);
      expect(xform.model!.records[1]).toEqual([{ type: "x", value: 1 }, { type: "m" }]);
    });

    it("should default x value to 0 when v is absent", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "x", attributes: {} });
      xform.parseClose("x");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "x", value: 0 }]);
    });

    it("should default s value to empty string when v is absent", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "s", attributes: {} });
      xform.parseClose("s");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model!.records[0]).toEqual([{ type: "s", value: "" }]);
    });
  });

  // ---------------------------------------------------------------------------
  // Render tests (loaded model)
  // ---------------------------------------------------------------------------

  describe("renderLoaded", () => {
    it("should render empty records", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [],
        count: 0,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain("<pivotCacheRecords");
      expect(xml).toContain('count="0"');
      // Empty records renders as self-closing tag — valid XML
      expect(xml).toMatch(/count="0"\s*\/>/);
    });

    it("should render shared item index values", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [[{ type: "x", value: 3 }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<x v="3" />');
    });

    it("should render numeric values", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [[{ type: "n", value: 42.5 }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<n v="42.5" />');
    });

    it("should render NaN/Infinity as missing value (B4 fix)", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [
          [{ type: "n", value: NaN }],
          [{ type: "n", value: Infinity }],
          [{ type: "n", value: -Infinity }]
        ],
        count: 3,
        isLoaded: true
      } as ParsedCacheRecords);

      // NaN/Infinity should render as <m /> not <n v="NaN" />
      expect(xml).not.toContain("NaN");
      expect(xml).not.toContain("Infinity");
      expect(xml).toContain("<m />");
    });

    it("should render string values with XML encoding", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [[{ type: "s", value: "A & B" }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<s v="A &amp; B" />');
    });

    it("should render boolean values", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [
          [
            { type: "b", value: true },
            { type: "b", value: false }
          ]
        ],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<b v="1" />');
      expect(xml).toContain('<b v="0" />');
    });

    it("should render missing values", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [[{ type: "m" }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain("<m />");
    });

    it("should render date values", () => {
      const xform = new PivotCacheRecordsXform();
      const date = new Date("2024-06-15T12:30:00.000Z");
      const xml = xform.toXml({
        records: [[{ type: "d", value: date }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<d v="2024-06-15T12:30:00" />');
    });

    it("should render error values", () => {
      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({
        records: [[{ type: "e", value: "DIV/0!" }]],
        count: 1,
        isLoaded: true
      } as ParsedCacheRecords);

      expect(xml).toContain('<e v="DIV/0!" />');
    });
  });

  // ---------------------------------------------------------------------------
  // Render tests (new model — from source)
  // ---------------------------------------------------------------------------

  describe("renderNew", () => {
    function createSource(
      headers: string[],
      rows: unknown[][]
    ): {
      source: {
        getSheetValues(): unknown[][];
        getColumn(n: number): { values: unknown[] };
        name: string;
        dimensions: { shortRange: string };
        getRow(n: number): { values: unknown[] };
      };
      cacheFields: CacheField[];
    } {
      const sheetValues: unknown[][] = [];
      sheetValues[1] = [undefined, ...headers];
      for (let i = 0; i < rows.length; i++) {
        sheetValues[i + 2] = [undefined, ...rows[i]];
      }

      return {
        source: {
          name: "Sheet1",
          dimensions: { shortRange: "A1:C3" },
          getSheetValues: () => sheetValues,
          getRow: (n: number) => ({ values: sheetValues[n] ?? [] }),
          getColumn: (n: number) => {
            const values: unknown[] = [];
            for (let i = 1; i < sheetValues.length; i++) {
              if (sheetValues[i]) {
                values[i] = (sheetValues[i] as unknown[])[n];
              }
            }
            return { values };
          }
        },
        cacheFields: headers.map(name => ({
          name,
          sharedItems: null
        }))
      };
    }

    it("should render records from source data", () => {
      const { source, cacheFields } = createSource(
        ["Name", "Value"],
        [
          ["Alice", 100],
          ["Bob", 200]
        ]
      );

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).toContain('count="2"');
      expect(xml).toContain('<s v="Alice" />');
      expect(xml).toContain('<n v="100" />');
      expect(xml).toContain('<s v="Bob" />');
      expect(xml).toContain('<n v="200" />');
    });

    it("should render null values as missing", () => {
      const { source, cacheFields } = createSource(
        ["Name", "Value"],
        [
          ["Alice", null],
          [null, 200]
        ]
      );

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).toContain("<m />");
    });

    it("should render Infinity as missing (B4 fix)", () => {
      const { source, cacheFields } = createSource(["Value"], [[Infinity], [-Infinity], [NaN]]);

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).not.toContain("Infinity");
      expect(xml).not.toContain("NaN");
      expect(xml).toContain("<m />");
    });

    it("should render boolean values correctly", () => {
      const { source, cacheFields } = createSource(["Flag"], [[true], [false]]);

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).toContain('<b v="1" />');
      expect(xml).toContain('<b v="0" />');
    });

    it("should render Date values correctly", () => {
      const d = new Date("2024-03-01T00:00:00.000Z");
      const { source, cacheFields } = createSource(["Date"], [[d]]);

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).toContain('<d v="2024-03-01T00:00:00" />');
    });

    it("should render PivotErrorValue as <e>", () => {
      const err = pivotError("VALUE!");
      const { source, cacheFields } = createSource(["Result"], [[err]]);

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      expect(xml).toContain('<e v="VALUE!" />');
    });

    it("should use shared item index when sharedItems is present", () => {
      const { source } = createSource(
        ["Category", "Value"],
        [
          ["A", 10],
          ["B", 20],
          ["A", 30]
        ]
      );
      const cacheFields: CacheField[] = [
        { name: "Category", sharedItems: ["A", "B"] },
        { name: "Value", sharedItems: null }
      ];

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      // "A" → index 0, "B" → index 1
      expect(xml).toContain('<x v="0" />');
      expect(xml).toContain('<x v="1" />');
    });

    it("should pad missing columns with <m /> (OOXML requirement)", () => {
      // Source has 1 column but cacheFields expects 2
      const { source } = createSource(["Name"], [["Alice"]]);
      const cacheFields: CacheField[] = [
        { name: "Name", sharedItems: null },
        { name: "Extra", sharedItems: null }
      ];

      const xform = new PivotCacheRecordsXform();
      const xml = xform.toXml({ source, cacheFields });

      // Should have the string value AND a missing pad value
      expect(xml).toContain('<s v="Alice" />');
      expect(xml).toContain("<m />");
    });
  });

  // ---------------------------------------------------------------------------
  // Parse → Render roundtrip
  // ---------------------------------------------------------------------------

  describe("roundtrip", () => {
    it("should preserve record values through parse → renderLoaded cycle", () => {
      const xform = new PivotCacheRecordsXform();

      // Parse
      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "2" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "x", attributes: { v: "0" } });
      xform.parseClose("x");
      xform.parseOpen({ name: "n", attributes: { v: "99.5" } });
      xform.parseClose("n");
      xform.parseOpen({ name: "s", attributes: { v: "test" } });
      xform.parseClose("s");
      xform.parseClose("r");
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "m", attributes: {} });
      xform.parseClose("m");
      xform.parseOpen({ name: "b", attributes: { v: "1" } });
      xform.parseClose("b");
      xform.parseOpen({ name: "e", attributes: { v: "N/A" } });
      xform.parseClose("e");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      // Render
      const xml = xform.toXml(xform.model!);

      // Verify all values are preserved
      expect(xml).toContain('<x v="0" />');
      expect(xml).toContain('<n v="99.5" />');
      expect(xml).toContain('<s v="test" />');
      expect(xml).toContain("<m />");
      expect(xml).toContain('<b v="1" />');
      expect(xml).toContain('<e v="N/A" />');
      expect(xml).toContain('count="2"');
    });

    it("should handle empty records roundtrip", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "0" } });
      xform.parseClose("pivotCacheRecords");

      const xml = xform.toXml(xform.model!);
      expect(xml).toContain('count="0"');
      expect(xml).toMatch(/count="0"\s*\/>/);
    });
  });

  // ---------------------------------------------------------------------------
  // Reset
  // ---------------------------------------------------------------------------

  describe("reset", () => {
    it("should clear model on reset", () => {
      const xform = new PivotCacheRecordsXform();

      xform.parseOpen({ name: "pivotCacheRecords", attributes: { count: "1" } });
      xform.parseOpen({ name: "r", attributes: {} });
      xform.parseOpen({ name: "n", attributes: { v: "42" } });
      xform.parseClose("n");
      xform.parseClose("r");
      xform.parseClose("pivotCacheRecords");

      expect(xform.model).not.toBeNull();
      xform.reset();
      expect(xform.model).toBeNull();
    });
  });
});
