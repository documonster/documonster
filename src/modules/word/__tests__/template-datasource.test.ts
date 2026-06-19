/**
 * DOCX Module - Template Data Source Tests
 */

import { describe, it, expect } from "vitest";

import {
  createJsonDataSource,
  createXmlDataSource,
  createCsvDataSource,
  createCompositeDataSource,
  fillTemplateFromSource
} from "../template/template-datasource";
import type { DocxDocument } from "../types";

// =============================================================================
// Helpers
// =============================================================================

function makeMinimalDoc(bodyText: string): DocxDocument {
  return {
    body: [
      {
        type: "paragraph",
        children: [{ content: [{ type: "text", text: bodyText }] }]
      }
    ]
  } as unknown as DocxDocument;
}

// =============================================================================
// createJsonDataSource
// =============================================================================

describe("createJsonDataSource", () => {
  it("constructs from a plain object", () => {
    const source = createJsonDataSource({ name: "Alice", age: 30 });
    expect(source.getData()).toEqual({ name: "Alice", age: 30 });
  });

  it("constructs from a JSON string", () => {
    const source = createJsonDataSource('{"city":"Tokyo","pop":14000000}');
    expect(source.getData()).toEqual({ city: "Tokyo", pop: 14000000 });
  });

  it("throws for non-object JSON (array)", () => {
    expect(() => createJsonDataSource("[1,2,3]")).toThrow();
  });

  it("throws for non-object JSON (null)", () => {
    expect(() => createJsonDataSource("null")).toThrow();
  });

  it("getValue resolves nested paths with dot notation", () => {
    const source = createJsonDataSource({
      user: { name: "Bob", address: { city: "Berlin" } }
    });
    expect(source.getValue("user.name")).toBe("Bob");
    expect(source.getValue("user.address.city")).toBe("Berlin");
  });

  it("getValue returns undefined for missing paths", () => {
    const source = createJsonDataSource({ a: 1 });
    expect(source.getValue("b")).toBeUndefined();
    expect(source.getValue("a.b.c")).toBeUndefined();
  });

  it("getArray returns arrays", () => {
    const source = createJsonDataSource({ items: [1, 2, 3] });
    expect(source.getArray("items")).toEqual([1, 2, 3]);
  });

  it("getArray returns empty array for non-array values", () => {
    const source = createJsonDataSource({ name: "Alice" });
    expect(source.getArray("name")).toEqual([]);
  });

  it("getArray supports nested paths", () => {
    const source = createJsonDataSource({ data: { tags: ["a", "b"] } });
    expect(source.getArray("data.tags")).toEqual(["a", "b"]);
  });
});

// =============================================================================
// createXmlDataSource
// =============================================================================

describe("createXmlDataSource", () => {
  it("parses simple XML elements as key-value pairs", () => {
    const xml = "<root><name>John</name><age>25</age></root>";
    const source = createXmlDataSource(xml);
    expect(source.getValue("name")).toBe("John");
    expect(source.getValue("age")).toBe("25");
  });

  it("parses nested elements with dot-path access", () => {
    const xml = "<root><address><city>NY</city><zip>10001</zip></address></root>";
    const source = createXmlDataSource(xml);
    expect(source.getValue("address.city")).toBe("NY");
    expect(source.getValue("address.zip")).toBe("10001");
  });

  it("parses attributes with @ prefix", () => {
    const xml = '<root id="123"><name>Test</name></root>';
    const source = createXmlDataSource(xml);
    expect(source.getValue("@id")).toBe("123");
    expect(source.getValue("name")).toBe("Test");
  });

  it("collects repeated elements into arrays", () => {
    const xml = "<root><item>A</item><item>B</item><item>C</item></root>";
    const source = createXmlDataSource(xml);
    expect(source.getArray("item")).toEqual(["A", "B", "C"]);
  });

  it("getArray returns empty array for non-array element", () => {
    const xml = "<root><name>Alice</name></root>";
    const source = createXmlDataSource(xml);
    expect(source.getArray("name")).toEqual([]);
  });

  it("handles XML with declaration and comments", () => {
    const xml = `<?xml version="1.0"?><!-- comment --><root><val>42</val></root>`;
    const source = createXmlDataSource(xml);
    expect(source.getValue("val")).toBe("42");
  });

  it("handles empty XML gracefully", () => {
    const source = createXmlDataSource("");
    expect(source.getData()).toEqual({});
  });

  it("getData returns flat record", () => {
    const xml = "<root><x>1</x><y>2</y></root>";
    const source = createXmlDataSource(xml);
    const data = source.getData();
    expect(data.x).toBe("1");
    expect(data.y).toBe("2");
  });
});

// =============================================================================
// createCsvDataSource
// =============================================================================

describe("createCsvDataSource", () => {
  it("parses basic CSV with headers as keys", () => {
    const csv = "name,age\nAlice,30\nBob,25\n";
    const source = createCsvDataSource(csv);
    const data = source.getData();
    expect(data.name).toEqual(["Alice", "Bob"]);
    expect(data.age).toEqual(["30", "25"]);
  });

  it("provides rows as array of objects under 'rows' key", () => {
    const csv = "name,age\nAlice,30\nBob,25\n";
    const source = createCsvDataSource(csv);
    expect(source.getArray("rows")).toEqual([
      { name: "Alice", age: "30" },
      { name: "Bob", age: "25" }
    ]);
  });

  it("supports custom delimiter", () => {
    const csv = "name;age\nAlice;30\n";
    const source = createCsvDataSource(csv, { delimiter: ";" });
    expect(source.getValue("name")).toEqual(["Alice"]);
  });

  it("supports custom rowsKey", () => {
    const csv = "x,y\n1,2\n";
    const source = createCsvDataSource(csv, { rowsKey: "data" });
    expect(source.getArray("data")).toEqual([{ x: "1", y: "2" }]);
  });

  it("handles empty CSV", () => {
    const source = createCsvDataSource("");
    expect(source.getArray("rows")).toEqual([]);
  });

  it("handles CSV with only headers (no data rows)", () => {
    const csv = "a,b,c\n";
    const source = createCsvDataSource(csv);
    expect(source.getArray("rows")).toEqual([]);
    expect(source.getArray("a")).toEqual([]);
  });

  it("getValue with dot path on row objects", () => {
    const csv = "name,age\nAlice,30\n";
    const source = createCsvDataSource(csv);
    // rows.0.name — dot path into array index won't work here since rows is an array
    // but rows is accessible as array
    expect(source.getValue("rows")).toBeInstanceOf(Array);
  });
});

// =============================================================================
// createCompositeDataSource
// =============================================================================

describe("createCompositeDataSource", () => {
  it("merges data from multiple sources", () => {
    const s1 = createJsonDataSource({ a: 1, b: 2 });
    const s2 = createJsonDataSource({ c: 3 });
    const composite = createCompositeDataSource([s1, s2]);
    expect(composite.getData()).toEqual({ a: 1, b: 2, c: 3 });
  });

  it("later sources override earlier ones for conflicting keys", () => {
    const s1 = createJsonDataSource({ name: "Alice", role: "dev" });
    const s2 = createJsonDataSource({ name: "Bob" });
    const composite = createCompositeDataSource([s1, s2]);
    expect(composite.getValue("name")).toBe("Bob");
    expect(composite.getValue("role")).toBe("dev");
  });

  it("getValue returns last defined value", () => {
    const s1 = createJsonDataSource({ x: 1 });
    const s2 = createJsonDataSource({ y: 2 });
    const composite = createCompositeDataSource([s1, s2]);
    expect(composite.getValue("x")).toBe(1);
    expect(composite.getValue("y")).toBe(2);
    expect(composite.getValue("z")).toBeUndefined();
  });

  it("getArray returns array from last source with non-empty array", () => {
    const s1 = createJsonDataSource({ items: [1, 2] });
    const s2 = createJsonDataSource({ items: [3, 4] });
    const composite = createCompositeDataSource([s1, s2]);
    expect(composite.getArray("items")).toEqual([3, 4]);
  });

  it("mergeArrays option concatenates arrays", () => {
    const s1 = createJsonDataSource({ tags: ["a", "b"] });
    const s2 = createJsonDataSource({ tags: ["c"] });
    const composite = createCompositeDataSource([s1, s2], { mergeArrays: true });
    expect(composite.getArray("tags")).toEqual(["a", "b", "c"]);
  });

  it("mergeArrays in getData concatenates array values", () => {
    const s1 = createJsonDataSource({ items: [1, 2] });
    const s2 = createJsonDataSource({ items: [3] });
    const composite = createCompositeDataSource([s1, s2], { mergeArrays: true });
    expect(composite.getData().items).toEqual([1, 2, 3]);
  });

  it("works with mixed source types", () => {
    const json = createJsonDataSource({ format: "json", value: 42 });
    const csv = createCsvDataSource("col\nA\nB\n");
    const composite = createCompositeDataSource([json, csv]);
    expect(composite.getValue("format")).toBe("json");
    expect(composite.getArray("col")).toEqual(["A", "B"]);
  });
});

// =============================================================================
// getValue with dot-path notation
// =============================================================================

describe("getValue dot-path notation", () => {
  it("resolves deeply nested paths", () => {
    const source = createJsonDataSource({
      a: { b: { c: { d: "deep" } } }
    });
    expect(source.getValue("a.b.c.d")).toBe("deep");
  });

  it("returns undefined for partially valid paths", () => {
    const source = createJsonDataSource({ a: { b: 1 } });
    expect(source.getValue("a.b.c")).toBeUndefined();
  });

  it("returns undefined when traversing through null", () => {
    const source = createJsonDataSource({ a: null } as unknown as Record<string, unknown>);
    expect(source.getValue("a.b")).toBeUndefined();
  });

  it("returns the value itself for single-segment paths", () => {
    const source = createJsonDataSource({ key: "value" });
    expect(source.getValue("key")).toBe("value");
  });
});

// =============================================================================
// getArray
// =============================================================================

describe("getArray", () => {
  it("returns arrays for array values", () => {
    const source = createJsonDataSource({ list: [10, 20, 30] });
    expect(source.getArray("list")).toEqual([10, 20, 30]);
  });

  it("returns empty array for missing key", () => {
    const source = createJsonDataSource({ a: 1 });
    expect(source.getArray("missing")).toEqual([]);
  });

  it("returns empty array for non-array value", () => {
    const source = createJsonDataSource({ obj: { x: 1 } });
    expect(source.getArray("obj")).toEqual([]);
  });
});

// =============================================================================
// Edge cases
// =============================================================================

describe("Edge cases", () => {
  it("createJsonDataSource with empty object", () => {
    const source = createJsonDataSource({});
    expect(source.getData()).toEqual({});
    expect(source.getValue("anything")).toBeUndefined();
    expect(source.getArray("anything")).toEqual([]);
  });

  it("createXmlDataSource with self-closing tag", () => {
    const xml = "<root><empty/></root>";
    const source = createXmlDataSource(xml);
    expect(source.getValue("empty")).toBe("");
  });

  it("createCsvDataSource with quoted fields", () => {
    const csv = 'name,desc\n"Alice","Has a ""nickname"""\n';
    const source = createCsvDataSource(csv);
    const rows = source.getArray("rows") as Record<string, string>[];
    expect(rows[0].name).toBe("Alice");
    expect(rows[0].desc).toBe('Has a "nickname"');
  });

  it("createCompositeDataSource with empty sources array", () => {
    const composite = createCompositeDataSource([]);
    expect(composite.getData()).toEqual({});
    expect(composite.getValue("x")).toBeUndefined();
    expect(composite.getArray("x")).toEqual([]);
  });

  it("missing keys return undefined across all source types", () => {
    const json = createJsonDataSource({ a: 1 });
    const xml = createXmlDataSource("<root><b>2</b></root>");
    const csv = createCsvDataSource("c\n3\n");

    expect(json.getValue("missing")).toBeUndefined();
    expect(xml.getValue("missing")).toBeUndefined();
    expect(csv.getValue("missing")).toBeUndefined();
  });
});

// =============================================================================
// fillTemplateFromSource
// =============================================================================

describe("fillTemplateFromSource", () => {
  it("fills template placeholders from data source", () => {
    const doc = makeMinimalDoc("Hello {{name}}!");
    const source = createJsonDataSource({ name: "World" });
    const result = fillTemplateFromSource(doc, source);
    const para = result.body[0] as unknown as { children: { content: { text: string }[] }[] };
    expect(para.children[0].content[0].text).toBe("Hello World!");
  });

  it("returns document unchanged when no placeholders match", () => {
    const doc = makeMinimalDoc("No placeholders here");
    const source = createJsonDataSource({ name: "Test" });
    const result = fillTemplateFromSource(doc, source);
    const para = result.body[0] as unknown as { children: { content: { text: string }[] }[] };
    expect(para.children[0].content[0].text).toBe("No placeholders here");
  });
});
