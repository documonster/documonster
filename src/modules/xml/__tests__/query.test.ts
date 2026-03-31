import { describe, it, expect } from "vitest";
import { parseXml, textContent, attr } from "@xml/dom";
import { query, queryAll } from "@xml/query";

// Helper to build a test document
const WORKSHEET_XML = [
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  "<sheetData>",
  '<row r="1">',
  '<c r="A1" t="s"><v>0</v></c>',
  '<c r="B1" t="n"><v>42</v></c>',
  "</row>",
  '<row r="2">',
  '<c r="A2" t="s"><v>1</v></c>',
  '<c r="B2" t="n"><v>99</v></c>',
  "</row>",
  "</sheetData>",
  "</worksheet>"
].join("");

const NESTED_XML = [
  "<root>",
  '  <a id="1">',
  '    <b id="10"><c>deep1</c></b>',
  '    <b id="20"><c>deep2</c></b>',
  "  </a>",
  '  <a id="2">',
  '    <b id="30"><c>deep3</c></b>',
  "  </a>",
  '  <d id="99"/>',
  "</root>"
].join("");

describe("query", () => {
  describe("simple paths", () => {
    it("should find direct child", () => {
      const doc = parseXml(WORKSHEET_XML);
      const sheetData = query(doc.root, "sheetData");
      expect(sheetData).toBeDefined();
      expect(sheetData!.name).toBe("sheetData");
    });

    it("should find nested path", () => {
      const doc = parseXml(WORKSHEET_XML);
      const v = query(doc.root, "sheetData/row/c/v");
      expect(v).toBeDefined();
      expect(textContent(v!)).toBe("0");
    });

    it("should return undefined for non-existent path", () => {
      const doc = parseXml(WORKSHEET_XML);
      expect(query(doc.root, "nonexistent")).toBeUndefined();
      expect(query(doc.root, "sheetData/nonexistent")).toBeUndefined();
    });
  });

  describe("attribute filter", () => {
    it("should filter by attribute value", () => {
      const doc = parseXml(WORKSHEET_XML);
      const cell = query(doc.root, "sheetData/row/c[@r='B1']");
      expect(cell).toBeDefined();
      expect(attr(cell!, "r")).toBe("B1");
    });

    it("should filter row by attribute", () => {
      const doc = parseXml(WORKSHEET_XML);
      const row2 = query(doc.root, "sheetData/row[@r='2']");
      expect(row2).toBeDefined();
      expect(attr(row2!, "r")).toBe("2");
    });

    it("should return undefined when attribute doesn't match", () => {
      const doc = parseXml(WORKSHEET_XML);
      expect(query(doc.root, "sheetData/row[@r='999']")).toBeUndefined();
    });
  });

  describe("wildcard", () => {
    it("should match any element with *", () => {
      const doc = parseXml(NESTED_XML);
      // root/*/b — match any child of root, then b under it
      const b = query(doc.root, "*/b");
      expect(b).toBeDefined();
      expect(attr(b!, "id")).toBe("10");
    });
  });

  describe("recursive descent //", () => {
    it("should find element at any depth", () => {
      const doc = parseXml(WORKSHEET_XML);
      const v = query(doc.root, "//v");
      expect(v).toBeDefined();
      expect(textContent(v!)).toBe("0");
    });

    it("should find nested element via recursive descent", () => {
      const doc = parseXml(NESTED_XML);
      const c = query(doc.root, "//c");
      expect(c).toBeDefined();
      expect(textContent(c!)).toBe("deep1");
    });

    it("should combine recursive descent with attribute filter", () => {
      const doc = parseXml(NESTED_XML);
      const b = query(doc.root, "//b[@id='30']");
      expect(b).toBeDefined();
      expect(attr(b!, "id")).toBe("30");
    });
  });

  describe("index filter", () => {
    it("should select by index", () => {
      const doc = parseXml(WORKSHEET_XML);
      const secondRow = query(doc.root, "sheetData/row[1]");
      expect(secondRow).toBeDefined();
      expect(attr(secondRow!, "r")).toBe("2");
    });

    it("should return undefined for out-of-bounds index", () => {
      const doc = parseXml(WORKSHEET_XML);
      expect(query(doc.root, "sheetData/row[99]")).toBeUndefined();
    });
  });
});

describe("queryAll", () => {
  it("should find all matching elements", () => {
    const doc = parseXml(WORKSHEET_XML);
    const rows = queryAll(doc.root, "sheetData/row");
    expect(rows.length).toBe(2);
    expect(attr(rows[0], "r")).toBe("1");
    expect(attr(rows[1], "r")).toBe("2");
  });

  it("should find all cells", () => {
    const doc = parseXml(WORKSHEET_XML);
    const cells = queryAll(doc.root, "sheetData/row/c");
    expect(cells.length).toBe(4);
    expect(attr(cells[0], "r")).toBe("A1");
    expect(attr(cells[3], "r")).toBe("B2");
  });

  it("should find all via recursive descent", () => {
    const doc = parseXml(NESTED_XML);
    const cs = queryAll(doc.root, "//c");
    expect(cs.length).toBe(3);
    expect(textContent(cs[0])).toBe("deep1");
    expect(textContent(cs[1])).toBe("deep2");
    expect(textContent(cs[2])).toBe("deep3");
  });

  it("should find all with attribute filter", () => {
    const doc = parseXml(WORKSHEET_XML);
    const numCells = queryAll(doc.root, "sheetData/row/c[@t='n']");
    expect(numCells.length).toBe(2);
    expect(attr(numCells[0], "r")).toBe("B1");
    expect(attr(numCells[1], "r")).toBe("B2");
  });

  it("should return empty array for no matches", () => {
    const doc = parseXml(WORKSHEET_XML);
    expect(queryAll(doc.root, "sheetData/nonexistent")).toEqual([]);
  });

  it("should find all with wildcard", () => {
    const doc = parseXml(NESTED_XML);
    // root/* — all direct children of root
    const children = queryAll(doc.root, "*");
    expect(children.length).toBe(3); // a, a, d
  });

  describe("combined path expressions", () => {
    it("should chain multiple features", () => {
      const doc = parseXml(NESTED_XML);
      // Find all c under b[@id='20']
      const result = queryAll(doc.root, "a/b[@id='20']/c");
      expect(result.length).toBe(1);
      expect(textContent(result[0])).toBe("deep2");
    });

    it("should use recursive descent with path continuation", () => {
      const doc = parseXml(WORKSHEET_XML);
      // Find all v elements anywhere
      const values = queryAll(doc.root, "//v");
      expect(values.length).toBe(4);
    });
  });
});
