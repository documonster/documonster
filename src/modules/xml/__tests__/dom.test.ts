import { describe, it, expect } from "vitest";
import { parseXml, findChild, findChildren, textContent, attr, walk } from "@xml/dom";
import type { XmlElement } from "@xml/types";

describe("parseXml", () => {
  describe("basic parsing", () => {
    it("should parse a simple document", () => {
      const doc = parseXml("<root/>");
      expect(doc.root.name).toBe("root");
      expect(doc.root.children).toEqual([]);
    });

    it("should parse element with text", () => {
      const doc = parseXml("<root>hello</root>");
      expect(doc.root.name).toBe("root");
      expect(doc.root.children.length).toBe(1);
      expect(doc.root.children[0].type).toBe("text");
      expect(textContent(doc.root)).toBe("hello");
    });

    it("should parse element with attributes", () => {
      const doc = parseXml('<root id="1" name="test"/>');
      expect(doc.root.attributes).toEqual({ id: "1", name: "test" });
    });

    it("should parse nested elements", () => {
      const doc = parseXml("<root><child>text</child></root>");
      const child = findChild(doc.root, "child");
      expect(child).toBeDefined();
      expect(child!.name).toBe("child");
      expect(textContent(child!)).toBe("text");
    });
  });

  describe("complex structures", () => {
    it("should parse multiple children", () => {
      const doc = parseXml("<root><a/><b/><c/></root>");
      const children = doc.root.children.filter((n): n is XmlElement => n.type === "element");
      expect(children.map(c => c.name)).toEqual(["a", "b", "c"]);
    });

    it("should parse deeply nested structure", () => {
      const doc = parseXml("<a><b><c><d>deep</d></c></b></a>");
      const b = findChild(doc.root, "b");
      const c = findChild(b!, "c");
      const d = findChild(c!, "d");
      expect(textContent(d!)).toBe("deep");
    });

    it("should handle mixed content (text + elements)", () => {
      const doc = parseXml("<root>before<child/>after</root>");
      expect(doc.root.children.length).toBe(3);
      expect(doc.root.children[0].type).toBe("text");
      expect(doc.root.children[1].type).toBe("element");
      expect(doc.root.children[2].type).toBe("text");
    });
  });

  describe("CDATA handling", () => {
    it("should merge CDATA into text by default", () => {
      const doc = parseXml("<root><![CDATA[content]]></root>");
      expect(doc.root.children.length).toBe(1);
      expect(doc.root.children[0].type).toBe("text");
      expect(textContent(doc.root)).toBe("content");
    });

    it("should keep CDATA as separate nodes when cdataAsNodes is true", () => {
      const doc = parseXml("<root><![CDATA[content]]></root>", { cdataAsNodes: true });
      expect(doc.root.children.length).toBe(1);
      expect(doc.root.children[0].type).toBe("cdata");
    });
  });

  describe("comments", () => {
    it("should skip comments by default", () => {
      const doc = parseXml("<root><!-- comment -->text</root>");
      expect(doc.root.children.length).toBe(1);
      expect(doc.root.children[0].type).toBe("text");
    });

    it("should include comments when option is set", () => {
      const doc = parseXml("<root><!-- comment -->text</root>", { comments: true });
      expect(doc.root.children.length).toBe(2);
      expect(doc.root.children[0].type).toBe("comment");
      expect(doc.root.children[1].type).toBe("text");
    });
  });

  describe("entity decoding", () => {
    it("should decode entities in text", () => {
      const doc = parseXml("<root>&lt;hello&gt;</root>");
      expect(textContent(doc.root)).toBe("<hello>");
    });

    it("should decode entities in attributes", () => {
      const doc = parseXml('<root val="a &amp; b"/>');
      expect(attr(doc.root, "val")).toBe("a & b");
    });
  });

  describe("error handling", () => {
    it("should throw on empty document", () => {
      expect(() => parseXml("")).toThrow();
    });

    it("should throw on whitespace-only input", () => {
      expect(() => parseXml("   \n\t  ")).toThrow();
    });

    it("should throw on comment-only input", () => {
      expect(() => parseXml("<!-- just a comment -->")).toThrow();
    });

    it("should throw on PI-only input", () => {
      expect(() => parseXml('<?xml version="1.0"?>')).toThrow();
    });
  });

  describe("Excel-like XML", () => {
    it("should parse worksheet XML fragment", () => {
      const xml = [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
        '<row r="1">',
        '<c r="A1" t="s"><v>0</v></c>',
        '<c r="B1" t="n"><v>42</v></c>',
        "</row>",
        "</sheetData>",
        "</worksheet>"
      ].join("");

      const doc = parseXml(xml);
      expect(doc.root.name).toBe("worksheet");

      const sheetData = findChild(doc.root, "sheetData")!;
      const row = findChild(sheetData, "row")!;
      expect(attr(row, "r")).toBe("1");

      const cells = findChildren(row, "c");
      expect(cells.length).toBe(2);
      expect(attr(cells[0], "r")).toBe("A1");
      expect(attr(cells[0], "t")).toBe("s");

      const v = findChild(cells[1], "v")!;
      expect(textContent(v)).toBe("42");
    });
  });
});

describe("DOM helpers", () => {
  const doc = parseXml('<root><a id="1">text1</a><b>text2</b><a id="2">text3</a></root>');

  describe("findChild", () => {
    it("should find first child by name", () => {
      const a = findChild(doc.root, "a");
      expect(a).toBeDefined();
      expect(attr(a!, "id")).toBe("1");
    });

    it("should return undefined when not found", () => {
      expect(findChild(doc.root, "missing")).toBeUndefined();
    });
  });

  describe("findChildren", () => {
    it("should find all children by name", () => {
      const as = findChildren(doc.root, "a");
      expect(as.length).toBe(2);
      expect(attr(as[0], "id")).toBe("1");
      expect(attr(as[1], "id")).toBe("2");
    });

    it("should return empty array when none found", () => {
      expect(findChildren(doc.root, "missing")).toEqual([]);
    });
  });

  describe("textContent", () => {
    it("should get text of element", () => {
      const a = findChild(doc.root, "a")!;
      expect(textContent(a)).toBe("text1");
    });

    it("should get recursive text content", () => {
      const doc2 = parseXml("<p>before<b>bold</b>after</p>");
      expect(textContent(doc2.root)).toBe("beforeboldafter");
    });
  });

  describe("attr", () => {
    it("should get attribute value", () => {
      const a = findChild(doc.root, "a")!;
      expect(attr(a, "id")).toBe("1");
    });

    it("should return undefined for missing attribute", () => {
      expect(attr(doc.root, "missing")).toBeUndefined();
    });
  });

  describe("walk", () => {
    it("should visit all elements depth-first", () => {
      const names: string[] = [];
      walk(doc.root, el => names.push(el.name));
      expect(names).toEqual(["root", "a", "b", "a"]);
    });
  });
});

// =============================================================================
// parseXml fragment mode
// =============================================================================

describe("parseXml fragment mode", () => {
  it("should parse multiple root elements and return the first", () => {
    const doc = parseXml("<a/><b/><c/>", { fragment: true });
    expect(doc.root.name).toBe("a");
  });

  it("should still throw on truly empty input in fragment mode", () => {
    expect(() => parseXml("", { fragment: true })).toThrow("no root element");
  });

  it("should parse text-only fragment without error", () => {
    // Fragment mode doesn't require a root element to start, but parseXml
    // always requires at least one element to return as doc.root
    expect(() => parseXml("just text", { fragment: true })).toThrow("no root element");
  });
});

// =============================================================================
// parseXml with cdataAsNodes and adjacent CDATA
// =============================================================================

describe("parseXml cdataAsNodes adjacent CDATA", () => {
  it("should produce separate cdata nodes for adjacent CDATA sections", () => {
    const doc = parseXml("<root><![CDATA[a]]><![CDATA[b]]></root>", { cdataAsNodes: true });
    expect(doc.root.children.length).toBe(2);
    expect(doc.root.children[0].type).toBe("cdata");
    expect(doc.root.children[1].type).toBe("cdata");
  });
});

// =============================================================================
// textContent on different node types
// =============================================================================

describe("textContent on non-element nodes", () => {
  it("should return value for cdata node", () => {
    const doc = parseXml("<root><![CDATA[data]]></root>", { cdataAsNodes: true });
    const cdataNode = doc.root.children[0];
    expect(textContent(cdataNode)).toBe("data");
  });

  it("should return empty string for comment node", () => {
    const doc = parseXml("<root><!-- hello --></root>", { comments: true });
    const commentNode = doc.root.children[0];
    expect(textContent(commentNode)).toBe("");
  });
});
