import { describe, it, expect } from "vitest";
import {
  parseXml,
  findChild,
  findChildren,
  textContent,
  attr,
  walk,
  toPlainObject
} from "@xml/dom";
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

  it("should expose all root elements via roots", () => {
    const doc = parseXml("<a/><b/><c/>", { fragment: true });
    expect(doc.roots.length).toBe(3);
    expect(doc.roots[0].name).toBe("a");
    expect(doc.roots[1].name).toBe("b");
    expect(doc.roots[2].name).toBe("c");
  });

  it("roots should have exactly one element in non-fragment mode", () => {
    const doc = parseXml("<root/>");
    expect(doc.roots.length).toBe(1);
    expect(doc.roots[0]).toBe(doc.root);
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
// parseXml prologue (top-level comments / PIs)
// =============================================================================

describe("parseXml prologue", () => {
  it("should collect top-level comments when enabled", () => {
    const doc = parseXml("<!-- before --><root/><!-- after -->", { comments: true });
    expect(doc.prologue.length).toBe(2);
    expect(doc.prologue[0].type).toBe("comment");
    expect((doc.prologue[0] as any).value).toBe(" before ");
    expect((doc.prologue[1] as any).value).toBe(" after ");
  });

  it("should collect top-level PIs when enabled", () => {
    const doc = parseXml('<?style type="xsl"?><root/>', { processingInstructions: true });
    expect(doc.prologue.length).toBe(1);
    expect(doc.prologue[0].type).toBe("processing-instruction");
    expect((doc.prologue[0] as any).target).toBe("style");
  });

  it("should return empty prologue when comments/PIs are disabled", () => {
    const doc = parseXml("<!-- hidden --><root/>");
    expect(doc.prologue.length).toBe(0);
  });

  it("should not include XML declaration in prologue", () => {
    const doc = parseXml('<?xml version="1.0"?><root/>', { processingInstructions: true });
    // xml declaration is captured in doc.declaration, not prologue
    expect(doc.prologue.length).toBe(0);
    expect(doc.declaration).toBeDefined();
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

// =============================================================================
// toPlainObject
// =============================================================================

describe("toPlainObject", () => {
  describe("basic conversion", () => {
    it("should convert empty element", () => {
      const doc = parseXml("<root/>");
      expect(toPlainObject(doc.root)).toEqual({ root: "" });
    });

    it("should convert text-only element", () => {
      const doc = parseXml("<root>hello</root>");
      expect(toPlainObject(doc.root)).toEqual({ root: "hello" });
    });

    it("should convert element with attributes", () => {
      const doc = parseXml('<root id="1" name="test"/>');
      expect(toPlainObject(doc.root)).toEqual({
        root: { "@_id": "1", "@_name": "test" }
      });
    });

    it("should convert element with attributes and text", () => {
      const doc = parseXml('<root id="1">hello</root>');
      expect(toPlainObject(doc.root)).toEqual({
        root: { "@_id": "1", "#text": "hello" }
      });
    });
  });

  describe("nested elements", () => {
    it("should convert single child element", () => {
      const doc = parseXml("<root><child>text</child></root>");
      expect(toPlainObject(doc.root)).toEqual({
        root: { child: "text" }
      });
    });

    it("should convert deeply nested structure", () => {
      const doc = parseXml("<a><b><c>deep</c></b></a>");
      expect(toPlainObject(doc.root)).toEqual({
        a: { b: { c: "deep" } }
      });
    });

    it("should merge repeated siblings into arrays", () => {
      const doc = parseXml("<list><item>a</item><item>b</item><item>c</item></list>");
      expect(toPlainObject(doc.root)).toEqual({
        list: { item: ["a", "b", "c"] }
      });
    });

    it("should handle mix of unique and repeated children", () => {
      const doc = parseXml("<root><name>test</name><tag>a</tag><tag>b</tag></root>");
      expect(toPlainObject(doc.root)).toEqual({
        root: { name: "test", tag: ["a", "b"] }
      });
    });
  });

  describe("attributes on nested elements", () => {
    it("should preserve attributes on child elements", () => {
      const doc = parseXml('<root><child id="1">text</child></root>');
      expect(toPlainObject(doc.root)).toEqual({
        root: { child: { "@_id": "1", "#text": "text" } }
      });
    });

    it("should handle attributes at multiple levels", () => {
      const doc = parseXml('<root attr="r"><child attr="c"/></root>');
      expect(toPlainObject(doc.root)).toEqual({
        root: { "@_attr": "r", child: { "@_attr": "c" } }
      });
    });
  });

  describe("options", () => {
    it("should use custom attribute prefix", () => {
      const doc = parseXml('<root id="1"/>');
      expect(toPlainObject(doc.root, { attributePrefix: "$" })).toEqual({
        root: { $id: "1" }
      });
    });

    it("should use empty attribute prefix", () => {
      const doc = parseXml('<root id="1"/>');
      expect(toPlainObject(doc.root, { attributePrefix: "" })).toEqual({
        root: { id: "1" }
      });
    });

    it("should use custom text key", () => {
      const doc = parseXml('<root id="1">hello</root>');
      expect(toPlainObject(doc.root, { textKey: "_text" })).toEqual({
        root: { "@_id": "1", _text: "hello" }
      });
    });

    it("should always wrap in arrays when alwaysArray is true", () => {
      const doc = parseXml("<root><child>text</child></root>");
      expect(toPlainObject(doc.root, { alwaysArray: true })).toEqual({
        root: { child: ["text"] }
      });
    });

    it("alwaysArray should still work with repeated siblings", () => {
      const doc = parseXml("<root><item>a</item><item>b</item></root>");
      expect(toPlainObject(doc.root, { alwaysArray: true })).toEqual({
        root: { item: ["a", "b"] }
      });
    });
  });

  describe("CDATA handling", () => {
    it("should include CDATA text by default (merged by parseXml)", () => {
      const doc = parseXml("<root><![CDATA[content]]></root>");
      // parseXml merges CDATA into text by default
      expect(toPlainObject(doc.root)).toEqual({ root: "content" });
    });

    it("should include CDATA when preserveCData is true and cdataAsNodes", () => {
      const doc = parseXml("<root><![CDATA[content]]></root>", { cdataAsNodes: true });
      expect(toPlainObject(doc.root, { preserveCData: true })).toEqual({ root: "content" });
    });

    it("should skip CDATA when preserveCData is false and cdataAsNodes", () => {
      const doc = parseXml("<root><![CDATA[content]]></root>", { cdataAsNodes: true });
      expect(toPlainObject(doc.root, { preserveCData: false })).toEqual({ root: "" });
    });
  });

  describe("Excel-like XML", () => {
    it("should convert worksheet XML to plain object", () => {
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
      const obj = toPlainObject(doc.root);

      expect(obj).toEqual({
        worksheet: {
          "@_xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          sheetData: {
            row: {
              "@_r": "1",
              c: [
                { "@_r": "A1", "@_t": "s", v: "0" },
                { "@_r": "B1", "@_t": "n", v: "42" }
              ]
            }
          }
        }
      });
    });
  });

  describe("edge cases", () => {
    it("should handle mixed content (text + elements)", () => {
      const doc = parseXml("<p>before<b>bold</b>after</p>");
      expect(toPlainObject(doc.root)).toEqual({
        p: { "#text": "beforeafter", b: "bold" }
      });
    });

    it("should handle element with only empty children", () => {
      const doc = parseXml("<root><a/><b/></root>");
      expect(toPlainObject(doc.root)).toEqual({
        root: { a: "", b: "" }
      });
    });

    it("should handle single element child with alwaysArray", () => {
      const doc = parseXml("<root><only/></root>");
      expect(toPlainObject(doc.root, { alwaysArray: true })).toEqual({
        root: { only: [""] }
      });
    });
  });

  describe("whitespace handling", () => {
    it("should ignore whitespace-only text nodes by default (pretty-printed XML)", () => {
      const xml = "<root>\n  <child>text</child>\n</root>";
      const doc = parseXml(xml);
      expect(toPlainObject(doc.root)).toEqual({
        root: { child: "text" }
      });
    });

    it("should preserve whitespace-only text when ignoreWhitespaceText is false", () => {
      const xml = "<root>\n  <child>text</child>\n</root>";
      const doc = parseXml(xml);
      expect(toPlainObject(doc.root, { ignoreWhitespaceText: false })).toEqual({
        root: { "#text": "\n  \n", child: "text" }
      });
    });

    it("should preserve non-whitespace text even when ignoreWhitespaceText is true", () => {
      const xml = "<root> meaningful text </root>";
      const doc = parseXml(xml);
      expect(toPlainObject(doc.root)).toEqual({
        root: " meaningful text "
      });
    });

    it("should handle pretty-printed nested XML", () => {
      const xml = ["<root>", "  <a>", "    <b>deep</b>", "  </a>", "</root>"].join("\n");
      const doc = parseXml(xml);
      expect(toPlainObject(doc.root)).toEqual({
        root: { a: { b: "deep" } }
      });
    });

    it("should preserve whitespace-only text in leaf elements", () => {
      const doc = parseXml("<pre>   </pre>");
      expect(toPlainObject(doc.root)).toEqual({ pre: "   " });
    });

    it("should preserve whitespace-only text in leaf with attributes", () => {
      const doc = parseXml('<space xml:space="preserve"> \t </space>');
      expect(toPlainObject(doc.root)).toEqual({
        space: { "@_xml:space": "preserve", "#text": " \t " }
      });
    });
  });
});
