import { describe, it, expect } from "vitest";
import { parseXml, toPlainObject } from "@xml/dom";
import { parseXmlToObject } from "@xml/to-object";

// =============================================================================
// Helper: verify parseXmlToObject matches parseXml + toPlainObject
// =============================================================================

function expectSameOutput(xml: string, options?: Parameters<typeof toPlainObject>[1]) {
  const doc = parseXml(xml);
  const fromDom = toPlainObject(doc.root, options);
  const direct = parseXmlToObject(xml, options);
  expect(direct).toEqual(fromDom);
}

// =============================================================================
// Tests
// =============================================================================

describe("parseXmlToObject", () => {
  describe("basic conversion", () => {
    it("should convert empty element", () => {
      const obj = parseXmlToObject("<root/>");
      expect(obj).toEqual({ root: "" });
    });

    it("should convert text-only element", () => {
      const obj = parseXmlToObject("<root>hello</root>");
      expect(obj).toEqual({ root: "hello" });
    });

    it("should convert element with attributes", () => {
      const obj = parseXmlToObject('<root id="1" name="test"/>');
      expect(obj).toEqual({ root: { "@_id": "1", "@_name": "test" } });
    });

    it("should convert element with attributes and text", () => {
      const obj = parseXmlToObject('<root id="1">hello</root>');
      expect(obj).toEqual({ root: { "@_id": "1", "#text": "hello" } });
    });
  });

  describe("nested elements", () => {
    it("should convert single child element", () => {
      const obj = parseXmlToObject("<root><child>text</child></root>");
      expect(obj).toEqual({ root: { child: "text" } });
    });

    it("should convert deeply nested structure", () => {
      const obj = parseXmlToObject("<a><b><c>deep</c></b></a>");
      expect(obj).toEqual({ a: { b: { c: "deep" } } });
    });

    it("should merge repeated siblings into arrays", () => {
      const obj = parseXmlToObject("<list><item>a</item><item>b</item><item>c</item></list>");
      expect(obj).toEqual({ list: { item: ["a", "b", "c"] } });
    });

    it("should handle mix of unique and repeated children", () => {
      const obj = parseXmlToObject("<root><name>test</name><tag>a</tag><tag>b</tag></root>");
      expect(obj).toEqual({ root: { name: "test", tag: ["a", "b"] } });
    });
  });

  describe("attributes on nested elements", () => {
    it("should preserve attributes on child elements", () => {
      const obj = parseXmlToObject('<root><child id="1">text</child></root>');
      expect(obj).toEqual({ root: { child: { "@_id": "1", "#text": "text" } } });
    });

    it("should handle attributes at multiple levels", () => {
      const obj = parseXmlToObject('<root attr="r"><child attr="c"/></root>');
      expect(obj).toEqual({
        root: { "@_attr": "r", child: { "@_attr": "c" } }
      });
    });
  });

  describe("options", () => {
    it("should use custom attribute prefix", () => {
      const obj = parseXmlToObject('<root id="1"/>', { attributePrefix: "$" });
      expect(obj).toEqual({ root: { $id: "1" } });
    });

    it("should use empty attribute prefix", () => {
      const obj = parseXmlToObject('<root id="1"/>', { attributePrefix: "" });
      expect(obj).toEqual({ root: { id: "1" } });
    });

    it("should use custom text key", () => {
      const obj = parseXmlToObject('<root id="1">hello</root>', { textKey: "_text" });
      expect(obj).toEqual({ root: { "@_id": "1", _text: "hello" } });
    });

    it("should always wrap in arrays when alwaysArray is true", () => {
      const obj = parseXmlToObject("<root><child>text</child></root>", { alwaysArray: true });
      expect(obj).toEqual({ root: { child: ["text"] } });
    });

    it("alwaysArray should still work with repeated siblings", () => {
      const obj = parseXmlToObject("<root><item>a</item><item>b</item></root>", {
        alwaysArray: true
      });
      expect(obj).toEqual({ root: { item: ["a", "b"] } });
    });

    it("isArray callback should wrap matching single children", () => {
      const obj = parseXmlToObject("<root><item>a</item><single>b</single></root>", {
        isArray: name => name === "item"
      });
      expect(obj).toEqual({ root: { item: ["a"], single: "b" } });
    });

    it("isArray should not affect repeated siblings", () => {
      const obj = parseXmlToObject("<root><item>a</item><item>b</item></root>", {
        isArray: name => name === "item"
      });
      expect(obj).toEqual({ root: { item: ["a", "b"] } });
    });

    it("isArray combined with alwaysArray false", () => {
      const obj = parseXmlToObject("<root><item>a</item><other>b</other></root>", {
        alwaysArray: false,
        isArray: name => name === "item"
      });
      expect(obj).toEqual({ root: { item: ["a"], other: "b" } });
    });

    it("document root is never wrapped by isArray", () => {
      const obj = parseXmlToObject("<root><child>text</child></root>", {
        isArray: () => true
      });
      expect(obj).toEqual({ root: { child: ["text"] } });
    });

    it("isArray should work on nested elements", () => {
      const obj = parseXmlToObject("<root><parent><child>a</child></parent></root>", {
        isArray: name => name === "child"
      });
      expect(obj).toEqual({ root: { parent: { child: ["a"] } } });
    });
  });

  describe("ignoreAttributes", () => {
    it("should discard all attributes when ignoreAttributes is true", () => {
      const obj = parseXmlToObject('<root id="1" name="test"><child attr="v">text</child></root>', {
        ignoreAttributes: true
      });
      expect(obj).toEqual({ root: { child: "text" } });
    });

    it("should collapse to string when ignoreAttributes removes all non-text content", () => {
      const obj = parseXmlToObject('<root id="1">hello</root>', { ignoreAttributes: true });
      expect(obj).toEqual({ root: "hello" });
    });

    it("should return empty string for attribute-only element", () => {
      const obj = parseXmlToObject('<root id="1"/>', { ignoreAttributes: true });
      expect(obj).toEqual({ root: "" });
    });
  });

  describe("whitespace handling", () => {
    it("should ignore whitespace-only text nodes by default", () => {
      const obj = parseXmlToObject("<root>\n  <child>text</child>\n</root>");
      expect(obj).toEqual({ root: { child: "text" } });
    });

    it("should preserve whitespace-only text when ignoreWhitespaceText is false", () => {
      const obj = parseXmlToObject("<root>\n  <child>text</child>\n</root>", {
        ignoreWhitespaceText: false
      });
      expect(obj).toEqual({ root: { "#text": "\n  \n", child: "text" } });
    });

    it("should preserve non-whitespace text even when ignoreWhitespaceText is true", () => {
      const obj = parseXmlToObject("<root> meaningful text </root>");
      expect(obj).toEqual({ root: " meaningful text " });
    });

    it("should preserve whitespace-only text in leaf elements", () => {
      const obj = parseXmlToObject("<pre>   </pre>");
      expect(obj).toEqual({ pre: "   " });
    });

    it("should handle pretty-printed nested XML", () => {
      const xml = ["<root>", "  <a>", "    <b>deep</b>", "  </a>", "</root>"].join("\n");
      const obj = parseXmlToObject(xml);
      expect(obj).toEqual({ root: { a: { b: "deep" } } });
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

      const obj = parseXmlToObject(xml);
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
      const obj = parseXmlToObject("<p>before<b>bold</b>after</p>");
      expect(obj).toEqual({ p: { "#text": "beforeafter", b: "bold" } });
    });

    it("should handle element with only empty children", () => {
      const obj = parseXmlToObject("<root><a/><b/></root>");
      expect(obj).toEqual({ root: { a: "", b: "" } });
    });

    it("should throw on malformed XML", () => {
      expect(() => parseXmlToObject("")).toThrow();
    });
  });

  // ===========================================================================
  // Equivalence: parseXmlToObject must match parseXml + toPlainObject
  // ===========================================================================

  describe("equivalence with toPlainObject", () => {
    it("empty element", () => expectSameOutput("<root/>"));
    it("text element", () => expectSameOutput("<root>hello</root>"));
    it("attributes", () => expectSameOutput('<root id="1" name="test"/>'));
    it("nested", () => expectSameOutput("<a><b><c>deep</c></b></a>"));
    it("repeated siblings", () => expectSameOutput("<list><item>a</item><item>b</item></list>"));

    it("pretty-printed", () =>
      expectSameOutput(["<root>", "  <a>", "    <b>deep</b>", "  </a>", "</root>"].join("\n")));

    it("with alwaysArray", () =>
      expectSameOutput("<root><child>text</child></root>", { alwaysArray: true }));

    it("with custom prefix", () => expectSameOutput('<root id="1"/>', { attributePrefix: "$" }));

    it("with ignoreWhitespaceText false", () =>
      expectSameOutput("<root>\n  <child/>\n</root>", { ignoreWhitespaceText: false }));

    it("complex Excel-like XML", () =>
      expectSameOutput(
        [
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          "  <sheetData>",
          '    <row r="1">',
          '      <c r="A1" t="s"><v>0</v></c>',
          '      <c r="B1" t="n"><v>42</v></c>',
          "    </row>",
          "  </sheetData>",
          "</worksheet>"
        ].join("\n")
      ));

    it("with isArray callback", () =>
      expectSameOutput("<root><item>a</item><single>b</single></root>", {
        isArray: name => name === "item"
      }));

    it("with isArray on nested elements", () =>
      expectSameOutput("<root><parent><child>a</child></parent></root>", {
        isArray: name => name === "child"
      }));

    it("with ignoreAttributes", () =>
      expectSameOutput('<root id="1"><child attr="v">text</child></root>', {
        ignoreAttributes: true
      }));
  });
});
