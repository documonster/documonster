/**
 * Misc parser behavior tests for parse-sax.ts
 */

import { describe } from "vitest";
import { test } from "@excel/utils/__tests__/parse-sax.test-utils";

describe("SaxParser", () => {
  describe("unicode", () => {
    test({
      name: "cyrillic",
      xml: "<Р>тест</Р>",
      expect: [
        ["opentag", { name: "Р", attributes: {}, isSelfClosing: false }],
        ["text", "тест"],
        ["closetag", { name: "Р", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "emoji direct",
      xml: "<a>💩</a>",
      expect: [
        ["opentag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["text", "💩"],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "chinese characters",
      xml: "<中文>你好世界</中文>",
      expect: [
        ["opentag", { name: "中文", attributes: {}, isSelfClosing: false }],
        ["text", "你好世界"],
        ["closetag", { name: "中文", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "unicode sliced",
      xml: ["<a>💩", "</a>"],
      expect: [
        ["opentag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["text", "💩"],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("BOM handling", () => {
    test({
      name: "BOM at start is stripped (XML spec)",
      xml: "\uFEFF<P></P>",
      expect: [
        ["opentag", { name: "P", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "P", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "BOM in contents preserved, BOM at start stripped",
      xml: '\uFEFF<P BOM="\uFEFF">\uFEFFStarts and ends with BOM\uFEFF</P>',
      expect: [
        ["opentag", { name: "P", attributes: { BOM: "\uFEFF" }, isSelfClosing: false }],
        ["text", "\uFEFFStarts and ends with BOM\uFEFF"],
        ["closetag", { name: "P", attributes: { BOM: "\uFEFF" }, isSelfClosing: false }]
      ]
    });
  });

  describe("errors", () => {
    test({
      name: "unclosed root",
      xml: "<root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["error", "1:6: unclosed tag: root"]
      ]
    });

    test({
      name: "unclosed root without position",
      xml: "<doc>",
      expect: [
        ["opentag", { name: "doc", attributes: {}, isSelfClosing: false }],
        ["error", "unclosed tag: doc"]
      ],
      opt: { position: false }
    });

    test({
      name: "unclosed root with fileName",
      xml: "<doc>",
      expect: [
        ["opentag", { name: "doc", attributes: {}, isSelfClosing: false }],
        ["error", "foobar.xml:1:5: unclosed tag: doc"]
      ],
      opt: { fileName: "foobar.xml" }
    });

    test({
      name: "unclosed nested tag",
      xml: "<root><child></root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["error", "1:20: unclosed tag: child"],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "trailing non-whitespace text",
      xml: "<span>Welcome,</span> to monkey land",
      expect: [
        ["opentag", { name: "span", attributes: {}, isSelfClosing: false }],
        ["text", "Welcome,"],
        ["closetag", { name: "span", attributes: {}, isSelfClosing: false }],
        ["error", "1:36: text data outside of root node."],
        ["text", " to monkey land"]
      ]
    });
  });

  describe("EOL handling", () => {
    test({
      name: "LF normalization",
      xml: "<r>line1\nline2</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "line1\nline2"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "CR normalization",
      xml: "<r>line1\rline2</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "line1\nline2"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "CRLF normalization",
      xml: "<r>line1\r\nline2</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "line1\nline2"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "attribute CR normalization",
      xml: '<r attr="line1\rline2"/>',
      expect: [
        ["opentag", { name: "r", attributes: { attr: "line1 line2" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "line1 line2" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "attribute LF normalization",
      xml: '<r attr="line1\nline2"/>',
      expect: [
        ["opentag", { name: "r", attributes: { attr: "line1 line2" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "line1 line2" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "attribute TAB normalization",
      xml: '<r attr="a\tb"/>',
      expect: [
        ["opentag", { name: "r", attributes: { attr: "a b" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "a b" }, isSelfClosing: true }]
      ]
    });
  });

  describe("chunked parsing", () => {
    test({
      name: "tag split across chunks",
      xml: ["<roo", "t>text</root>"],
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["text", "text"],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "attribute split across chunks",
      xml: ["<root att", 'r="value">text</root>'],
      expect: [
        ["opentag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }],
        ["text", "text"],
        ["closetag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "attribute value split across chunks",
      xml: ['<root attr="val', 'ue">text</root>'],
      expect: [
        ["opentag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }],
        ["text", "text"],
        ["closetag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "entity split across chunks",
      xml: ["<r>&am", "p;</r>"],
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "&"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "char-by-char parsing",
      xml: "<r>text</r>".split(""),
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "text"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("namespace-prefixed tags", () => {
    test({
      name: "prefixed element",
      xml: "<ns:root><ns:child/></ns:root>",
      expect: [
        ["opentag", { name: "ns:root", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "ns:child", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "ns:child", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "ns:root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "prefixed attribute",
      xml: '<root ns:attr="value"/>',
      expect: [
        ["opentag", { name: "root", attributes: { "ns:attr": "value" }, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: { "ns:attr": "value" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "xmlns declaration (treated as normal attribute)",
      xml: '<root xmlns="http://example.com"/>',
      expect: [
        [
          "opentag",
          { name: "root", attributes: { xmlns: "http://example.com" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "root", attributes: { xmlns: "http://example.com" }, isSelfClosing: true }
        ]
      ]
    });

    test({
      name: "xmlns prefix declaration",
      xml: '<ns:root xmlns:ns="http://example.com"/>',
      expect: [
        [
          "opentag",
          { name: "ns:root", attributes: { "xmlns:ns": "http://example.com" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "ns:root", attributes: { "xmlns:ns": "http://example.com" }, isSelfClosing: true }
        ]
      ]
    });
  });

  describe("Excel XML patterns", () => {
    test({
      name: "worksheet cell",
      xml: '<c r="A1" s="1"><v>42</v></c>',
      expect: [
        ["opentag", { name: "c", attributes: { r: "A1", s: "1" }, isSelfClosing: false }],
        ["opentag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["text", "42"],
        ["closetag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "c", attributes: { r: "A1", s: "1" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "shared string",
      xml: "<si><t>Hello World</t></si>",
      expect: [
        ["opentag", { name: "si", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "t", attributes: {}, isSelfClosing: false }],
        ["text", "Hello World"],
        ["closetag", { name: "t", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "si", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "formula cell",
      xml: '<c r="B1" s="2"><f>A1+1</f><v>43</v></c>',
      expect: [
        ["opentag", { name: "c", attributes: { r: "B1", s: "2" }, isSelfClosing: false }],
        ["opentag", { name: "f", attributes: {}, isSelfClosing: false }],
        ["text", "A1+1"],
        ["closetag", { name: "f", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["text", "43"],
        ["closetag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "c", attributes: { r: "B1", s: "2" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "row with multiple cells",
      xml: '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>',
      expect: [
        ["opentag", { name: "row", attributes: { r: "1" }, isSelfClosing: false }],
        ["opentag", { name: "c", attributes: { r: "A1" }, isSelfClosing: false }],
        ["opentag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["text", "1"],
        ["closetag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "c", attributes: { r: "A1" }, isSelfClosing: false }],
        ["opentag", { name: "c", attributes: { r: "B1" }, isSelfClosing: false }],
        ["opentag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["text", "2"],
        ["closetag", { name: "v", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "c", attributes: { r: "B1" }, isSelfClosing: false }],
        ["closetag", { name: "row", attributes: { r: "1" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "style element",
      xml: '<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>',
      expect: [
        ["opentag", { name: "font", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "sz", attributes: { val: "11" }, isSelfClosing: true }],
        ["closetag", { name: "sz", attributes: { val: "11" }, isSelfClosing: true }],
        ["opentag", { name: "color", attributes: { theme: "1" }, isSelfClosing: true }],
        ["closetag", { name: "color", attributes: { theme: "1" }, isSelfClosing: true }],
        ["opentag", { name: "name", attributes: { val: "Calibri" }, isSelfClosing: true }],
        ["closetag", { name: "name", attributes: { val: "Calibri" }, isSelfClosing: true }],
        ["closetag", { name: "font", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("fragment mode", () => {
    test({
      name: "fragment - text only",
      xml: "just some text",
      expect: [["text", "just some text"]],
      opt: { fragment: true }
    });

    test({
      name: "fragment - multiple roots",
      xml: "<a/><b/><c/>",
      expect: [
        ["opentag", { name: "a", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: true }],
        ["opentag", { name: "b", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "b", attributes: {}, isSelfClosing: true }],
        ["opentag", { name: "c", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "c", attributes: {}, isSelfClosing: true }]
      ],
      opt: { fragment: true }
    });

    test({
      name: "fragment - mixed content",
      xml: "text1<tag>text2</tag>text3",
      expect: [
        ["text", "text1"],
        ["opentag", { name: "tag", attributes: {}, isSelfClosing: false }],
        ["text", "text2"],
        ["closetag", { name: "tag", attributes: {}, isSelfClosing: false }],
        ["text", "text3"]
      ],
      opt: { fragment: true }
    });
  });

  describe("edge cases", () => {
    test({
      name: "empty document (just root)",
      xml: "<r/>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: true }]
      ]
    });

    test({
      name: "very long attribute value",
      xml: `<r attr="${"a".repeat(10000)}"/>`,
      expect: [
        ["opentag", { name: "r", attributes: { attr: "a".repeat(10000) }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "a".repeat(10000) }, isSelfClosing: true }]
      ]
    });

    test({
      name: "very long text content",
      xml: `<r>${"x".repeat(10000)}</r>`,
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "x".repeat(10000)],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "deeply nested elements",
      xml: "<a><b><c><d><e>deep</e></d></c></b></a>",
      expect: [
        ["opentag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "c", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "d", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "e", attributes: {}, isSelfClosing: false }],
        ["text", "deep"],
        ["closetag", { name: "e", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "d", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "c", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "mixed content (text and elements)",
      xml: "<p>Hello <b>world</b>!</p>",
      expect: [
        ["opentag", { name: "p", attributes: {}, isSelfClosing: false }],
        ["text", "Hello "],
        ["opentag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["text", "world"],
        ["closetag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["text", "!"],
        ["closetag", { name: "p", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "whitespace-only text",
      xml: "<r>   \n\t  </r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "   \n\t  "],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "numeric tag name",
      xml: "<_123>test</_123>",
      expect: [
        ["opentag", { name: "_123", attributes: {}, isSelfClosing: false }],
        ["text", "test"],
        ["closetag", { name: "_123", attributes: {}, isSelfClosing: false }]
      ]
    });
  });
});
