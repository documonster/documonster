/**
 * Unit tests for the SAX XML parser (parse-sax.ts)
 */

import { test } from "@excel/utils/__tests__/parse-sax.test-utils";
import { describe } from "vitest";

describe("SaxParser", () => {
  describe("basic parsing", () => {
    test({
      name: "simple element",
      xml: "<root></root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "element with text",
      xml: "<root>hello world</root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["text", "hello world"],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "nested elements",
      xml: "<root><child>text</child></root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["text", "text"],
        ["closetag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "multiple children",
      xml: "<root><a>1</a><b>2</b><c>3</c></root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["text", "1"],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["text", "2"],
        ["closetag", { name: "b", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "c", attributes: {}, isSelfClosing: false }],
        ["text", "3"],
        ["closetag", { name: "c", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("self-closing tags", () => {
    test({
      name: "self-closing tag",
      xml:
        "<root>   " +
        "<haha /> " +
        "<haha/>  " +
        "<monkey> " +
        "=(|)    " +
        "</monkey>" +
        "</root>  ",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["text", "   "],
        ["opentag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["text", " "],
        ["opentag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["text", "  "],
        ["opentag", { name: "monkey", attributes: {}, isSelfClosing: false }],
        ["text", " =(|)    "],
        ["closetag", { name: "monkey", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["text", "  "]
      ]
    });

    test({
      name: "self-closing child",
      xml:
        "<root>" +
        "<child>" +
        "<haha />" +
        "</child>" +
        "<monkey>" +
        "=(|)" +
        "</monkey>" +
        "</root>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "haha", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "child", attributes: {}, isSelfClosing: false }],
        ["opentag", { name: "monkey", attributes: {}, isSelfClosing: false }],
        ["text", "=(|)"],
        ["closetag", { name: "monkey", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "simple self-closing",
      xml: "<br/>",
      expect: [
        ["opentag", { name: "br", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "br", attributes: {}, isSelfClosing: true }]
      ]
    });

    test({
      name: "self-closing with space",
      xml: "<br />",
      expect: [
        ["opentag", { name: "br", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "br", attributes: {}, isSelfClosing: true }]
      ]
    });
  });

  describe("attributes", () => {
    test({
      name: "single attribute",
      xml: '<root attr="value"></root>',
      expect: [
        ["opentag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "multiple attributes",
      xml: '<root a="1" b="2" c="3"></root>',
      expect: [
        ["opentag", { name: "root", attributes: { a: "1", b: "2", c: "3" }, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: { a: "1", b: "2", c: "3" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "single quotes",
      xml: "<root attr='value'></root>",
      expect: [
        ["opentag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: { attr: "value" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "mixed quotes",
      xml: `<root a="1" b='2' c="3"></root>`,
      expect: [
        ["opentag", { name: "root", attributes: { a: "1", b: "2", c: "3" }, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: { a: "1", b: "2", c: "3" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "attributes separated by a space",
      xml: '<root attr1="first" attr2="second"/>',
      expect: [
        [
          "opentag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ]
      ]
    });

    test({
      name: "attributes separated by a newline",
      xml: '<root attr1="first"\nattr2="second"/>',
      expect: [
        [
          "opentag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ]
      ]
    });

    test({
      name: "attributes separated by multiple spaces",
      xml: '<root attr1="first"   attr2="second"/>',
      expect: [
        [
          "opentag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ]
      ]
    });

    test({
      name: "attribute with empty value",
      xml: '<root attr=""></root>',
      expect: [
        ["opentag", { name: "root", attributes: { attr: "" }, isSelfClosing: false }],
        ["closetag", { name: "root", attributes: { attr: "" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "attributes on self-closing",
      xml: '<item id="123" name="test"/>',
      expect: [
        ["opentag", { name: "item", attributes: { id: "123", name: "test" }, isSelfClosing: true }],
        ["closetag", { name: "item", attributes: { id: "123", name: "test" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "duplicate attribute",
      xml: '<span id="hello" id="there"></span>',
      expect: [
        ["error", "1:28: duplicate attribute: id"],
        ["opentag", { name: "span", attributes: { id: "there" }, isSelfClosing: false }],
        ["closetag", { name: "span", attributes: { id: "there" }, isSelfClosing: false }]
      ]
    });

    test({
      name: "attributes without a space (should error but still parse)",
      xml: '<root attr1="first"attr2="second"/>',
      expect: [
        ["error", "1:20: no whitespace between attributes"],
        [
          "opentag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ],
        [
          "closetag",
          { name: "root", attributes: { attr1: "first", attr2: "second" }, isSelfClosing: true }
        ]
      ]
    });
  });

  describe("entities", () => {
    test({
      name: "built-in entities",
      xml: "<r>&amp; &lt; &gt; ></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "& < > >"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "numeric decimal entity",
      xml: "<r>&#65;&#66;&#67;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "ABC"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "numeric hex entity",
      xml: "<r>&#x41;&#x42;&#x43;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "ABC"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "numeric hex entity lowercase",
      xml: "<r>&#xa;&#xA;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "\n\n"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "entity in attribute",
      xml: '<r attr="&amp;&lt;&gt;&quot;&apos;"/>',
      expect: [
        ["opentag", { name: "r", attributes: { attr: "&<>\"'" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "&<>\"'" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "single quote in double-quoted attribute",
      xml: '<r attr="it&apos;s"/>',
      expect: [
        ["opentag", { name: "r", attributes: { attr: "it's" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: "it's" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "double quote in single-quoted attribute",
      xml: "<r attr='&quot;quoted&quot;'/>",
      expect: [
        ["opentag", { name: "r", attributes: { attr: '"quoted"' }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { attr: '"quoted"' }, isSelfClosing: true }]
      ]
    });

    test({
      name: "emoji via numeric entity",
      xml: "<a>&#x1f525;</a>",
      expect: [
        ["opentag", { name: "a", attributes: {}, isSelfClosing: false }],
        ["text", "\ud83d\udd25"],
        ["closetag", { name: "a", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("bad entities", () => {
    test({
      name: "empty entity",
      xml: "<r>&;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["error", "1:5: empty entity"],
        ["text", "&;"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "empty decimal entity",
      xml: "<r>&#;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["error", "1:6: invalid character entity: &#;"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "empty hex entity",
      xml: "<r>&#x;</r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["error", "1:7: invalid character entity: &#x;"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("CDATA", () => {
    test({
      name: "cdata basic",
      xml: "<r><![CDATA[ this is character data  ]]></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", " this is character data  "],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "cdata empty",
      xml: "<r><![CDATA[]]></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "cdata with special chars",
      xml: "<r><![CDATA[<>&\"']]></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", "<>&\"'"],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "cdata end in attribute",
      xml: "<r foo=']]>'/>",
      expect: [
        ["opentag", { name: "r", attributes: { foo: "]]>" }, isSelfClosing: true }],
        ["closetag", { name: "r", attributes: { foo: "]]>" }, isSelfClosing: true }]
      ]
    });

    test({
      name: "cdata surrounded by whitespace",
      xml: `<content:encoded>
          <![CDATA[spacetime is four dimensional]]>
  </content:encoded>`,
      expect: [
        ["opentag", { name: "content:encoded", attributes: {}, isSelfClosing: false }],
        ["text", "\n          "],
        ["text", "spacetime is four dimensional"],
        ["text", "\n  "],
        ["closetag", { name: "content:encoded", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "cdata chunked",
      xml: ["<r><![CDATA[ this is ", "character data  ", "]]></r>"],
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["text", " this is character data  "],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("comments", () => {
    test({
      name: "comment basic",
      xml: "<r><!--foo--></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "comment empty",
      xml: "<r><!----></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "comment with dashes",
      xml: "<r><!-- foo - bar - baz --></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });

    test({
      name: "multiple comments",
      xml: "<r><!--a--><!--b--></r>",
      expect: [
        ["opentag", { name: "r", attributes: {}, isSelfClosing: false }],
        ["closetag", { name: "r", attributes: {}, isSelfClosing: false }]
      ]
    });
  });

  describe("processing instructions", () => {
    test({
      name: "xml declaration",
      xml: '<?xml version="1.0"?><root/>',
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: true }]
      ]
    });

    test({
      name: "xml declaration with encoding",
      xml: '<?xml version="1.0" encoding="UTF-8"?><root/>',
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: true }]
      ]
    });

    test({
      name: "processing instruction",
      xml: "<?foo bar?><root/>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: true }]
      ]
    });
  });

  describe("DOCTYPE", () => {
    test({
      name: "simple doctype",
      xml: "<!DOCTYPE root><root/>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: true }]
      ]
    });

    test({
      name: "doctype with internal subset",
      xml: "<!DOCTYPE root [<!ELEMENT root (#PCDATA)>]><root/>",
      expect: [
        ["opentag", { name: "root", attributes: {}, isSelfClosing: true }],
        ["closetag", { name: "root", attributes: {}, isSelfClosing: true }]
      ]
    });
  });
});
