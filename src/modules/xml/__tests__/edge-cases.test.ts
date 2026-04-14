/**
 * XML Module Edge Case Tests
 *
 * Comprehensive edge case coverage for:
 * - Entity handling boundaries
 * - Unicode edge cases (surrogates, BOM, CJK)
 * - CDATA edge cases
 * - Malformed XML error recovery
 * - Large/deep structures
 * - Attribute edge cases
 * - Namespace edge cases
 * - Writer edge cases
 */

import { parseXml, findChild, findChildren, textContent, attr, walk } from "@xml/dom";
import { xmlEncode, xmlDecode } from "@xml/encode";
import { query, queryAll } from "@xml/query";
import { SaxParser, parseSax } from "@xml/sax";
import { XmlStreamWriter } from "@xml/stream-writer";
import type { SaxTag, WritableTarget } from "@xml/types";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

// =============================================================================
// Entity Handling Edge Cases
// =============================================================================

describe("Entity handling edge cases", () => {
  it("should handle entity in attribute value", () => {
    const doc = parseXml('<root attr="a &amp; b"/>');
    expect(attr(doc.root, "attr")).toBe("a & b");
  });

  it("should NOT expand entities inside CDATA", () => {
    const doc = parseXml("<root><![CDATA[&amp; &lt;]]></root>");
    expect(textContent(doc.root)).toBe("&amp; &lt;");
  });

  it("should handle empty entity reference &;", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&;</root>");
    parser.close();
    expect(errors.some(e => e.includes("empty entity"))).toBe(true);
  });

  it("should handle undefined entity gracefully", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    parser.write("<root>&nonexistent;</root>");
    parser.close();
    expect(errors.some(e => e.includes("undefined entity"))).toBe(true);
  });

  it("should handle all 5 predefined entities", () => {
    const doc = parseXml("<r>&lt;&gt;&amp;&quot;&apos;</r>");
    expect(textContent(doc.root)).toBe("<>&\"'");
  });

  it("should handle mixed text and entities", () => {
    const doc = parseXml("<r>hello &amp; &lt;world&gt;</r>");
    expect(textContent(doc.root)).toBe("hello & <world>");
  });

  it("should handle hex entity case insensitivity", () => {
    const doc = parseXml("<r>&#x41;&#X42;&#x61;</r>");
    expect(textContent(doc.root)).toBe("ABa");
  });

  it("should handle entity at end of text", () => {
    const doc = parseXml("<r>text&amp;</r>");
    expect(textContent(doc.root)).toBe("text&");
  });

  it("should handle entity at start of text", () => {
    const doc = parseXml("<r>&amp;text</r>");
    expect(textContent(doc.root)).toBe("&text");
  });

  it("should handle consecutive entities", () => {
    const doc = parseXml("<r>&lt;&gt;&amp;</r>");
    expect(textContent(doc.root)).toBe("<>&");
  });
});

// =============================================================================
// Unicode Edge Cases
// =============================================================================

describe("Unicode edge cases", () => {
  it("should handle surrogate pairs in text content", () => {
    // U+1F600 = 😀 (surrogate pair in UTF-16)
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>\uD83D\uDE00</root>");
    parser.close();
    expect(texts[0]).toBe("😀");
  });

  it("should handle CJK characters in tag names", () => {
    const doc = parseXml("<根><子>text</子></根>");
    expect(doc.root.name).toBe("根");
    expect(findChild(doc.root, "子")).toBeDefined();
  });

  it("should handle CJK characters in attribute values", () => {
    const doc = parseXml('<root name="中文"/>');
    expect(attr(doc.root, "name")).toBe("中文");
  });

  it("should handle emoji in text", () => {
    const doc = parseXml("<root>🎉🎊</root>");
    expect(textContent(doc.root)).toBe("🎉🎊");
  });

  it("should handle high unicode numeric entities", () => {
    const doc = parseXml("<root>&#x1F600;</root>"); // 😀
    expect(textContent(doc.root)).toBe("😀");
  });

  it("should encode and decode round-trip for special chars", () => {
    const original = '<tag attr="v&al">';
    const encoded = xmlEncode(original);
    const decoded = xmlDecode(encoded);
    expect(decoded).toBe(original);
  });

  it("should handle empty string encoding/decoding", () => {
    expect(xmlEncode("")).toBe("");
    expect(xmlDecode("")).toBe("");
  });

  it("should strip invalid control characters in encoding", () => {
    expect(xmlEncode("hello\x00world")).toBe("helloworld");
    expect(xmlEncode("a\x01b\x08c")).toBe("abc");
    expect(xmlEncode("a\x7Fb")).toBe("ab");
  });

  it("should preserve valid whitespace in encoding", () => {
    expect(xmlEncode("a\tb\nc\rd")).toBe("a\tb\nc\rd");
  });
});

// =============================================================================
// CDATA Edge Cases
// =============================================================================

describe("CDATA edge cases", () => {
  it("should handle empty CDATA", () => {
    const doc = parseXml("<root><![CDATA[]]></root>");
    // Empty CDATA produces empty text node (CDATA handler still fires with "")
    expect(doc.root.children.length).toBeLessThanOrEqual(1);
  });

  it("should handle CDATA with XML-like content", () => {
    const doc = parseXml("<root><![CDATA[<script>alert(1)</script>]]></root>");
    expect(textContent(doc.root)).toBe("<script>alert(1)</script>");
  });

  it("should handle CDATA with entity-like content", () => {
    const doc = parseXml("<root><![CDATA[&amp; &lt; &gt;]]></root>");
    expect(textContent(doc.root)).toBe("&amp; &lt; &gt;");
  });

  it("should handle multiple adjacent CDATA sections", () => {
    const doc = parseXml("<root><![CDATA[a]]><![CDATA[b]]></root>");
    expect(textContent(doc.root)).toBe("ab");
  });

  it("should handle CDATA as explicit nodes", () => {
    const doc = parseXml("<root><![CDATA[data]]></root>", { cdataAsNodes: true });
    expect(doc.root.children[0].type).toBe("cdata");
  });

  it("should handle CDATA with ]] inside", () => {
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    parser.write("<root><![CDATA[a]]b]]></root>");
    parser.close();
    expect(cdatas[0]).toBe("a]]b");
  });
});

// =============================================================================
// Malformed XML Error Handling
// =============================================================================

describe("Malformed XML error handling", () => {
  it("should report unclosed tag", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root><child>");
    parser.close();
    expect(errors.some(e => e.includes("unclosed tag"))).toBe(true);
  });

  it("should report mismatched tags", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<a></b>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });

  it("should handle truncated XML gracefully", () => {
    // Truncated XML should not crash — may or may not report errors depending on state
    expect(() => {
      const parser = new SaxParser();
      parser.on("error", () => {});
      parser.write('<root attr="val');
      parser.close();
    }).not.toThrow();
  });

  it("should handle multiple root elements (non-fragment)", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    // After first root closes, text between will trigger error
    parser.write("<a/>text<b/>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });

  it("should detect text outside root element", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    // After root closes, a second root element will have text data reported as outside root
    parser.write("text before <root/>");
    parser.close();
    expect(errors.some(e => e.includes("outside of root") || e.includes("text data"))).toBe(true);
  });

  it("should parseXml throw on empty input", () => {
    expect(() => parseXml("")).toThrow();
  });

  it("should handle invalid tag name start character", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<123/>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });
});

// =============================================================================
// Large/Deep Structures
// =============================================================================

describe("Large structures", () => {
  it("should handle 100 levels of nesting", () => {
    let xml = "";
    for (let i = 0; i < 100; i++) {
      xml += "<a>";
    }
    xml += "deep";
    for (let i = 0; i < 100; i++) {
      xml += "</a>";
    }
    const doc = parseXml(xml);
    expect(doc.root.name).toBe("a");
  });

  it("should handle many siblings", () => {
    let xml = "<root>";
    for (let i = 0; i < 10000; i++) {
      xml += `<item id="${i}"/>`;
    }
    xml += "</root>";
    const doc = parseXml(xml);
    const items = findChildren(doc.root, "item");
    expect(items.length).toBe(10000);
    expect(attr(items[9999], "id")).toBe("9999");
  });

  it("should handle many attributes on one element", () => {
    let xml = "<root";
    for (let i = 0; i < 1000; i++) {
      xml += ` a${i}="v${i}"`;
    }
    xml += "/>";
    const doc = parseXml(xml);
    expect(attr(doc.root, "a0")).toBe("v0");
    expect(attr(doc.root, "a999")).toBe("v999");
  });

  it("should handle long attribute values", () => {
    const longVal = "x".repeat(100000);
    const xml = `<root attr="${longVal}"/>`;
    const doc = parseXml(xml);
    expect(attr(doc.root, "attr")).toBe(longVal);
  });

  it("should handle long text content", () => {
    const longText = "x".repeat(100000);
    const xml = `<root>${longText}</root>`;
    const doc = parseXml(xml);
    expect(textContent(doc.root)).toBe(longText);
  });
});

// =============================================================================
// Attribute Edge Cases
// =============================================================================

describe("Attribute edge cases", () => {
  it("should handle single-quoted attributes", () => {
    const parser = new SaxParser();
    let attrs: Record<string, string> = {};
    parser.on("opentag", tag => {
      attrs = tag.attributes;
    });
    parser.write("<root attr='value'/>");
    parser.close();
    expect(attrs.attr).toBe("value");
  });

  it("should handle empty attribute value", () => {
    const doc = parseXml('<root attr=""/>');
    expect(attr(doc.root, "attr")).toBe("");
  });

  it("should handle attribute with entity", () => {
    const doc = parseXml('<root attr="a&amp;b"/>');
    expect(attr(doc.root, "attr")).toBe("a&b");
  });

  it("should handle attribute with numeric entity", () => {
    const doc = parseXml('<root attr="&#65;"/>');
    expect(attr(doc.root, "attr")).toBe("A");
  });
});

// =============================================================================
// Namespace Edge Cases
// =============================================================================

describe("Namespace edge cases", () => {
  it("should handle namespace undeclaration", () => {
    const parser = new SaxParser({ xmlns: true });
    const tags: SaxTag[] = [];
    parser.on("opentag", tag => tags.push({ ...tag }));
    parser.write('<root xmlns="http://default"><child xmlns="">text</child></root>');
    parser.close();
    expect(tags[0].uri).toBe("http://default");
    expect(tags[1].uri).toBe(""); // undeclared
  });

  it("should handle element with multiple colons as a name", () => {
    // Multi-colon QNames are invalid in namespace mode and produce an error,
    // but parsing continues with error recovery
    const parser = new SaxParser({ xmlns: true });
    const tags: SaxTag[] = [];
    const errors: string[] = [];
    parser.on("opentag", tag => tags.push({ ...tag }));
    parser.on("error", e => errors.push(e.message));
    parser.write("<a:b:c/>");
    parser.close();
    // Should parse without crashing; prefix is "a", local is "b:c"
    expect(tags.length).toBe(1);
    // Should report the multi-colon error
    expect(errors.some(e => e.includes("local part must not contain"))).toBe(true);
  });

  it("should handle very long namespace URI", () => {
    const longUri = "http://example.com/" + "a".repeat(10000);
    const doc = parseXml(`<root xmlns="${longUri}"/>`, { xmlns: true });
    expect(doc.root.uri).toBe(longUri);
  });
});

// =============================================================================
// Comment Edge Cases
// =============================================================================

describe("Comment edge cases", () => {
  it("should handle empty comment", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!----></root>");
    parser.close();
    expect(comments[0]).toBe("");
  });

  it("should handle comment with dashes", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    const errors: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.on("error", e => errors.push(e.message));
    parser.write("<root><!-- -- --></root>");
    parser.close();
    // Double dash in comment is technically malformed per XML spec
    expect(errors.length).toBeGreaterThan(0);
  });

  it("should include comments in DOM when option is set", () => {
    const doc = parseXml("<root><!-- comment --><child/></root>", { comments: true });
    expect(doc.root.children[0].type).toBe("comment");
    expect(doc.root.children[1].type).toBe("element");
  });

  it("should skip comments in DOM by default", () => {
    const doc = parseXml("<root><!-- comment --><child/></root>");
    expect(doc.root.children.length).toBe(1);
    expect(doc.root.children[0].type).toBe("element");
  });
});

// =============================================================================
// Processing Instruction Edge Cases
// =============================================================================

describe("Processing instruction edge cases", () => {
  it("should handle PI without body", () => {
    const parser = new SaxParser();
    const pis: Array<{ target: string; body: string }> = [];
    parser.on("pi", (target, body) => pis.push({ target, body }));
    parser.write("<?target?><root/>");
    parser.close();
    expect(pis[0]).toEqual({ target: "target", body: "" });
  });

  it("should handle PI with body", () => {
    const parser = new SaxParser();
    const pis: Array<{ target: string; body: string }> = [];
    parser.on("pi", (target, body) => pis.push({ target, body }));
    parser.write('<?xml-stylesheet type="text/xsl" href="style.xsl"?><root/>');
    parser.close();
    expect(pis[0].target).toBe("xml-stylesheet");
    expect(pis[0].body).toContain("text/xsl");
  });
});

// =============================================================================
// Writer Edge Cases
// =============================================================================

describe("Writer edge cases", () => {
  it("should handle boolean and number attribute values", () => {
    const w = new XmlWriter();
    w.openNode("root", { flag: true, count: 42 });
    w.closeNode();
    expect(w.xml).toBe('<root flag="true" count="42"/>');
  });

  it("should skip undefined attribute values", () => {
    const w = new XmlWriter();
    w.openNode("root", { a: "1", b: undefined, c: "3" });
    w.closeNode();
    expect(w.xml).toBe('<root a="1" c="3"/>');
  });

  it("should handle writeText with number 0", () => {
    const w = new XmlWriter();
    w.openNode("v");
    w.writeText(0);
    w.closeNode();
    expect(w.xml).toBe("<v>0</v>");
  });

  it("should handle deeply nested XML generation", () => {
    const w = new XmlWriter();
    for (let i = 0; i < 100; i++) {
      w.openNode("level");
    }
    w.writeText("deep");
    w.closeAll();
    expect(w.xml).toContain("<level>");
    expect(w.xml).toContain("deep");
    expect(w.xml).toContain("</level>");
  });

  it("should handle XmlStreamWriter parity for all content types", () => {
    const chunks: string[] = [];
    const target: WritableTarget = { write: (c: string | Uint8Array) => chunks.push(String(c)) };
    const sw = new XmlStreamWriter(target);
    const w = new XmlWriter();

    const ops = (s: any) => {
      s.openXml();
      s.openNode("root");
      s.writeCData("cdata content");
      s.writeComment(" comment ");
      s.leafNode("item", { id: "1" }, "text & more");
      s.writeRaw("<raw/>");
      s.closeNode();
    };

    ops(sw);
    ops(w);
    expect(chunks.join("")).toBe(w.xml);
  });
});

// =============================================================================
// Query Edge Cases
// =============================================================================

describe("Query edge cases", () => {
  const xml = "<root><a><b><c>text</c></b></a><d/></root>";

  it("should handle empty path gracefully", () => {
    const doc = parseXml(xml);
    // Empty path — implementation-defined, should not crash
    expect(() => queryAll(doc.root, "")).not.toThrow();
  });

  it("should handle path with only wildcard", () => {
    const doc = parseXml(xml);
    const result = queryAll(doc.root, "*");
    expect(result.length).toBe(2); // a and d
  });

  it("should handle path to non-existent deep child", () => {
    const doc = parseXml(xml);
    expect(query(doc.root, "a/b/c/d/e/f")).toBeUndefined();
  });

  it("should handle query on leaf element", () => {
    const doc = parseXml(xml);
    const c = query(doc.root, "a/b/c");
    expect(queryAll(c!, "*")).toEqual([]); // c has no child elements
  });

  it("should handle recursive descent on flat structure", () => {
    const doc = parseXml("<root><a/><b/><c/></root>");
    const result = queryAll(doc.root, "//b");
    expect(result.length).toBe(1);
    expect(result[0].name).toBe("b");
  });

  it("should handle index 0", () => {
    const doc = parseXml("<root><a/><a/><a/></root>");
    const result = query(doc.root, "a[0]");
    expect(result).toBeDefined();
    expect(result!.name).toBe("a");
  });

  it("should handle walk with single element", () => {
    const doc = parseXml("<root/>");
    const names: string[] = [];
    walk(doc.root, el => names.push(el.name));
    expect(names).toEqual(["root"]);
  });
});

// =============================================================================
// Chunked Parsing Edge Cases
// =============================================================================

describe("Chunked parsing edge cases", () => {
  it("should handle input split in middle of tag name", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.write("<roo");
    parser.write("t/>");
    parser.close();
    expect(tags).toEqual(["root"]);
  });

  it("should handle input split in middle of attribute value", () => {
    const parser = new SaxParser();
    let attrs: Record<string, string> = {};
    parser.on("opentag", tag => {
      attrs = tag.attributes;
    });
    parser.write('<root attr="hel');
    parser.write('lo world"/>');
    parser.close();
    expect(attrs.attr).toBe("hello world");
  });

  it("should handle input split in middle of entity", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>&am");
    parser.write("p;</root>");
    parser.close();
    expect(texts.join("")).toBe("&");
  });

  it("should handle input split in middle of CDATA", () => {
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    parser.write("<root><![CDA");
    parser.write("TA[content]]></root>");
    parser.close();
    expect(cdatas[0]).toBe("content");
  });

  it("should handle input split in middle of comment", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!-- com");
    parser.write("ment --></root>");
    parser.close();
    expect(comments[0]).toBe(" comment ");
  });

  it("should handle single-character chunks", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.on("closetag", tag => tags.push("/" + tag.name));
    const xml = "<root><child/></root>";
    for (const ch of xml) {
      parser.write(ch);
    }
    parser.close();
    expect(tags).toEqual(["root", "child", "/child", "/root"]);
  });
});

// =============================================================================
// parseSax async generator edge cases
// =============================================================================

describe("parseSax edge cases", () => {
  it("should handle empty async iterable", async () => {
    async function* empty(): AsyncGenerator<string> {
      // nothing
    }
    // Empty input will cause parser.close() which reports "no root element"
    // The parseSax generator should propagate this as a thrown error
    await expect(async () => {
      for await (const _batch of parseSax(empty())) {
        // consume
      }
    }).rejects.toThrow();
  });

  it("should handle async iterable of Uint8Array with BOM", async () => {
    const encoder = new TextEncoder();
    async function* chunks(): AsyncGenerator<Uint8Array> {
      yield encoder.encode("<root>text</root>");
    }
    const events: any[] = [];
    for await (const batch of parseSax(chunks())) {
      events.push(...batch);
    }
    expect(events.some((e: any) => e.eventType === "opentag" && e.value.name === "root")).toBe(
      true
    );
  });
});

// =============================================================================
// BOM Handling
// =============================================================================

describe("BOM handling", () => {
  it("should ignore UTF-8 BOM at start of input", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.write("\uFEFF<root/>");
    parser.close();
    expect(tags).toEqual(["root"]);
  });

  it("should ignore BOM when it arrives as first chunk", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.write("\uFEFF");
    parser.write("<root/>");
    parser.close();
    expect(tags).toEqual(["root"]);
  });

  it("should work via parseXml with BOM", () => {
    const doc = parseXml("\uFEFF<root><child/></root>");
    expect(doc.root.name).toBe("root");
  });

  it("should NOT strip BOM that appears mid-stream", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>before\uFEFFafter</root>");
    parser.close();
    expect(texts.join("")).toContain("\uFEFF");
  });
});

// =============================================================================
// Chunk Boundary Regression Tests
// =============================================================================

describe("Chunk boundary regressions", () => {
  it("should handle chunk split right after <", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.on("closetag", tag => tags.push("/" + tag.name));
    parser.write("<");
    parser.write("root/>");
    parser.close();
    expect(tags).toEqual(["root", "/root"]);
  });

  it("should handle chunk split inside <! (CDATA opening)", () => {
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    parser.write("<root><!");
    parser.write("[CDATA[content]]></root>");
    parser.close();
    expect(cdatas[0]).toBe("content");
  });

  it("should handle chunk split inside <!-- (comment opening)", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!");
    parser.write("-- comment --></root>");
    parser.close();
    expect(comments[0]).toBe(" comment ");
  });

  it("should handle chunk split between ]] and > in CDATA", () => {
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    parser.write("<root><![CDATA[content]]");
    parser.write("></root>");
    parser.close();
    expect(cdatas[0]).toBe("content");
  });

  it("should handle chunk split between ] and ]> in CDATA", () => {
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    parser.write("<root><![CDATA[content]");
    parser.write("]></root>");
    parser.close();
    expect(cdatas[0]).toBe("content");
  });

  it("should handle chunk split between -- and > in comment", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!-- comment --");
    parser.write("></root>");
    parser.close();
    expect(comments[0]).toBe(" comment ");
  });

  it("should handle chunk split between - and -> in comment", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!-- comment -");
    parser.write("-></root>");
    parser.close();
    expect(comments[0]).toBe(" comment ");
  });

  it("should handle chunk split between ? and > in PI", () => {
    const parser = new SaxParser();
    const pis: Array<{ target: string; body: string }> = [];
    parser.on("pi", (target, body) => pis.push({ target, body }));
    parser.write("<?target body?");
    parser.write("><root/>");
    parser.close();
    expect(pis[0].target).toBe("target");
    expect(pis[0].body).toBe("body");
  });

  it("should handle chunk split inside <!DOCTYPE", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.write("<!");
    parser.write("DOCTYPE html>");
    parser.write("<root/>");
    parser.close();
    expect(tags).toEqual(["root"]);
  });

  it("should not produce 'undefined' text on chunk boundary in CDATA end", () => {
    // Regression: charFromCode(-1) was producing "undefined" string
    const parser = new SaxParser();
    const cdatas: string[] = [];
    parser.on("cdata", t => cdatas.push(t));
    // Split right at the ] in ]]>
    parser.write("<root><![CDATA[hello]");
    parser.write("]></root>");
    parser.close();
    expect(cdatas.join("")).toBe("hello");
    expect(cdatas.join("")).not.toContain("undefined");
  });

  it("should not produce 'undefined' text on chunk boundary in comment end", () => {
    const parser = new SaxParser();
    const comments: string[] = [];
    parser.on("comment", c => comments.push(c));
    parser.write("<root><!-- hello -");
    parser.write("-></root>");
    parser.close();
    expect(comments.join("")).toBe(" hello ");
    expect(comments.join("")).not.toContain("undefined");
  });

  it("should handle trailing CR carried across chunks", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    // Chunk ends with CR, next chunk starts with LF
    parser.write("<root>line1\r");
    parser.write("\nline2</root>");
    parser.close();
    expect(texts.join("")).toBe("line1\nline2");
  });

  it("should handle synchronous iterable in parseSax", async () => {
    const chunks = ["<root>", "<item/>", "</root>"];
    const events: any[] = [];
    for await (const batch of parseSax(chunks)) {
      events.push(...batch);
    }
    const names = events
      .filter((e: any) => e.eventType === "opentag")
      .map((e: any) => e.value.name);
    expect(names).toEqual(["root", "item"]);
  });

  it("should handle chunk boundary between / and > in self-closing tag", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    const errors: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.on("error", e => errors.push(e.message));
    parser.write("<root/");
    parser.write(">");
    parser.close();
    expect(tags).toEqual(["root"]);
    expect(errors).toEqual([]);
  });

  it("should handle entity spanning chunk boundary in attribute value", () => {
    const parser = new SaxParser();
    let attrs: Record<string, string> = {};
    parser.on("opentag", tag => {
      attrs = tag.attributes;
    });
    parser.write('<root attr="a&am');
    parser.write('p;b"/>');
    parser.close();
    expect(attrs.attr).toBe("a&b");
  });

  it("should emit closetag events for auto-closed intermediate tags", () => {
    const parser = new SaxParser();
    const closeTags: string[] = [];
    parser.on("error", () => {});
    parser.on("closetag", tag => closeTags.push(tag.name));
    parser.write("<a><b><c></a>");
    parser.close();
    // c and b are auto-closed when </a> is encountered
    expect(closeTags).toEqual(["c", "b", "a"]);
  });
});

// =============================================================================
// Writer Round-Trip Integration
// =============================================================================

describe("Writer round-trip through parseXml", () => {
  it("should produce parseable output with complex content", () => {
    const w = new XmlWriter();
    w.openXml();
    w.openNode("root", { version: "1.0", lang: "en" });
    w.leafNode("title", undefined, 'Say "hello" & <goodbye>');
    w.openNode("data");
    w.writeCData("raw ]]> content");
    w.closeNode();
    w.leafNode("count", undefined, 0);
    w.leafNode("empty");
    w.closeNode();

    const doc = parseXml(w.xml);
    expect(doc.root.name).toBe("root");
    expect(attr(doc.root, "version")).toBe("1.0");
    expect(textContent(findChild(doc.root, "title")!)).toBe('Say "hello" & <goodbye>');
    expect(textContent(findChild(doc.root, "data")!)).toBe("raw ]]> content");
    expect(textContent(findChild(doc.root, "count")!)).toBe("0");
    expect(findChild(doc.root, "empty")).toBeDefined();
  });
});

// =============================================================================
// Writer Parity for CDATA Split
// =============================================================================

describe("Writer parity for CDATA splitting", () => {
  it("should produce identical CDATA-split output from both writers", () => {
    const chunks: string[] = [];
    const target: WritableTarget = { write: (c: string | Uint8Array) => chunks.push(String(c)) };
    const sw = new XmlStreamWriter(target);
    const w = new XmlWriter();

    const ops = (s: any) => {
      s.openNode("root");
      s.writeCData("a]]>b]]>c");
      s.closeNode();
    };

    ops(sw);
    ops(w);
    expect(chunks.join("")).toBe(w.xml);
  });
});

// =============================================================================
// UTF-8 Multibyte Across Uint8Array Chunk Boundary
// =============================================================================

describe("parseSax UTF-8 multibyte across chunk boundary", () => {
  it("should handle a 3-byte UTF-8 character split across Uint8Array chunks", async () => {
    // € = U+20AC = 0xE2 0x82 0xAC in UTF-8
    const full = new TextEncoder().encode("<r>€</r>");
    // Split right in the middle of the € encoding
    const chunk1 = full.slice(0, 4); // <r> + 0xE2
    const chunk2 = full.slice(4); // 0x82 0xAC + </r>

    const events: any[] = [];
    for await (const batch of parseSax([chunk1, chunk2])) {
      events.push(...batch);
    }
    const texts = events.filter((e: any) => e.eventType === "text").map((e: any) => e.value);
    expect(texts.join("")).toBe("€");
  });

  it("should handle a 4-byte UTF-8 character split across Uint8Array chunks", async () => {
    // 😀 = U+1F600 = 0xF0 0x9F 0x98 0x80 in UTF-8
    const full = new TextEncoder().encode("<r>😀</r>");
    const chunk1 = full.slice(0, 5); // <r> + 0xF0 0x9F
    const chunk2 = full.slice(5); // 0x98 0x80 + </r>

    const events: any[] = [];
    for await (const batch of parseSax([chunk1, chunk2])) {
      events.push(...batch);
    }
    const texts = events.filter((e: any) => e.eventType === "text").map((e: any) => e.value);
    expect(texts.join("")).toBe("😀");
  });
});

// =============================================================================
// XML Declaration with Single Quotes
// =============================================================================

describe("parseXml declaration with single quotes", () => {
  it("should capture declaration attributes using single quotes", () => {
    const doc = parseXml("<?xml version='1.0' encoding='UTF-8'?><root/>");
    expect(doc.declaration).toBeDefined();
    expect(doc.declaration!.version).toBe("1.0");
    expect(doc.declaration!.encoding).toBe("UTF-8");
  });
});

// =============================================================================
// SAX Chunk Boundary — Additional Edge Cases
// =============================================================================

describe("SAX chunk boundary additional edge cases", () => {
  it("attribute name split across chunks", () => {
    const parser = new SaxParser();
    const tags: any[] = [];
    parser.on("opentag", t => tags.push({ name: t.name, attributes: { ...t.attributes } }));
    parser.write("<root attr");
    parser.write('Name="val"/>');
    parser.close();
    expect(tags).toHaveLength(1);
    expect(tags[0].attributes.attrName).toBe("val");
  });

  it("close tag name split across chunks", () => {
    const parser = new SaxParser();
    const closed: string[] = [];
    parser.on("closetag", t => closed.push(t.name));
    parser.write("<root></ro");
    parser.write("ot>");
    parser.close();
    expect(closed).toEqual(["root"]);
  });

  it("attribute value closing quote at chunk boundary", () => {
    const parser = new SaxParser();
    const tags: any[] = [];
    parser.on("opentag", t => tags.push({ ...t.attributes }));
    parser.write('<root attr="val');
    parser.write('"/>');
    parser.close();
    expect(tags[0].attr).toBe("val");
  });

  it("empty chunk between meaningful chunks", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>");
    parser.write("");
    parser.write("hello");
    parser.write("");
    parser.write("</root>");
    parser.close();
    expect(texts.join("")).toBe("hello");
  });

  it("chunk boundary right at '=' between attribute name and value", () => {
    const parser = new SaxParser();
    const tags: any[] = [];
    parser.on("opentag", t => tags.push({ ...t.attributes }));
    parser.write("<root attr=");
    parser.write('"val"/>');
    parser.close();
    expect(tags[0].attr).toBe("val");
  });

  it("surrogate pair split across string chunks", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>\uD83D");
    parser.write("\uDE00</root>");
    parser.close();
    expect(texts.join("")).toBe("\uD83D\uDE00");
  });

  it("entity numeric reference split at '#' boundary", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>&#x4");
    parser.write("1;</root>");
    parser.close();
    expect(texts.join("")).toBe("A");
  });

  it("entity numeric reference split at 'x' boundary", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>&#");
    parser.write("x41;</root>");
    parser.close();
    expect(texts.join("")).toBe("A");
  });
});

// =============================================================================
// SAX Error Recovery
// =============================================================================

describe("SAX error recovery", () => {
  it("continues parsing after invalid character in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const tags: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.on("opentag", t => tags.push(t.name));
    parser.write("<root>\x01<child/></root>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
    expect(tags).toContain("child");
  });

  it("continues parsing after entity expansion limit", () => {
    const parser = new SaxParser({ maxEntityExpansions: 2 });
    parser.ENTITIES.x = "expanded";
    const errors: string[] = [];
    const tags: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.on("opentag", t => tags.push(t.name));
    parser.write("<root>&x;&x;&x;<child/></root>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
    expect(tags).toContain("child");
  });
});

// =============================================================================
// XmlStreamWriter Pending Tag Edge Cases
// =============================================================================

describe("XmlStreamWriter pending tag edge cases", () => {
  it("addAttribute after writeText should throw", () => {
    const sw = new XmlStreamWriter({ write() {} });
    sw.openNode("r");
    sw.writeText("text");
    expect(() => sw.addAttribute("a", "1")).toThrow();
  });

  it("multiple addAttributes accumulate in pending buffer", () => {
    const chunks: string[] = [];
    const sw = new XmlStreamWriter({ write: (c: string) => chunks.push(c) });
    sw.openNode("r");
    sw.addAttribute("a", "1");
    sw.addAttribute("b", "2");
    sw.closeNode();
    const xml = chunks.join("");
    expect(xml).toBe('<r a="1" b="2"/>');
  });

  it("writeComment when open tag is pending flushes the tag", () => {
    const chunks: string[] = [];
    const sw = new XmlStreamWriter({ write: (c: string) => chunks.push(c) });
    sw.openNode("r");
    sw.writeComment("note");
    sw.closeNode();
    const xml = chunks.join("");
    expect(xml).toBe("<r><!--note--></r>");
  });

  it("writeCData with empty string", () => {
    const chunks: string[] = [];
    const sw = new XmlStreamWriter({ write: (c: string) => chunks.push(c) });
    sw.openNode("r");
    sw.writeCData("");
    sw.closeNode();
    const xml = chunks.join("");
    expect(xml).toBe("<r><![CDATA[]]></r>");
  });
});

// =============================================================================
// SAX: Second root element rejection
// =============================================================================

describe("SAX second root element rejection", () => {
  it("rejects second root element in non-fragment mode", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.write("<a/><b/>");
    parser.close();
    expect(errors.some(e => e.includes("one root element"))).toBe(true);
  });

  it("accepts multiple roots in fragment mode", () => {
    const parser = new SaxParser({ fragment: true });
    const errors: string[] = [];
    const tags: string[] = [];
    parser.on("error", e => errors.push(e.message));
    parser.on("opentag", t => tags.push(t.name));
    parser.write("<a/><b/>");
    parser.close();
    expect(errors).toEqual([]);
    expect(tags).toEqual(["a", "b"]);
  });
});

// =============================================================================
// DOM: Second root element rejection
// =============================================================================

describe("parseXml second root element rejection", () => {
  it("rejects document with two root elements", () => {
    expect(() => parseXml("<a/><b/>")).toThrow(/root/i);
  });

  it("accepts multiple roots in fragment mode", () => {
    const doc = parseXml("<a/><b/>", { fragment: true });
    expect(doc.root.name).toBe("a");
  });
});

// =============================================================================
// Fatal UTF-8 decoding in parseSax
// =============================================================================

describe("parseSax fatal UTF-8 decoding", () => {
  it("rejects invalid UTF-8 byte sequence", async () => {
    const invalidUtf8 = new Uint8Array([
      0x3c,
      0x72,
      0x6f,
      0x6f,
      0x74,
      0x3e, // <root>
      0xff,
      0xfe, // invalid UTF-8 bytes
      0x3c,
      0x2f,
      0x72,
      0x6f,
      0x6f,
      0x74,
      0x3e // </root>
    ]);
    const chunks = [invalidUtf8];

    await expect(async () => {
      for await (const _events of parseSax(chunks)) {
        // consume
      }
    }).rejects.toThrow();
  });

  it("handles valid multibyte UTF-8 across chunks", async () => {
    // "\u4f60\u597d" in UTF-8: E4BDA0 E5A5BD
    const chunk1 = new Uint8Array([
      0x3c,
      0x72,
      0x3e, // <r>
      0xe4,
      0xbd // first 2 bytes of \u4f60
    ]);
    const chunk2 = new Uint8Array([
      0xa0, // last byte of \u4f60
      0xe5,
      0xa5,
      0xbd, // \u597d
      0x3c,
      0x2f,
      0x72,
      0x3e // </r>
    ]);

    const texts: string[] = [];
    for await (const events of parseSax([chunk1, chunk2])) {
      for (const evt of events) {
        if (evt.eventType === "text") {
          texts.push(evt.value);
        }
      }
    }
    expect(texts.join("")).toBe("\u4f60\u597d");
  });
});
