import { parseXml, textContent } from "@xml/dom";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

describe("XmlWriter", () => {
  describe("basic XML generation", () => {
    it("should generate XML declaration", () => {
      const w = new XmlWriter();
      w.openXml();
      expect(w.xml).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    });

    it("should generate XML declaration with custom attributes", () => {
      const w = new XmlWriter();
      w.openXml({ version: "1.0", encoding: "UTF-16" });
      expect(w.xml).toBe('<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\n');
    });

    it("should generate a self-closing element", () => {
      const w = new XmlWriter();
      w.openNode("br");
      w.closeNode();
      expect(w.xml).toBe("<br/>");
    });

    it("should generate an element with text", () => {
      const w = new XmlWriter();
      w.openNode("p");
      w.writeText("hello");
      w.closeNode();
      expect(w.xml).toBe("<p>hello</p>");
    });

    it("should generate an element with attributes", () => {
      const w = new XmlWriter();
      w.openNode("div", { id: "main", class: "container" });
      w.closeNode();
      expect(w.xml).toBe('<div id="main" class="container"/>');
    });

    it("should handle nested elements", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.openNode("child");
      w.writeText("text");
      w.closeNode();
      w.closeNode();
      expect(w.xml).toBe("<root><child>text</child></root>");
    });
  });

  describe("leafNode", () => {
    it("should generate a leaf element without text", () => {
      const w = new XmlWriter();
      w.leafNode("br");
      expect(w.xml).toBe("<br/>");
    });

    it("should generate a leaf element with text", () => {
      const w = new XmlWriter();
      w.leafNode("span", { class: "bold" }, "text");
      expect(w.xml).toBe('<span class="bold">text</span>');
    });

    it("should generate a leaf element with zero value", () => {
      const w = new XmlWriter();
      w.leafNode("v", undefined, 0);
      expect(w.xml).toBe("<v>0</v>");
    });
  });

  describe("addAttribute / addAttributes", () => {
    it("should add a single attribute after openNode", () => {
      const w = new XmlWriter();
      w.openNode("tag");
      w.addAttribute("key", "value");
      w.closeNode();
      expect(w.xml).toBe('<tag key="value"/>');
    });

    it("should add multiple attributes after openNode", () => {
      const w = new XmlWriter();
      w.openNode("tag");
      w.addAttributes({ a: "1", b: "2" });
      w.closeNode();
      expect(w.xml).toBe('<tag a="1" b="2"/>');
    });

    it("should throw when adding attribute with no open element", () => {
      const w = new XmlWriter();
      expect(() => w.addAttribute("key", "value")).toThrow("no element is open");
    });

    it("should skip undefined attribute values in addAttributes", () => {
      const w = new XmlWriter();
      w.openNode("tag");
      w.addAttributes({ a: "1", b: undefined, c: "3" });
      w.closeNode();
      expect(w.xml).toBe('<tag a="1" c="3"/>');
    });
  });

  describe("writeRaw", () => {
    it("should write raw XML content", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeRaw("<pre-built>content</pre-built>");
      w.closeNode();
      expect(w.xml).toBe("<root><pre-built>content</pre-built></root>");
    });
  });

  describe("writeCData", () => {
    it("should write CDATA section", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeCData("some <special> & content");
      w.closeNode();
      expect(w.xml).toBe("<root><![CDATA[some <special> & content]]></root>");
    });
  });

  describe("writeComment", () => {
    it("should write comment", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeComment(" a comment ");
      w.closeNode();
      expect(w.xml).toBe("<root><!-- a comment --></root>");
    });
  });

  describe("XML encoding", () => {
    it("should encode special characters in text", () => {
      const w = new XmlWriter();
      w.openNode("p");
      w.writeText('hello <world> & "friends"');
      w.closeNode();
      expect(w.xml).toBe("<p>hello &lt;world&gt; &amp; &quot;friends&quot;</p>");
    });

    it("should encode special characters in attributes", () => {
      const w = new XmlWriter();
      w.openNode("div", { title: 'say "hello" & <goodbye>' });
      w.closeNode();
      expect(w.xml).toBe('<div title="say &quot;hello&quot; &amp; &lt;goodbye&gt;"/>');
    });
  });

  describe("closeAll", () => {
    it("should close all open elements", () => {
      const w = new XmlWriter();
      w.openNode("a");
      w.openNode("b");
      w.openNode("c");
      w.closeAll();
      expect(w.xml).toBe("<a><b><c/></b></a>");
    });
  });

  describe("state queries", () => {
    it("should track depth", () => {
      const w = new XmlWriter();
      expect(w.depth).toBe(0);
      w.openNode("a");
      expect(w.depth).toBe(1);
      w.openNode("b");
      expect(w.depth).toBe(2);
      w.closeNode();
      expect(w.depth).toBe(1);
      w.closeNode();
      expect(w.depth).toBe(0);
    });

    it("should track currentElement", () => {
      const w = new XmlWriter();
      expect(w.currentElement).toBeUndefined();
      w.openNode("root");
      expect(w.currentElement).toBe("root");
      w.openNode("child");
      expect(w.currentElement).toBe("child");
      w.closeNode();
      expect(w.currentElement).toBe("root");
    });

    it("should track cursor", () => {
      const w = new XmlWriter();
      const c0 = w.cursor;
      w.openNode("root");
      expect(w.cursor).toBeGreaterThan(c0);
    });
  });

  describe("rollback / transaction", () => {
    it("should support save/commit", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.save();
      w.leafNode("child", undefined, "text");
      w.commit();
      w.closeNode();
      expect(w.xml).toBe("<root><child>text</child></root>");
    });

    it("should support save/rollback", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.save();
      w.leafNode("unwanted");
      w.rollback();
      w.leafNode("wanted");
      w.closeNode();
      expect(w.xml).toBe("<root><wanted/></root>");
    });

    it("should support nested save/rollback", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.save();
      w.openNode("outer");
      w.save();
      w.leafNode("inner-bad");
      w.rollback();
      w.leafNode("inner-good");
      w.closeNode();
      w.commit();
      w.closeNode();
      expect(w.xml).toBe("<root><outer><inner-good/></outer></root>");
    });

    it("should throw on commit without save", () => {
      const w = new XmlWriter();
      expect(() => w.commit()).toThrow("no snapshot to commit");
    });

    it("should throw on rollback without save", () => {
      const w = new XmlWriter();
      expect(() => w.rollback()).toThrow("no snapshot to rollback");
    });
  });

  describe("reset", () => {
    it("should reset to empty state", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.leafNode("child");
      w.reset();
      expect(w.xml).toBe("");
      expect(w.depth).toBe(0);
    });
  });

  describe("closeNode error", () => {
    it("should throw when no element to close", () => {
      const w = new XmlWriter();
      expect(() => w.closeNode()).toThrow("no element is open");
    });
  });

  describe("writeCData safety", () => {
    it("should handle ]]> inside CDATA by splitting", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeCData("a]]>b");
      w.closeNode();
      // Should produce valid XML that a parser can read back
      const xml = w.xml;
      expect(xml).not.toBe("<root><![CDATA[a]]>b]]></root>"); // NOT the naive broken output
      expect(xml).toContain("CDATA");
      // Verify round-trip: parse it back
      const doc = parseXml(xml);
      expect(textContent(doc.root)).toBe("a]]>b");
    });

    it("should handle multiple ]]> occurrences in CDATA", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeCData("x]]>y]]>z");
      w.closeNode();
      const doc = parseXml(w.xml);
      expect(textContent(doc.root)).toBe("x]]>y]]>z");
    });

    it("should handle CDATA that is just ]]>", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.writeCData("]]>");
      w.closeNode();
      const doc = parseXml(w.xml);
      expect(textContent(doc.root)).toBe("]]>");
    });
  });

  describe("writeComment safety", () => {
    it("should reject -- inside comment text", () => {
      const w = new XmlWriter();
      w.openNode("root");
      expect(() => w.writeComment("a--b")).toThrow();
    });

    it("should reject comment text ending with -", () => {
      const w = new XmlWriter();
      w.openNode("root");
      expect(() => w.writeComment("text-")).toThrow();
    });

    it("should accept valid comment text", () => {
      const w = new XmlWriter();
      w.openNode("root");
      expect(() => w.writeComment(" valid comment ")).not.toThrow();
      w.closeNode();
      expect(w.xml).toBe("<root><!-- valid comment --></root>");
    });

    it("should accept empty comment", () => {
      const w = new XmlWriter();
      w.openNode("root");
      expect(() => w.writeComment("")).not.toThrow();
      w.closeNode();
      expect(w.xml).toBe("<root><!----></root>");
    });
  });

  describe("rollback edge cases", () => {
    it("should restore _leaf and _open state after rollback", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.save();
      w.writeText("text");
      w.rollback();
      // After rollback, _leaf should be true and _open should be true
      // So closeNode produces self-closing tag
      w.closeNode();
      expect(w.xml).toBe("<root/>");
    });

    it("should handle double save/rollback", () => {
      const w = new XmlWriter();
      w.openNode("root");
      w.save();
      w.leafNode("a");
      w.save();
      w.leafNode("b");
      w.rollback(); // removes b
      w.rollback(); // removes a
      w.leafNode("c");
      w.closeNode();
      expect(w.xml).toBe("<root><c/></root>");
    });

    it("should handle closeAll on empty writer", () => {
      const w = new XmlWriter();
      w.closeAll(); // no-op, should not throw
      expect(w.xml).toBe("");
    });
  });
});

describe("XmlWriter — complex Excel-like XML", () => {
  it("should generate worksheet-like XML", () => {
    const w = new XmlWriter();
    w.openXml();
    w.openNode("worksheet", {
      xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    });
    w.openNode("sheetData");
    w.openNode("row", { r: 1 });
    w.openNode("c", { r: "A1", t: "s" });
    w.leafNode("v", undefined, 0);
    w.closeNode(); // c
    w.closeNode(); // row
    w.closeNode(); // sheetData
    w.closeNode(); // worksheet
    expect(w.xml).toContain("<worksheet");
    expect(w.xml).toContain('<row r="1">');
    expect(w.xml).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(w.xml).toContain("</sheetData></worksheet>");
  });
});
