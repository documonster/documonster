import { describe, it, expect } from "vitest";
import { XmlStreamWriter } from "@xml/stream-writer";
import { parseXml, textContent } from "@xml/dom";
import type { WritableTarget } from "@xml/types";

/** Helper to capture stream output. */
function createCapture(): { target: WritableTarget; output: () => string } {
  const chunks: string[] = [];
  return {
    target: { write: (chunk: string | Uint8Array) => chunks.push(String(chunk)) },
    output: () => chunks.join("")
  };
}

describe("XmlStreamWriter", () => {
  describe("basic XML generation", () => {
    it("should generate XML declaration", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openXml();
      expect(output()).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n');
    });

    it("should generate a self-closing element", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("br");
      sw.closeNode();
      expect(output()).toBe("<br/>");
    });

    it("should generate an element with text", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("p");
      sw.writeText("hello");
      sw.closeNode();
      expect(output()).toBe("<p>hello</p>");
    });

    it("should generate nested elements", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.openNode("child");
      sw.writeText("text");
      sw.closeNode();
      sw.closeNode();
      expect(output()).toBe("<root><child>text</child></root>");
    });
  });

  describe("attributes", () => {
    it("should write attributes on open", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("div", { id: "main", class: "box" });
      sw.closeNode();
      expect(output()).toBe('<div id="main" class="box"/>');
    });

    it("should add attribute after open", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("tag");
      sw.addAttribute("key", "value");
      sw.closeNode();
      expect(output()).toBe('<tag key="value"/>');
    });

    it("should throw when adding attribute with no open element", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      expect(() => sw.addAttribute("key", "value")).toThrow("no element is open");
    });
  });

  describe("leafNode", () => {
    it("should write leaf with text", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.leafNode("span", { class: "bold" }, "text");
      expect(output()).toBe('<span class="bold">text</span>');
    });

    it("should write leaf without text", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.leafNode("br");
      expect(output()).toBe("<br/>");
    });
  });

  describe("special content", () => {
    it("should write CDATA", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.writeCData("special & <content>");
      sw.closeNode();
      expect(output()).toBe("<root><![CDATA[special & <content>]]></root>");
    });

    it("should write comment", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.writeComment(" note ");
      sw.closeNode();
      expect(output()).toBe("<root><!-- note --></root>");
    });

    it("should write raw XML", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.writeRaw("<pre>built</pre>");
      sw.closeNode();
      expect(output()).toBe("<root><pre>built</pre></root>");
    });
  });

  describe("state queries", () => {
    it("should track depth and currentElement", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      expect(sw.depth).toBe(0);
      expect(sw.currentElement).toBeUndefined();
      sw.openNode("a");
      expect(sw.depth).toBe(1);
      expect(sw.currentElement).toBe("a");
      sw.openNode("b");
      expect(sw.depth).toBe(2);
      expect(sw.currentElement).toBe("b");
      sw.closeNode();
      expect(sw.depth).toBe(1);
      sw.closeNode();
      expect(sw.depth).toBe(0);
    });
  });

  describe("closeAll", () => {
    it("should close all open elements", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("a");
      sw.openNode("b");
      sw.openNode("c");
      sw.closeAll();
      expect(output()).toBe("<a><b><c/></b></a>");
    });
  });

  describe("closeNode error", () => {
    it("should throw when no element to close", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      expect(() => sw.closeNode()).toThrow("no element is open");
    });
  });

  describe("XmlSink interface parity with XmlWriter", () => {
    it("should produce identical output to XmlWriter for same operations", async () => {
      // Import XmlWriter dynamically to verify parity
      const { XmlWriter } = await import("@xml/writer");

      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      const w = new XmlWriter();

      // Run identical operations on both
      const ops = (sink: any) => {
        sink.openXml();
        sink.openNode("root", { version: "1.0" });
        sink.leafNode("empty");
        sink.openNode("data");
        sink.writeText("hello & world");
        sink.closeNode();
        sink.leafNode("item", { id: "1" }, "value");
        sink.closeNode();
      };

      ops(sw);
      ops(w);

      expect(output()).toBe(w.xml);
    });
  });

  describe("writeCData safety", () => {
    it("should handle ]]> inside CDATA by splitting", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.writeCData("a]]>b");
      sw.closeNode();
      const xml = output();
      expect(xml).not.toBe("<root><![CDATA[a]]>b]]></root>");
      const doc = parseXml(xml);
      expect(textContent(doc.root)).toBe("a]]>b");
    });
  });

  describe("writeComment safety", () => {
    it("should reject -- inside comment text", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      expect(() => sw.writeComment("a--b")).toThrow();
    });

    it("should reject comment text ending with -", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      expect(() => sw.writeComment("text-")).toThrow();
    });

    it("should accept valid comment text", () => {
      const { target, output } = createCapture();
      const sw = new XmlStreamWriter(target);
      sw.openNode("root");
      sw.writeComment(" valid ");
      sw.closeNode();
      expect(output()).toBe("<root><!-- valid --></root>");
    });
  });

  describe("addAttributes error", () => {
    it("should throw when addAttributes called with no open element", () => {
      const { target } = createCapture();
      const sw = new XmlStreamWriter(target);
      expect(() => sw.addAttributes({ a: "1" })).toThrow("no element is open");
    });
  });
});
