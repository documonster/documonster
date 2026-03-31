import { describe, it, expect } from "vitest";
import { SaxParser, parseSax } from "@xml/sax";
import type { SaxTag, SaxEventAny } from "@xml/types";

describe("SaxParser", () => {
  describe("basic parsing", () => {
    it("should parse a simple element", () => {
      const parser = new SaxParser();
      const tags: string[] = [];
      parser.on("opentag", tag => tags.push(`open:${tag.name}`));
      parser.on("closetag", tag => tags.push(`close:${tag.name}`));
      parser.write("<root></root>");
      parser.close();
      expect(tags).toEqual(["open:root", "close:root"]);
    });

    it("should parse self-closing element", () => {
      const parser = new SaxParser();
      const events: string[] = [];
      parser.on("opentag", tag => {
        events.push(`open:${tag.name}:selfClosing=${tag.isSelfClosing}`);
      });
      parser.on("closetag", tag => events.push(`close:${tag.name}`));
      parser.write("<br/>");
      parser.close();
      expect(events).toEqual(["open:br:selfClosing=true", "close:br"]);
    });

    it("should parse text content", () => {
      const parser = new SaxParser();
      const texts: string[] = [];
      parser.on("text", text => texts.push(text));
      parser.write("<p>hello world</p>");
      parser.close();
      expect(texts).toEqual(["hello world"]);
    });

    it("should parse attributes", () => {
      const parser = new SaxParser();
      let attrs: Record<string, string> = {};
      parser.on("opentag", tag => {
        attrs = tag.attributes;
      });
      parser.write('<div id="main" class="container"/>');
      parser.close();
      expect(attrs).toEqual({ id: "main", class: "container" });
    });
  });

  describe("entity handling", () => {
    it("should decode XML entities in text", () => {
      const parser = new SaxParser();
      const texts: string[] = [];
      parser.on("text", text => texts.push(text));
      parser.write("<p>&lt;hello&gt; &amp; &quot;world&quot;</p>");
      parser.close();
      expect(texts).toEqual(['<hello> & "world"']);
    });

    it("should decode entities in attribute values", () => {
      const parser = new SaxParser();
      let attrs: Record<string, string> = {};
      parser.on("opentag", tag => {
        attrs = tag.attributes;
      });
      parser.write('<tag attr="a &amp; b"/>');
      parser.close();
      expect(attrs.attr).toBe("a & b");
    });

    it("should handle numeric character references", () => {
      const parser = new SaxParser();
      const texts: string[] = [];
      parser.on("text", text => texts.push(text));
      parser.write("<p>&#65;&#x42;</p>");
      parser.close();
      expect(texts).toEqual(["AB"]);
    });
  });

  describe("nested elements", () => {
    it("should track nesting correctly", () => {
      const parser = new SaxParser();
      const events: string[] = [];
      parser.on("opentag", tag => events.push(`open:${tag.name}`));
      parser.on("closetag", tag => events.push(`close:${tag.name}`));
      parser.on("text", text => events.push(`text:${text}`));
      parser.write("<root><child>text</child></root>");
      parser.close();
      expect(events).toEqual(["open:root", "open:child", "text:text", "close:child", "close:root"]);
    });
  });

  describe("CDATA sections", () => {
    it("should emit cdata event when handler is registered", () => {
      const parser = new SaxParser();
      const cdatas: string[] = [];
      parser.on("cdata", text => cdatas.push(text));
      parser.write("<root><![CDATA[some <special> content]]></root>");
      parser.close();
      expect(cdatas).toEqual(["some <special> content"]);
    });

    it("should emit as text when no cdata handler", () => {
      const parser = new SaxParser();
      const texts: string[] = [];
      parser.on("text", text => texts.push(text));
      parser.write("<root><![CDATA[content]]></root>");
      parser.close();
      expect(texts).toEqual(["content"]);
    });
  });

  describe("comments", () => {
    it("should emit comment event", () => {
      const parser = new SaxParser();
      const comments: string[] = [];
      parser.on("comment", text => comments.push(text));
      parser.write("<root><!-- a comment --></root>");
      parser.close();
      expect(comments).toEqual([" a comment "]);
    });
  });

  describe("processing instructions", () => {
    it("should emit pi event", () => {
      const parser = new SaxParser();
      const pis: Array<{ target: string; body: string }> = [];
      parser.on("pi", (target, body) => pis.push({ target, body }));
      parser.write('<?xml version="1.0"?><root/>');
      parser.close();
      expect(pis.length).toBe(1);
      expect(pis[0].target).toBe("xml");
    });
  });

  describe("chunked input", () => {
    it("should handle input split across multiple write calls", () => {
      const parser = new SaxParser();
      const events: string[] = [];
      parser.on("opentag", tag => events.push(`open:${tag.name}`));
      parser.on("text", text => events.push(`text:${text}`));
      parser.on("closetag", tag => events.push(`close:${tag.name}`));

      parser.write("<roo");
      parser.write("t>hel");
      parser.write("lo</root>");
      parser.close();

      expect(events).toEqual(["open:root", "text:hello", "close:root"]);
    });
  });

  describe("error handling", () => {
    it("should throw on malformed XML without error handler", () => {
      const parser = new SaxParser();
      // Writing null to close, unclosed tag should trigger error
      expect(() => {
        parser.write("<root><unclosed>");
        parser.close();
      }).toThrow();
    });

    it("should call error handler instead of throwing", () => {
      const parser = new SaxParser();
      const errors: Error[] = [];
      parser.on("error", err => errors.push(err));
      parser.write("<root><unclosed>");
      parser.close();
      expect(errors.length).toBeGreaterThan(0);
    });
  });

  describe("position tracking", () => {
    it("should track line and column", () => {
      const parser = new SaxParser({ position: true });
      parser.on("opentag", () => {});
      parser.write("<root/>");
      expect(parser.line).toBe(1);
    });
  });

  describe("fragment mode", () => {
    it("should parse fragments without root element requirement", () => {
      const parser = new SaxParser({ fragment: true });
      const events: string[] = [];
      parser.on("opentag", tag => events.push(`open:${tag.name}`));
      parser.on("closetag", tag => events.push(`close:${tag.name}`));
      parser.write("<a/><b/>");
      parser.close();
      expect(events).toEqual(["open:a", "close:a", "open:b", "close:b"]);
    });
  });

  describe("off", () => {
    it("should unregister event handler", () => {
      const parser = new SaxParser();
      const texts: string[] = [];
      parser.on("text", text => texts.push(text));
      parser.off("text");
      parser.write("<root>hello</root>");
      parser.close();
      expect(texts).toEqual([]);
    });
  });
});

describe("parseSax", () => {
  it("should parse async iterable of string chunks", async () => {
    async function* chunks(): AsyncGenerator<string> {
      yield "<root>";
      yield "<child>text</child>";
      yield "</root>";
    }

    const allEvents: SaxEventAny[] = [];
    for await (const batch of parseSax(chunks())) {
      allEvents.push(...batch);
    }

    const openTags = allEvents
      .filter(e => e.eventType === "opentag")
      .map(e => (e.value as SaxTag).name);
    expect(openTags).toEqual(["root", "child"]);

    const texts = allEvents.filter(e => e.eventType === "text").map(e => e.value);
    expect(texts).toEqual(["text"]);
  });

  it("should parse async iterable of Uint8Array chunks", async () => {
    const encoder = new TextEncoder();
    async function* chunks(): AsyncGenerator<Uint8Array> {
      yield encoder.encode("<root>");
      yield encoder.encode("<item>data</item>");
      yield encoder.encode("</root>");
    }

    const allEvents: SaxEventAny[] = [];
    for await (const batch of parseSax(chunks())) {
      allEvents.push(...batch);
    }

    const openTags = allEvents
      .filter(e => e.eventType === "opentag")
      .map(e => (e.value as SaxTag).name);
    expect(openTags).toEqual(["root", "item"]);
  });
});
