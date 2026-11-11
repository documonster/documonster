/**
 * Well-formedness and resource rules the SAX parser enforces, each with the verdict of expat
 * (Python's `xml.parsers.expat`, in namespace mode for the namespace cases) and, where it checks
 * the rule, libxml2 (`xmllint --noout --nonet`) — and the same verdict however the input is chunked.
 */
import { isUint8Array } from "@utils/binary";
import { XmlArgumentError, XmlError, XmlLimitExceededError, XmlParseError } from "@xml/errors";
import { Xml } from "@xml/index";
import { SaxParser } from "@xml/sax";
import { parseXmlToObject } from "@xml/to-object";
import type { XmlSaxOptions } from "@xml/types";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();

/** Whether `doc` parses, written in pieces of every size, as text and as bytes: one verdict, or a failure. */
function verdict(doc: string, options?: XmlSaxOptions): boolean {
  const outcomes = new Set<boolean>();
  const bytes = encoder.encode(doc);
  for (const input of [doc, bytes] as const) {
    for (let size = 1; size <= input.length; size++) {
      const parser = new SaxParser(options);
      try {
        for (let at = 0; at < input.length; at += size) {
          parser.write(input.slice(at, at + size));
        }
        parser.close();
        outcomes.add(true);
      } catch (err) {
        expect(err).toBeInstanceOf(XmlParseError);
        outcomes.add(false);
      }
    }
  }
  expect([...outcomes], doc).toHaveLength(1);
  return [...outcomes][0];
}

describe("SaxParser — the XML declaration", () => {
  it.each([
    ["<?xml?><R/>", false],
    [' <?xml version="1.0"?><R/>', false],
    ["<R/><?XML x?>", false],
    ['<?Xml version="1.0"?><R/>', false],
    ['<?xml version="1.0" junk?><R/>', false],
    ['<?xml version="2.0"?><R/>', false],
    ['<?xml version="1.0" standalone="yes" encoding="UTF-8"?><R/>', false],
    ['<R><?xml version="1.0"?></R>', false],
    ['<?xml version="1.0"?><R/>', true],
    ["<?xml version='1.1' encoding='UTF-8' standalone='no' ?><R/>", true],
    ['\uFEFF<?xml version="1.0"?><R/>', true],
    ['<?xml-stylesheet href="a"?><R/>', true]
  ])("%s", (doc, wellFormed) => {
    expect(verdict(doc)).toBe(wellFormed);
  });
});

describe("SaxParser — character data", () => {
  it.each([
    ["<R>]]></R>", false],
    ["<R>a]]>b</R>", false],
    ["<R>]]]></R>", false],
    ["<R>]] ></R>", true],
    ["<R>]&#93;></R>", true],
    ["<R>&#93;&#93;></R>", true],
    ["<R>x > y</R>", true],
    ["<R><![CDATA[x]]>]></R>", true]
  ])("%s", (doc, wellFormed) => {
    expect(verdict(doc)).toBe(wellFormed);
  });

  it("reports text outside the root where it is, before anything after it", () => {
    for (const doc of ["x \u0000<R/>", "x\n\n <R/>"]) {
      for (const size of [1, doc.length]) {
        const parser = new SaxParser();
        const error = (() => {
          try {
            for (let at = 0; at < doc.length; at += size) {
              parser.write(encoder.encode(doc.slice(at, at + size)));
            }
          } catch (err) {
            return err as XmlParseError;
          }
          return undefined;
        })();
        expect(error?.message).toContain("text data outside of root node");
        expect([error?.byte, error?.line, error?.column]).toEqual([0, 1, 1]);
      }
    }
  });

  it("counts columns in UTF-16 units however a supplementary character arrives", () => {
    for (const doc of [
      "<𐀀></𐀀>!",
      '<r a="😀😀">😀x<!--😀--><𐀀 b="1"/></r>!',
      "<r>\n😀<𐀀/>😀</r>  !"
    ]) {
      const places = new Set<string>();
      for (let size = 1; size <= doc.length; size++) {
        const parser = new SaxParser();
        try {
          for (let at = 0; at < doc.length; at += size) {
            parser.write(doc.slice(at, at + size));
          }
          parser.close();
        } catch (err) {
          places.add(`${(err as XmlParseError).line}:${(err as XmlParseError).column}`);
        }
      }
      expect([...places], doc).toHaveLength(1);
    }
  });
});

describe("SaxParser — namespaces", () => {
  it.each([
    ["<constructor:r/>", false],
    ["<toString:r/>", false],
    ['<r xmlns="http://www.w3.org/2000/xmlns/"/>', false],
    ['<r xmlns="http://www.w3.org/XML/1998/namespace"/>', false],
    ["<:r/>", false],
    ['<r: xmlns:r="u"/>', false],
    ['<r:1 xmlns:r="u"/>', false],
    ['<r xmlns:a:b="u"/>', false],
    ['<r xmlns:r="u" r:a:b="1"/>', false],
    ['<a:r xmlns:a="u" b:c="1" xmlns:b="v"/>', true]
  ])("%s", (doc, wellFormed) => {
    expect(verdict(doc, { xmlns: true })).toBe(wellFormed);
  });

  it("binds a prefix only by a declaration, never by what Object.prototype holds", () => {
    const proto = Xml.parse('<__proto__:r xmlns:__proto__="urn:test"/>', { xmlns: true }).root;
    expect(proto.uri).toBe("urn:test");
    const inner = Xml.parse('<r xmlns:constructor="urn:x"><constructor:c/></r>', { xmlns: true })
      .root.children[0];
    expect((inner as { uri?: unknown }).uri).toBe("urn:x");
  });
});

describe("SaxParser — the DOCTYPE internal subset", () => {
  it("never takes a declaration unchecked because it is long", () => {
    // Too long to hold: a malformed one is refused for what it is, any other as past the limit.
    const bogus = `<!DOCTYPE r [<!BOGUS ${"x".repeat(70000)}>]><r/>`;
    expect(() => new SaxParser().write(bogus).close()).toThrow(/expected ELEMENT, ATTLIST/);
    const attlist = `<!DOCTYPE r [<!ATTLIST r a CDATA "${"x".repeat(70000)}">]><r/>`;
    expect(() => new SaxParser().write(attlist).close()).toThrow(XmlLimitExceededError);
    // Within a raised bound, both are checked whole.
    const raised = { maxContextBytes: 1 << 20 };
    expect(() => new SaxParser(raised).write(bogus).close()).toThrow(/expected ELEMENT, ATTLIST/);
    expect(() => new SaxParser(raised).write(attlist).close()).not.toThrow();
    const broken = `<!DOCTYPE r [<!ATTLIST r a CDATA "${"x".repeat(70000)}" junk>]><r/>`;
    expect(() => Xml.parse(broken)).toThrow(XmlParseError);
  });
});

describe("SaxParser — entity expansion", () => {
  /** A chain of entities each doubling the last, `e0` being `length` characters. */
  function doubling(length: number, steps: number): string {
    let dtd = `<!ENTITY e0 "${"A".repeat(length)}">`;
    for (let i = 1; i <= steps; i++) {
      dtd += `<!ENTITY e${i} "&e${i - 1};&e${i - 1};">`;
    }
    return dtd;
  }

  it("refuses an invalid character reference before building what the entity refers to", () => {
    // `&#38;#` declares `&#`: read where the entity is used, the references run together into one
    // that is no character — which once counted as a negative length and let `&e12;` be built.
    const doc = `<!DOCTYPE r [${doubling(140000, 12)}<!ENTITY bad "${"&#38;#".repeat(24000)}&e12;">]><r>&bad;</r>`;
    expect(() => parseXmlToObject(doc, { internalEntities: true })).toThrow(
      /invalid character reference/
    );
    // Recovering from errors, the length counted is never less than what is given out.
    const small = `<!DOCTYPE r [${doubling(128, 8)}<!ENTITY bad "${"&#38;#".repeat(300)}&e8;">]><r>&bad;&e8;</r>`;
    const parser = new SaxParser({ internalEntities: true, maxEntityExpansionLength: 1 });
    let length = 0;
    const limits: unknown[] = [];
    parser.on("error", err => limits.push((err as XmlParseError).limit));
    parser.on("text", text => void (length += text.length));
    parser.write(small).close();
    expect(length).toBe(0);
    expect(limits).toContain("entityExpansionLength");
  });

  it("expands a long chain of entities, each written straight into the text", () => {
    let dtd = `<!ENTITY e0 "${"A".repeat(1000)}">`;
    for (let i = 1; i <= 1000; i++) {
      dtd += `<!ENTITY e${i} "&e${i - 1};B">`;
    }
    const object = parseXmlToObject(`<!DOCTYPE r [${dtd}]><r a="&e1000;">&e1000;</r>`, {
      internalEntities: true
    });
    const expected = "A".repeat(1000) + "B".repeat(1000);
    expect(object).toEqual({ r: { "@_a": expected, "#text": expected } });
  });

  it("counts a caller's own entities toward maxEntityExpansionLength", () => {
    const parser = new SaxParser({ maxEntityExpansionLength: 3 });
    parser.ENTITIES.x = "1234";
    expect(() => parser.write("<r>&x;</r>").close()).toThrow(XmlLimitExceededError);
    const within = new SaxParser({ maxEntityExpansionLength: 4 });
    within.ENTITIES.x = "1234";
    let text = "";
    within.on("text", t => void (text += t));
    within.write("<r>&x;</r>").close();
    expect(text).toBe("1234");
  });
});

describe("SaxParser — bytes", () => {
  it("takes a malformed UTF-8 prefix at the end of the input as malformed, not as cut short", () => {
    const bytes = Uint8Array.of(0x3c, 0x72, 0x2f, 0x3e, 0xe0, 0x80);
    const outcome = (invalidCharHandling: "error" | "skip" | "replace"): string => {
      try {
        new SaxParser({ invalidCharHandling }).write(bytes).close();
        return "ok";
      } catch (err) {
        return (err as Error).message;
      }
    };
    expect(outcome("error")).toMatch(/invalid UTF-8/);
    expect(outcome("skip")).toBe("ok");
    expect(outcome("replace")).toMatch(/text data outside of root node/);
    // A character that may still arrive whole is held, not taken for malformed.
    const parser = new SaxParser();
    let text = "";
    parser.on("text", t => void (text += t));
    parser.write(Uint8Array.of(0x3c, 0x72, 0x3e, 0xe2));
    parser.write(Uint8Array.of(0x82, 0xac, 0x3c, 0x2f, 0x72, 0x3e));
    parser.close();
    expect(text).toBe("€");
  });

  it("knows a Uint8Array by what it is, not by what its tag says", () => {
    const spoofed = new Uint16Array([0x523c, 0x3e2f]);
    Object.defineProperty(spoofed, Symbol.toStringTag, { value: "Uint8Array" });
    expect(isUint8Array(spoofed)).toBe(false);
    expect(() => new SaxParser().write(spoofed as never)).toThrow(XmlArgumentError);
    // A Uint8Array that says it is something else is one still.
    const relabelled = new Uint8Array(1);
    Object.defineProperty(relabelled, Symbol.toStringTag, { value: "Bytes" });
    expect(isUint8Array(relabelled)).toBe(true);
    class Bytes extends Uint8Array {}
    expect(isUint8Array(new Bytes(1))).toBe(true);
  });
});

describe("SaxParser — stop", () => {
  it("reads nothing after stop() in an error's handler, and reports no tag after it", () => {
    const events: string[] = [];
    const parser = new SaxParser();
    parser.on("opentag", tag => events.push(`open ${tag.name}`));
    parser.on("text", text => events.push(`text ${JSON.stringify(text)}`));
    parser.on("error", () => {
      events.push("error");
      parser.stop();
    });
    parser.write("<r>a\u0000after<x/>");
    expect(events).toEqual(["open r", "error"]);

    const tags: string[] = [];
    const duplicate = new SaxParser();
    duplicate.on("opentag", tag => tags.push(tag.name));
    duplicate.on("error", () => duplicate.stop());
    duplicate.write('<r a="1" a="2"><x/>');
    expect(tags).toEqual([]);
  });

  it("reports no end of a self-closing tag whose start's handler stopped the parse", () => {
    const events: string[] = [];
    const parser = new SaxParser();
    parser.on("opentag", () => {
      events.push("open");
      parser.stop();
    });
    parser.on("closetag", () => events.push("close"));
    parser.write("<R/>");
    expect(events).toEqual(["open"]);
  });
});

describe("Xml.parseSaxStream and Xml.feedSax", () => {
  const bad = Uint8Array.of(0x3c, 0x72, 0x3e, 0xff, 0x3c, 0x2f, 0x72, 0x3e);

  it("give bytes to the parser as bytes: invalidCharHandling, byteOffset and positions apply", async () => {
    const texts: string[] = [];
    for await (const batch of Xml.parseSaxStream([bad], { invalidCharHandling: "replace" })) {
      for (const event of batch) {
        if (event.eventType === "text") {
          texts.push(event.value);
        }
      }
    }
    expect(texts).toEqual(["\uFFFD"]);
    const parser = new SaxParser({ byteOffset: 100 });
    const starts: unknown[] = [];
    parser.on("opentag", () => starts.push(parser.startByte));
    await Xml.feedSax(parser, [encoder.encode("<r/>")]);
    expect(starts).toEqual([100]);
  });

  it("refuse a chunk that is not text or bytes as an argument, and text mixed with bytes", async () => {
    const drain = async (chunks: unknown[]): Promise<unknown> => {
      try {
        for await (const _ of Xml.parseSaxStream(chunks as never)) {
          // drain
        }
        return undefined;
      } catch (err) {
        return err;
      }
    };
    expect(await drain([42])).toMatchObject({ field: "iterable" });
    expect(await drain(["<r>", encoder.encode("</r>")])).toBeInstanceOf(XmlArgumentError);
  });

  it("use the options checked when called, and stop at the first error", async () => {
    const options = { maxDepth: 1 };
    const events = Xml.parseSaxStream(["<r><x/></r>"], options);
    options.maxDepth = 0;
    await expect(
      (async () => {
        for await (const _ of events) {
          // drain
        }
      })()
    ).rejects.toBeInstanceOf(XmlLimitExceededError);
    const parser = new SaxParser({ maxDepth: 1 });
    const opened: string[] = [];
    parser.on("opentag", tag => opened.push(tag.name));
    await expect(Xml.feedSax(parser, ["<r><x><z/></x></r>"])).rejects.toBeInstanceOf(
      XmlLimitExceededError
    );
    expect(opened).toEqual(["r"]);
  });

  it("give the parser its own error handler back, however the feed ends", async () => {
    const handler = (): void => {};
    const parser = new SaxParser();
    parser.on("error", handler);
    await Xml.feedSax(parser, ["<r/>"]);
    expect(parser.getHandler("error")).toBe(handler);
    await Xml.feedSax(parser, ["<r>"]).catch(() => undefined);
    expect(parser.getHandler("error")).toBe(handler);
  });
});

describe("SaxParser — a stopped parse, a handler writing, text nobody reads", () => {
  /** The events of `doc` read by a parser whose error handler stops it at the first error. */
  function eventsStoppedAtFirstError(doc: string, options?: XmlSaxOptions): string[] {
    const parser = new SaxParser(options);
    const events: string[] = [];
    parser.on("error", err => {
      events.push(`error ${err.message}`);
      parser.stop();
    });
    parser.on("opentag", tag => events.push(`open ${tag.name}`));
    parser.on("closetag", tag => events.push(`close ${tag.name}`));
    parser.on("text", text => events.push(`text ${text}`));
    parser.write(doc).close();
    return events;
  }

  it.each<[string, string, XmlSaxOptions?]>([
    ["an attribute value with two bad characters", '<r a="\u0000after\u0001"/>'],
    ["a tag with three equal attributes", '<r a="1" a="2" a="3"/>'],
    ["two bad namespace declarations", '<r xmlns:p="" xmlns:q=""/>', { xmlns: true }],
    ["two ]]> in one run of text", "<r>]]>after]]><x/></r>"],
    ["a prolog context past its limit", "<root/>", { internalEntities: true, maxContextBytes: 26 }],
    ["an element left open at the end", "<r><a>hello"]
  ])("reports nothing after %s once stopped", (_label, doc, options) => {
    const events = eventsStoppedAtFirstError(doc, options);
    const first = events.findIndex(event => event.startsWith("error"));
    expect(first).toBeGreaterThanOrEqual(0);
    expect(events.slice(first + 1)).toEqual([]);
  });

  it("closes no more elements once a close tag's handler stops it", () => {
    const parser = new SaxParser();
    const events: string[] = [];
    parser.on("error", err => events.push(`error ${err.message}`));
    parser.on("closetag", tag => {
      events.push(`close ${tag.name}`);
      parser.stop();
    });
    parser.write("<r><a></r>").close();
    expect(events).toEqual(["close a"]);
  });

  it("places a ]]> in text at its own byte, however the input is chunked", () => {
    const bytes = encoder.encode("<r>éa]]></r>");
    for (let size = 1; size <= bytes.length; size++) {
      const parser = new SaxParser();
      let byte: number | undefined;
      parser.on("error", err => {
        byte ??= (err as XmlParseError).byte;
      });
      for (let at = 0; at < bytes.length; at += size) {
        parser.write(bytes.subarray(at, at + size));
      }
      parser.close();
      expect(byte, `size ${size}`).toBe(8);
    }
  });

  it("refuses a write from a handler while it parses, and goes on with its own", () => {
    const parser = new SaxParser();
    const opened: string[] = [];
    let refused: unknown;
    parser.on("opentag", tag => {
      opened.push(tag.name);
      try {
        parser.write("<x/>");
      } catch (err) {
        refused ??= err;
      }
    });
    parser.write("<r><a/></r>").close();
    expect(refused).toBeInstanceOf(XmlError);
    expect(String(refused)).toMatch(/called from a handler/);
    expect(opened).toEqual(["r", "a"]);
  });

  it("keeps no text of references when nobody listens for text", () => {
    const parser = new SaxParser();
    parser.write("<r>");
    for (let i = 0; i < 1000; i++) {
      parser.write("&amp;&#65;");
    }
    expect((parser as unknown as { text: string }).text).toBe("");
    // A listener still gets all of it.
    const listened = new SaxParser();
    let text = "";
    listened.on("text", t => {
      text += t;
    });
    listened.write("<r>");
    for (let i = 0; i < 3; i++) {
      listened.write("&amp;&#65;");
    }
    listened.write("</r>").close();
    expect(text).toBe("&A&A&A");
  });

  it("counts a caller's one-character entity toward the expansion length", () => {
    const parser = new SaxParser({ maxEntityExpansions: 0, maxEntityExpansionLength: 1 });
    parser.ENTITIES.x = "x";
    expect(() => parser.write("<r>&x;&x;</r>").close()).toThrow(XmlLimitExceededError);
    // The five predefined entities, with their own values, are not counted.
    const predefined = new SaxParser({ maxEntityExpansions: 0, maxEntityExpansionLength: 1 });
    expect(() => predefined.write("<r>&lt;&gt;&amp;&quot;&apos;</r>").close()).not.toThrow();
  });
});
