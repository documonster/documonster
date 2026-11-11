/**
 * What the SAX parser reports for markup it used to misread: a DOCTYPE internal subset whose comments, processing
 * instructions or literals hold `]`, `>` or a quote; input that ends inside a construct; and the smaller rules each
 * construct has. Every document is parsed whole and one character at a time, since a state machine that is right for
 * one chunk can still be wrong across a boundary.
 */

import { parseXml } from "@xml/dom";
import { XmlError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import type { XmlSaxOptions } from "@xml/types";
import { describe, it, expect } from "vitest";

/** The events and errors `xml` produces, written whole or one character at a time. */
function events(xml: string, chunked: boolean, options?: XmlSaxOptions): string[] {
  const out: string[] = [];
  const parser = new SaxParser({ position: false, ...options });
  parser.on("opentag", tag => out.push(`open ${tag.name} ${JSON.stringify(tag.attributes)}`));
  parser.on("closetag", tag => out.push(`close ${tag.name}`));
  parser.on("text", text => out.push(`text ${JSON.stringify(text)}`));
  parser.on("pi", (target, body) =>
    out.push(`pi ${JSON.stringify(target)} ${JSON.stringify(body)}`)
  );
  parser.on("error", err => out.push(`error ${err.message}`));
  if (chunked) {
    for (const ch of xml) {
      parser.write(ch);
    }
  } else {
    parser.write(xml);
  }
  parser.close();
  return out;
}

function parse(xml: string, options?: XmlSaxOptions): string[] {
  const whole = events(xml, false, options);
  expect(events(xml, true, options)).toEqual(whole);
  return whole;
}

const ROOT_ONLY = ['open R {"a":"1"}', "close R"];

describe("DOCTYPE internal subset", () => {
  it.each([
    ["a ] inside a comment", "<!DOCTYPE R [ <!-- ] --> ]>"],
    ["a ] and > inside a comment", '<!DOCTYPE R [ <!-- ] > <F b="x"/> --> ]>'],
    ["an apostrophe inside a comment", "<!DOCTYPE R [ <!-- it's --> ]>"],
    ["a quotation mark inside a comment", '<!DOCTYPE R [ <!-- say "hi --> ]>'],
    ["a single - inside a comment", "<!DOCTYPE R [ <!-- x - y --> ]>"],
    ["a ] and > inside a processing instruction", '<!DOCTYPE R [ <?pi ] > <F b="x"/> ?> ]>'],
    ["an apostrophe inside a processing instruction", "<!DOCTYPE R [ <?pi it's ?> ]>"],
    ["a ?? before the end of a processing instruction", "<!DOCTYPE R [ <?pi x ??> ]>"],
    ["a ] and > inside an entity literal", "<!DOCTYPE R [ <!ENTITY e \"]>\"> <!ENTITY f ']>'> ]>"],
    ["an apostrophe inside a double-quoted literal", '<!DOCTYPE R [ <!ENTITY e "it\'s"> ]>'],
    ["an attribute-list default with ]", '<!DOCTYPE R [ <!ATTLIST R a CDATA "]"> ]>'],
    [
      "every kind of declaration",
      '<!DOCTYPE R SYSTEM "r.dtd" [ <!ELEMENT R (#PCDATA)> <!NOTATION n SYSTEM "x]>"> %pe; ]>'
    ],
    ["whitespace before the closing >", "<!DOCTYPE R [ <!-- ] --> ] >"]
  ])("skips %s", (_, doctype) => {
    // A processing instruction there is still given to the application (XML 1.0 §2.6).
    expect(parse(`${doctype}<R a="1"/>`).filter(event => !event.startsWith("pi "))).toEqual(
      ROOT_ONLY
    );
  });

  it("gives the application a processing instruction of the internal subset, with its own range", () => {
    expect(parse('<!DOCTYPE R [ <?pi ] > <F b="x"/> ?> <?q?> ]><R a="1"/>')).toEqual([
      'pi "pi" "] > <F b=\\"x\\"/> "',
      'pi "q" ""',
      ...ROOT_ONLY
    ]);
    const bytes = new TextEncoder().encode("<!DOCTYPE r [<?inside data?>]><r/>");
    for (let size = 1; size <= bytes.length; size++) {
      const parser = new SaxParser();
      const ranges: number[][] = [];
      parser.on("pi", () => ranges.push([parser.startByte!, parser.endByte!]));
      for (let at = 0; at < bytes.length; at += size) {
        parser.write(bytes.subarray(at, at + size));
      }
      parser.close();
      expect(ranges, `size ${size}`).toEqual([[13, 28]]);
    }
    // It belongs to the DTD, not to the document's prologue.
    expect(
      parseXml("<!DOCTYPE r [<?p x?>]><r/>", { processingInstructions: true }).prologue
    ).toEqual([]);
  });

  it("still ends at the first ] outside markup", () => {
    expect(parse('<!DOCTYPE R [ <!ENTITY e "x"> ]><R a="1"/>')).toEqual(ROOT_ONLY);
  });

  it("Xml.parse reads the document after such a subset", () => {
    const doc = parseXml('<!DOCTYPE R [ <!-- it\'s ] here --> ]><R a="1"><c/></R>');
    expect(doc.root.name).toBe("R");
    expect(doc.root.attributes.a).toBe("1");
  });

  it("reports a subset the input ends inside", () => {
    expect(parse("<!DOCTYPE R [ <!-- ]> <R/>")).toEqual([
      "error unexpected end of document inside the DOCTYPE declaration",
      "error document must contain a root element"
    ]);
  });
});

describe("DOCTYPE placement", () => {
  it("refuses a second DOCTYPE", () => {
    expect(parse('<!DOCTYPE R><!DOCTYPE R><R a="1"/>')).toEqual([
      "error a document may contain only one DOCTYPE declaration",
      ...ROOT_ONLY
    ]);
  });

  it.each([
    ["inside the root", '<R a="1"><!DOCTYPE R></R>', ['open R {"a":"1"}']],
    ["after the root", '<R a="1"/><!DOCTYPE R>', ROOT_ONLY]
  ])("refuses a DOCTYPE %s", (_, xml, before) => {
    expect(parse(xml).slice(0, before.length + 1)).toEqual([
      ...before,
      "error the DOCTYPE declaration must come before the root element"
    ]);
  });
});

describe("input that ends inside a construct", () => {
  it.each([
    ["a comment", "<R/><!-- x", "a comment"],
    ["a CDATA section", "<R><![CDATA[x", "a CDATA section"],
    ["a processing instruction", "<R/><?pi x", "a processing instruction"],
    ["an attribute value", '<R a="1', "a start tag"],
    ["a start tag", "<R a", "a start tag"],
    ["an end tag", "<R></R", "an end tag"],
    ["an entity reference", "<R>&amp", "an entity reference"],
    ["markup", "<R/><", "markup"]
  ])("reports %s and does not emit its characters as text", (_, xml, construct) => {
    const out = parse(xml);
    expect(out).toContain(`error unexpected end of document inside ${construct}`);
    expect(out.filter(e => e.startsWith("text"))).toEqual([]);
  });

  it("Xml.parse refuses a document cut off after its root", () => {
    expect(() => parseXml("<R/><!-- x")).toThrow(XmlError);
    expect(() => parseXml("<R/><?pi")).toThrow(XmlError);
  });
});

describe("processing instructions", () => {
  it.each([
    ["a space", "<?pi body?>", "pi", "body"],
    ["a tab", "<?pi\tbody?>", "pi", "body"],
    ["a newline", "<?pi\nbody?>", "pi", "body"],
    ["several spaces", "<?pi   a  b ?>", "pi", "a  b "],
    ["no body", "<?pi?>", "pi", ""],
    ["only whitespace", "<?pi   ?>", "pi", ""]
  ])("splits the target from the body at %s", (_, pi, target, body) => {
    expect(parse(`${pi}<R a="1"/>`)).toEqual([
      `pi ${JSON.stringify(target)} ${JSON.stringify(body)}`,
      ...ROOT_ONLY
    ]);
  });

  it.each([
    ["<? x?>", "processing instruction without a target"],
    ["<??>", "processing instruction without a target"],
    ["<?1x?>", "invalid processing instruction target: 1x"]
  ])("refuses %j", (pi, message) => {
    expect(parse(`${pi}<R a="1"/>`)[0]).toBe(`error ${message}`);
  });
});

describe("comments", () => {
  it("refuses a comment ending in --->", () => {
    expect(parse('<R a="1"><!-- a ---></R>')).toEqual([
      'open R {"a":"1"}',
      "error malformed comment",
      "close R"
    ]);
  });

  it("still accepts a single - before -->", () => {
    expect(parse('<R a="1"><!-- a - b --></R>')).toEqual(ROOT_ONLY);
  });
});

describe("CDATA outside the root", () => {
  it("is refused in a document", () => {
    expect(parse('<![CDATA[x]]><R a="1"/>')[0]).toBe("error CDATA section outside of root node");
  });

  it("is accepted in a fragment", () => {
    expect(parse("<![CDATA[x]]><R/>", { fragment: true })).not.toContainEqual(
      expect.stringMatching(/^error/)
    );
  });
});

describe("<! that can be nothing", () => {
  it("is refused at the first character that rules out every construct, keeping the markup after it", () => {
    expect(parse('<R a="1"><!x></R>')).toEqual([
      'open R {"a":"1"}',
      "error incorrect syntax",
      'text ">"',
      "close R"
    ]);
  });
});

describe("entity references", () => {
  it.each(["constructor", "toString", "__proto__", "hasOwnProperty"])(
    "does not resolve &%s; from Object.prototype",
    name => {
      expect(parse(`<R>&${name};</R>`)).toEqual([
        "open R {}",
        `error undefined entity: ${name}`,
        `text "&${name};"`,
        "close R"
      ]);
      expect(parse(`<R a="&${name};"/>`)[0]).toBe(`error undefined entity: ${name}`);
    }
  );

  it("still resolves an entity a caller defined under such a name", () => {
    const parser = new SaxParser();
    parser.ENTITIES["constructor"] = "C";
    let text = "";
    parser.on("text", t => (text += t));
    parser.write("<R>&constructor;</R>").close();
    expect(text).toBe("C");
  });

  it.each(["#65abc", "#x41g", "#", "#x", "#1.5", "#x-41"])(
    "refuses the character reference &%s;",
    ref => {
      expect(parse(`<R>&${ref};</R>`)).toContain(`error invalid character entity: &${ref};`);
    }
  );

  it("reads well-formed character references", () => {
    expect(parse("<R>&#65;&#x42;&#X43;&#0068;</R>")).toEqual([
      "open R {}",
      'text "ABCD"',
      "close R"
    ]);
  });

  it("does not accumulate expansions when nothing reads the text", () => {
    const parser = new SaxParser();
    parser.write("<R>");
    for (let i = 0; i < 1000; i++) {
      parser.write("<a>&amp;&lt;</a>");
    }
    // Reaching into the buffer is the only way to see a leak that changes no event.
    expect((parser as unknown as { text: string }).text).toBe("");
    parser.write("</R>").close();
  });
});

describe("namespace well-formedness", () => {
  it("refuses undeclaring a prefix", () => {
    expect(parse('<R xmlns:p=""/>', { xmlns: true })[0]).toBe(
      'error the "p" prefix must not be bound to an empty namespace name'
    );
  });

  it("still allows undeclaring the default namespace", () => {
    expect(parse('<R xmlns="urn:x"><c xmlns=""/></R>', { xmlns: true })).not.toContainEqual(
      expect.stringMatching(/^error/)
    );
  });

  it("refuses an element with the xmlns prefix", () => {
    expect(parse("<xmlns:R/>", { xmlns: true })[0]).toBe(
      'error invalid element name "xmlns:R": the "xmlns" prefix is reserved for declarations'
    );
  });
});
