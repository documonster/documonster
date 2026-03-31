/**
 * XML Module Security Tests
 *
 * Tests for all known XML attack vectors. Covers:
 * - XXE (external entity injection)
 * - Billion Laughs (entity expansion DoS)
 * - Numeric entity validation
 * - Prototype pollution
 * - Nesting depth limits
 * - Entity name length limits
 * - Quadratic blowup
 */

import { describe, it, expect } from "vitest";
import { SaxParser } from "@xml/sax";
import { parseXml } from "@xml/dom";
import { xmlDecode, xmlEncode } from "@xml/encode";
import { XmlError, XmlParseError, XmlWriteError, isXmlError, isXmlParseError } from "@xml/errors";

// =============================================================================
// XXE (XML External Entity) Prevention
// =============================================================================

describe("XXE prevention", () => {
  it("should not resolve external SYSTEM entities", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    // DOCTYPE is parsed but SYSTEM entities are never resolved
    parser.write('<?xml version="1.0"?>');
    parser.write("<root>&xxe;</root>");
    parser.close();
    // The entity should be undefined, not resolved
    expect(errors.some(e => e.includes("undefined entity"))).toBe(true);
    expect(texts.join("")).not.toContain("/etc/passwd");
  });

  it("should not resolve external parameter entities", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&xxe;</root>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });
});

// =============================================================================
// Billion Laughs / Entity Expansion DoS
// =============================================================================

describe("Billion Laughs protection", () => {
  it("should enforce entity expansion limit", () => {
    const parser = new SaxParser({ maxEntityExpansions: 5 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    // Add a custom entity and reference it many times
    parser.ENTITIES["x"] = "boom";
    parser.write("<root>&x;&x;&x;&x;&x;&x;&x;&x;&x;&x;</root>");
    parser.close();
    expect(errors.some(e => e.includes("entity expansion limit"))).toBe(true);
  });

  it("should count each entity expansion individually", () => {
    const parser = new SaxParser({ maxEntityExpansions: 3 });
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    parser.ENTITIES["a"] = "A";
    parser.write("<root>&a;&a;&a;&a;</root>");
    parser.close();
    // First 3 should resolve, 4th should trigger limit
    expect(errors.some(e => e.includes("entity expansion limit"))).toBe(true);
  });

  it("should allow exactly N expansions when limit is N", () => {
    const parser = new SaxParser({ maxEntityExpansions: 3 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.ENTITIES["a"] = "A";
    parser.write("<root>&a;&a;&a;</root>");
    parser.close();
    // Exactly 3 — should NOT trigger limit
    expect(errors.filter(e => e.includes("entity expansion limit"))).toEqual([]);
  });

  it("should default to 10000 expansion limit", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.ENTITIES["a"] = "A";
    // Generate 10001 entity references
    let xml = "<root>";
    for (let i = 0; i <= 10000; i++) {
      xml += "&a;";
    }
    xml += "</root>";
    parser.write(xml);
    parser.close();
    expect(errors.some(e => e.includes("entity expansion limit"))).toBe(true);
  });

  it("should allow disabling limit with 0", () => {
    const parser = new SaxParser({ maxEntityExpansions: 0 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.ENTITIES["a"] = "A";
    let xml = "<root>";
    for (let i = 0; i < 100; i++) {
      xml += "&a;";
    }
    xml += "</root>";
    parser.write(xml);
    parser.close();
    expect(errors.filter(e => e.includes("entity expansion limit"))).toEqual([]);
  });

  it("should not count predefined entities against expansion limit", () => {
    const parser = new SaxParser({ maxEntityExpansions: 3 });
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    // 8 predefined entity refs with a limit of 3 — should all resolve fine
    parser.write("<root>&lt;&gt;&amp;&quot;&apos;&lt;&gt;&amp;</root>");
    parser.close();
    expect(errors.filter(e => e.includes("entity expansion limit"))).toEqual([]);
    expect(texts.join("")).toBe("<>&\"'<>&");
  });
});

// =============================================================================
// Numeric Entity Security
// =============================================================================

describe("Numeric entity security", () => {
  it("should reject null character &#0;", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&#0;</root>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });

  it("should reject surrogate halves &#xD800;", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&#xD800;</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid character entity"))).toBe(true);
  });

  it("should reject surrogate halves &#xDFFF;", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&#xDFFF;</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid character entity"))).toBe(true);
  });

  it("should reject out-of-range &#x110000;", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&#x110000;</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid character entity"))).toBe(true);
  });

  it("should reject very large numeric entities", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>&#9999999;</root>");
    parser.close();
    expect(errors.length).toBeGreaterThan(0);
  });

  it("should NOT crash with RangeError on invalid code points", () => {
    // This is the exact CVE from a popular parser
    expect(() => {
      const parser = new SaxParser();
      parser.on("error", () => {}); // suppress
      parser.write("<root>&#xFFFFFF;</root>");
      parser.close();
    }).not.toThrow();
  });

  it("should accept valid high code point &#x10FFFF;", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>&#x10FFFF;</root>");
    parser.close();
    expect(texts[0]).toBe(String.fromCodePoint(0x10ffff));
  });

  it("should accept normal numeric entities", () => {
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.write("<root>&#65;&#x42;&#8364;</root>"); // A, B, Euro sign
    parser.close();
    expect(texts.join("")).toBe("AB\u20AC");
  });
});

// =============================================================================
// xmlDecode security
// =============================================================================

describe("xmlDecode security", () => {
  it("should not crash on invalid code points", () => {
    expect(xmlDecode("&#0;")).toBe("&#0;");
    expect(xmlDecode("&#xD800;")).toBe("&#xD800;");
    expect(xmlDecode("&#xDFFF;")).toBe("&#xDFFF;");
    expect(xmlDecode("&#x110000;")).toBe("&#x110000;");
    expect(xmlDecode("&#9999999;")).toBe("&#9999999;");
  });

  it("should reject surrogate halves", () => {
    expect(xmlDecode("&#55296;")).toBe("&#55296;"); // 0xD800 in decimal
    expect(xmlDecode("&#57343;")).toBe("&#57343;"); // 0xDFFF in decimal
  });

  it("should handle valid code points", () => {
    expect(xmlDecode("&#65;")).toBe("A");
    expect(xmlDecode("&#x10FFFF;")).toBe(String.fromCodePoint(0x10ffff));
  });
});

// =============================================================================
// Prototype Pollution Prevention
// =============================================================================

describe("Prototype pollution prevention", () => {
  it("should not pollute Object.prototype via __proto__ element name", () => {
    const before = ({} as any).polluted;
    const doc = parseXml("<root><__proto__><polluted>hacked</polluted></__proto__></root>");
    const after = ({} as any).polluted;
    expect(before).toBeUndefined();
    expect(after).toBeUndefined();
    // The element should still exist in the tree but not pollute prototypes
    expect(doc.root.children.length).toBe(1);
  });

  it("should not pollute via constructor.prototype", () => {
    const doc = parseXml(
      "<root><constructor><prototype><isAdmin>true</isAdmin></prototype></constructor></root>"
    );
    expect(({} as any).isAdmin).toBeUndefined();
    expect(doc.root.children.length).toBe(1);
  });

  it("should strip __proto__ from attribute names", () => {
    const doc = parseXml('<root __proto__="polluted"/>');
    expect(doc.root.attributes.__proto__).toBeUndefined();
    expect(doc.root.attributes["__proto__"]).toBeUndefined();
  });

  it("should strip constructor from attribute names", () => {
    const doc = parseXml('<root constructor="polluted"/>');
    // The attribute key "constructor" is dangerous and should be filtered
    expect(doc.root.attributes["constructor"]).toBeUndefined();
  });

  it("should use null-prototype objects for attributes", () => {
    const doc = parseXml('<root a="1"/>');
    // Object.create(null) has no prototype methods
    expect(Object.getPrototypeOf(doc.root.attributes)).toBeNull();
  });
});

// =============================================================================
// Nesting Depth Limit
// =============================================================================

describe("Nesting depth limit", () => {
  it("should enforce default depth limit (256)", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));

    let xml = "";
    for (let i = 0; i < 300; i++) {
      xml += "<a>";
    }
    for (let i = 0; i < 300; i++) {
      xml += "</a>";
    }
    parser.write(xml);
    parser.close();
    expect(errors.some(e => e.includes("nesting depth"))).toBe(true);
  });

  it("should allow configurable depth limit", () => {
    const parser = new SaxParser({ maxDepth: 5 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));

    let xml = "";
    for (let i = 0; i < 10; i++) {
      xml += "<a>";
    }
    for (let i = 0; i < 10; i++) {
      xml += "</a>";
    }
    parser.write(xml);
    parser.close();
    expect(errors.some(e => e.includes("nesting depth"))).toBe(true);
  });

  it("should allow exactly maxDepth levels", () => {
    const parser = new SaxParser({ maxDepth: 3 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<a><b><c/></b></a>");
    parser.close();
    expect(errors.filter(e => e.includes("nesting depth"))).toEqual([]);
  });

  it("should reject maxDepth+1 levels", () => {
    const parser = new SaxParser({ maxDepth: 3 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<a><b><c><d/></c></b></a>");
    parser.close();
    expect(errors.some(e => e.includes("nesting depth"))).toBe(true);
  });

  it("should allow disabling depth limit with 0", () => {
    const parser = new SaxParser({ maxDepth: 0 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));

    let xml = "";
    for (let i = 0; i < 500; i++) {
      xml += "<a>";
    }
    for (let i = 0; i < 500; i++) {
      xml += "</a>";
    }
    parser.write(xml);
    parser.close();
    expect(errors.filter(e => e.includes("nesting depth"))).toEqual([]);
  });

  it("should enforce depth limit via parseXml", () => {
    expect(() => {
      let xml = "";
      for (let i = 0; i < 300; i++) {
        xml += "<a>";
      }
      for (let i = 0; i < 300; i++) {
        xml += "</a>";
      }
      parseXml(xml);
    }).toThrow(/nesting depth/);
  });
});

// =============================================================================
// Entity Name Length Limit
// =============================================================================

describe("Entity name length limit", () => {
  it("should reject entity names longer than 64 characters", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    const longName = "a".repeat(100);
    parser.write(`<root>&${longName};</root>`);
    parser.close();
    expect(errors.some(e => e.includes("entity name too long"))).toBe(true);
  });

  it("should accept entity names up to 64 characters", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.ENTITIES["a".repeat(64)] = "ok";
    parser.write(`<root>&${"a".repeat(64)};</root>`);
    parser.close();
    // Should resolve without "too long" error
    expect(errors.filter(e => e.includes("entity name too long"))).toEqual([]);
  });
});

// =============================================================================
// Regex Safety
// =============================================================================

describe("Regex safety (ReDoS prevention)", () => {
  it("should not hang on long whitespace sequences in attributes", () => {
    const start = Date.now();
    const parser = new SaxParser();
    parser.on("error", () => {});
    // Crafted input with many spaces
    parser.write(`<root attr="${" ".repeat(10000)}" />`);
    parser.close();
    expect(Date.now() - start).toBeLessThan(1000); // Should finish in < 1s
  });

  it("should not hang on many ampersands", () => {
    const start = Date.now();
    const parser = new SaxParser();
    parser.on("error", () => {});
    parser.write("<root>" + "&".repeat(10000) + "</root>");
    parser.close();
    expect(Date.now() - start).toBeLessThan(1000);
  });
});

// =============================================================================
// No User Input in RegExp
// =============================================================================

describe("No dynamic RegExp construction", () => {
  it("should not allow entity names with regex metacharacters to affect parsing", () => {
    // This was a critical vulnerability in a popular parser
    // Entity name containing "." should not act as regex wildcard
    const parser = new SaxParser();
    const texts: string[] = [];
    parser.on("text", t => texts.push(t));
    parser.on("error", () => {});
    // Even if someone adds a custom entity with "." in the name,
    // it should only match exactly, not as a regex
    parser.ENTITIES["l."] = "INJECTED";
    parser.write("<root>&lt;</root>");
    parser.close();
    // &lt; should resolve to "<", NOT to "INJECTED"
    expect(texts[0]).toBe("<");
  });
});

// =============================================================================
// Entity Expansion in Attribute Values
// =============================================================================

describe("Entity expansion in attribute values", () => {
  it("should count attribute entity expansions toward the limit", () => {
    const parser = new SaxParser({ maxEntityExpansions: 3 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.ENTITIES["x"] = "X";
    // 4 entity refs in attribute values should exceed limit of 3
    parser.write('<root a="&x;&x;&x;&x;"/>');
    parser.close();
    expect(errors.some(e => e.includes("entity expansion limit"))).toBe(true);
  });
});

// =============================================================================
// < in Attribute Values
// =============================================================================

describe("< in attribute values", () => {
  it("should report error for < in attribute value", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write('<root attr="a<b"/>');
    parser.close();
    expect(errors.some(e => e.includes("<") || e.includes("not allowed"))).toBe(true);
  });
});

// =============================================================================
// Nesting Depth with Self-Closing Tags
// =============================================================================

describe("Nesting depth with self-closing tags", () => {
  it("should enforce maxDepth at boundary with self-closing tag", () => {
    const parser = new SaxParser({ maxDepth: 3 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    // 3 levels deep, then self-closing at level 4 — should trigger
    parser.write("<a><b><c><d/></c></b></a>");
    parser.close();
    expect(errors.some(e => e.includes("nesting depth"))).toBe(true);
  });
});

// =============================================================================
// Prototype Pollution — Extended Dangerous Keys
// =============================================================================

describe("Prototype pollution — extended dangerous keys", () => {
  it("should strip __defineGetter__ from attribute names", () => {
    const doc = parseXml('<root __defineGetter__="val"/>');
    expect(doc.root.attributes["__defineGetter__"]).toBeUndefined();
  });

  it("should strip __defineSetter__ from attribute names", () => {
    const doc = parseXml('<root __defineSetter__="val"/>');
    expect(doc.root.attributes["__defineSetter__"]).toBeUndefined();
  });

  it("should strip __lookupGetter__ from attribute names", () => {
    const doc = parseXml('<root __lookupGetter__="val"/>');
    expect(doc.root.attributes["__lookupGetter__"]).toBeUndefined();
  });

  it("should strip __lookupSetter__ from attribute names", () => {
    const doc = parseXml('<root __lookupSetter__="val"/>');
    expect(doc.root.attributes["__lookupSetter__"]).toBeUndefined();
  });

  it("should strip prototype from attribute names", () => {
    const doc = parseXml('<root prototype="val"/>');
    expect(doc.root.attributes["prototype"]).toBeUndefined();
  });
});

// =============================================================================
// Type Guards
// =============================================================================

describe("Error type guards", () => {
  it("isXmlError should return true for XmlError", () => {
    expect(isXmlError(new XmlError("test"))).toBe(true);
  });

  it("isXmlError should return true for XmlParseError", () => {
    expect(isXmlError(new XmlParseError("test"))).toBe(true);
  });

  it("isXmlError should return true for XmlWriteError", () => {
    expect(isXmlError(new XmlWriteError("op", "state"))).toBe(true);
  });

  it("isXmlError should return false for plain Error", () => {
    expect(isXmlError(new Error("test"))).toBe(false);
  });

  it("isXmlError should return false for non-errors", () => {
    expect(isXmlError("string")).toBe(false);
    expect(isXmlError(null)).toBe(false);
    expect(isXmlError(undefined)).toBe(false);
  });

  it("isXmlParseError should return true for XmlParseError", () => {
    expect(isXmlParseError(new XmlParseError("test"))).toBe(true);
  });

  it("isXmlParseError should return false for XmlError", () => {
    expect(isXmlParseError(new XmlError("test"))).toBe(false);
  });

  it("isXmlParseError should return false for plain Error", () => {
    expect(isXmlParseError(new Error("test"))).toBe(false);
  });
});

// =============================================================================
// XmlParseError Construction
// =============================================================================

describe("XmlParseError construction", () => {
  it("should include line/column in message when provided", () => {
    const err = new XmlParseError("test", { line: 5, column: 10 });
    expect(err.message).toContain("5:10");
    expect(err.line).toBe(5);
    expect(err.column).toBe(10);
  });

  it("should include fileName in message when provided", () => {
    const err = new XmlParseError("test", { fileName: "input.xml", line: 1, column: 0 });
    expect(err.message).toContain("input.xml");
    expect(err.fileName).toBe("input.xml");
  });

  it("should work without context", () => {
    const err = new XmlParseError("test message");
    expect(err.message).toBe("test message");
    expect(err.line).toBeUndefined();
    expect(err.column).toBeUndefined();
    expect(err.fileName).toBeUndefined();
  });
});

// =============================================================================
// XmlWriteError Construction
// =============================================================================

describe("XmlWriteError construction", () => {
  it("should expose operation and state", () => {
    const err = new XmlWriteError("close node", "no element is open");
    expect(err.operation).toBe("close node");
    expect(err.state).toBe("no element is open");
    expect(err.message).toContain("close node");
    expect(err.message).toContain("no element is open");
  });
});

// =============================================================================
// Quadratic Blowup Protection
// =============================================================================

describe("Quadratic blowup protection", () => {
  it("should enforce expansion limit for large entity values", () => {
    const parser = new SaxParser({ maxEntityExpansions: 5 });
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    // Large entity value, referenced many times
    parser.ENTITIES["big"] = "A".repeat(10000);
    parser.write("<root>&big;&big;&big;&big;&big;&big;&big;</root>");
    parser.close();
    expect(errors.some(e => e.includes("entity expansion limit"))).toBe(true);
  });
});

// =============================================================================
// Text After Root Close
// =============================================================================

describe("Text after root close", () => {
  it("should detect non-whitespace text after root close tag", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root/>text after");
    parser.close();
    expect(errors.some(e => e.includes("outside of root") || e.includes("text data"))).toBe(true);
  });
});

// =============================================================================
// Close Tag Name Validation
// =============================================================================

describe("Close tag name validation", () => {
  it("should reject close tag starting with a digit", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root></123></root>");
    parser.close();
    expect(errors.some(e => e.includes("unexpected character") || e.includes("unmatched"))).toBe(
      true
    );
  });

  it("should accept close tag starting with a letter", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<child></child>");
    parser.close();
    expect(errors.filter(e => e.includes("unexpected character"))).toEqual([]);
  });
});

// =============================================================================
// Attribute < Recovery
// =============================================================================

describe("Attribute < error recovery", () => {
  it("should recover and continue parsing after < in attribute value", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const tags: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("opentag", tag => tags.push(tag.name));
    // After the < error in attr, parser should recover and find <next/>
    parser.write('<root attr="a<next/>"');
    parser.close();
    // Should have reported the < error
    expect(errors.some(e => e.includes("<") || e.includes("not allowed"))).toBe(true);
    // Should have attempted to parse 'next' as a tag
    expect(tags).toContain("next");
  });
});

// =============================================================================
// write(null) equivalence
// =============================================================================

describe("write(null) equivalence", () => {
  it("write(null) should be equivalent to close()", () => {
    const parser = new SaxParser();
    const tags: string[] = [];
    parser.on("opentag", tag => tags.push(tag.name));
    parser.write("<root/>");
    parser.write(null);
    expect(parser.closed).toBe(true);
    expect(tags).toEqual(["root"]);
  });
});

// =============================================================================
// Lone Surrogate in xmlEncode — security
// =============================================================================

describe("xmlEncode lone surrogate security", () => {
  it("should strip lone surrogates that would produce invalid XML", () => {
    const malicious = "data" + String.fromCharCode(0xd800) + "more";
    const encoded = xmlEncode(malicious);
    expect(encoded).not.toContain(String.fromCharCode(0xd800));
    expect(encoded).toBe("datamore");
  });
});

// =============================================================================
// Invalid XML 1.0 Character Rejection
// =============================================================================

describe("Invalid XML 1.0 character rejection", () => {
  it("should report error for null byte in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a\x00b</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should report error for control character 0x01 in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a\x01b</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should report error for 0x0B (vertical tab) in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a\x0Bb</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should report error for 0x0C (form feed) in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a\x0Cb</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should report error for DEL (0x7F) in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a\x7Fb</root>");
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should report error for control character in attribute value", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write('<root attr="a\x01b"/>');
    parser.close();
    expect(errors.some(e => e.includes("invalid XML character"))).toBe(true);
  });

  it("should accept valid whitespace (TAB, LF, CR) in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    parser.write("<root>a\tb\nc\rd</root>");
    parser.close();
    expect(errors.filter(e => e.includes("invalid XML character"))).toEqual([]);
    // CR is normalized to LF per XML spec
    expect(texts.join("")).toBe("a\tb\nc\nd");
  });

  it("should report error for lone high surrogate in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a" + String.fromCharCode(0xd800) + "b</root>");
    parser.close();
    expect(errors.some(e => e.includes("lone surrogate"))).toBe(true);
  });

  it("should report error for lone low surrogate in text", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.write("<root>a" + String.fromCharCode(0xdc00) + "b</root>");
    parser.close();
    expect(errors.some(e => e.includes("lone surrogate"))).toBe(true);
  });

  it("should accept valid surrogate pair", () => {
    const parser = new SaxParser();
    const errors: string[] = [];
    const texts: string[] = [];
    parser.on("error", err => errors.push(err.message));
    parser.on("text", t => texts.push(t));
    parser.write("<root>\uD83D\uDE00</root>"); // 😀
    parser.close();
    expect(
      errors.filter(e => e.includes("invalid XML character") || e.includes("surrogate"))
    ).toEqual([]);
    expect(texts.join("")).toBe("😀");
  });
});
