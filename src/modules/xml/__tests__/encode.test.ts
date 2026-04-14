import { xmlEncode, xmlDecode, xmlEncodeAttr } from "@xml/encode";
import { describe, it, expect } from "vitest";

describe("xmlEncode", () => {
  it("should return unchanged string when no special characters", () => {
    expect(xmlEncode("hello world")).toBe("hello world");
    expect(xmlEncode("abc123")).toBe("abc123");
    expect(xmlEncode("")).toBe("");
  });

  it("should encode XML entities", () => {
    expect(xmlEncode("<")).toBe("&lt;");
    expect(xmlEncode(">")).toBe("&gt;");
    expect(xmlEncode("&")).toBe("&amp;");
    expect(xmlEncode('"')).toBe("&quot;");
    expect(xmlEncode("'")).toBe("&apos;");
  });

  it("should encode mixed content", () => {
    expect(xmlEncode('<tag attr="val">text & more</tag>')).toBe(
      "&lt;tag attr=&quot;val&quot;&gt;text &amp; more&lt;/tag&gt;"
    );
  });

  it("should strip invalid control characters", () => {
    expect(xmlEncode("hello\x00world")).toBe("helloworld");
    expect(xmlEncode("test\x01\x02\x03")).toBe("test");
    expect(xmlEncode("a\x7Fb")).toBe("ab");
  });

  it("should preserve valid whitespace", () => {
    expect(xmlEncode("hello\nworld")).toBe("hello\nworld");
    expect(xmlEncode("hello\tworld")).toBe("hello\tworld");
    expect(xmlEncode("hello\rworld")).toBe("hello\rworld");
  });
});

describe("xmlDecode", () => {
  it("should return unchanged string when no entities", () => {
    expect(xmlDecode("hello world")).toBe("hello world");
    expect(xmlDecode("")).toBe("");
  });

  it("should decode named entities", () => {
    expect(xmlDecode("&lt;")).toBe("<");
    expect(xmlDecode("&gt;")).toBe(">");
    expect(xmlDecode("&amp;")).toBe("&");
    expect(xmlDecode("&quot;")).toBe('"');
    expect(xmlDecode("&apos;")).toBe("'");
  });

  it("should decode decimal numeric references", () => {
    expect(xmlDecode("&#60;")).toBe("<");
    expect(xmlDecode("&#62;")).toBe(">");
    expect(xmlDecode("&#65;")).toBe("A");
  });

  it("should decode hex numeric references", () => {
    expect(xmlDecode("&#x3C;")).toBe("<");
    expect(xmlDecode("&#x3E;")).toBe(">");
    expect(xmlDecode("&#x41;")).toBe("A");
  });

  it("should decode mixed content", () => {
    expect(xmlDecode("&lt;tag&gt; &amp; &#65;")).toBe("<tag> & A");
  });

  it("should leave unknown entities unchanged", () => {
    expect(xmlDecode("&unknown;")).toBe("&unknown;");
  });
});

describe("xmlEncodeAttr", () => {
  it("should encode attribute values", () => {
    expect(xmlEncodeAttr('value with "quotes"')).toBe("value with &quot;quotes&quot;");
    expect(xmlEncodeAttr("value with <angle>")).toBe("value with &lt;angle&gt;");
  });
});

// =============================================================================
// Lone Surrogate Handling
// =============================================================================

describe("xmlEncode lone surrogate handling", () => {
  it("should strip lone high surrogate", () => {
    const input = "a" + String.fromCharCode(0xd800) + "b";
    expect(xmlEncode(input)).toBe("ab");
  });

  it("should strip lone low surrogate", () => {
    const input = "a" + String.fromCharCode(0xdc00) + "b";
    expect(xmlEncode(input)).toBe("ab");
  });

  it("should preserve valid surrogate pair", () => {
    const emoji = "\uD83D\uDE00"; // 😀
    expect(xmlEncode(emoji)).toBe(emoji);
  });

  it("should strip multiple lone surrogates", () => {
    const input =
      String.fromCharCode(0xd800) +
      String.fromCharCode(0xdbff) +
      "text" +
      String.fromCharCode(0xdc00);
    expect(xmlEncode(input)).toBe("text");
  });

  it("should handle string of only control characters", () => {
    expect(xmlEncode("\x00\x01\x02\x08")).toBe("");
  });
});

// =============================================================================
// xmlDecode Edge Cases
// =============================================================================

describe("xmlDecode edge cases", () => {
  it("should handle &#x; (hex with no digits)", () => {
    // parseInt("", 16) returns NaN — should return original
    expect(xmlDecode("&#x;")).toBe("&#x;");
  });

  it("should handle &#; (decimal with no digits)", () => {
    // parseInt("", 10) returns NaN — should return original
    expect(xmlDecode("&#;")).toBe("&#;");
  });

  it("should handle entity-like patterns without semicolon", () => {
    expect(xmlDecode("&amp no semicolon")).toBe("&amp no semicolon");
  });

  it("should handle consecutive entities", () => {
    expect(xmlDecode("&lt;&gt;&amp;")).toBe("<>&");
  });

  it("should handle hex entity case insensitivity", () => {
    expect(xmlDecode("&#x41;")).toBe("A");
    expect(xmlDecode("&#X41;")).toBe("A");
  });

  it("should handle high unicode code point", () => {
    expect(xmlDecode("&#x1F600;")).toBe("😀");
  });
});
