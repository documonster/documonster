import { xmlEncode, xmlDecode, xmlEncodeAttr, encodeCData } from "@xml/encode";
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

  it("should encode CR / LF / tab as numeric character references", () => {
    // XML 1.0 §3.3.3 normalises literal whitespace inside attribute
    // values to a single space at parse time. Numeric character
    // references survive verbatim, so the only way to preserve tabs /
    // newlines / carriage returns through round-trip is to encode them.
    expect(xmlEncodeAttr("first\nsecond")).toBe("first&#xA;second");
    expect(xmlEncodeAttr("col1\tcol2")).toBe("col1&#x9;col2");
    expect(xmlEncodeAttr("dos\r\nline")).toBe("dos&#xD;&#xA;line");
  });

  it("should combine entity encoding with whitespace preservation", () => {
    expect(xmlEncodeAttr('a "quote"\nwith newline')).toBe("a &quot;quote&quot;&#xA;with newline");
  });

  it("should fast-path values with no whitespace to encode", () => {
    expect(xmlEncodeAttr("plain value")).toBe("plain value");
    expect(xmlEncodeAttr('has "quote"')).toBe("has &quot;quote&quot;");
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

  // Regression — xmlDecode must not let a numeric character reference
  // re-introduce a control character that XML 1.0 forbids. Previously
  // these tests lived in the Excel chart suite as incidental coverage;
  // they belong in the xml module's own test file where the helper
  // is defined.
  it("refuses numeric references to illegal XML chars (C0 controls, noncharacters)", () => {
    // Forbidden C0 control (BEL) — decimal and hex forms.
    expect(xmlDecode("&#7;")).toBe("&#7;");
    expect(xmlDecode("&#x7;")).toBe("&#x7;");
    expect(xmlDecode("&#8;")).toBe("&#8;");
    // Noncharacter — must be rejected.
    expect(xmlDecode("&#xFFFE;")).toBe("&#xFFFE;");
    expect(xmlDecode("&#xFFFF;")).toBe("&#xFFFF;");
    // Allowed whitespace — must still decode.
    expect(xmlDecode("&#9;")).toBe("\t");
    expect(xmlDecode("&#10;")).toBe("\n");
    expect(xmlDecode("&#13;")).toBe("\r");
  });
});

// =============================================================================
// encodeCData
// =============================================================================

describe("encodeCData", () => {
  it("strips illegal XML chars before wrapping", () => {
    // BEL inside user text — must be stripped, not wrapped.
    expect(encodeCData("before\u0007after")).toBe("<![CDATA[beforeafter]]>");
  });

  it("splits `]]>` sequences across adjacent CDATA sections", () => {
    // The CDATA end marker cannot appear inside a CDATA section.
    expect(encodeCData("a]]>b")).toBe("<![CDATA[a]]]]><![CDATA[>b]]>");
  });

  it("preserves legal whitespace and surrogate pairs", () => {
    // Tabs / newlines / CR are legal XML chars; surrogate pairs
    // encoding a valid supplementary code point stay together.
    expect(encodeCData("tab\tnewline\ncarriage\rend")).toBe(
      "<![CDATA[tab\tnewline\ncarriage\rend]]>"
    );
    // 😀 is U+1F600, encoded as surrogate pair D83D DE00.
    expect(encodeCData("😀")).toBe("<![CDATA[😀]]>");
  });

  it("strips DEL (0x7F) as a project-policy extension", () => {
    // DEL is technically legal XML but we strip it to match the rest
    // of the sanitiser chain (some downstream consumers choke on it).
    expect(encodeCData("a\u007Fb")).toBe("<![CDATA[ab]]>");
  });

  it("strips lone UTF-16 surrogate halves", () => {
    // Lone high surrogate (D800) and lone low surrogate (DC00) are
    // invalid on their own.
    const loneHigh = String.fromCharCode(0xd800);
    const loneLow = String.fromCharCode(0xdc00);
    expect(encodeCData(`a${loneHigh}b`)).toBe("<![CDATA[ab]]>");
    expect(encodeCData(`a${loneLow}b`)).toBe("<![CDATA[ab]]>");
  });
});
