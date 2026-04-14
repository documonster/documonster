import { parseXml } from "@xml/dom";
import { SaxParser } from "@xml/sax";
import { parseXmlToObject } from "@xml/to-object";
import { describe, it, expect } from "vitest";

// =============================================================================
// Test Helpers
// =============================================================================

/** Collect all text events from a SAX parse. */
function collectText(xml: string, options?: ConstructorParameters<typeof SaxParser>[0]): string[] {
  const parser = new SaxParser(options);
  const texts: string[] = [];
  parser.on("text", text => texts.push(text));
  parser.on("error", () => {}); // swallow errors so parsing continues in "error" mode
  parser.write(xml);
  parser.close();
  return texts;
}

/** Collect all error messages from a SAX parse. */
function collectErrors(
  xml: string,
  options?: ConstructorParameters<typeof SaxParser>[0]
): string[] {
  const parser = new SaxParser(options);
  const errors: string[] = [];
  parser.on("error", err => errors.push(err.message));
  parser.write(xml);
  parser.close();
  return errors;
}

// =============================================================================
// SaxParser — invalidCharHandling
// =============================================================================

describe("SaxParser invalidCharHandling", () => {
  // ---------------------------------------------------------------------------
  // "error" mode (default)
  // ---------------------------------------------------------------------------

  describe('mode: "error" (default)', () => {
    it("should throw on 0x7F (DEL) by default", () => {
      const parser = new SaxParser();
      parser.on("text", () => {});
      expect(() => {
        parser.write("<root>hello\x7fworld</root>");
        parser.close();
      }).toThrow("invalid XML character: 0x7f");
    });

    it("should throw on ASCII control char 0x01", () => {
      const parser = new SaxParser();
      parser.on("text", () => {});
      expect(() => {
        parser.write("<root>hello\x01world</root>");
        parser.close();
      }).toThrow("invalid XML character: 0x1");
    });

    it("should report error via handler when registered", () => {
      const errors = collectErrors("<root>hello\x7fworld</root>");
      expect(errors.some(e => e.includes("0x7f"))).toBe(true);
    });
  });

  // ---------------------------------------------------------------------------
  // "skip" mode
  // ---------------------------------------------------------------------------

  describe('mode: "skip"', () => {
    it("should skip 0x7F (DEL) in text content", () => {
      const texts = collectText("<root>hello\x7fworld</root>", {
        invalidCharHandling: "skip"
      });
      expect(texts).toEqual(["helloworld"]);
    });

    it("should skip ASCII control char 0x01 in text content", () => {
      const texts = collectText("<root>hello\x01world</root>", {
        invalidCharHandling: "skip"
      });
      expect(texts).toEqual(["helloworld"]);
    });

    it("should skip multiple invalid chars", () => {
      const texts = collectText("<root>\x01a\x02b\x7fc</root>", {
        invalidCharHandling: "skip"
      });
      expect(texts).toEqual(["abc"]);
    });

    it("should skip 0x00 (NUL)", () => {
      const texts = collectText("<root>hello\x00world</root>", {
        invalidCharHandling: "skip"
      });
      expect(texts).toEqual(["helloworld"]);
    });

    it("should not produce any error events", () => {
      const errors = collectErrors("<root>hello\x7fworld</root>", {
        invalidCharHandling: "skip"
      });
      expect(errors).toEqual([]);
    });

    it("should skip invalid chars in attribute values", () => {
      const parser = new SaxParser({ invalidCharHandling: "skip" });
      let attrValue = "";
      parser.on("opentag", tag => {
        attrValue = tag.attributes.val ?? "";
      });
      parser.write('<root val="hello\x7fworld"/>');
      parser.close();
      expect(attrValue).toBe("helloworld");
    });

    it("should skip invalid chars in attribute with non-ASCII neighbors", () => {
      const parser = new SaxParser({ invalidCharHandling: "skip" });
      let attrValue = "";
      parser.on("opentag", tag => {
        attrValue = tag.attributes.val ?? "";
      });
      parser.write('<root val="日本\x7f語"/>');
      parser.close();
      expect(attrValue).toBe("日本語");
    });

    it("should skip invalid chars in attribute with emoji neighbors", () => {
      const parser = new SaxParser({ invalidCharHandling: "skip" });
      let attrValue = "";
      parser.on("opentag", tag => {
        attrValue = tag.attributes.val ?? "";
      });
      parser.write('<root val="😀\x01😀"/>');
      parser.close();
      expect(attrValue).toBe("😀😀");
    });

    it("should handle text that is entirely invalid chars", () => {
      const texts = collectText("<root>\x01\x02\x03</root>", {
        invalidCharHandling: "skip"
      });
      // Text handler may not be called, or may be called with empty string
      const combined = texts.join("");
      expect(combined).toBe("");
    });

    it("should preserve valid XML characters around skipped ones", () => {
      const texts = collectText("<root>abc\x0bdef\x0cghi</root>", {
        invalidCharHandling: "skip"
      });
      expect(texts).toEqual(["abcdefghi"]);
    });

    it("should handle a long run of consecutive invalid chars without stack overflow", () => {
      // 50000 consecutive NUL bytes — would overflow stack with recursive approach
      const badChars = "\x00".repeat(50000);
      const texts = collectText(`<root>before${badChars}after</root>`, {
        invalidCharHandling: "skip"
      });
      expect(texts.join("")).toBe("beforeafter");
    });
  });

  // ---------------------------------------------------------------------------
  // "replace" mode
  // ---------------------------------------------------------------------------

  describe('mode: "replace"', () => {
    it("should replace 0x7F with U+FFFD in text content", () => {
      const texts = collectText("<root>hello\x7fworld</root>", {
        invalidCharHandling: "replace"
      });
      expect(texts).toEqual(["hello\uFFFDworld"]);
    });

    it("should replace ASCII control char 0x01 with U+FFFD", () => {
      const texts = collectText("<root>hello\x01world</root>", {
        invalidCharHandling: "replace"
      });
      expect(texts).toEqual(["hello\uFFFDworld"]);
    });

    it("should replace multiple invalid chars", () => {
      const texts = collectText("<root>\x01a\x02b\x7fc</root>", {
        invalidCharHandling: "replace"
      });
      expect(texts).toEqual(["\uFFFDa\uFFFDb\uFFFDc"]);
    });

    it("should not produce any error events", () => {
      const errors = collectErrors("<root>hello\x7fworld</root>", {
        invalidCharHandling: "replace"
      });
      expect(errors).toEqual([]);
    });

    it("should replace invalid chars in attribute values", () => {
      const parser = new SaxParser({ invalidCharHandling: "replace" });
      let attrValue = "";
      parser.on("opentag", tag => {
        attrValue = tag.attributes.val ?? "";
      });
      parser.write('<root val="hello\x7fworld"/>');
      parser.close();
      expect(attrValue).toBe("hello\uFFFDworld");
    });

    it("should replace invalid chars in attribute with non-ASCII neighbors", () => {
      const parser = new SaxParser({ invalidCharHandling: "replace" });
      let attrValue = "";
      parser.on("opentag", tag => {
        attrValue = tag.attributes.val ?? "";
      });
      parser.write('<root val="日本\x7f語"/>');
      parser.close();
      expect(attrValue).toBe("日本\uFFFD語");
    });
  });

  // ---------------------------------------------------------------------------
  // Valid characters should be unaffected in all modes
  // ---------------------------------------------------------------------------

  describe("valid characters unaffected", () => {
    const modes = ["error", "skip", "replace"] as const;

    for (const mode of modes) {
      it(`should preserve TAB, LF, CR in "${mode}" mode`, () => {
        const texts = collectText("<root>a\tb\nc\rd</root>", {
          invalidCharHandling: mode
        });
        // CR is normalized to LF by XML spec
        const combined = texts.join("");
        expect(combined).toContain("a\tb");
        expect(combined).toContain("\n");
      });

      it(`should preserve non-ASCII characters in "${mode}" mode`, () => {
        const texts = collectText("<root>日本語</root>", {
          invalidCharHandling: mode
        });
        expect(texts).toEqual(["日本語"]);
      });

      it(`should preserve emoji (surrogate pairs) in "${mode}" mode`, () => {
        const texts = collectText("<root>hello 😀 world</root>", {
          invalidCharHandling: mode
        });
        expect(texts).toEqual(["hello 😀 world"]);
      });
    }
  });
});

// =============================================================================
// parseXml (DOM) — invalidCharHandling propagation
// =============================================================================

describe("parseXml invalidCharHandling", () => {
  it("should throw by default on invalid chars", () => {
    expect(() => parseXml("<root>hello\x7fworld</root>")).toThrow("invalid XML character");
  });

  it('should skip invalid chars with "skip"', () => {
    const doc = parseXml("<root>hello\x7fworld</root>", { invalidCharHandling: "skip" });
    const text = doc.root.children[0];
    expect(text.type).toBe("text");
    if (text.type === "text") {
      expect(text.value).toBe("helloworld");
    }
  });

  it('should replace invalid chars with "replace"', () => {
    const doc = parseXml("<root>hello\x7fworld</root>", { invalidCharHandling: "replace" });
    const text = doc.root.children[0];
    expect(text.type).toBe("text");
    if (text.type === "text") {
      expect(text.value).toBe("hello\uFFFDworld");
    }
  });
});

// =============================================================================
// parseXmlToObject — invalidCharHandling propagation
// =============================================================================

describe("parseXmlToObject invalidCharHandling", () => {
  it("should throw by default on invalid chars", () => {
    expect(() => parseXmlToObject("<root>hello\x7fworld</root>")).toThrow("invalid XML character");
  });

  it('should skip invalid chars with "skip"', () => {
    const obj = parseXmlToObject("<root>hello\x7fworld</root>", {
      invalidCharHandling: "skip"
    });
    expect(obj).toEqual({ root: "helloworld" });
  });

  it('should replace invalid chars with "replace"', () => {
    const obj = parseXmlToObject("<root>hello\x7fworld</root>", {
      invalidCharHandling: "replace"
    });
    expect(obj).toEqual({ root: "hello\uFFFDworld" });
  });
});

// =============================================================================
// Chunk boundary & edge cases
// =============================================================================

describe("invalidCharHandling chunk boundaries", () => {
  it("skip: invalid char at end of text before closing tag (inside root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "skip" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>hello\x7f</root>");
    parser.close();
    expect(texts).toEqual(["hello"]);
  });

  it("replace: invalid char at end of text before closing tag (inside root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "replace" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>hello\x7f</root>");
    parser.close();
    expect(texts).toEqual(["hello\uFFFD"]);
  });

  it("skip: chunk ends with invalid char (outside root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "skip", fragment: true });
    parser.on("text", t => texts.push(t));
    parser.write("hello\x7f");
    parser.write("<root/>");
    parser.close();
    expect(texts.join("")).toBe("hello");
  });

  it("replace: chunk ends with invalid char (outside root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "replace", fragment: true });
    parser.on("text", t => texts.push(t));
    parser.write("hello\x7f");
    parser.write("<root/>");
    parser.close();
    expect(texts.join("")).toBe("hello\uFFFD");
  });

  it("skip: multiple chunks with invalid chars at boundaries", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "skip" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>abc\x01");
    parser.write("def\x7fghi</root>");
    parser.close();
    expect(texts.join("")).toBe("abcdefghi");
  });

  it("replace: multiple chunks with invalid chars at boundaries", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "replace" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>abc\x01");
    parser.write("def\x7fghi</root>");
    parser.close();
    expect(texts.join("")).toBe("abc\uFFFDdef\uFFFDghi");
  });

  it("skip: invalid char at very end of chunk (text in root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "skip" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>hello\x7f");
    parser.write("world</root>");
    parser.close();
    expect(texts.join("")).toBe("helloworld");
  });

  it("replace: invalid char at very end of chunk (text in root)", () => {
    const texts: string[] = [];
    const parser = new SaxParser({ invalidCharHandling: "replace" });
    parser.on("text", t => texts.push(t));
    parser.write("<root>hello\x7f");
    parser.write("world</root>");
    parser.close();
    expect(texts.join("")).toBe("hello\uFFFDworld");
  });

  it("skip: invalid char in attribute at chunk boundary", () => {
    const parser = new SaxParser({ invalidCharHandling: "skip" });
    let attrValue = "";
    parser.on("opentag", tag => {
      attrValue = tag.attributes.v ?? "";
    });
    parser.write('<root v="abc\x01');
    parser.write('def"/>');
    parser.close();
    expect(attrValue).toBe("abcdef");
  });

  it("replace: invalid char in attribute at chunk boundary", () => {
    const parser = new SaxParser({ invalidCharHandling: "replace" });
    let attrValue = "";
    parser.on("opentag", tag => {
      attrValue = tag.attributes.v ?? "";
    });
    parser.write('<root v="abc\x01');
    parser.write('def"/>');
    parser.close();
    expect(attrValue).toBe("abc\uFFFDdef");
  });
});
