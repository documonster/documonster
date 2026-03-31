import { describe, it, expect } from "vitest";
import { SaxParser } from "@xml/sax";
import { parseXml, findChild, attr } from "@xml/dom";
import type { SaxTag } from "@xml/types";

describe("Namespace support", () => {
  describe("SaxParser with xmlns", () => {
    it("should parse default namespace", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<root xmlns="http://example.com"/>');
      parser.close();

      expect(tags[0].prefix).toBe("");
      expect(tags[0].local).toBe("root");
      expect(tags[0].uri).toBe("http://example.com");
      expect(tags[0].ns).toEqual({ "": "http://example.com" });
    });

    it("should parse prefixed namespace", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<x:root xmlns:x="http://example.com"/>');
      parser.close();

      expect(tags[0].name).toBe("x:root");
      expect(tags[0].prefix).toBe("x");
      expect(tags[0].local).toBe("root");
      expect(tags[0].uri).toBe("http://example.com");
      expect(tags[0].ns).toEqual({ x: "http://example.com" });
    });

    it("should inherit namespace from parent", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<x:root xmlns:x="http://example.com"><x:child/></x:root>');
      parser.close();

      expect(tags[1].name).toBe("x:child");
      expect(tags[1].prefix).toBe("x");
      expect(tags[1].local).toBe("child");
      expect(tags[1].uri).toBe("http://example.com");
      // Child does not declare its own ns
      expect(tags[1].ns).toBeUndefined();
    });

    it("should handle multiple namespace prefixes", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<root xmlns:a="http://a.com" xmlns:b="http://b.com"><a:foo/><b:bar/></root>');
      parser.close();

      expect(tags[0].ns).toEqual({ a: "http://a.com", b: "http://b.com" });
      expect(tags[1].prefix).toBe("a");
      expect(tags[1].uri).toBe("http://a.com");
      expect(tags[2].prefix).toBe("b");
      expect(tags[2].uri).toBe("http://b.com");
    });

    it("should handle namespace override in child", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<x:root xmlns:x="http://v1"><x:child xmlns:x="http://v2"/></x:root>');
      parser.close();

      expect(tags[0].uri).toBe("http://v1");
      expect(tags[1].uri).toBe("http://v2");
      expect(tags[1].ns).toEqual({ x: "http://v2" });
    });

    it("should handle unprefixed elements with default namespace", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<root xmlns="http://default"><child/></root>');
      parser.close();

      expect(tags[0].prefix).toBe("");
      expect(tags[0].uri).toBe("http://default");
      expect(tags[1].prefix).toBe("");
      expect(tags[1].uri).toBe("http://default");
    });

    it("should not set namespace fields when xmlns is false", () => {
      const parser = new SaxParser({ xmlns: false });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      parser.write('<x:root xmlns:x="http://example.com"/>');
      parser.close();

      expect(tags[0].prefix).toBeUndefined();
      expect(tags[0].local).toBeUndefined();
      expect(tags[0].uri).toBeUndefined();
    });

    it("should handle OOXML-style namespaces", () => {
      const parser = new SaxParser({ xmlns: true });
      const tags: SaxTag[] = [];
      parser.on("opentag", tag => tags.push({ ...tag }));
      const xml = [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheetData>",
        '<row r="1"><c r="A1" t="s"><v>0</v></c></row>',
        "</sheetData>",
        "</worksheet>"
      ].join("");
      parser.write(xml);
      parser.close();

      const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      expect(tags[0].uri).toBe(main); // worksheet
      expect(tags[1].uri).toBe(main); // sheetData
      expect(tags[2].uri).toBe(main); // row
      expect(tags[3].uri).toBe(main); // c
      expect(tags[4].uri).toBe(main); // v
    });

    it("should handle self-closing tags with namespace", () => {
      const parser = new SaxParser({ xmlns: true });
      const openTags: SaxTag[] = [];
      const closeTags: SaxTag[] = [];
      parser.on("opentag", tag => openTags.push({ ...tag }));
      parser.on("closetag", tag => closeTags.push({ ...tag }));
      parser.write('<root xmlns:x="http://example.com"><x:self/><x:after/></root>');
      parser.close();

      // x:self is self-closing — namespace should still resolve
      expect(openTags[1].prefix).toBe("x");
      expect(openTags[1].uri).toBe("http://example.com");
      // x:after comes after x:self — should still resolve (scope wasn't leaked)
      expect(openTags[2].prefix).toBe("x");
      expect(openTags[2].uri).toBe("http://example.com");
    });
  });

  describe("parseXml with xmlns", () => {
    it("should propagate namespace info to DOM elements", () => {
      const doc = parseXml(
        '<x:root xmlns:x="http://example.com"><x:child>text</x:child></x:root>',
        { xmlns: true }
      );

      expect(doc.root.prefix).toBe("x");
      expect(doc.root.local).toBe("root");
      expect(doc.root.uri).toBe("http://example.com");
      expect(doc.root.ns).toEqual({ x: "http://example.com" });

      const child = findChild(doc.root, "x:child");
      expect(child).toBeDefined();
      expect(child!.prefix).toBe("x");
      expect(child!.local).toBe("child");
      expect(child!.uri).toBe("http://example.com");
    });

    it("should not set namespace fields when xmlns is false", () => {
      const doc = parseXml('<x:root xmlns:x="http://example.com"/>');
      expect(doc.root.prefix).toBeUndefined();
      expect(doc.root.uri).toBeUndefined();
    });

    it("should handle OOXML worksheet with namespaces", () => {
      const xml = [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
        '<row r="1"><c r="A1"><v>42</v></c></row>',
        "</sheetData>",
        "</worksheet>"
      ].join("");

      const doc = parseXml(xml, { xmlns: true });
      const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
      expect(doc.root.uri).toBe(main);

      const sheetData = findChild(doc.root, "sheetData")!;
      expect(sheetData.uri).toBe(main);

      const row = findChild(sheetData, "row")!;
      expect(attr(row, "r")).toBe("1");
      expect(row.uri).toBe(main);
    });
  });
});
