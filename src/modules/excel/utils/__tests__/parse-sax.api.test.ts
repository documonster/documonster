/**
 * API tests for the SAX XML parser (parse-sax.ts)
 */

import { describe, it, expect } from "vitest";
import type { SaxTag } from "@xml/types";
import { SaxParser } from "@xml/sax";

describe("SaxParser", () => {
  describe("API", () => {
    it("should support write().close() chaining", () => {
      const parser = new SaxParser();
      const events: string[] = [];
      parser.on("opentag", tag => events.push(tag.name));
      parser.write("<root/>").close();
      expect(events).toEqual(["root"]);
    });

    it("should support off() to remove handlers", () => {
      const parser = new SaxParser();
      const events: string[] = [];
      const handler = (tag: SaxTag) => events.push(tag.name);
      parser.on("opentag", handler);
      parser.write("<root><a/>");
      parser.off("opentag");
      parser.write("<b/></root>").close();
      // Only "root" and "a" should be captured; "b" fires after off()
      expect(events).toEqual(["root", "a"]);
    });

    it("should track position", () => {
      const parser = new SaxParser({ position: true });
      parser.write("<root>\n  <child/>\n</root>");
      expect(parser.line).toBe(3);
    });
  });
});
