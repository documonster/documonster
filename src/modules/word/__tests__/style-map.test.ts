/**
 * DOCX Module - Style Map Tests
 */

import { describe, it, expect } from "vitest";

import { Styles } from "../index";

describe("Style mapping DSL", () => {
  describe("parseStyleMap", () => {
    it("parses basic rules", () => {
      const map = Styles.parse(
        `
        p[style-name='Custom'] => div.custom
      `,
        { includeDefaults: false }
      );

      expect(map.rules.length).toBe(1);
      expect(map.rules[0]!.source).toBe("p");
      expect(map.rules[0]!.conditions[0]!.attribute).toBe("style-name");
      expect(map.rules[0]!.conditions[0]!.value).toBe("Custom");
      expect(map.rules[0]!.target.tagName).toBe("div");
      expect(map.rules[0]!.target.className).toBe("custom");
    });

    it("parses multiple rules", () => {
      const map = Styles.parse(
        `
        p[style-name='H1'] => h1
        r[style-name='Bold'] => strong
      `,
        { includeDefaults: false }
      );

      expect(map.rules.length).toBe(2);
    });

    it("skips comments and blank lines", () => {
      const map = Styles.parse(
        `
        # This is a comment
        // Another comment

        p[style-name='X'] => span
      `,
        { includeDefaults: false }
      );

      expect(map.rules.length).toBe(1);
    });
  });

  describe("matchStyleMap", () => {
    it("matches source element with correct attributes", () => {
      const map = Styles.parse(`p[style-name='Quote'] => blockquote`, {
        includeDefaults: false
      });

      const target = Styles.match(map, "p", { "style-name": "Quote" });
      expect(target).toBeDefined();
      expect(target!.tagName).toBe("blockquote");
    });

    it("returns undefined when no match", () => {
      const map = Styles.parse(`p[style-name='Quote'] => blockquote`, {
        includeDefaults: false
      });

      const target = Styles.match(map, "p", { "style-name": "Normal" });
      expect(target).toBeUndefined();
    });

    it("does not match wrong source type", () => {
      const map = Styles.parse(`p[style-name='Code'] => pre`, {
        includeDefaults: false
      });

      const target = Styles.match(map, "r", { "style-name": "Code" });
      expect(target).toBeUndefined();
    });
  });

  describe("nested targets (pre > code)", () => {
    it("parses nested target into child property", () => {
      const map = Styles.parse(`p[style-name='Code'] => pre > code`, {
        includeDefaults: false
      });

      expect(map.rules[0]!.target.tagName).toBe("pre");
      expect(map.rules[0]!.target.child).toBeDefined();
      expect(map.rules[0]!.target.child!.tagName).toBe("code");
    });
  });

  describe("mergeStyleMaps", () => {
    it("merges two style maps with priority ordering", () => {
      const map1 = Styles.create([
        {
          source: "p",
          conditions: [{ attribute: "style-name", value: "A" }],
          target: { tagName: "div" },
          priority: 1
        }
      ]);
      const map2 = Styles.create([
        {
          source: "p",
          conditions: [{ attribute: "style-name", value: "B" }],
          target: { tagName: "section" },
          priority: 1
        }
      ]);

      const merged = Styles.merge(map1, map2);
      expect(merged.rules.length).toBe(2);
      // map2 rules should have higher effective priority (later map = higher)
      expect(merged.rules[0]!.target.tagName).toBe("section");
    });
  });

  describe("DEFAULT_STYLE_MAP", () => {
    it("contains heading mappings", () => {
      const h1 = Styles.match(Styles.DEFAULT, "p", { "style-name": "Heading 1" });
      expect(h1).toBeDefined();
      expect(h1!.tagName).toBe("h1");

      const h3 = Styles.match(Styles.DEFAULT, "p", { "style-name": "Heading 3" });
      expect(h3!.tagName).toBe("h3");
    });

    it("contains run style mappings", () => {
      const strong = Styles.match(Styles.DEFAULT, "r", { "style-name": "Strong" });
      expect(strong).toBeDefined();
      expect(strong!.tagName).toBe("strong");
    });
  });

  describe("createStyleMap", () => {
    it("creates a sorted style map from rules", () => {
      const map = Styles.create([
        { source: "p", conditions: [], target: { tagName: "p" }, priority: 1 },
        {
          source: "p",
          conditions: [{ attribute: "style-name", value: "H1" }],
          target: { tagName: "h1" },
          priority: 10
        }
      ]);

      expect(map.rules[0]!.priority).toBe(10);
      expect(map.rules[1]!.priority).toBe(1);
    });
  });
});
