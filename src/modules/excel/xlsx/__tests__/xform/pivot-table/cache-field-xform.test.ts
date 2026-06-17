import { isPivotError, type PivotError } from "@excel/pivot-table";
import { renderCacheField } from "@excel/xlsx/xform/pivot-table/cache-field";
import { CacheFieldXform } from "@excel/xlsx/xform/pivot-table/cache-field-xform";
import { describe, it, expect } from "vitest";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Parse a cacheField by driving SAX events into the xform and return the model. */
function parseCacheField(
  attrs: Record<string, string>,
  children: (xform: CacheFieldXform) => void = () => {}
) {
  const xform = new CacheFieldXform();
  xform.parseOpen({ name: "cacheField", attributes: attrs });
  children(xform);
  xform.parseClose("cacheField");
  return xform.model!;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("CacheFieldXform", () => {
  // =========================================================================
  // Basic parsing
  // =========================================================================

  describe("basic cacheField parsing", () => {
    it("should parse name attribute", () => {
      const model = parseCacheField({ name: "Category", numFmtId: "0" });
      expect(model.name).toBe("Category");
    });

    it("should parse numFmtId attribute", () => {
      const model = parseCacheField({ name: "Date", numFmtId: "14" });
      expect(model.numFmtId).toBe("14");
    });

    it("should default name to empty string when absent", () => {
      const model = parseCacheField({ numFmtId: "0" });
      expect(model.name).toBe("");
    });

    it("should set isLoaded to true on all parsed fields", () => {
      const model = parseCacheField({ name: "Field", numFmtId: "0" });
      expect(model.isLoaded).toBe(true);
    });

    it("should have null sharedItems when no <sharedItems> element is present", () => {
      const model = parseCacheField({ name: "Bare", numFmtId: "0" });
      expect(model.sharedItems).toBeNull();
    });
  });

  // =========================================================================
  // sharedItems parsing
  // =========================================================================

  describe("sharedItems parsing", () => {
    it("should parse string shared items (<s>)", () => {
      const model = parseCacheField({ name: "Cat", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "2" } });
        xform.parseOpen({ name: "s", attributes: { v: "Apple" } });
        xform.parseClose("s");
        xform.parseOpen({ name: "s", attributes: { v: "Banana" } });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual(["Apple", "Banana"]);
    });

    it("should parse numeric shared items (<n>)", () => {
      const model = parseCacheField({ name: "Val", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsNumber: "1", containsString: "0", count: "2" }
        });
        xform.parseOpen({ name: "n", attributes: { v: "42.5" } });
        xform.parseClose("n");
        xform.parseOpen({ name: "n", attributes: { v: "-10" } });
        xform.parseClose("n");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([42.5, -10]);
    });

    it("should parse boolean shared items (<b>)", () => {
      const model = parseCacheField({ name: "Flag", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "2" } });
        xform.parseOpen({ name: "b", attributes: { v: "1" } });
        xform.parseClose("b");
        xform.parseOpen({ name: "b", attributes: { v: "0" } });
        xform.parseClose("b");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([true, false]);
    });

    it("should parse date shared items (<d>)", () => {
      const model = parseCacheField({ name: "Date", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "d", attributes: { v: "2024-01-15T00:00:00" } });
        xform.parseClose("d");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toHaveLength(1);
      const dateVal = model.sharedItems![0] as Date;
      expect(dateVal).toBeInstanceOf(Date);
      expect(dateVal.getUTCFullYear()).toBe(2024);
      expect(dateVal.getUTCMonth()).toBe(0);
      expect(dateVal.getUTCDate()).toBe(15);
    });

    it("should parse error shared items (<e>)", () => {
      const model = parseCacheField({ name: "Err", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "e", attributes: { v: "REF!" } });
        xform.parseClose("e");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toHaveLength(1);
      expect(isPivotError(model.sharedItems![0])).toBe(true);
      expect((model.sharedItems![0] as PivotError).code).toBe("REF!");
    });

    it("should parse missing value items (<m>)", () => {
      const model = parseCacheField({ name: "Nullable", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "m", attributes: {} });
        xform.parseClose("m");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([null]);
    });

    it("should handle absent v attribute on <n> as null", () => {
      const model = parseCacheField({ name: "Val", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "n", attributes: {} });
        xform.parseClose("n");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([null]);
    });

    it("should handle empty v attribute on <n> as null", () => {
      const model = parseCacheField({ name: "Val", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "n", attributes: { v: "" } });
        xform.parseClose("n");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([null]);
    });

    it("should handle absent v attribute on <d> as null", () => {
      const model = parseCacheField({ name: "Date", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "d", attributes: {} });
        xform.parseClose("d");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([null]);
    });

    it("should default <s> v to empty string when absent", () => {
      const model = parseCacheField({ name: "Str", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "s", attributes: {} });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([""]);
    });

    it("should default <e> v to empty string when absent", () => {
      const model = parseCacheField({ name: "Err", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "e", attributes: {} });
        xform.parseClose("e");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toHaveLength(1);
      expect(isPivotError(model.sharedItems![0])).toBe(true);
      expect((model.sharedItems![0] as PivotError).code).toBe("");
    });

    it("should parse empty sharedItems (count=0) as empty array", () => {
      const model = parseCacheField({ name: "Empty", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "0" } });
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([]);
    });

    it("should parse self-closing sharedItems as empty array", () => {
      const model = parseCacheField({ name: "SelfClose", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: {} });
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual([]);
    });

    it("should parse mixed type shared items in order", () => {
      const model = parseCacheField({ name: "Mixed", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "4" } });
        xform.parseOpen({ name: "s", attributes: { v: "text" } });
        xform.parseClose("s");
        xform.parseOpen({ name: "n", attributes: { v: "99" } });
        xform.parseClose("n");
        xform.parseOpen({ name: "m", attributes: {} });
        xform.parseClose("m");
        xform.parseOpen({ name: "b", attributes: { v: "1" } });
        xform.parseClose("b");
        xform.parseClose("sharedItems");
      });

      expect(model.sharedItems).toEqual(["text", 99, null, true]);
    });
  });

  // =========================================================================
  // sharedItems attributes preservation
  // =========================================================================

  describe("sharedItems attributes preservation", () => {
    it("should preserve containsNumber and containsInteger", () => {
      const model = parseCacheField({ name: "Num", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsNumber: "1", containsInteger: "1" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsNumber).toBe("1");
      expect(model.containsInteger).toBe("1");
    });

    it("should preserve minValue and maxValue", () => {
      const model = parseCacheField({ name: "Range", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: {
            containsNumber: "1",
            minValue: "5.5",
            maxValue: "100.25"
          }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.minValue).toBe(5.5);
      expect(model.maxValue).toBe(100.25);
    });

    it("should preserve containsSemiMixedTypes", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsSemiMixedTypes: "0" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsSemiMixedTypes).toBe("0");
    });

    it("should preserve containsNonDate", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsNonDate: "0" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsNonDate).toBe("0");
    });

    it("should preserve containsString", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsString: "0" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsString).toBe("0");
    });

    it("should preserve containsBlank", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsBlank: "1" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsBlank).toBe("1");
    });

    it("should preserve containsDate", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsDate: "1" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsDate).toBe("1");
    });

    it("should preserve containsMixedTypes", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: { containsMixedTypes: "1" }
        });
        xform.parseClose("sharedItems");
      });

      expect(model.containsMixedTypes).toBe("1");
    });

    it("should not set containsNumber/containsInteger when attributes are absent", () => {
      const model = parseCacheField({ name: "StrOnly", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "s", attributes: { v: "text" } });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
      });

      expect(model.containsNumber).toBeUndefined();
      expect(model.containsInteger).toBeUndefined();
      expect(model.minValue).toBeUndefined();
      expect(model.maxValue).toBeUndefined();
    });

    it("should preserve all typical numeric field attributes together", () => {
      const model = parseCacheField({ name: "Amount", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: {
            containsSemiMixedTypes: "0",
            containsString: "0",
            containsNumber: "1",
            containsInteger: "1",
            minValue: "5",
            maxValue: "45"
          }
        });
        xform.parseOpen({ name: "n", attributes: { v: "5" } });
        xform.parseClose("n");
        xform.parseOpen({ name: "n", attributes: { v: "45" } });
        xform.parseClose("n");
        xform.parseClose("sharedItems");
      });

      expect(model.containsSemiMixedTypes).toBe("0");
      expect(model.containsString).toBe("0");
      expect(model.containsNumber).toBe("1");
      expect(model.containsInteger).toBe("1");
      expect(model.minValue).toBe(5);
      expect(model.maxValue).toBe(45);
      expect(model.sharedItems).toEqual([5, 45]);
    });
  });

  // =========================================================================
  // Absent sharedItems (B8 fix)
  // =========================================================================

  describe("absent sharedItems (B8 fix)", () => {
    it("should set sharedItems to null and isLoaded to true when no <sharedItems> element", () => {
      const model = parseCacheField({ name: "NoShared", numFmtId: "0" });
      expect(model.sharedItems).toBeNull();
      expect(model.isLoaded).toBe(true);
    });

    it("should distinguish absent sharedItems from empty sharedItems", () => {
      // Absent: no <sharedItems> element at all
      const absent = parseCacheField({ name: "A", numFmtId: "0" });

      // Empty: <sharedItems /> present but no children
      const empty = parseCacheField({ name: "B", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: {} });
        xform.parseClose("sharedItems");
      });

      expect(absent.sharedItems).toBeNull();
      expect(empty.sharedItems).toEqual([]);
    });
  });

  // =========================================================================
  // fieldGroup XML collection
  // =========================================================================

  describe("fieldGroup XML collection", () => {
    it("should collect fieldGroup XML for roundtrip preservation", () => {
      const model = parseCacheField({ name: "Grouped", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "s", attributes: { v: "A" } });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
        xform.parseOpen({ name: "fieldGroup", attributes: { base: "0" } });
        xform.parseOpen({ name: "rangePr", attributes: { groupBy: "months" } });
        xform.parseClose("rangePr");
        xform.parseOpen({ name: "groupItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "s", attributes: { v: "Jan" } });
        xform.parseClose("s");
        xform.parseClose("groupItems");
        xform.parseClose("fieldGroup");
      });

      expect(model.fieldGroupXml).toBeDefined();
      expect(model.fieldGroupXml).toContain("<fieldGroup");
      expect(model.fieldGroupXml).toContain('base="0"');
      expect(model.fieldGroupXml).toContain('groupBy="months"');
      expect(model.fieldGroupXml).toContain("</fieldGroup>");
    });

    it("should not set fieldGroupXml when no fieldGroup element is present", () => {
      const model = parseCacheField({ name: "Plain", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
        xform.parseOpen({ name: "s", attributes: { v: "A" } });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
      });

      expect(model.fieldGroupXml).toBeUndefined();
    });
  });

  // =========================================================================
  // extraAttrs bag
  // =========================================================================

  describe("extraAttrs bag", () => {
    it("should collect unknown cacheField attributes into extraAttrs", () => {
      const model = parseCacheField({
        name: "CalcField",
        numFmtId: "0",
        caption: "My Field",
        formula: "'A' * 'B'",
        databaseField: "0"
      });

      expect(model.extraAttrs).toBeDefined();
      expect(model.extraAttrs!["caption"]).toBe("My Field");
      expect(model.extraAttrs!["formula"]).toBe("'A' * 'B'");
      expect(model.extraAttrs!["databaseField"]).toBe("0");
    });

    it("should not include known attributes (name, numFmtId) in extraAttrs", () => {
      const model = parseCacheField({
        name: "Field",
        numFmtId: "14"
      });

      // No extra attrs → undefined
      expect(model.extraAttrs).toBeUndefined();
    });

    it("should set extraAttrs to undefined when all attributes are known", () => {
      const model = parseCacheField({ name: "F", numFmtId: "0" });
      expect(model.extraAttrs).toBeUndefined();
    });
  });

  // =========================================================================
  // Reset
  // =========================================================================

  describe("reset", () => {
    it("should clear model on reset", () => {
      const xform = new CacheFieldXform();
      xform.parseOpen({ name: "cacheField", attributes: { name: "F", numFmtId: "0" } });
      xform.parseClose("cacheField");
      expect(xform.model).not.toBeNull();

      xform.reset();
      expect(xform.model).toBeNull();
    });

    it("should not leak state between parses", () => {
      const xform = new CacheFieldXform();

      // First parse with sharedItems
      xform.parseOpen({ name: "cacheField", attributes: { name: "A", numFmtId: "0" } });
      xform.parseOpen({ name: "sharedItems", attributes: { count: "1" } });
      xform.parseOpen({ name: "s", attributes: { v: "X" } });
      xform.parseClose("s");
      xform.parseClose("sharedItems");
      xform.parseClose("cacheField");
      expect(xform.model!.sharedItems).toEqual(["X"]);

      // Second parse without sharedItems — should not leak
      xform.parseOpen({ name: "cacheField", attributes: { name: "B", numFmtId: "0" } });
      xform.parseClose("cacheField");
      expect(xform.model!.name).toBe("B");
      expect(xform.model!.sharedItems).toBeNull();
    });
  });

  // =========================================================================
  // Parse → Render roundtrip
  // =========================================================================

  describe("parse → render roundtrip", () => {
    it("should roundtrip string shared items", () => {
      const model = parseCacheField({ name: "Category", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: { count: "2" } });
        xform.parseOpen({ name: "s", attributes: { v: "Apple" } });
        xform.parseClose("s");
        xform.parseOpen({ name: "s", attributes: { v: "Banana" } });
        xform.parseClose("s");
        xform.parseClose("sharedItems");
      });

      const xml = renderCacheField(model);
      expect(xml).toContain('name="Category"');
      expect(xml).toContain('count="2"');
      expect(xml).toContain('<s v="Apple" />');
      expect(xml).toContain('<s v="Banana" />');
    });

    it("should roundtrip numeric shared items with attributes", () => {
      const model = parseCacheField({ name: "Amount", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: {
            containsSemiMixedTypes: "0",
            containsString: "0",
            containsNumber: "1",
            containsInteger: "1",
            minValue: "5",
            maxValue: "45"
          }
        });
        xform.parseOpen({ name: "n", attributes: { v: "5" } });
        xform.parseClose("n");
        xform.parseOpen({ name: "n", attributes: { v: "45" } });
        xform.parseClose("n");
        xform.parseClose("sharedItems");
      });

      const xml = renderCacheField(model);
      expect(xml).toContain('containsSemiMixedTypes="0"');
      expect(xml).toContain('containsString="0"');
      expect(xml).toContain('containsNumber="1"');
      expect(xml).toContain('containsInteger="1"');
      expect(xml).toContain('minValue="5"');
      expect(xml).toContain('maxValue="45"');
      expect(xml).toContain('<n v="5" />');
      expect(xml).toContain('<n v="45" />');
    });

    it("should roundtrip absent sharedItems without injecting <sharedItems>", () => {
      const model = parseCacheField({ name: "NoShared", numFmtId: "0" });
      const xml = renderCacheField(model);

      expect(xml).toContain('name="NoShared"');
      // Should NOT contain <sharedItems> since it was absent in the original
      expect(xml).not.toContain("<sharedItems");
    });

    it("should roundtrip empty sharedItems as self-closing", () => {
      const model = parseCacheField({ name: "Empty", numFmtId: "0" }, xform => {
        xform.parseOpen({ name: "sharedItems", attributes: {} });
        xform.parseClose("sharedItems");
      });

      const xml = renderCacheField(model);
      expect(xml).toMatch(/<sharedItems\s*\/>/);
    });

    it("should roundtrip extraAttrs", () => {
      const model = parseCacheField({
        name: "Calc",
        numFmtId: "0",
        caption: "My Caption",
        databaseField: "0"
      });

      const xml = renderCacheField(model);
      expect(xml).toContain('caption="My Caption"');
      expect(xml).toContain('databaseField="0"');
    });

    it("should roundtrip numFmtId", () => {
      const model = parseCacheField({ name: "Date", numFmtId: "14" });
      const xml = renderCacheField(model);
      expect(xml).toContain('numFmtId="14"');
    });

    it("should roundtrip empty sharedItems with containsBlank", () => {
      const model = parseCacheField({ name: "Blank", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: {
            containsSemiMixedTypes: "0",
            containsString: "0",
            containsBlank: "1"
          }
        });
        xform.parseClose("sharedItems");
      });

      const xml = renderCacheField(model);
      expect(xml).toContain('containsSemiMixedTypes="0"');
      expect(xml).toContain('containsString="0"');
      expect(xml).toContain('containsBlank="1"');
    });

    it("should roundtrip mixed types (string + number + null)", () => {
      const model = parseCacheField({ name: "Mix", numFmtId: "0" }, xform => {
        xform.parseOpen({
          name: "sharedItems",
          attributes: {
            containsNumber: "1",
            containsBlank: "1",
            minValue: "42",
            maxValue: "42"
          }
        });
        xform.parseOpen({ name: "s", attributes: { v: "text" } });
        xform.parseClose("s");
        xform.parseOpen({ name: "n", attributes: { v: "42" } });
        xform.parseClose("n");
        xform.parseOpen({ name: "m", attributes: {} });
        xform.parseClose("m");
        xform.parseClose("sharedItems");
      });

      const xml = renderCacheField(model);
      expect(xml).toContain('<s v="text" />');
      expect(xml).toContain('<n v="42" />');
      expect(xml).toContain("<m />");
      expect(xml).toContain('count="3"');
    });
  });
});
