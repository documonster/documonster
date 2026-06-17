import { pivotError } from "@excel/pivot-table";
import { renderCacheField } from "@excel/xlsx/xform/pivot-table/cache-field";
import { describe, it, expect } from "vitest";

describe("CacheField", () => {
  describe("render", () => {
    describe("string sharedItems", () => {
      it("should render string sharedItems with <s> elements", () => {
        const xml = renderCacheField({
          name: "Category",
          sharedItems: ["Apple", "Banana", "Cherry"]
        });

        expect(xml).toContain('name="Category"');
        expect(xml).toContain('count="3"');
        expect(xml).toContain('<s v="Apple" />');
        expect(xml).toContain('<s v="Banana" />');
        expect(xml).toContain('<s v="Cherry" />');
        // Should NOT contain numeric attributes
        expect(xml).not.toContain("containsNumber");
      });

      it("should escape XML special characters in string values", () => {
        const xml = renderCacheField({
          name: "Data",
          sharedItems: ["A & B", "C < D", 'E "F"']
        });

        expect(xml).toContain('<s v="A &amp; B" />');
        expect(xml).toContain('<s v="C &lt; D" />');
        expect(xml).toContain('<s v="E &quot;F&quot;" />');
      });

      it("should escape XML special characters in field name", () => {
        const xml = renderCacheField({
          name: "A & B",
          sharedItems: ["value"]
        });

        expect(xml).toContain('name="A &amp; B"');
      });
    });

    describe("numeric sharedItems", () => {
      it("should render integer sharedItems with <n> elements", () => {
        const xml = renderCacheField({
          name: "Amount",
          sharedItems: [5, 24, 35, 45]
        });

        expect(xml).toContain('name="Amount"');
        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain('containsInteger="1"');
        expect(xml).toContain('minValue="5"');
        expect(xml).toContain('maxValue="45"');
        expect(xml).toContain('count="4"');
        expect(xml).toContain('<n v="5" />');
        expect(xml).toContain('<n v="24" />');
        expect(xml).toContain('<n v="35" />');
        expect(xml).toContain('<n v="45" />');
        // Should NOT use string format
        expect(xml).not.toContain('<s v="5"');
        expect(xml).not.toContain('<s v="24"');
      });

      it("should render float sharedItems without containsInteger", () => {
        const xml = renderCacheField({
          name: "Price",
          sharedItems: [5.5, 10.25, 24.75]
        });

        expect(xml).toContain('containsNumber="1"');
        expect(xml).not.toContain('containsInteger="1"');
        expect(xml).toContain('minValue="5.5"');
        expect(xml).toContain('maxValue="24.75"');
        expect(xml).toContain('<n v="5.5" />');
        expect(xml).toContain('<n v="10.25" />');
        expect(xml).toContain('<n v="24.75" />');
      });

      it("should handle negative numbers correctly", () => {
        const xml = renderCacheField({
          name: "Value",
          sharedItems: [-10, 0, 20]
        });

        expect(xml).toContain('minValue="-10"');
        expect(xml).toContain('maxValue="20"');
        expect(xml).toContain('<n v="-10" />');
        expect(xml).toContain('<n v="0" />');
        expect(xml).toContain('<n v="20" />');
      });

      it("should handle single numeric value", () => {
        const xml = renderCacheField({
          name: "Single",
          sharedItems: [42]
        });

        expect(xml).toContain('minValue="42"');
        expect(xml).toContain('maxValue="42"');
        expect(xml).toContain('count="1"');
        expect(xml).toContain('<n v="42" />');
      });
    });

    describe("mixed types in sharedItems", () => {
      it("should render mixed string/number with native elements", () => {
        const xml = renderCacheField({
          name: "Mixed",
          sharedItems: ["text", 123, "another"]
        });

        // Each type rendered natively
        expect(xml).toContain('<s v="text" />');
        expect(xml).toContain('<n v="123" />');
        expect(xml).toContain('<s v="another" />');
        expect(xml).toContain('containsMixedTypes="1"');
        expect(xml).toContain('containsNumber="1"');
      });
    });

    describe("edge cases", () => {
      it("should handle empty sharedItems array", () => {
        const xml = renderCacheField({
          name: "Empty",
          sharedItems: []
        });

        // Empty array should render as string type with count=0
        expect(xml).toContain('name="Empty"');
        expect(xml).toContain('count="0"');
        // Should NOT have numeric attributes (especially not Infinity/-Infinity)
        expect(xml).not.toContain("containsNumber");
        expect(xml).not.toContain("minValue");
        expect(xml).not.toContain("maxValue");
        expect(xml).not.toContain("Infinity");
      });

      it("should handle zero-only values", () => {
        const xml = renderCacheField({
          name: "Zeros",
          sharedItems: [0, 0, 0]
        });

        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain('containsInteger="1"');
        expect(xml).toContain('minValue="0"');
        expect(xml).toContain('maxValue="0"');
        expect(xml).toContain('<n v="0" />');
      });

      it("should handle very large numbers", () => {
        const xml = renderCacheField({
          name: "Large",
          sharedItems: [1e10, 1e15, Number.MAX_SAFE_INTEGER]
        });

        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain(`<n v="${1e10}" />`);
        expect(xml).toContain(`<n v="${1e15}" />`);
        expect(xml).toContain(`<n v="${Number.MAX_SAFE_INTEGER}" />`);
      });

      it("should handle very small decimal numbers", () => {
        const xml = renderCacheField({
          name: "Small",
          sharedItems: [0.001, 0.0001, 1e-10]
        });

        expect(xml).toContain('containsNumber="1"');
        expect(xml).not.toContain('containsInteger="1"');
        expect(xml).toContain(`minValue="${1e-10}"`);
      });

      it("should treat NaN and Infinity as missing values", () => {
        const xml = renderCacheField({
          name: "Special",
          sharedItems: [NaN, Infinity, -Infinity]
        });

        // NaN and Infinity are not valid in OOXML, treated as missing (<m />)
        expect(xml).not.toContain("containsNumber");
        expect(xml).not.toContain("<s v=");
        expect(xml).toContain("containsBlank");
        expect(xml).toContain("<m />");
      });

      it("should handle Unicode characters in field name", () => {
        const xml = renderCacheField({
          name: "日本語フィールド",
          sharedItems: ["値1", "値2"]
        });

        expect(xml).toContain('name="日本語フィールド"');
        expect(xml).toContain('<s v="値1" />');
        expect(xml).toContain('<s v="値2" />');
      });

      it("should handle emoji in values", () => {
        const xml = renderCacheField({
          name: "Emoji",
          sharedItems: ["😀", "🎉", "👍"]
        });

        expect(xml).toContain('<s v="😀" />');
        expect(xml).toContain('<s v="🎉" />');
        expect(xml).toContain('<s v="👍" />');
      });

      it("should render boolean values as <b> elements", () => {
        const xml = renderCacheField({
          name: "Boolean",
          sharedItems: [true, false] as any[]
        });

        // Booleans rendered as <b> per OOXML spec
        expect(xml).not.toContain("containsNumber");
        expect(xml).toContain('<b v="1" />');
        expect(xml).toContain('<b v="0" />');
        expect(xml).toContain('containsSemiMixedTypes="0"');
        expect(xml).toContain('containsString="0"');
      });

      it("should handle single string value", () => {
        const xml = renderCacheField({
          name: "Single",
          sharedItems: ["OnlyOne"]
        });

        expect(xml).toContain('count="1"');
        expect(xml).toContain('<s v="OnlyOne" />');
      });

      it("should handle whitespace-only values", () => {
        const xml = renderCacheField({
          name: "Whitespace",
          sharedItems: [" ", "  ", "\t", "\n"]
        });

        expect(xml).toContain('<s v=" " />');
        expect(xml).toContain('<s v="  " />');
      });

      it("should handle empty string value", () => {
        const xml = renderCacheField({
          name: "EmptyString",
          sharedItems: ["", "nonempty"]
        });

        expect(xml).toContain('<s v="" />');
        expect(xml).toContain('<s v="nonempty" />');
      });
    });

    describe("null sharedItems (unused or value-only fields)", () => {
      it("should render empty sharedItems for unused field", () => {
        const xml = renderCacheField({
          name: "Unused",
          sharedItems: null
        });

        expect(xml).toContain('name="Unused"');
        expect(xml).toContain("<sharedItems />");
        // Should NOT contain numeric attributes for unused fields
        expect(xml).not.toContain("containsNumber");
        expect(xml).not.toContain("minValue");
        expect(xml).not.toContain("maxValue");
      });

      it("should render numeric attributes for value-only field with minMax (integers)", () => {
        const xml = renderCacheField({
          name: "ValueOnly",
          sharedItems: null,
          minValue: 10,
          maxValue: 100,
          containsInteger: "1"
        });

        expect(xml).toContain('name="ValueOnly"');
        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain('containsInteger="1"');
        expect(xml).toContain('minValue="10"');
        expect(xml).toContain('maxValue="100"');
        // Should be self-closing (no child elements)
        expect(xml).toMatch(/<sharedItems[^>]+\/>/);
      });

      it("should NOT render containsInteger for value-only field with decimal minMax", () => {
        const xml = renderCacheField({
          name: "DecimalField",
          sharedItems: null,
          minValue: 10.5,
          maxValue: 99.9
        });

        expect(xml).toContain('name="DecimalField"');
        expect(xml).toContain('containsNumber="1"');
        expect(xml).not.toContain('containsInteger="1"');
        expect(xml).toContain('minValue="10.5"');
        expect(xml).toContain('maxValue="99.9"');
        expect(xml).toMatch(/<sharedItems[^>]+\/>/);
      });
    });

    describe("Date values in sharedItems", () => {
      it("should render Date sharedItems with <d> elements", () => {
        const d1 = new Date("2024-01-15T00:00:00.000Z");
        const d2 = new Date("2024-06-30T12:30:00.000Z");
        const xml = renderCacheField({
          name: "DateField",
          sharedItems: [d1, d2]
        });

        expect(xml).toContain('containsDate="1"');
        expect(xml).toContain('containsSemiMixedTypes="0"');
        expect(xml).toContain('containsString="0"');
        expect(xml).toContain(`<d v="${d1.toISOString().replace(/\.\d{3}Z$/, "")}" />`);
        expect(xml).toContain(`<d v="${d2.toISOString().replace(/\.\d{3}Z$/, "")}" />`);
        expect(xml).toContain('count="2"');
      });

      it("should render mixed string and Date with containsMixedTypes", () => {
        const d1 = new Date("2024-01-01T00:00:00.000Z");
        const xml = renderCacheField({
          name: "MixedDateField",
          sharedItems: ["pending", d1]
        });

        expect(xml).toContain('containsDate="1"');
        // Has both string and date types
        expect(xml).toContain('containsMixedTypes="1"');
        expect(xml).toContain('<s v="pending" />');
        expect(xml).toContain(`<d v="${d1.toISOString().replace(/\.\d{3}Z$/, "")}" />`);
      });
    });

    describe("null values within sharedItems array", () => {
      it("should render null values as <m /> with containsBlank", () => {
        const xml = renderCacheField({
          name: "Nullable",
          sharedItems: ["A", null, "B"]
        });

        expect(xml).toContain('containsBlank="1"');
        expect(xml).toContain('<s v="A" />');
        expect(xml).toContain("<m />");
        expect(xml).toContain('<s v="B" />');
        expect(xml).toContain('count="3"');
      });

      it("should handle null-only sharedItems", () => {
        const xml = renderCacheField({
          name: "AllNull",
          sharedItems: [null, null]
        });

        expect(xml).toContain('containsBlank="1"');
        expect(xml).toContain('containsSemiMixedTypes="0"');
        expect(xml).toContain('containsString="0"');
        expect(xml).toContain("<m />");
        expect(xml).toContain('count="2"');
      });

      it("should handle mixed null and number", () => {
        const xml = renderCacheField({
          name: "NullableNumber",
          sharedItems: [10, null, 30]
        });

        expect(xml).toContain('containsBlank="1"');
        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain('containsSemiMixedTypes="0"');
        expect(xml).toContain('containsString="0"');
        expect(xml).toContain('<n v="10" />');
        expect(xml).toContain("<m />");
        expect(xml).toContain('<n v="30" />');
      });
    });

    describe("numFmtId preservation (bug #6)", () => {
      it("should use provided numFmtId instead of default '0'", () => {
        const xml = renderCacheField({
          name: "DateCol",
          sharedItems: null,
          numFmtId: "14"
        });

        expect(xml).toContain('numFmtId="14"');
        expect(xml).not.toContain('numFmtId="0"');
      });

      it("should default numFmtId to '0' when not specified", () => {
        const xml = renderCacheField({
          name: "Plain",
          sharedItems: null
        });

        expect(xml).toContain('numFmtId="0"');
      });

      it("should preserve numFmtId for fields with sharedItems", () => {
        const xml = renderCacheField({
          name: "Formatted",
          sharedItems: ["A", "B"],
          numFmtId: "164"
        });

        expect(xml).toContain('numFmtId="164"');
      });

      it("should preserve numFmtId for value-only fields with minMax", () => {
        const xml = renderCacheField({
          name: "Currency",
          sharedItems: null,
          minValue: 100,
          maxValue: 9999,
          containsInteger: "1",
          numFmtId: "44"
        });

        expect(xml).toContain('numFmtId="44"');
        expect(xml).toContain('containsNumber="1"');
      });
    });

    describe("PivotErrorValue rendering (bug #8)", () => {
      it("should render PivotErrorValue as <e> elements, not <s>", () => {
        const xml = renderCacheField({
          name: "ErrorField",
          sharedItems: [pivotError("REF!"), pivotError("VALUE!")]
        });

        expect(xml).toContain('<e v="REF!" />');
        expect(xml).toContain('<e v="VALUE!" />');
        // Must NOT render as string
        expect(xml).not.toContain('<s v="REF!"');
        expect(xml).not.toContain('<s v="#REF!"');
        expect(xml).not.toContain('<s v="VALUE!"');
        // Error-only: no strings → containsSemiMixedTypes="0", containsString="0"
        expect(xml).toContain('containsSemiMixedTypes="0"');
        expect(xml).toContain('containsString="0"');
      });

      it("should render mixed error and string with containsMixedTypes", () => {
        const xml = renderCacheField({
          name: "MixedErrors",
          sharedItems: ["valid", pivotError("N/A"), "also valid"]
        });

        expect(xml).toContain('<s v="valid" />');
        expect(xml).toContain('<e v="N/A" />');
        expect(xml).toContain('<s v="also valid" />');
        // Has strings AND errors = containsMixedTypes
        expect(xml).toContain('containsMixedTypes="1"');
      });

      it("should handle error with XML special characters in code", () => {
        // Unlikely in practice but should be safe
        const xml = renderCacheField({
          name: "SpecialError",
          sharedItems: [pivotError("A&B")]
        });

        expect(xml).toContain('<e v="A&amp;B" />');
      });

      it("should handle mixed error, string, number, and null", () => {
        const xml = renderCacheField({
          name: "AllTypes",
          sharedItems: ["text", 42, pivotError("DIV/0!"), null]
        });

        expect(xml).toContain('<s v="text" />');
        expect(xml).toContain('<n v="42" />');
        expect(xml).toContain('<e v="DIV/0!" />');
        expect(xml).toContain("<m />");
        expect(xml).toContain('containsMixedTypes="1"');
        expect(xml).toContain('containsNumber="1"');
        expect(xml).toContain('containsBlank="1"');
        expect(xml).toContain('count="4"');
      });
    });
  });

  // ===========================================================================
  // Round 6 Bug B: cacheField extraAttrs bag
  // ===========================================================================

  describe("extraAttrs roundtrip (R6-BugB)", () => {
    it("should render extra attributes like caption, formula, databaseField", () => {
      const xml = renderCacheField({
        name: "CalcField",
        sharedItems: null,
        numFmtId: "0",
        extraAttrs: {
          caption: "My Calc Field",
          formula: "'Amount' * 'Rate'",
          databaseField: "0"
        }
      });

      expect(xml).toContain('caption="My Calc Field"');
      expect(xml).toContain("formula=");
      expect(xml).toContain('databaseField="0"');
      // Should still have the standard attributes
      expect(xml).toContain('name="CalcField"');
      expect(xml).toContain('numFmtId="0"');
    });

    it("should render extra attributes for field with sharedItems", () => {
      const xml = renderCacheField({
        name: "CaptionField",
        sharedItems: ["A", "B"],
        extraAttrs: { caption: "Nice Caption" }
      });

      expect(xml).toContain('caption="Nice Caption"');
      expect(xml).toContain('name="CaptionField"');
      expect(xml).toContain('<s v="A" />');
    });

    it("should XML-encode extra attribute values", () => {
      const xml = renderCacheField({
        name: "Field",
        sharedItems: null,
        extraAttrs: { formula: "A & B < C" }
      });

      expect(xml).toContain('formula="A &amp; B &lt; C"');
    });

    it("should not emit extra attributes when extraAttrs is undefined", () => {
      const xml = renderCacheField({
        name: "Plain",
        sharedItems: null
      });

      // Only standard attributes
      expect(xml).toMatch(/<cacheField name="Plain" numFmtId="0">/);
    });

    it("should render extra attributes on loaded fields with empty sharedItems", () => {
      const xml = renderCacheField({
        name: "LoadedCalc",
        sharedItems: [],
        isLoaded: true,
        extraAttrs: { caption: "Loaded Caption" }
      });

      expect(xml).toContain('caption="Loaded Caption"');
    });
  });
});
