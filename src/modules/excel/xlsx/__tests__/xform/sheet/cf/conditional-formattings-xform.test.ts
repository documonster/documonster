import { ConditionalFormattingsXform } from "@excel/xlsx/xform/sheet/cf/conditional-formattings-xform";
import { describe, it, expect } from "vitest";

describe("ConditionalFormattingsXform", () => {
  describe("prepare", () => {
    it("should add default cfvo and color for dataBar rules without them", () => {
      const xform = new ConditionalFormattingsXform();
      const model: any[] = [
        {
          ref: "A1:A10",
          rules: [
            {
              type: "dataBar"
              // No cfvo or color specified - this was causing the issue
            }
          ]
        }
      ];

      const options = {
        styles: {
          addDxfStyle: () => 0
        }
      };

      xform.prepare(model, options);

      // Check that default cfvo was added
      expect(model[0].rules[0].cfvo).toEqual([{ type: "min" }, { type: "max" }]);
      // Check that default color was added
      expect(model[0].rules[0].color).toEqual({ argb: "FF638EC6" });
      // Check that priority was assigned
      expect(model[0].rules[0].priority).toBe(1);
    });

    it("should preserve user-specified cfvo and color for dataBar rules", () => {
      const xform = new ConditionalFormattingsXform();
      const model: any[] = [
        {
          ref: "A1:A10",
          rules: [
            {
              type: "dataBar",
              cfvo: [
                { type: "num", value: 5 },
                { type: "num", value: 20 }
              ],
              color: { argb: "FFFF0000" }
            }
          ]
        }
      ];

      const options = {
        styles: {
          addDxfStyle: () => 0
        }
      };

      xform.prepare(model, options);

      // Check that user-specified cfvo is preserved
      expect(model[0].rules[0].cfvo).toEqual([
        { type: "num", value: 5 },
        { type: "num", value: 20 }
      ]);
      // Check that user-specified color is preserved
      expect(model[0].rules[0].color).toEqual({ argb: "FFFF0000" });
    });

    it("should replace incomplete cfvo arrays for dataBar rules", () => {
      const xform = new ConditionalFormattingsXform();
      const model: any[] = [
        {
          ref: "A1:A10",
          rules: [
            {
              type: "dataBar",
              cfvo: [{ type: "min" }] // Only one cfvo, need at least 2
            }
          ]
        }
      ];

      const options = {
        styles: {
          addDxfStyle: () => 0
        }
      };

      xform.prepare(model, options);

      // Check that default cfvo replaces incomplete array
      expect(model[0].rules[0].cfvo).toEqual([{ type: "min" }, { type: "max" }]);
    });

    it("should not modify non-dataBar rules", () => {
      const xform = new ConditionalFormattingsXform();
      const model: any[] = [
        {
          ref: "A1:A10",
          rules: [
            {
              type: "colorScale",
              cfvo: [{ type: "min" }, { type: "max" }],
              color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }]
            }
          ]
        }
      ];

      const options = {
        styles: {
          addDxfStyle: () => 0
        }
      };

      const originalCfvo = [...model[0].rules[0].cfvo];
      const originalColor = [...model[0].rules[0].color];

      xform.prepare(model, options);

      // Check that colorScale rule is not modified
      expect(model[0].rules[0].cfvo).toEqual(originalCfvo);
      expect(model[0].rules[0].color).toEqual(originalColor);
    });
  });
});
