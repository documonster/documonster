import { describe, it, expect } from "vitest";
import { SheetFormatPropertiesXform } from "@excel/xlsx/xform/sheet/sheet-format-properties-xform";
import { XmlWriter } from "@xml/writer";

describe("dyDescent behavior", () => {
  describe("1. New worksheet creation - should NOT have dyDescent unless explicitly provided", () => {
    it("should not output dyDescent when not provided in model", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      // Model without dyDescent (simulates new worksheet creation)
      const model = {
        defaultRowHeight: 15,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      // Should NOT contain x14ac:dyDescent
      expect(xmlStream.xml).not.toContain("x14ac:dyDescent");
      expect(xmlStream.xml).toContain("defaultRowHeight");
    });

    it("should not output dyDescent when dyDescent is 0", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      const model = {
        defaultRowHeight: 15,
        dyDescent: 0,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      // dyDescent: 0 should be treated as "not set"
      expect(xmlStream.xml).not.toContain("x14ac:dyDescent");
    });

    it("should not output dyDescent when dyDescent is undefined", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      const model = {
        defaultRowHeight: 15,
        dyDescent: undefined,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      expect(xmlStream.xml).not.toContain("x14ac:dyDescent");
    });
  });

  describe("2. Read then write - should preserve dyDescent from original file", () => {
    it("should output dyDescent when parsed from XML (read-write roundtrip)", () => {
      const xform = new SheetFormatPropertiesXform();

      // Step 1: Parse XML that contains dyDescent (simulates reading from file)
      xform.parseOpen({
        name: "sheetFormatPr",
        attributes: {
          defaultRowHeight: "14.4",
          customHeight: "1",
          "x14ac:dyDescent": "0.55"
        }
      });

      const parsedModel = xform.model;

      // Verify parsing preserved dyDescent
      expect(parsedModel.dyDescent).toBe(0.55);

      // Step 2: Render back to XML (simulates writing to file)
      const xmlStream = new XmlWriter();
      xform.render(xmlStream, parsedModel);

      // Should contain the same dyDescent value
      expect(xmlStream.xml).toContain('x14ac:dyDescent="0.55"');
    });

    it("should preserve dyDescent value 0.25 through roundtrip", () => {
      const xform = new SheetFormatPropertiesXform();

      xform.parseOpen({
        name: "sheetFormatPr",
        attributes: {
          defaultRowHeight: "15",
          "x14ac:dyDescent": "0.25"
        }
      });

      const parsedModel = xform.model;
      expect(parsedModel.dyDescent).toBe(0.25);

      const xmlStream = new XmlWriter();
      xform.render(xmlStream, parsedModel);

      expect(xmlStream.xml).toContain('x14ac:dyDescent="0.25"');
    });
  });

  describe("3. Manual input - should output dyDescent when explicitly provided by user", () => {
    it("should output dyDescent when explicitly set to 0.25", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      // User explicitly provides dyDescent
      const model = {
        defaultRowHeight: 15,
        dyDescent: 0.25,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      expect(xmlStream.xml).toContain('x14ac:dyDescent="0.25"');
    });

    it("should output dyDescent when explicitly set to 0.55", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      const model = {
        defaultRowHeight: 14.4,
        dyDescent: 0.55,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      expect(xmlStream.xml).toContain('x14ac:dyDescent="0.55"');
    });

    it("should output dyDescent with custom value like 0.3", () => {
      const xform = new SheetFormatPropertiesXform();
      const xmlStream = new XmlWriter();

      const model = {
        defaultRowHeight: 15,
        dyDescent: 0.3,
        outlineLevelRow: 0,
        outlineLevelCol: 0
      };

      xform.render(xmlStream, model);

      expect(xmlStream.xml).toContain('x14ac:dyDescent="0.3"');
    });
  });
});
