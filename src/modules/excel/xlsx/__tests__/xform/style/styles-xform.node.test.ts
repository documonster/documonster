import fs from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { testXformHelper, normalizeXml } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Helper function to add apply* flags to styles based on their non-zero IDs and alignments
function addApplyFlags(model: any): any {
  const result = JSON.parse(JSON.stringify(model));
  if (result.styles) {
    // Map of style property to apply flag
    const flagMappings: Array<[string, string]> = [
      ["fontId", "applyFont"],
      ["fillId", "applyFill"],
      ["borderId", "applyBorder"],
      ["numFmtId", "applyNumberFormat"],
      ["alignment", "applyAlignment"],
      ["protection", "applyProtection"]
    ];
    result.styles = result.styles.map((style: any) => {
      const newStyle = { ...style };
      for (const [prop, flag] of flagMappings) {
        if (style[prop]) {
          newStyle[flag] = true;
        }
      }
      return newStyle;
    });
  }
  return result;
}

const expectations = [
  {
    title: "Styles with fonts",
    create() {
      return new StylesXform();
    },
    preparedModel: JSON.parse(fs.readFileSync(join(__dirname, "data/styles.1.1.json")).toString()),
    xml: fs.readFileSync(join(__dirname, "data/styles.1.2.xml")).toString(),
    get parsedModel() {
      // parsedModel includes apply* flags from the XML
      const model = addApplyFlags(this.preparedModel);
      // An empty border (<border><left/><right/><top/><bottom/><diagonal/></border>)
      // parses to undefined rather than {}, because there are no edges with
      // style/color data. This is correct — cells referencing borderId 0 should
      // not get a truthy border property.
      model.borders[0] = undefined;
      // The parser now preserves the named-style collections for round-trip
      // fidelity: the implicit "Normal" cellStyle and its base cellStyleXf.
      model.cellStyleXfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }];
      model.cellStyles = [{ name: "Normal", xfId: 0, builtinId: 0 }];
      return model;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("StylesXform", () => {
  testXformHelper(expectations);

  describe("As StyleManager", () => {
    it("Renders empty model", () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs.readFileSync(join(__dirname, "data/styles.2.2.xml")).toString();

      const xmlStream = new XmlWriter();
      stylesXform.render(xmlStream);

      // Use normalizeXml from test-xform-helper for consistent XML comparison
      expect(normalizeXml(xmlStream.xml)).toBe(normalizeXml(expectedXml));
    });
  });
});
