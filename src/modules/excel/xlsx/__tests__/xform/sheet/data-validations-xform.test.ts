import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { DataValidationsXform } from "@excel/xlsx/xform/sheet/data-validations-xform";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

const expectations = [
  {
    title: "list type",
    create: () => new DataValidationsXform(),
    preparedModel: {
      E1: {
        type: "list",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ["Ducks"]
      }
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1">
          <formula1>Ducks</formula1>
        </dataValidation>
      </dataValidations>
    `,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "whole type",
    create: () => new DataValidationsXform(),
    preparedModel: {
      A1: {
        type: "whole",
        operator: "between",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [5, 10]
      }
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "decimal type",
    create: () => new DataValidationsXform(),
    preparedModel: {
      A1: {
        type: "decimal",
        operator: "notBetween",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [5, 10]
      }
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "custom type",
    create: () => new DataValidationsXform(),
    preparedModel: {
      A1: {
        type: "custom",
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ["OR(C21=5,C21=7)"]
      }
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>OR(C21=5,C21=7)</formula1>
        </dataValidation>
      </dataValidations>
    `,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "parse open office",
    create: () => new DataValidationsXform(),
    preparedModel: {
      A1: {
        type: "whole",
        operator: "between",
        allowBlank: true,
        showInputMessage: false,
        formulae: [5, 10]
      }
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="true" showInputMessage="false" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    parsedModel: {
      A1: {
        type: "whole",
        operator: "between",
        allowBlank: true,
        showInputMessage: false,
        formulae: [5, 10]
      }
    },
    tests: ["parse"]
  },
  {
    title: "optimised",
    create: () => new DataValidationsXform(),
    preparedModel: {
      A1: { type: "whole", operator: "between", formulae: [5, 10] },
      A2: { type: "whole", operator: "between", formulae: [5, 10] },
      B1: { type: "whole", operator: "between", formulae: [5, 10] },
      B2: { type: "whole", operator: "between", formulae: [5, 10] }
    },
    parsedModel: {
      "range:A1:B2": { type: "whole", operator: "between", formulae: [5, 10] }
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" sqref="A1:B2">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ["render", "parse"]
  }
];

describe("DataValidationsXform", () => {
  testXformHelper(expectations);

  describe("optimisation with double-digit rows", () => {
    it("should merge C5:C15 into a single validation range", () => {
      // This is the exact data-validation edge case:
      // With localeCompare, "C10" < "C2" (string comparison),
      // which broke the downward expansion and produced overlapping ranges.
      const validation = { type: "list", formulae: ["Yes,No"] };
      const model: Record<string, any> = {};
      for (let row = 5; row <= 15; row++) {
        model[`C${row}`] = validation;
      }

      const xform = new DataValidationsXform();
      const xmlStream = new XmlWriter();
      xform.render(xmlStream, model);

      // Should produce exactly one <dataValidation> with sqref="C5:C15"
      expect(xmlStream.xml).toContain('sqref="C5:C15"');
      // Should NOT contain multiple dataValidation nodes
      const dvCount = (xmlStream.xml.match(/<dataValidation /g) || []).length;
      expect(dvCount).toBe(1);
    });

    it("should merge a 2D block with double-digit rows", () => {
      const validation = { type: "whole", operator: "between", formulae: [1, 100] };
      const model: Record<string, any> = {};
      // A5:B12 — 2 columns × 8 rows crossing the single/double digit boundary
      for (const col of ["A", "B"]) {
        for (let row = 5; row <= 12; row++) {
          model[`${col}${row}`] = validation;
        }
      }

      const xform = new DataValidationsXform();
      const xmlStream = new XmlWriter();
      xform.render(xmlStream, model);

      expect(xmlStream.xml).toContain('sqref="A5:B12"');
      const dvCount = (xmlStream.xml.match(/<dataValidation /g) || []).length;
      expect(dvCount).toBe(1);
    });
  });
});
