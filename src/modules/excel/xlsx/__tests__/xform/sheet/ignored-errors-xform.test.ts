import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { IgnoredErrorsXform } from "@excel/xlsx/xform/sheet/ignored-errors-xform";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

const expectations = [
  {
    title: "empty ignoredErrors",
    create() {
      return new IgnoredErrorsXform();
    },
    preparedModel: [],
    xml: "",
    tests: ["render"]
  },
  {
    title: "single ignoredError - numberStoredAsText",
    create() {
      return new IgnoredErrorsXform();
    },
    preparedModel: [{ ref: "A1:XFD1048576", numberStoredAsText: true }],
    xml: '<ignoredErrors><ignoredError sqref="A1:XFD1048576" numberStoredAsText="1"/></ignoredErrors>',
    parsedModel: [{ ref: "A1:XFD1048576", numberStoredAsText: true }],
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "multiple boolean attributes",
    create() {
      return new IgnoredErrorsXform();
    },
    preparedModel: [{ ref: "A1:B10", numberStoredAsText: true, formula: true, evalError: true }],
    xml: '<ignoredErrors><ignoredError sqref="A1:B10" numberStoredAsText="1" formula="1" evalError="1"/></ignoredErrors>',
    parsedModel: [{ ref: "A1:B10", numberStoredAsText: true, formula: true, evalError: true }],
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "multiple ignoredError entries",
    create() {
      return new IgnoredErrorsXform();
    },
    preparedModel: [
      { ref: "A1:A100", numberStoredAsText: true },
      { ref: "B1:B100", formula: true }
    ],
    xml: '<ignoredErrors><ignoredError sqref="A1:A100" numberStoredAsText="1"/><ignoredError sqref="B1:B100" formula="1"/></ignoredErrors>',
    parsedModel: [
      { ref: "A1:A100", numberStoredAsText: true },
      { ref: "B1:B100", formula: true }
    ],
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "all boolean attributes",
    create() {
      return new IgnoredErrorsXform();
    },
    preparedModel: [
      {
        ref: "A1:Z100",
        numberStoredAsText: true,
        formula: true,
        formulaRange: true,
        unlockedFormula: true,
        emptyCellReference: true,
        listDataValidation: true,
        calculatedColumn: true,
        evalError: true,
        twoDigitTextYear: true
      }
    ],
    xml: '<ignoredErrors><ignoredError sqref="A1:Z100" numberStoredAsText="1" formula="1" formulaRange="1" unlockedFormula="1" emptyCellReference="1" listDataValidation="1" calculatedColumn="1" evalError="1" twoDigitTextYear="1"/></ignoredErrors>',
    parsedModel: [
      {
        ref: "A1:Z100",
        numberStoredAsText: true,
        formula: true,
        formulaRange: true,
        unlockedFormula: true,
        emptyCellReference: true,
        listDataValidation: true,
        calculatedColumn: true,
        evalError: true,
        twoDigitTextYear: true
      }
    ],
    tests: ["render", "renderIn", "parse"]
  }
];

describe("IgnoredErrorsXform", () => {
  testXformHelper(expectations);

  it("does not render when model is undefined", () => {
    const xform = new IgnoredErrorsXform();
    const xmlStream = new XmlWriter();
    xform.render(xmlStream, undefined);
    expect(xmlStream.xml).toBe("");
  });

  it("skips false-y boolean attributes during render", () => {
    const xform = new IgnoredErrorsXform();
    const xmlStream = new XmlWriter();
    xform.render(xmlStream, [{ ref: "A1:B10", numberStoredAsText: true, formula: false }]);
    const xml = xmlStream.xml;
    expect(xml).toContain('numberStoredAsText="1"');
    expect(xml).not.toContain("formula");
  });
});
