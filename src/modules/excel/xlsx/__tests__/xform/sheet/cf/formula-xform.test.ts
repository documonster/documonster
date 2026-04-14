import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { FormulaXform } from "@excel/xlsx/xform/sheet/cf/formula-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "formula",
    create() {
      return new FormulaXform();
    },
    preparedModel: "ROW()",
    xml: "<formula>ROW()</formula>",
    parsedModel: "ROW()",
    tests: ["render", "parse"]
  }
];

describe("FormulaXform", () => {
  testXformHelper(expectations);
});
