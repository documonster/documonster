import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { WorkbookCalcPropertiesXform } from "@excel/xlsx/xform/book/workbook-calc-properties-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "default",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: {},
    xml: '<calcPr calcId="171027"/>',
    parsedModel: {},
    tests: ["render", "renderIn"]
  },
  {
    title: "fullCalcOnLoad",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { fullCalcOnLoad: true },
    xml: '<calcPr calcId="171027" fullCalcOnLoad="1"/>',
    parsedModel: {},
    tests: ["render", "renderIn", "parse"]
  }
];

describe("WorkbookCalcPropertiesXform", () => {
  testXformHelper(expectations);
});
