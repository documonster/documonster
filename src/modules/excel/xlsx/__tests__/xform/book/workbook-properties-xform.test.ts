import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { WorkbookPropertiesXform } from "@excel/xlsx/xform/book/workbook-properties-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "default",
    create() {
      return new WorkbookPropertiesXform();
    },
    preparedModel: {},
    xml: '<workbookPr filterPrivacy="1"/>',
    parsedModel: {},
    tests: ["render", "renderIn"]
  },
  {
    title: "date1904",
    create() {
      return new WorkbookPropertiesXform();
    },
    preparedModel: { date1904: true },
    xml: '<workbookPr date1904="1" filterPrivacy="1"/>',
    parsedModel: { date1904: true },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("WorkbookPropertiesXform", () => {
  testXformHelper(expectations);
});
