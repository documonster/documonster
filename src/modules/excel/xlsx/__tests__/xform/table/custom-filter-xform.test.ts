import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CustomFilterXform } from "@excel/xlsx/xform/table/custom-filter-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "custom filter",
    create() {
      return new CustomFilterXform();
    },
    preparedModel: { val: "*brandywine*" },
    xml: '<customFilter val="*brandywine*"/>',
    parsedModel: { val: "*brandywine*" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "custom filter with operator",
    create() {
      return new CustomFilterXform();
    },
    preparedModel: { operator: "notEqual", val: "4" },
    xml: '<customFilter operator="notEqual" val="4"/>',
    parsedModel: { operator: "notEqual", val: "4" },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("CustomFilterXform", () => {
  testXformHelper(expectations);
});
