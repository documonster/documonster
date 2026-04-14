import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { AutoFilterXform } from "@excel/xlsx/xform/sheet/auto-filter-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "Range",
    create() {
      return new AutoFilterXform();
    },
    preparedModel: "A1:C1",
    xml: '<autoFilter ref="A1:C1"/>',
    parsedModel: "A1:C1",
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "Row and Column Address",
    create() {
      return new AutoFilterXform();
    },
    preparedModel: { from: { row: 1, col: 1 }, to: { row: 1, col: 3 } },
    xml: '<autoFilter ref="A1:C1"/>',
    tests: ["render", "renderIn"]
  },
  {
    title: "String address",
    create() {
      return new AutoFilterXform();
    },
    preparedModel: { from: "A1", to: "C1" },
    xml: '<autoFilter ref="A1:C1"/>',
    tests: ["render", "renderIn"]
  }
];

describe("AutoFilterXform", () => {
  testXformHelper(expectations);
});
