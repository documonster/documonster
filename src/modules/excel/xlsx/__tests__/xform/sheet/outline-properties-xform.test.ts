import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { OutlinePropertiesXform } from "@excel/xlsx/xform/sheet/outline-properties-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "empty",
    create() {
      return new OutlinePropertiesXform();
    },
    preparedModel: {},
    xml: "",
    parsedModel: {},
    tests: ["render", "renderIn"]
  },
  {
    title: "summaryBelow",
    create() {
      return new OutlinePropertiesXform();
    },
    preparedModel: { summaryBelow: false },
    xml: '<outlinePr summaryBelow="0"/>',
    parsedModel: { summaryBelow: false },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "summaryRight",
    create() {
      return new OutlinePropertiesXform();
    },
    preparedModel: { summaryRight: false },
    xml: '<outlinePr summaryRight="0"/>',
    parsedModel: { summaryRight: false },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "summaryBelow + summaryRight",
    create() {
      return new OutlinePropertiesXform();
    },
    preparedModel: { summaryBelow: true, summaryRight: false },
    xml: '<outlinePr summaryBelow="1" summaryRight="0"/>',
    parsedModel: { summaryBelow: true, summaryRight: false },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("OutlinePropertiesXform", () => {
  testXformHelper(expectations);
});
