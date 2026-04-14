import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { FloatXform } from "@excel/xlsx/xform/simple/float-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "five",
    create() {
      return new FloatXform({ tag: "float", attr: "val" });
    },
    preparedModel: 5,
    xml: '<float val="5"/>',
    parsedModel: 5,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "pi",
    create() {
      return new FloatXform({ tag: "float", attr: "val" });
    },
    preparedModel: 3.14,
    xml: '<float val="3.14"/>',
    parsedModel: 3.14,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "zero",
    create() {
      return new FloatXform({ tag: "float", attr: "val" });
    },
    preparedModel: 0,
    xml: '<float val="0"/>',
    parsedModel: 0,
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "undefined",
    create() {
      return new FloatXform({ tag: "float", attr: "val" });
    },
    preparedModel: undefined,
    xml: "",
    tests: ["render", "renderIn"]
  }
];

describe("FloatXform", () => {
  testXformHelper(expectations);
});
