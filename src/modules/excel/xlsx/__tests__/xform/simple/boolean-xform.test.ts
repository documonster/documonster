import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { BooleanXform } from "@excel/xlsx/xform/simple/boolean-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "true",
    create() {
      return new BooleanXform({ tag: "boolean", attr: "val" });
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: "<boolean/>",
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "false",
    create() {
      return new BooleanXform({ tag: "boolean", attr: "val" });
    },
    preparedModel: false,
    xml: "",
    tests: ["render", "renderIn"]
  },
  {
    title: "undefined",
    create() {
      return new BooleanXform({ tag: "boolean", attr: "val" });
    },
    preparedModel: undefined,
    xml: "",
    tests: ["render", "renderIn"]
  }
];

describe("BooleanXform", () => {
  testXformHelper(expectations);
});
