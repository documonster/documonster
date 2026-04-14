import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CfvoXform } from "@excel/xlsx/xform/sheet/cf/cfvo-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "min",
    create() {
      return new CfvoXform();
    },
    preparedModel: { type: "min" },
    xml: '<cfvo type="min" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "parse"]
  },
  {
    title: "percent",
    create() {
      return new CfvoXform();
    },
    preparedModel: { type: "percent", value: 12.5 },
    xml: '<cfvo type="percent" val="12.5" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "parse"]
  }
];

describe("CfvoXform", () => {
  testXformHelper(expectations);
});
