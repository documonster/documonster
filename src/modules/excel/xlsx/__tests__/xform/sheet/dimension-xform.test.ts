import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { DimensionXform } from "@excel/xlsx/xform/sheet/dimension-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "Dimension",
    create() {
      return new DimensionXform();
    },
    preparedModel: "A1:F5",
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<dimension ref="A1:F5"/>',
    tests: ["render", "renderIn", "parse"]
  }
];

describe("DimensionXform", () => {
  testXformHelper(expectations);
});
