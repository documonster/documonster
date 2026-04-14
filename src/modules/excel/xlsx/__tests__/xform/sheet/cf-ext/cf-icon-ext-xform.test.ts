import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CfIconExtXform } from "@excel/xlsx/xform/sheet/cf-ext/cf-icon-ext-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "range",
    create() {
      return new CfIconExtXform();
    },
    preparedModel: { iconSet: "3Triangles", iconId: 7 },
    xml: '<x14:cfIcon iconSet="3Triangles" iconId="7" />',
    parsedModel: { iconSet: "3Triangles", iconId: 7 },
    tests: ["render", "parse"]
  }
];

describe("CfIconExtXform", () => {
  testXformHelper(expectations);
});
