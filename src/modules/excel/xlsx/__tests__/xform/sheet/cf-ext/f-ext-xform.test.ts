import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { FExtXform } from "@excel/xlsx/xform/sheet/cf-ext/f-ext-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "formula",
    create() {
      return new FExtXform();
    },
    preparedModel: "7",
    xml: "<xm:f>7</xm:f>",
    parsedModel: "7",
    tests: ["render", "parse"]
  }
];

describe("FExtXform", () => {
  testXformHelper(expectations);
});
