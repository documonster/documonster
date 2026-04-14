import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { SqrefExtXform } from "@excel/xlsx/xform/sheet/cf-ext/sqref-ext-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "range",
    create() {
      return new SqrefExtXform();
    },
    preparedModel: "A1:C3",
    xml: "<xm:sqref>A1:C3</xm:sqref>",
    parsedModel: "A1:C3",
    tests: ["render", "parse"]
  }
];

describe("SqrefExtXform", () => {
  testXformHelper(expectations);
});
