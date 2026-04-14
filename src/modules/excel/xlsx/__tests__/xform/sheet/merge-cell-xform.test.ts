import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { MergeCellXform } from "@excel/xlsx/xform/sheet/merge-cell-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "Merge",
    create() {
      return new MergeCellXform();
    },
    preparedModel: "B2:C4",
    xml: '<mergeCell ref="B2:C4"/>',
    parsedModel: "B2:C4",
    tests: ["render", "renderIn", "parse"]
  }
];

describe("MergeCellXform", () => {
  testXformHelper(expectations);
});
