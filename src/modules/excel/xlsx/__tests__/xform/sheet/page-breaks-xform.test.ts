import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { PageBreaksXform } from "@excel/xlsx/xform/sheet/page-breaks-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "one page break",
    create() {
      return new PageBreaksXform();
    },
    initialModel: { id: 2, max: 3, min: 1, man: 1 },
    preparedModel: { id: 2, max: 3, min: 1, man: 1 },
    xml: '<brk id="2" max="3" min="1" man="1"/>',
    parsedModel: { id: 2, max: 3, min: 1, man: 1 },
    tests: ["prepare", "render", "renderIn"]
  }
];

describe("PageBreaksXform", () => {
  testXformHelper(expectations);
});
