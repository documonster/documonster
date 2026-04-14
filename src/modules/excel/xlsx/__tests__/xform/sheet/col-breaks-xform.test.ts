import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { ColBreaksXform } from "@excel/xlsx/xform/sheet/col-breaks-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "empty column breaks",
    create() {
      return new ColBreaksXform();
    },
    initialModel: [],
    preparedModel: [],
    xml: "",
    parsedModel: undefined,
    tests: ["render"]
  },
  {
    title: "single column break",
    create() {
      return new ColBreaksXform();
    },
    initialModel: [{ id: 3, max: 1048575, man: 1 }],
    preparedModel: [{ id: 3, max: 1048575, man: 1 }],
    xml: '<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>',
    parsedModel: [{ id: 3, max: 1048575, man: 1 }],
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "multiple column breaks",
    create() {
      return new ColBreaksXform();
    },
    initialModel: [
      { id: 3, max: 1048575, man: 1 },
      { id: 6, max: 1048575, man: 1 },
      { id: 9, max: 1048575, man: 1 }
    ],
    preparedModel: [
      { id: 3, max: 1048575, man: 1 },
      { id: 6, max: 1048575, man: 1 },
      { id: 9, max: 1048575, man: 1 }
    ],
    xml: '<colBreaks count="3" manualBreakCount="3"><brk id="3" max="1048575" man="1"/><brk id="6" max="1048575" man="1"/><brk id="9" max="1048575" man="1"/></colBreaks>',
    parsedModel: [
      { id: 3, max: 1048575, man: 1 },
      { id: 6, max: 1048575, man: 1 },
      { id: 9, max: 1048575, man: 1 }
    ],
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "column break with min constraint",
    create() {
      return new ColBreaksXform();
    },
    initialModel: [{ id: 5, max: 100, min: 10, man: 1 }],
    preparedModel: [{ id: 5, max: 100, min: 10, man: 1 }],
    xml: '<colBreaks count="1" manualBreakCount="1"><brk id="5" max="100" min="10" man="1"/></colBreaks>',
    parsedModel: [{ id: 5, max: 100, min: 10, man: 1 }],
    tests: ["render", "renderIn", "parse"]
  }
];

describe("ColBreaksXform", () => {
  testXformHelper(expectations);
});
