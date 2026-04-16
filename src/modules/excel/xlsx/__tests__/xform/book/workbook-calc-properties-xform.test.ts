import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { WorkbookCalcPropertiesXform } from "@excel/xlsx/xform/book/workbook-calc-properties-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "default",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: {},
    xml: '<calcPr calcId="171027"/>',
    parsedModel: { fullCalcOnLoad: false },
    tests: ["render", "renderIn"]
  },
  {
    title: "fullCalcOnLoad",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { fullCalcOnLoad: true },
    xml: '<calcPr calcId="171027" fullCalcOnLoad="1"/>',
    parsedModel: { fullCalcOnLoad: true },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "iterate",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { iterate: true },
    xml: '<calcPr calcId="171027" iterate="1"/>',
    parsedModel: { fullCalcOnLoad: false, iterate: true },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "iterateCount",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { iterate: true, iterateCount: 50 },
    xml: '<calcPr calcId="171027" iterate="1" iterateCount="50"/>',
    parsedModel: { fullCalcOnLoad: false, iterate: true, iterateCount: 50 },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "iterateDelta",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { iterate: true, iterateCount: 100, iterateDelta: 0.001 },
    xml: '<calcPr calcId="171027" iterate="1" iterateCount="100" iterateDelta="0.001"/>',
    parsedModel: {
      fullCalcOnLoad: false,
      iterate: true,
      iterateCount: 100,
      iterateDelta: 0.001
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "all properties",
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: { fullCalcOnLoad: true, iterate: true, iterateCount: 200, iterateDelta: 0.01 },
    xml: '<calcPr calcId="171027" fullCalcOnLoad="1" iterate="1" iterateCount="200" iterateDelta="0.01"/>',
    parsedModel: {
      fullCalcOnLoad: true,
      iterate: true,
      iterateCount: 200,
      iterateDelta: 0.01
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("WorkbookCalcPropertiesXform", () => {
  testXformHelper(expectations);
});
