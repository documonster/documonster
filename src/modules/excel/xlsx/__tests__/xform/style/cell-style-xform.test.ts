import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CellStyleXform } from "@excel/xlsx/xform/style/cell-style-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "builtin named style",
    create: () => new CellStyleXform(),
    preparedModel: { name: "Heading 1", xfId: 1, builtinId: 16 },
    xml: '<cellStyle name="Heading 1" xfId="1" builtinId="16"/>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "custom named style (no builtinId)",
    create: () => new CellStyleXform(),
    preparedModel: { name: "Accent", xfId: 2 },
    xml: '<cellStyle name="Accent" xfId="2"/>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "preserves hidden / customBuiltin / iLevel",
    create: () => new CellStyleXform(),
    preparedModel: {
      name: "RowLevel_1",
      xfId: 3,
      builtinId: 1,
      iLevel: 2,
      hidden: true,
      customBuiltin: true
    },
    xml: '<cellStyle name="RowLevel_1" xfId="3" builtinId="1" iLevel="2" hidden="1" customBuiltin="1"/>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("CellStyleXform", () => {
  testXformHelper(expectations);
});
