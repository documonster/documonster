import { Enums } from "@excel/core/enums";
import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { RowXform } from "@excel/xlsx/xform/sheet/row-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { describe } from "vitest";

const fakeStyles = {
  addStyleModel(style: any) {
    return style && JSON.stringify(style) !== "{}" ? 1 : 0;
  },
  getStyleModel(styleId: number) {
    return styleId ? { numFmt: "#" } : undefined;
  }
};

const fakeHyperlinkMap = {
  getHyperlink() {}
};

const expectations = [
  {
    title: "Plain",
    create: () => new RowXform(),
    initialModel: {
      number: 1,
      min: 1,
      max: 1,
      style: {},
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    get preparedModel() {
      return this.initialModel;
    },
    xml: '<row r="1" spans="1:1"><c r="A1"><v>5</v></c></row>',
    parsedModel: {
      number: 1,
      min: 1,
      max: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    reconciledModel: {
      number: 1,
      min: 1,
      max: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }],
      style: {}
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "No spans",
    create: () => new RowXform(),
    initialModel: {
      number: 1,
      style: {},
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    get preparedModel() {
      return this.initialModel;
    },
    xml: '<row r="1"><c r="A1"><v>5</v></c></row>',
    parsedModel: {
      number: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    reconciledModel: {
      number: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }],
      style: {}
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Styled",
    create: () => new RowXform(),
    initialModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }]
    },
    preparedModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      styleId: 1
    },
    xml: '<row r="2" spans="1:1" s="1" customFormat="1"><c r="A2"><v>5</v></c></row>',
    parsedModel: {
      number: 2,
      min: 1,
      max: 1,
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      styleId: 1
    },
    reconciledModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }]
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Outline",
    create: () => new RowXform(),
    initialModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      outlineLevel: 1,
      collapsed: true
    },
    preparedModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      outlineLevel: 1,
      styleId: 1,
      collapsed: true
    },
    xml: '<row r="2" spans="1:1" s="1" customFormat="1" outlineLevel="1" collapsed="1"><c r="A2"><v>5</v></c></row>',
    parsedModel: {
      number: 2,
      min: 1,
      max: 1,
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      outlineLevel: 1,
      collapsed: true,
      styleId: 1
    },
    reconciledModel: {
      number: 2,
      min: 1,
      max: 1,
      style: { numFmt: "#" },
      cells: [{ address: "A2", type: Enums.ValueType.Number, value: 5 }],
      outlineLevel: 1,
      collapsed: true
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Custom Height",
    create: () => new RowXform(),
    initialModel: {
      number: 1,
      min: 1,
      max: 1,
      style: {},
      height: 30,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    get preparedModel() {
      return this.initialModel;
    },
    xml: '<row r="1" ht="30" customHeight="1" spans="1:1"><c r="A1"><v>5</v></c></row>',
    parsedModel: {
      number: 1,
      min: 1,
      max: 1,
      height: 30,
      customHeight: true,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    reconciledModel: {
      number: 1,
      min: 1,
      max: 1,
      height: 30,
      customHeight: true,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }],
      style: {}
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Height zero (auto-height)",
    create: () => new RowXform(),
    initialModel: {
      number: 1,
      min: 1,
      max: 1,
      style: {},
      height: 0,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    get preparedModel() {
      return this.initialModel;
    },
    xml: '<row r="1" ht="1" spans="1:1"><c r="A1"><v>5</v></c></row>',
    parsedModel: {
      number: 1,
      min: 1,
      max: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    reconciledModel: {
      number: 1,
      min: 1,
      max: 1,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }],
      style: {}
    },
    tests: ["render", "renderIn"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Height with customHeight=false (ht without customHeight attribute)",
    create: () => new RowXform(),
    initialModel: {
      number: 1,
      min: 1,
      max: 1,
      style: {},
      height: 20,
      customHeight: false,
      cells: [{ address: "A1", type: Enums.ValueType.Number, value: 5 }]
    },
    get preparedModel() {
      return this.initialModel;
    },
    xml: '<row r="1" ht="20" spans="1:1"><c r="A1"><v>5</v></c></row>',
    tests: ["render", "renderIn"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  },
  {
    title: "Parse height without customHeight",
    create: () => new RowXform(),
    xml: '<row r="3" ht="25" spans="1:1"><c r="A3"><v>10</v></c></row>',
    parsedModel: {
      number: 3,
      min: 1,
      max: 1,
      height: 25,
      cells: [{ address: "A3", type: Enums.ValueType.Number, value: 10 }]
    },
    tests: ["parse"],
    options: {
      sharedStrings: new SharedStringsXform(),
      styles: fakeStyles,
      hyperlinkMap: fakeHyperlinkMap
    }
  }
];

describe("RowXform", () => {
  testXformHelper(expectations);
});
