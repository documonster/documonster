import { Enums } from "@excel/enums";
import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CellXform } from "@excel/xlsx/xform/sheet/cell-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { describe } from "vitest";

const fakeStyles = {
  addStyleModel(style: any, effectiveType: any) {
    if (effectiveType === Enums.ValueType.Date) {
      return 1;
    }
    return 0;
  },
  getStyleModel(styleId: number) {
    switch (styleId) {
      case 1:
        return { numFmt: "mm-dd-yy" };
      case 2:
        return { numFmt: "mmm-yy" };
      default:
        return null;
    }
  }
};

const fakeHyperlinkMap = {
  H1: "http://www.foo.com",
  H2: "http://www.rich.com"
};

const expectations = [
  {
    title: "Styled Null",
    create() {
      return new CellXform();
    },
    preparedModel: { address: "A1", type: Enums.ValueType.Null, styleId: 1 },
    xml: '<c r="A1" s="1" />',
    parsedModel: { address: "A1", type: Enums.ValueType.Null, styleId: 1 },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "Number",
    create() {
      return new CellXform();
    },
    preparedModel: { address: "A1", type: Enums.ValueType.Number, value: 5 },
    parsedModel: { address: "A1", type: Enums.ValueType.Number, value: 5 },
    xml: '<c r="A1"><v>5</v></c>',
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "Boolean",
    create() {
      return new CellXform();
    },
    preparedModel: { address: "A1", type: Enums.ValueType.Boolean, value: true },
    parsedModel: { address: "A1", type: Enums.ValueType.Boolean, value: true },
    xml: '<c r="A1" t="b"><v>1</v></c>',
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "Error",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.Error,
      value: { error: "#N/A" }
    },
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.Error,
      value: { error: "#N/A" }
    },
    xml: '<c r="A1" t="e"><v>#N/A</v></c>',
    tests: ["render", "renderIn", "parse"]
  },
  {
    // Strict OpenXML format stores dates as ISO strings with t="d"
    title: "Date (Strict OpenXML format)",
    create() {
      return new CellXform();
    },
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.Date,
      value: new Date("2024-02-09")
    },
    xml: '<c r="A1" t="d"><v>2024-02-09</v></c>',
    tests: ["parse"]
  },
  {
    title: "String",
    create() {
      return new CellXform();
    },
    initialModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
    preparedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
    xml: '<c r="A1" t="str"><v>Foo</v></c>',
    parsedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Foo"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles }
  },
  {
    title: "String with Invalid Number",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "6E1000"
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "6E1000"
    },
    xml: '<c r="A1" t="str"><v>6E1000</v></c>',
    parsedModel: { address: "A1", type: Enums.ValueType.String, value: "6E1000" },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "6E1000"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    title: "Inline String with plain text",
    create() {
      return new CellXform();
    },
    xml: '<c r="A1" t="inlineStr"><is><t>Foo</t></is></c>',
    parsedModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Foo"
    },
    tests: ["parse", "reconcile"],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles }
  },
  {
    title: "Inline String with OOXML escape (lowercase _x000a_)",
    create() {
      return new CellXform();
    },
    xml: '<c r="A1" t="inlineStr"><is><t>Col3_x000a_new line</t></is></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Col3\nnew line"
    },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Col3\nnew line"
    },
    tests: ["parse", "reconcile"],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles }
  },
  {
    title: "Inline String with RichText",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    xml: '<c r="A1" t="inlineStr"><is><r><rPr><color rgb="FF0000"/></rPr><t>red</t></r><r><rPr><color rgb="00FF00"/></rPr><t>green</t></r></is></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: { hyperlinkMap: fakeHyperlinkMap, styles: fakeStyles }
  },
  {
    title: "Shared String",
    create() {
      return new CellXform();
    },
    initialModel: { address: "A1", type: Enums.ValueType.String, value: "Foo" },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Foo",
      ssId: 0
    },
    xml: '<c r="A1" t="s"><v>0</v></c>',
    parsedModel: { address: "A1", type: Enums.ValueType.String, value: 0 },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: "Foo"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    title: "Shared String with RichText",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A1",
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      },
      ssId: 0
    },
    xml: '<c r="A1" t="s"><v>0</v></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.String,
      value: 0
    },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.RichText,
      value: {
        richText: [
          { font: { color: { argb: "FF0000" } }, text: "red" },
          { font: { color: { argb: "00FF00" } }, text: "green" }
        ]
      }
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    title: "Date",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A1",
      type: Enums.ValueType.Date,
      value: new Date("2016-06-09T00:00:00.000Z")
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.Date,
      value: new Date("2016-06-09T00:00:00.000Z"),
      styleId: 1
    },
    xml: '<c r="A1" s="1"><v>42530</v></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.Number,
      value: 42530,
      styleId: 1
    },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.Date,
      value: new Date("2016-06-09T00:00:00.000Z"),
      style: { numFmt: "mm-dd-yy" }
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    title: "Hyperlink",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "H1",
      type: Enums.ValueType.Hyperlink,
      hyperlink: "http://www.foo.com",
      text: "www.foo.com"
    },
    preparedModel: {
      address: "H1",
      type: Enums.ValueType.Hyperlink,
      hyperlink: "http://www.foo.com",
      text: "www.foo.com",
      ssId: 0
    },
    xml: '<c r="H1" t="s"><v>0</v></c>',
    parsedModel: { address: "H1", type: Enums.ValueType.String, value: 0 },
    reconciledModel: {
      address: "H1",
      type: Enums.ValueType.Hyperlink,
      text: "www.foo.com",
      hyperlink: "http://www.foo.com"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    // Regression test for issue #142:
    // A hyperlink whose display text is rich-text must round-trip with
    //   * `text` flattened to a plain string (public contract),
    //   * `richText` runs preserved for re-writing with formatting intact.
    title: "Hyperlink with rich-text display",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "H2",
      type: Enums.ValueType.Hyperlink,
      hyperlink: "http://www.rich.com",
      text: "hello world",
      richText: [{ text: "hello ", font: { bold: true } }, { text: "world" }]
    },
    preparedModel: {
      address: "H2",
      type: Enums.ValueType.Hyperlink,
      hyperlink: "http://www.rich.com",
      text: "hello world",
      richText: [{ text: "hello ", font: { bold: true } }, { text: "world" }],
      ssId: 0
    },
    xml: '<c r="H2" t="s"><v>0</v></c>',
    // When parsing the <c>, the shared-string index resolves later during
    // reconcile. The shared-string itself carries the rich-text payload.
    parsedModel: { address: "H2", type: Enums.ValueType.String, value: 0 },
    reconciledModel: {
      address: "H2",
      type: Enums.ValueType.Hyperlink,
      text: "hello world",
      hyperlink: "http://www.rich.com",
      richText: [{ text: "hello ", font: { bold: true } }, { text: "world" }]
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles
    }
  },
  {
    title: "String Formula",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: "Foo"
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: "Foo"
    },
    xml: '<c r="A1" t="str"><f>A2</f><v>Foo</v></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: "Foo"
    },
    reconciledModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: "Foo"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0
    }
  },
  {
    title: "Number Formula",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: 7
    },
    xml: '<c r="A1"><f>A2</f><v>7</v></c>',
    parsedModel: {
      address: "A1",
      type: Enums.ValueType.Formula,
      formula: "A2",
      result: 7
    },
    tests: ["render", "renderIn", "parse"],
    options: { formulae: {}, siFormulae: 0 }
  },
  {
    title: "Master Shared Formula",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "shared",
      ref: "A2:B2",
      formula: "A1",
      result: 2,
      si: 0
    },
    xml: '<c r="A2"><f t="shared" ref="A2:B2" si="0">A1</f><v>2</v></c>',
    parsedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "shared",
      ref: "A2:B2",
      formula: "A1",
      result: 2,
      si: "0"
    },
    reconciledModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "shared",
      ref: "A2:B2",
      formula: "A1",
      result: 2
    },
    tests: ["render", "renderIn", "parse", "reconcile"],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0
    }
  },
  {
    title: "Shared Formula Slave",
    create() {
      return new CellXform();
    },
    initialModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      sharedFormula: "A1",
      result: 2
    },
    preparedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      sharedFormula: "A1",
      result: 2,
      si: 0
    },
    xml: '<c r="A2"><f t="shared" si="0" /><v>2</v></c>',
    parsedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      result: 2,
      shareType: "shared",
      si: "0"
    },
    reconciledModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      result: 2,
      sharedFormula: "A1"
    },
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {
        A1: {
          address: "A1",
          type: Enums.ValueType.Formula,
          formula: "ROW()",
          result: 1,
          si: 0
        },
        0: "A1"
      },
      siFormulae: 1
    }
  },
  {
    title: "Array Shared Formula",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "A2:B2",
      formula: "A1",
      result: 2
    },
    xml: '<c r="A2"><f t="array" ref="A2:B2">A1</f><v>2</v></c>',
    parsedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "A2:B2",
      formula: "A1",
      result: 2
    },
    reconciledModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "A2:B2",
      formula: "A1",
      result: 2
    },
    tests: ["render", "renderIn", "parse", "reconcile"],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0
    }
  },
  // String formula result with date format should not be converted to date
  {
    title: "String Formula with Date Format",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      formula: "A1",
      result: "test",
      styleId: 2
    },
    xml: '<c r="A2" s="2" t="str"><f>A1</f><v>test</v></c>',
    parsedModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      formula: "A1",
      result: "test",
      styleId: 2
    },
    reconciledModel: {
      address: "A2",
      type: Enums.ValueType.Formula,
      formula: "A1",
      result: "test",
      style: { numFmt: "mmm-yy" }
    },
    tests: ["render", "renderIn", "parse", "reconcile"],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0
    }
  },
  // Dynamic Array Formula (Excel 365 FILTER/SORT/UNIQUE etc.)
  {
    title: "Dynamic Array Formula",
    create() {
      return new CellXform();
    },
    preparedModel: {
      address: "D1",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "D1",
      formula: "_xlfn._xlws.FILTER(A1:A10,B1:B10>5)",
      result: 3,
      isDynamicArray: true,
      cm: 1
    },
    xml: '<c r="D1" cm="1"><f t="array" ref="D1">_xlfn._xlws.FILTER(A1:A10,B1:B10&gt;5)</f><v>3</v></c>',
    parsedModel: {
      address: "D1",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "D1",
      formula: "_xlfn._xlws.FILTER(A1:A10,B1:B10>5)",
      result: 3,
      cm: 1
    },
    reconciledModel: {
      address: "D1",
      type: Enums.ValueType.Formula,
      shareType: "array",
      ref: "D1",
      formula: "_xlfn._xlws.FILTER(A1:A10,B1:B10>5)",
      result: 3,
      isDynamicArray: true
    },
    tests: ["render", "renderIn", "parse", "reconcile"],
    options: {
      styles: fakeStyles,
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      formulae: {},
      siFormulae: 0,
      hasDynamicArrayMetadata: true,
      dynamicArrayCmIndices: new Set([1])
    }
  }
];

describe("CellXform", () => {
  testXformHelper(expectations);
});
