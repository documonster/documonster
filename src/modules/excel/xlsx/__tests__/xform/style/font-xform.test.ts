import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { FontXform } from "@excel/xlsx/xform/style/font-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "green bold",
    create() {
      return new FontXform();
    },
    preparedModel: {
      bold: true,
      size: 14,
      color: { argb: "FF00FF00" },
      name: "Calibri",
      family: 2,
      scheme: "minor"
    },
    xml: '<font><b/><sz val="14"/><color rgb="FF00FF00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "rPr tag",
    create() {
      return new FontXform({ tagName: "rPr", fontNameTag: "rFont" });
    },
    preparedModel: {
      italic: true,
      size: 14,
      color: { argb: "FF00FF00" },
      name: "Calibri",
      family: 2,
      scheme: "minor"
    },
    xml: '<rPr><i/><sz val="14"/><color rgb="FF00FF00"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("FontXform", () => {
  testXformHelper(expectations);
});
