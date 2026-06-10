import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { BlipXform } from "@excel/xlsx/xform/drawing/blip-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "full",
    create() {
      return new BlipXform();
    },
    preparedModel: { rId: "rId1" },
    xml: '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />',
    parsedModel: { rId: "rId1" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "svg companion",
    create() {
      return new BlipXform();
    },
    preparedModel: { rId: "rId1", svgRId: "rId2" },
    xml:
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print">' +
      '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId2" />' +
      "</a:ext></a:extLst>" +
      "</a:blip>",
    parsedModel: { rId: "rId1", svgRId: "rId2" },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "alpha + svg companion",
    create() {
      return new BlipXform();
    },
    preparedModel: { rId: "rId1", alphaModFix: 50000, svgRId: "rId2" },
    xml:
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print">' +
      '<a:alphaModFix amt="50000" />' +
      '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId2" />' +
      "</a:ext></a:extLst>" +
      "</a:blip>",
    parsedModel: { rId: "rId1", alphaModFix: 50000, svgRId: "rId2" },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("BlipXform", () => {
  testXformHelper(expectations);
});
