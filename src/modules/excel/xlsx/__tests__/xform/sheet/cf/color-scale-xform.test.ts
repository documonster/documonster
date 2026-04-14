import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { ColorScaleXform } from "@excel/xlsx/xform/sheet/cf/color-scale-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "Colour Scale",
    create() {
      return new ColorScaleXform();
    },
    preparedModel: {
      cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
      color: [{ argb: "FFFF0000" }, { argb: "FF00FF00" }, { argb: "FF0000FF" }]
    },
    xml: `
      <colorScale>
        <cfvo type="min" />
        <cfvo type="percentile" val="50" />
        <cfvo type="max" />
        <color rgb="FFFF0000" />
        <color rgb="FF00FF00" />
        <color rgb="FF0000FF" />
      </colorScale>
    `,
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "parse"]
  }
];

describe("ColorScaleXform", () => {
  testXformHelper(expectations);
});
