import { ChartEx } from "@excel/chart";
import { themeXmlColors } from "@excel/chart/render/workbook-theme";
import { describe, it, expect } from "vitest";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";

function theme(open: string, p: string): string {
  return (
    `${open}<${p}themeElements><${p}clrScheme name="t">` +
    `<${p}accent1><${p}srgbClr val="FF1234"/></${p}accent1>` +
    `</${p}clrScheme></${p}themeElements></${p}theme>`
  );
}

function chartEx(open: string, p: string): string {
  return (
    `${open}<${p}chart><${p}title><${p}tx><${p}txData><${p}v>WRONG</${p}v></${p}txData></${p}tx></${p}title>` +
    `<${p}plotArea><${p}plotAreaRegion><${p}series layoutId="sunburst"/></${p}plotAreaRegion></${p}plotArea>` +
    `</${p}chart></${p}chartSpace>`
  );
}

describe("theme colours are read by namespace, not by prefix", () => {
  it("ignores a clrScheme outside DrawingML", () => {
    expect(themeXmlColors(theme(`<a:theme xmlns:a="urn:WRONG">`, "a:"))).toBeUndefined();
  });

  it("reads an alias prefix and a default namespace", () => {
    expect(themeXmlColors(theme(`<d:theme xmlns:d="${A}">`, "d:"))?.accent1).toBe("FF1234");
    expect(themeXmlColors(theme(`<theme xmlns="${A}">`, ""))?.accent1).toBe("FF1234");
  });
});

describe("ChartEx.fromXml checks the chartSpace namespace", () => {
  it("refuses a cx: root bound to another namespace", () => {
    expect(() => ChartEx.fromXml(chartEx(`<cx:chartSpace xmlns:cx="urn:WRONG">`, "cx:"))).toThrow(
      /chartSpace/
    );
  });

  it("accepts an alias prefix and a default namespace", () => {
    for (const [open, p] of [
      [`<c:chartSpace xmlns:c="${CX}">`, "c:"],
      [`<chartSpace xmlns="${CX}">`, ""]
    ]) {
      expect(JSON.stringify(ChartEx.fromXml(chartEx(open, p)).chartSpace.chart.title)).toContain(
        "WRONG"
      );
    }
  });
});
