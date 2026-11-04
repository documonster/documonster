/**
 * Every chart colour goes through one resolver, which applies the colour's DrawingML
 * transforms (`lumMod`, `lumOff`, `tint`, `shade`, `satMod`, `alpha`) with the engine every
 * module shares (`@utils/theme-colors` `applyColorTransforms`) and reads a scheme colour from
 * the workbook's theme when there is one.
 *
 * Checked against an Excel-saved chart (`workbook-roundtrip.xlsx`, Excel 16, the 2023 Office
 * theme): its gridlines and category axis are `tx1` lumMod 15 % / lumOff 85 % — Excel's
 * `#D9D9D9` — its axis labels `tx1` 65 % / 35 % — `#595959` — and its series `accent1`, which
 * that theme defines as `#156082`.
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { ChartClassic, ChartEx } from "@excel/chart/index";
import { resolveChartColor } from "@excel/chart/shared/chart-utils";
import { Chart, Workbook } from "@excel/index";
import { decompressPdfContent } from "@pdf/__tests__/test-helpers";
import { Pdf } from "@pdf/index";
import { describe, expect, it } from "vitest";

const FIXTURE = resolve(__dirname, "../../__tests__/data/workbook-roundtrip.xlsx");

async function fixtureChart() {
  const wb = await Workbook.read(readFileSync(FIXTURE));
  for (const ws of Workbook.getWorksheets(wb)) {
    const [chart] = Chart.list(ws);
    if (chart) {
      return chart;
    }
  }
  throw new Error("the fixture has a chart");
}

describe("an Excel-saved chart draws in the colours Excel shows", () => {
  it("gridlines and the category axis in #D9D9D9, labels in #595959, the series in the workbook's accent1", async () => {
    const svg = Chart.toSvg(await fixtureChart(), { width: 480, height: 300 }).toLowerCase();
    expect(svg).toMatch(/stroke="#d9d9d9"/);
    expect(svg).toMatch(/<text[^>]*fill="#595959"/);
    expect(svg).toContain('fill="#156082"');
    // None of it falls back to the untransformed theme colour, and the value axis Excel saved
    // with `<a:ln><a:noFill/>` draws no line.
    expect(svg).not.toMatch(/<polyline[^>]*stroke="#000000"/);
    expect(svg).not.toMatch(/<polyline[^>]*stroke="#444444"/);
  });
});

describe("the PDF converters read the workbook's theme too", () => {
  // `156082` as PDF operands: 0x15 / 255, 0x60 / 255, 0x82 / 255.
  const ACCENT1 = /0\.0824 0\.3765 0\.5098 rg/;
  it("Pdf.fromChart and Pdf.fromExcel fill the series in the workbook's accent1", async () => {
    const chart = await fixtureChart();
    expect(decompressPdfContent(await Pdf.fromChart(chart))).toMatch(ACCENT1);
    const wb = await Workbook.read(readFileSync(FIXTURE));
    expect(decompressPdfContent(await Pdf.fromExcel(wb))).toMatch(ACCENT1);
  });
});

describe("resolveChartColor", () => {
  it("applies the transforms in the order the writers emit them", () => {
    expect(resolveChartColor({ theme: 0, lumMod: 15000, lumOff: 85000 })).toBe("#D9D9D9");
    expect(resolveChartColor({ srgb: "FFFFFF", shade: 50000 })).toBe("#BDBDBD");
    // `tint` on the model is a fraction; a negative one is a shade (the writers' convention).
    expect(resolveChartColor({ srgb: "000000", tint: 0.5 })).toBe("#BDBDBD");
    expect(resolveChartColor({ srgb: "FFFFFF", tint: -0.5 })).toBe("#BDBDBD");
    expect(resolveChartColor({ srgb: "FF0000", satMod: 0 })).toBe("#808080");
    expect(resolveChartColor({ sysClr: "windowText", lumMod: 50000, lumOff: 50000 })).toBe(
      "#808080"
    );
    expect(resolveChartColor({ prstClr: "white", lumMod: 50000 })).toBe("#808080");
  });

  it("reads a scheme colour from the theme it is given, the default Office theme otherwise", () => {
    expect(resolveChartColor({ theme: 4 })).toBe("#4472C4");
    expect(resolveChartColor({ theme: 4 }, { accent1: "156082" })).toBe("#156082");
    expect(resolveChartColor({ theme: 4, lumMod: 75000 }, { accent1: "156082" })).not.toBe(
      "#156082"
    );
    // A slot the theme leaves out keeps the default.
    expect(resolveChartColor({ theme: 5 }, { accent1: "156082" })).toBe("#ED7D31");
  });
});

describe("ChartRenderOptions.themeColors", () => {
  const model = ChartClassic.createModel({
    type: "bar",
    series: [
      {
        name: "S",
        categories: "Sheet1!$A$1:$A$2",
        values: "Sheet1!$B$1:$B$2",
        spPr: { fill: { solid: { theme: 4 } } }
      }
    ]
  } as never);

  it("resolves the model's scheme colours, for a model with no workbook", () => {
    expect(ChartClassic.toSvg(model).toUpperCase()).toContain('FILL="#4472C4"');
    expect(
      ChartClassic.toSvg(model, { themeColors: { accent1: "156082" } }).toUpperCase()
    ).toContain('FILL="#156082"');
  });

  it("reaches the Office 2016 renderer", () => {
    const ex = ChartEx.createModel({
      type: "treemap",
      series: [
        { values: "Sheet1!$B$1:$B$2", literalValues: [10, 20], literalCategories: ["A", "B"] }
      ]
    } as never);
    ex.chartSpace.spPr = { fill: { solid: { theme: 4, lumMod: 20000, lumOff: 80000 } } };
    const area = (svg: string) => /<rect[^>]*stroke="#d9d9d9"[^>]*>/i.exec(svg)?.[0] ?? "";
    // accent1 at 20 % luminance plus 80 %: of the workbook's accent1, not the default one.
    const themed = area(ChartEx.toSvg(ex, { themeColors: { accent1: "156082" } }));
    const fallback = area(ChartEx.toSvg(ex));
    expect(themed).toMatch(/fill="#[0-9a-f]{6}"/i);
    expect(themed).not.toBe(fallback);
  });

  it("is checked", () => {
    expect(() => ChartClassic.toSvg(model, { themeColors: { accent1: "blue" } } as never)).toThrow(
      /options\.themeColors\.accent1/
    );
    expect(() => ChartClassic.toSvg(model, { themeColors: { nope: "FFFFFF" } } as never)).toThrow(
      /options\.themeColors/
    );
  });
});
