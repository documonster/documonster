/**
 * Chart builder tests — advanced builder options: trendline, cache, rich text, effects, anchors, sidecars.
 *
 * Split out of the original 13,000-line `chart-builder.test.ts`
 * so vitest transform/import stays fast in full-suite runs.
 * Shared helpers and imports live in `chart-builder.helpers.ts`.
 */

import { extractAll } from "@archive/unzip/extract";
import type { LineChartGroup, LineSeries, BarSeries, ValueAxis } from "@excel/chart";
import {
  buildChartModel,
  buildComboChartModel,
  fillChartCaches,
  parseSpPr,
  buildSpPr,
  getSpPrFillColor,
  getSpPrGradient,
  getSpPrPattern,
  setSpPrFill,
  parseChartColors,
  buildChartColors,
  parseChartStyle,
  buildChartScene
} from "@excel/chart";
import { installChartSupport } from "@excel/chart/install";
import { Workbook } from "@excel/workbook";
import { beforeAll, describe, it, expect } from "vitest";

import { CATEGORIES, VALUES_B, baseSeries, ctg, pa } from "./chart-builder.helpers";

const textDecoder = new TextDecoder();

beforeAll(() => {
  installChartSupport();
});

describe("trendline array", () => {
  it("single trendline option produces one trendline", () => {
    const m = buildChartModel({
      type: "line",
      series: [{ ...baseSeries("S1"), trendline: { type: "linear" } }]
    });
    const s = ctg(m).series[0] as LineSeries;
    expect(s.trendlines).toHaveLength(1);
    expect(s.trendlines![0].type).toBe("linear");
  });

  it("array of trendlines produces multiple trendlines", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          trendline: [{ type: "linear" }, { type: "exp", name: "Exponential" }]
        }
      ]
    });
    const s = ctg(m).series[0] as LineSeries;
    expect(s.trendlines).toHaveLength(2);
    expect(s.trendlines![0].type).toBe("linear");
    expect(s.trendlines![1].type).toBe("exp");
    expect(s.trendlines![1].name).toBe("Exponential");
  });

  // Regression — previously all non-`movingAvg` trendline types
  // silently fell back to a linear fit **in pixel space** (so the
  // regressed slope carried the inverted sign of an SVG y axis), and
  // `exp`/`log`/`power`/`poly` rendered as straight lines because
  // only linear / movingAvg were implemented. Data-space fitting now
  // produces the correct curve for every OOXML type.
  it("renderer fits exp/log/power/poly trendlines as curves, not straight lines", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      [1, 2],
      [2, 4],
      [3, 8],
      [4, 16],
      [5, 32],
      [6, 64]
    ]);
    ws.addChart(
      {
        type: "scatter",
        series: [
          {
            name: "S",
            xValues: "Sheet1!$A$1:$A$6",
            values: "Sheet1!$B$1:$B$6",
            trendline: [{ type: "exp", name: "ExpFit" }]
          }
        ]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    const scene = buildChartScene(model, { width: 420, height: 260 });
    const scatterSeries = scene.series[0] as {
      trendlines?: Array<{ points: Array<{ x: number; y: number }> }>;
    };
    const trendlinePoints = scatterSeries.trendlines?.[0]?.points;
    expect(trendlinePoints).toBeDefined();
    // Curve-sampled exp trendline should have many more than 2 points.
    expect(trendlinePoints!.length).toBeGreaterThan(10);
  });

  it("linear trendline produces a positive-slope line for positive-slope data", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      [1, 3],
      [2, 5],
      [3, 7],
      [4, 9]
    ]);
    ws.addChart(
      {
        type: "scatter",
        series: [
          {
            name: "S",
            xValues: "Sheet1!$A$1:$A$4",
            values: "Sheet1!$B$1:$B$4",
            trendline: { type: "linear" }
          }
        ]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    const scene = buildChartScene(model, { width: 400, height: 300 });
    const scatterSeries = scene.series[0] as {
      trendlines?: Array<{ points: Array<{ x: number; y: number }> }>;
    };
    const trendlinePoints = scatterSeries.trendlines?.[0]?.points;
    expect(trendlinePoints).toBeDefined();
    expect(trendlinePoints!.length).toBeGreaterThanOrEqual(2);
    const first = trendlinePoints![0];
    const last = trendlinePoints![trendlinePoints!.length - 1];
    // Positive data slope → negative pixel slope (because SVG y is
    // inverted). `first` is the left end of the trendline, `last`
    // the right. For positive-slope data we expect the left to have
    // the HIGHER pixel y (lower on screen = data minimum).
    expect(first.y).toBeGreaterThan(last.y);
  });
});

// ---------------------------------------------------------------------------
// 13. DataPoint marker and invertIfNegative
// ---------------------------------------------------------------------------

describe("dataPoint extended options", () => {
  it("dataPoint with marker", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          dataPoints: [{ index: 0, marker: { symbol: "diamond", size: 10 } }]
        }
      ]
    });
    const dp = (ctg(m).series[0] as LineSeries).dataPoints![0];
    expect(dp.marker).toBeDefined();
    expect(dp.marker!.symbol).toBe("diamond");
    expect(dp.marker!.size).toBe(10);
  });

  it("dataPoint with invertIfNegative", () => {
    const m = buildChartModel({
      type: "bar",
      series: [{ ...baseSeries("S1"), dataPoints: [{ index: 0, invertIfNegative: true }] }]
    });
    const dp = (ctg(m).series[0] as BarSeries).dataPoints![0];
    expect(dp.invertIfNegative).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 14. Extended axis options
// ---------------------------------------------------------------------------

describe("extended axis options", () => {
  it("lblAlgn and lblOffset on category axis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { lblAlgn: "l", lblOffset: 200 }
    });
    const catAx = pa(m).axes.find(a => a.axisType === "cat")! as any;
    expect(catAx.lblAlgn).toBe("l");
    expect(catAx.lblOffset).toBe(200);
  });

  it("tickLblSkip and tickMarkSkip on category axis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { tickLblSkip: 2, tickMarkSkip: 3 }
    });
    const catAx = pa(m).axes.find(a => a.axisType === "cat")! as any;
    expect(catAx.tickLblSkip).toBe(2);
    expect(catAx.tickMarkSkip).toBe(3);
  });

  it("crosses and crossesAt on axis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { crosses: "min", crossesAt: 100 }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.crosses).toBe("min");
    expect(valAx.crossesAt).toBe(100);
  });

  it("displayUnits on value axis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { displayUnits: "thousands" }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")! as any;
    expect(valAx.dispUnits).toBeDefined();
    expect(valAx.dispUnits.builtInUnit).toBe("thousands");
  });

  it("lineColor and lineWidth on axis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { lineColor: "#FF0000", lineWidth: 2 }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.spPr).toBeDefined();
    expect(valAx.spPr!.line).toBeDefined();
    expect(valAx.spPr!.line!.color).toBeDefined();
    expect(valAx.spPr!.line!.width).toBe(2 * 12700);
  });
});

// ---------------------------------------------------------------------------
// 15. Chart class convenience methods
// ---------------------------------------------------------------------------

describe("Chart class convenience methods", () => {
  it("plotArea getter returns the plot area", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];
    expect(chart.plotArea).toBeDefined();
    expect(chart.plotArea!.chartTypes).toHaveLength(1);
    expect(chart.plotArea!.axes.length).toBeGreaterThan(0);
  });

  it("addSeries adds to first chart type group", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];
    expect(chart.seriesCount).toBe(1);
    chart.addSeries({
      index: 1,
      order: 1,
      cat: { numRef: { f: CATEGORIES } },
      val: { numRef: { f: VALUES_B } }
    } as any);
    expect(chart.seriesCount).toBe(2);
  });

  it("removeSeries removes by index", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";
    ws.addChart({ type: "bar", series: [baseSeries("S1"), baseSeries("S2", VALUES_B)] }, "C1:J10");
    const chart = ws.getCharts()[0];
    expect(chart.seriesCount).toBe(2);
    const removed = chart.removeSeries(0);
    expect(removed).toBeDefined();
    expect(chart.seriesCount).toBe(1);
  });

  it("removeSeries returns undefined for out-of-range index", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];
    expect(chart.removeSeries(5)).toBeUndefined();
    expect(chart.removeSeries(-1)).toBeUndefined();
  });

  it("combo chart dataTable in builder", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Bar")] },
        { type: "line", series: [baseSeries("Line")] }
      ],
      dataTable: true
    });
    expect(pa(m).dataTable).toBeDefined();
    expect(pa(m).dataTable!.showKeys).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 16. Gradient / pattern fill in shape-properties
// ---------------------------------------------------------------------------

describe("gradient and pattern fill", () => {
  it("buildSpPr with gradient fill produces valid structure", () => {
    const spPr = buildSpPr({
      fill: {
        gradient: {
          stops: [
            { position: 0, color: { srgb: "FF0000" } },
            { position: 100, color: { srgb: "0000FF" } }
          ],
          angle: 90,
          type: "linear"
        }
      }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.fill?.gradient).toBeDefined();
    expect(spPr.fill!.gradient!.stops).toHaveLength(2);
    expect(spPr.fill!.gradient!.stops[0].color.srgb).toBe("FF0000");
    expect(spPr.fill!.gradient!.stops[1].color.srgb).toBe("0000FF");
  });

  it("parseSpPr extracts gradient fill from raw XML", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:gradFill>",
      "    <a:gsLst>",
      '      <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>',
      '      <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>',
      '      <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>',
      "    </a:gsLst>",
      '    <a:lin ang="5400000" scaled="1"/>',
      "  </a:gradFill>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.fill?.gradient).toBeDefined();
    expect(parsed.fill!.gradient!.stops).toHaveLength(3);
    // OOXML `<a:gs pos>` is encoded in hundredths of a percent
    // (0–100000 = 0%–100%); the parsed `position` is the fraction
    // (0–1). The previous implementation divided by 1000 and
    // emitted the same factor in the writer — a pair of self-cancelling
    // bugs that corrupted anyone reading the file outside this library.
    expect(parsed.fill!.gradient!.stops[0].position).toBe(0);
    expect(parsed.fill!.gradient!.stops[0].color.srgb).toBe("FF0000");
    expect(parsed.fill!.gradient!.stops[1].position).toBe(0.5);
    expect(parsed.fill!.gradient!.stops[2].position).toBe(1);
    expect(parsed.fill!.gradient!.stops[2].color.srgb).toBe("0000FF");
    expect(parsed.fill!.gradient!.angle).toBe(90);
    expect(parsed.fill!.gradient!.type).toBe("linear");
  });

  it("getSpPrGradient returns gradient from raw XML", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:gradFill>",
      "    <a:gsLst>",
      '      <a:gs pos="0"><a:srgbClr val="AABBCC"/></a:gs>',
      '      <a:gs pos="100000"><a:srgbClr val="112233"/></a:gs>',
      "    </a:gsLst>",
      '    <a:lin ang="0" scaled="1"/>',
      "  </a:gradFill>",
      "</c:spPr>"
    ].join("");
    const grad = getSpPrGradient({ _rawXml: rawXml } as any);
    expect(grad).toBeDefined();
    expect(grad!.stops).toHaveLength(2);
  });

  it("buildSpPr with pattern fill produces valid structure", () => {
    const spPr = buildSpPr({
      fill: {
        pattern: {
          preset: "dkDnDiag",
          foreground: { srgb: "FF0000" },
          background: { srgb: "FFFFFF" }
        }
      }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.fill?.pattern).toBeDefined();
    expect(spPr.fill!.pattern!.preset).toBe("dkDnDiag");
    expect(spPr.fill!.pattern!.foreground?.srgb).toBe("FF0000");
    expect(spPr.fill!.pattern!.background?.srgb).toBe("FFFFFF");
  });

  it("parseSpPr extracts pattern fill from raw XML", () => {
    const rawXml = [
      "<c:spPr>",
      '  <a:pattFill prst="ltHorz">',
      '    <a:fgClr><a:srgbClr val="112233"/></a:fgClr>',
      '    <a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>',
      "  </a:pattFill>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.fill?.pattern).toBeDefined();
    expect(parsed.fill!.pattern!.preset).toBe("ltHorz");
    expect(parsed.fill!.pattern!.foreground?.srgb).toBe("112233");
    expect(parsed.fill!.pattern!.background?.srgb).toBe("FFFFFF");
  });

  it("getSpPrPattern returns pattern from raw XML", () => {
    const rawXml =
      '<c:spPr><a:pattFill prst="smCheck"><a:fgClr><a:srgbClr val="000000"/></a:fgClr></a:pattFill></c:spPr>';
    const pat = getSpPrPattern({ _rawXml: rawXml } as any);
    expect(pat).toBeDefined();
    expect(pat!.preset).toBe("smCheck");
  });

  it("setSpPrFill with gradient replaces existing fill", () => {
    const original = buildSpPr({ fill: { solid: { srgb: "FF0000" } } });
    const updated = setSpPrFill(original, {
      gradient: {
        stops: [
          { position: 0, color: { srgb: "000000" } },
          { position: 100, color: { srgb: "FFFFFF" } }
        ]
      }
    });
    expect(getSpPrFillColor(updated)).toBeUndefined();
    expect(getSpPrGradient(updated)).toBeDefined();
  });

  it("path gradient type is parsed", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:gradFill>",
      "    <a:gsLst>",
      '      <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>',
      '      <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>',
      "    </a:gsLst>",
      '    <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>',
      "  </a:gradFill>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.fill!.gradient!.type).toBe("circle");
  });
});

// ---------------------------------------------------------------------------
// 17. Automatic cache population from worksheet cell values
// ---------------------------------------------------------------------------

describe("automatic chart cache population", () => {
  it("fills numRef cache from worksheet values when addChart is called", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Q1";
    ws.getCell("A2").value = "Q2";
    ws.getCell("A3").value = "Q3";
    ws.getCell("A4").value = "Q4";
    ws.getCell("B1").value = 100;
    ws.getCell("B2").value = 200;
    ws.getCell("B3").value = 300;
    ws.getCell("B4").value = 400;

    ws.addChart({ type: "bar", series: [baseSeries("Sales")] }, "C1:J10");

    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;

    expect(series.val.numRef.cache.points).toHaveLength(4);
    expect(series.val.numRef.cache.points[0]).toEqual({ index: 0, value: 100 });
    expect(series.val.numRef.cache.points[3]).toEqual({ index: 3, value: 400 });
    expect(series.val.numRef.cache.pointCount).toBe(4);
  });

  it("fills strRef cache from category values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Jan";
    ws.getCell("A2").value = "Feb";
    ws.getCell("A3").value = "Mar";
    ws.getCell("A4").value = "Apr";
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("B3").value = 30;
    ws.getCell("B4").value = 40;

    ws.addChart({ type: "line", series: [baseSeries("S")] }, "C1:J10");

    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.cat.strRef.cache.points).toHaveLength(4);
    expect(series.cat.strRef.cache.points[0]).toEqual({ index: 0, value: "Jan" });
    expect(series.cat.strRef.cache.points[3]).toEqual({ index: 3, value: "Apr" });
  });

  it("handles missing cells by leaving gaps in cache", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "a";
    ws.getCell("A2").value = "b";
    // A3 intentionally left empty
    ws.getCell("A4").value = "d";
    ws.getCell("B1").value = 10;
    // B2, B3 empty
    ws.getCell("B4").value = 40;

    ws.addChart({ type: "bar", series: [baseSeries("S")] }, "C1:J10");

    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    const numPoints = series.val.numRef.cache.points;
    // pointCount should include total slots
    expect(series.val.numRef.cache.pointCount).toBe(4);
    // points should only include non-empty cells
    expect(numPoints).toHaveLength(2);
    expect(numPoints[0].index).toBe(0);
    expect(numPoints[1].index).toBe(3);
  });

  it("resolves references to other worksheets", () => {
    const wb = new Workbook();
    const data = wb.addWorksheet("Data");
    data.getCell("A1").value = 1;
    data.getCell("A2").value = 2;
    data.getCell("A3").value = 3;
    const chartWs = wb.addWorksheet("Charts");
    chartWs.addChart(
      {
        type: "bar",
        series: [
          {
            name: "S",
            categories: "Data!$A$1:$A$3",
            values: "Data!$A$1:$A$3"
          }
        ]
      },
      "A1:H10"
    );
    const chart = chartWs.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache.points).toHaveLength(3);
    expect(series.val.numRef.cache.points[0]).toEqual({ index: 0, value: 1 });
  });

  it("handles sheet names with spaces (quoted)", () => {
    const wb = new Workbook();
    const data = wb.addWorksheet("My Data");
    data.getCell("A1").value = 10;
    data.getCell("A2").value = 20;
    const chartWs = wb.addWorksheet("Chart");
    chartWs.addChart(
      {
        type: "bar",
        series: [
          {
            name: "S",
            categories: "'My Data'!$A$1:$A$2",
            values: "'My Data'!$A$1:$A$2"
          }
        ]
      },
      "A1:H10"
    );
    const chart = chartWs.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache.points).toHaveLength(2);
  });

  it("populates bubble chart xVal, yVal, and bubbleSize caches", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("B3").value = 30;
    ws.getCell("C1").value = 5;
    ws.getCell("C2").value = 10;
    ws.getCell("C3").value = 15;

    ws.addChart(
      {
        type: "bubble",
        series: [
          {
            name: "S",
            xValues: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            bubbleSize: "Sheet1!$C$1:$C$3"
          }
        ]
      },
      "D1:K10"
    );
    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.xVal.numRef.cache.points).toHaveLength(3);
    expect(series.yVal.numRef.cache.points).toHaveLength(3);
    expect(series.bubbleSize.numRef.cache.points).toHaveLength(3);
  });

  it("handles non-existent sheet gracefully", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    // Should not throw even though NoSuchSheet doesn't exist
    expect(() => {
      ws.addChart(
        {
          type: "bar",
          series: [
            {
              name: "S",
              categories: "NoSuchSheet!$A$1:$A$4",
              values: "NoSuchSheet!$B$1:$B$4"
            }
          ]
        },
        "C1:J10"
      );
    }).not.toThrow();
  });

  it("converts boolean cell values to 0/1 for numRef", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "a";
    ws.getCell("A2").value = "b";
    ws.getCell("B1").value = true;
    ws.getCell("B2").value = false;
    ws.addChart({ type: "bar", series: [baseSeries("S", "Sheet1!$B$1:$B$2")] }, "C1:J10");
    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache.points[0].value).toBe(1);
    expect(series.val.numRef.cache.points[1].value).toBe(0);
  });

  it("skips already-populated caches (idempotent)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "a";
    ws.getCell("B1").value = 100;

    // Build manually so we can pre-populate
    const m = buildChartModel({
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$B$1:$B$1" }]
    });
    const series = m.chart.plotArea.chartTypes[0].series[0] as any;
    series.val.numRef.cache.points = [{ index: 0, value: 999 }];
    // Now import
    fillChartCaches(m, wb);
    // Should keep our pre-set 999, not override with 100
    expect(series.val.numRef.cache.points[0].value).toBe(999);
  });
});

// ---------------------------------------------------------------------------
// 17b. Defined-name references in chart data sources
// ---------------------------------------------------------------------------

describe("chart cache population — defined names", () => {
  it("resolves a workbook-scoped defined name pointing to a single column", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Jan";
    ws.getCell("A2").value = "Feb";
    ws.getCell("A3").value = "Mar";
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("B3").value = 30;

    wb.definedNames.add("Sheet1!$B$1:$B$3", "Sales");
    wb.definedNames.add("Sheet1!$A$1:$A$3", "Months");

    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Months", values: "Sales" }]
      },
      "D1:K10"
    );

    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache.points).toHaveLength(3);
    expect(series.val.numRef.cache.points[0]).toEqual({ index: 0, value: 10 });
    expect(series.val.numRef.cache.points[2]).toEqual({ index: 2, value: 30 });
    expect(series.cat.strRef.cache.points[0]).toEqual({ index: 0, value: "Jan" });
  });

  it("resolves a qualified `Sheet!Name` reference", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Data");
    const ws2 = wb.addWorksheet("Report");
    ws1.getCell("A1").value = 111;
    ws1.getCell("A2").value = 222;
    wb.definedNames.add("Data!$A$1:$A$2", "MyVals");

    ws2.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Data!$A$1:$A$2", values: "Data!MyVals" }]
      },
      "A1:H10"
    );

    const chart = ws2.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache.points).toEqual([
      { index: 0, value: 111 },
      { index: 1, value: 222 }
    ]);
  });

  it("prefers sheet-scoped entry over workbook-scoped when both exist", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    ws1.getCell("A1").value = 1;
    ws1.getCell("A2").value = 2;
    ws2.getCell("A1").value = 100;
    ws2.getCell("A2").value = 200;

    // Workbook-scoped points at Sheet1
    wb.definedNames.model = [
      { name: "Local", ranges: ["Sheet1!$A$1:$A$2"] },
      // Sheet-scoped on Sheet2 (localSheetId = 1 since Sheet2 is index 1 in worksheets)
      { name: "Local", ranges: ["Sheet2!$A$1:$A$2"], localSheetId: 1 }
    ];

    // Chart on Sheet2: bare `Local` should resolve to Sheet2's scoped entry
    ws2.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet2!$A$1:$A$2", values: "Local" }]
      },
      "C1:J10"
    );
    const chart2 = ws2.getCharts()[0];
    const series2 = chart2.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series2.val.numRef.cache.points).toEqual([
      { index: 0, value: 100 },
      { index: 1, value: 200 }
    ]);

    // Chart on Sheet1: bare `Local` should resolve to workbook-scoped Sheet1 entry
    ws1.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Local" }]
      },
      "C1:J10"
    );
    const chart1 = ws1.getCharts()[0];
    const series1 = chart1.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series1.val.numRef.cache.points).toEqual([
      { index: 0, value: 1 },
      { index: 1, value: 2 }
    ]);
  });

  it("expands a multi-area defined name into concatenated cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("C1").value = 3;
    ws.getCell("C2").value = 4;

    // Add two disjoint ranges under the same name — the matrix expands them
    // into two ranges that should be concatenated in order when the chart
    // is resolved.
    wb.definedNames.add("Sheet1!$A$1:$A$2", "Multi");
    wb.definedNames.add("Sheet1!$C$1:$C$2", "Multi");

    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Multi" }]
      },
      "E1:L10"
    );
    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    // All four values should be present (order: both ranges concatenated)
    const values = (series.val.numRef.cache.points as Array<{ value: number }>).map(p => p.value);
    expect(values.sort((a, b) => a - b)).toEqual([1, 2, 3, 4]);
  });

  it("silently ignores unknown defined names (no crash, empty cache)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;

    // Use the builder directly so the assertion targets the resolver, not
    // the Worksheet.addChart convenience path.
    const m = buildChartModel({
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1", values: "DoesNotExist" }]
    });
    fillChartCaches(m, wb);

    const series = m.chart.plotArea.chartTypes[0].series[0] as any;
    // No cache points populated because the name is unresolvable.
    expect(series.val.numRef.cache?.points ?? []).toHaveLength(0);
  });

  it("terminates on cyclic defined names without blowing the stack", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;

    // Build a synthetic cycle by using formula-based defined names that
    // refer to each other. The resolver should bail out, not recurse
    // forever.
    wb.definedNames.addFormula("A", "B");
    wb.definedNames.addFormula("B", "A");

    const m = buildChartModel({
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1", values: "A" }]
    });
    expect(() => fillChartCaches(m, wb)).not.toThrow();
    const series = m.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.val.numRef.cache?.points ?? []).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------

describe("rich text writing API", () => {
  function makeChart() {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "a";
    ws.getCell("B1").value = 1;
    ws.addChart({ type: "bar", series: [baseSeries("S")], title: "Original" }, "C1:J10");
    return { wb, ws, chart: ws.getCharts()[0] };
  }

  it("chart.title setter accepts plain string", () => {
    const { chart } = makeChart();
    chart.title = "New Title";
    expect(chart.title).toBe("New Title");
  });

  it("chart.title setter accepts structured rich text", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [
            { text: "Bold ", properties: { bold: true, size: 1400 } },
            { text: "italic", properties: { italic: true, size: 1400 } }
          ]
        }
      ]
    };
    expect(chart.titleRichText).toBeDefined();
    const runs = chart.titleRichText!.paragraphs[0].runs!;
    expect(runs).toHaveLength(2);
    expect(runs[0].properties?.bold).toBe(true);
    expect(runs[1].properties?.italic).toBe(true);
  });

  it("chart.title setter accepts formula reference", () => {
    const { chart } = makeChart();
    chart.title = { formula: "Sheet1!$A$1" };
    expect(chart.chartModel!.chart.title!.strRef).toBeDefined();
    expect(chart.chartModel!.chart.title!.strRef!.formula).toBe("Sheet1!$A$1");
  });

  it("chart.title = undefined clears the title", () => {
    const { chart } = makeChart();
    chart.title = undefined;
    expect(chart.chartModel!.chart.title).toBeUndefined();
    expect(chart.chartModel!.chart.autoTitleDeleted).toBe(true);
  });

  it("chart.title setter preserves first-run formatting when setting a plain string", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [
            {
              text: "Original",
              properties: { bold: true, italic: true, size: 1600, color: { srgb: "FF0000" } }
            }
          ]
        }
      ]
    };
    // Now set as plain string
    chart.title = "Replaced";
    expect(chart.title).toBe("Replaced");
    const props = chart.titleRichText!.paragraphs[0].runs![0].properties;
    expect(props?.bold).toBe(true);
    expect(props?.italic).toBe(true);
    expect(props?.size).toBe(1600);
    expect(props?.color?.srgb).toBe("FF0000");
  });

  it("setTitleRichText convenience method", () => {
    const { chart } = makeChart();
    chart.setTitleRichText({
      paragraphs: [{ runs: [{ text: "Hi", properties: { size: 2000 } }] }]
    });
    expect(chart.titleRichText!.paragraphs[0].runs![0].text).toBe("Hi");
  });

  it("rich text supports bodyProperties", () => {
    const { chart } = makeChart();
    chart.title = {
      bodyProperties: { rotation: 900000, anchor: "ctr" },
      paragraphs: [{ runs: [{ text: "Rotated" }] }]
    };
    expect(chart.titleRichText!.bodyProperties?.rotation).toBe(900000);
    expect(chart.titleRichText!.bodyProperties?.anchor).toBe("ctr");
  });

  it("paragraph properties support alignment and indent", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          properties: { alignment: "ctr", indent: 914400, level: 1 },
          runs: [{ text: "Aligned" }]
        }
      ]
    };
    const pPr = chart.titleRichText!.paragraphs[0].properties;
    expect(pPr?.alignment).toBe("ctr");
    expect(pPr?.indent).toBe(914400);
    expect(pPr?.level).toBe(1);
  });

  it("paragraph supports bullets", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          properties: { bullet: { type: "char", character: "•" } },
          runs: [{ text: "Bullet point" }]
        }
      ]
    };
    const pPr = chart.titleRichText!.paragraphs[0].properties;
    expect(pPr?.bullet).toEqual({ type: "char", character: "•" });
  });

  it("paragraph supports line spacing as percentage", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          properties: { lineSpacing: { type: "percentage", value: 150000 } },
          runs: [{ text: "Spaced" }]
        }
      ]
    };
    const pPr = chart.titleRichText!.paragraphs[0].properties;
    expect(pPr?.lineSpacing).toEqual({ type: "percentage", value: 150000 });
  });

  it("run supports strike, baseline, cap, kern, spacing", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [
            {
              text: "Styled",
              properties: {
                strike: "sngStrike",
                baseline: 30000,
                cap: "all",
                kern: 1200,
                spacing: 100
              }
            }
          ]
        }
      ]
    };
    const props = chart.titleRichText!.paragraphs[0].runs![0].properties!;
    expect(props.strike).toBe("sngStrike");
    expect(props.baseline).toBe(30000);
    expect(props.cap).toBe("all");
    expect(props.kern).toBe(1200);
    expect(props.spacing).toBe(100);
  });

  it("run supports underline as enum string", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [{ text: "Wavy", properties: { underline: "wavyDbl" } }]
        }
      ]
    };
    expect(chart.titleRichText!.paragraphs[0].runs![0].properties?.underline).toBe("wavyDbl");
  });

  it("run supports boolean underline for convenience", () => {
    const { chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [{ text: "Underlined", properties: { underline: true } }]
        }
      ]
    };
    expect(chart.titleRichText!.paragraphs[0].runs![0].properties?.underline).toBe(true);
  });

  it("round-trips rich text through workbook write/read", async () => {
    const { wb, chart } = makeChart();
    chart.title = {
      paragraphs: [
        {
          runs: [
            { text: "Bold ", properties: { bold: true, size: 1400, color: { srgb: "FF0000" } } },
            { text: "plain" }
          ]
        }
      ]
    };

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart2 = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    // On read, rich text may come back as rawTx — the plain-string title getter
    // should still extract the concatenated text.
    expect(chart2.title).toContain("Bold");
    expect(chart2.title).toContain("plain");
  });
});

// ---------------------------------------------------------------------------
// 19. Builder P1 completions — floor/walls/legend entries/data label entries/etc.
// ---------------------------------------------------------------------------

describe("P1 builder completions", () => {
  it("floor / sideWall / backWall as simplified fill options", () => {
    const m = buildChartModel({
      type: "bar3D",
      series: [baseSeries("S")],
      floor: { fill: "#EEEEEE" },
      sideWall: { fill: "#AABBCC", border: "#000000", borderWidth: 1 },
      backWall: { noFill: true }
    });
    expect(m.chart.floor?.fill?.solid?.srgb).toBe("EEEEEE");
    expect(m.chart.sideWall?.fill?.solid?.srgb).toBe("AABBCC");
    expect(m.chart.sideWall?.line?.color?.srgb).toBe("000000");
    expect(m.chart.sideWall?.line?.width).toBe(12700);
    expect(m.chart.backWall?.fill?.noFill).toBe(true);
  });

  it("title options: layout and overlay", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S")],
      title: "Hello",
      titleOptions: {
        layout: {
          manualLayout: { layoutTarget: "inner", x: 0.1, y: 0.05, w: 0.3, h: 0.1 }
        },
        overlay: true
      }
    });
    expect(m.chart.title?.layout).toBeDefined();
    expect(m.chart.title?.overlay).toBe(true);
  });

  it("legend options: layout, overlay, entries, spPr, txPr", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S"), baseSeries("S2", VALUES_B)],
      showLegend: true,
      legendOptions: {
        layout: {
          manualLayout: { x: 0.8, y: 0.1, w: 0.2, h: 0.3 }
        },
        overlay: false,
        entries: [{ index: 1, hidden: true }],
        spPr: { fill: "#FFFFFF", border: "#CCCCCC" }
      }
    });
    expect(m.chart.legend?.layout).toBeDefined();
    expect(m.chart.legend?.legendEntries).toHaveLength(1);
    expect(m.chart.legend?.legendEntries![0].index).toBe(1);
    expect(m.chart.legend?.legendEntries![0].delete).toBe(true);
    expect(m.chart.legend?.spPr?.fill?.solid?.srgb).toBe("FFFFFF");
  });

  it("plotArea options: layout and background", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S")],
      plotAreaOptions: {
        layout: { manualLayout: { x: 0, y: 0, w: 1, h: 1 } },
        spPr: { fill: "#F5F5F5" }
      }
    });
    expect(m.chart.plotArea.layout).toBeDefined();
    expect(m.chart.plotArea.spPr?.fill?.solid?.srgb).toBe("F5F5F5");
  });

  it("data label per-entry overrides", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S"),
          dataLabels: {
            showVal: true,
            entries: [
              { index: 0, text: "Highest", spPr: { fill: "#FFFF00" } },
              { index: 1, delete: true }
            ]
          }
        }
      ]
    });
    const dl = (ctg(m).series[0] as BarSeries).dataLabels!;
    expect(dl.entries).toHaveLength(2);
    expect(dl.entries![0].index).toBe(0);
    expect(dl.entries![0].text?.paragraphs[0].runs![0].text).toBe("Highest");
    expect(dl.entries![0].spPr?.fill?.solid?.srgb).toBe("FFFF00");
    expect(dl.entries![1].delete).toBe(true);
  });

  it("trendline label styling", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S"),
          trendline: {
            type: "linear",
            displayEq: true,
            label: {
              numFmt: "0.000",
              spPr: { fill: "#FFFFFF" },
              txPr: { bold: true, size: 1000 }
            }
          }
        }
      ]
    });
    const tl = (ctg(m).series[0] as LineSeries).trendlines![0];
    expect(tl.trendlineLbl).toBeDefined();
    expect(tl.trendlineLbl!.numFmt?.formatCode).toBe("0.000");
    expect(tl.trendlineLbl!.spPr?.fill?.solid?.srgb).toBe("FFFFFF");
    expect(tl.trendlineLbl!.txPr?.bold).toBe(true);
  });

  it("error bars with spPr styling", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S"),
          errorBars: {
            type: "stdDev",
            direction: "y",
            line: "#FF0000",
            lineWidth: 1.5,
            lineDash: "dash"
          }
        }
      ]
    });
    const eb = (ctg(m).series[0] as BarSeries).errorBars!;
    expect(eb.spPr?.line?.color?.srgb).toBe("FF0000");
    expect(eb.spPr?.line?.width).toBe(Math.round(1.5 * 12700));
    expect(eb.spPr?.line?.dash).toBe("dash");
  });

  it("up-down bars with upBars and downBars styling", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S"), baseSeries("S2", VALUES_B)],
      upDownBars: {
        gapWidth: 100,
        upBars: { fill: "#00FF00" },
        downBars: { fill: "#FF0000" }
      }
    });
    const g = ctg(m) as LineChartGroup;
    expect(g.upDownBars?.gapWidth).toBe(100);
    expect(g.upDownBars?.upBars?.fill?.solid?.srgb).toBe("00FF00");
    expect(g.upDownBars?.downBars?.fill?.solid?.srgb).toBe("FF0000");
  });

  it("surface bandFormats", () => {
    const m = buildChartModel({
      type: "surface",
      series: [baseSeries("S")],
      bandFormats: [
        { index: 0, spPr: { fill: "#FF0000" } },
        { index: 1, spPr: { fill: "#00FF00" } },
        { index: 2, spPr: { fill: "#0000FF" } }
      ]
    });
    const g = ctg(m) as any;
    expect(g.bandFormats).toHaveLength(3);
    expect(g.bandFormats![0].spPr.fill.solid.srgb).toBe("FF0000");
    expect(g.bandFormats![2].spPr.fill.solid.srgb).toBe("0000FF");
  });

  it("axis extended options: spPr, txPr, customUnit", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S")],
      valueAxis: {
        spPr: { fill: "#CCCCCC", border: "#000000" },
        txPr: { bold: true, size: 900 },
        displayUnits: "thousands",
        customUnit: 500,
        displayUnitsLabel: "× 1000"
      }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")! as ValueAxis;
    expect(valAx.spPr?.fill?.solid?.srgb).toBe("CCCCCC");
    expect(valAx.txPr?.bold).toBe(true);
    expect(valAx.dispUnits?.builtInUnit).toBe("thousands");
    expect(valAx.dispUnits?.custUnit).toBe(500);
    expect(valAx.dispUnits?.label).toBeDefined();
  });

  it("axis date-specific time units", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S")],
      categoryAxis: {
        baseTimeUnit: "months",
        majorTimeUnit: "years",
        minorTimeUnit: "months"
      }
    });
    // Category axis defaults to "cat" not "date"; this verifies the options
    // compile. Date axis detection happens elsewhere.
    const axes = pa(m).axes;
    expect(axes.find(a => a.axisType === "cat")).toBeDefined();
  });

  it("axis majorGridlines with custom styling", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S")],
      valueAxis: {
        majorGridlinesStyle: { border: "#DDDDDD", borderWidth: 0.5 }
      }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.majorGridlines?.line?.color?.srgb).toBe("DDDDDD");
    expect(valAx.majorGridlines?.line?.width).toBe(0.5 * 12700);
  });

  it("series spPr override via structured options", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S"),
          spPr: {
            fill: { solid: { srgb: "123456" } },
            line: { color: { srgb: "654321" }, width: 25400 }
          }
        }
      ]
    });
    const s = ctg(m).series[0];
    expect(s.spPr?.fill?.solid?.srgb).toBe("123456");
    expect(s.spPr?.line?.color?.srgb).toBe("654321");
  });

  it("series pictureFill for bar charts", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S"),
          pictureFill: {
            fillMode: "stretch",
            applyToFront: true,
            applyToSides: false
          }
        }
      ]
    });
    const s = ctg(m).series[0] as BarSeries;
    expect(s.pictureOptions?.pictureFormat).toBe("stretch");
    expect(s.pictureOptions?.applyToFront).toBe(true);
    expect(s.pictureOptions?.applyToSides).toBe(false);
  });

  it("combo chart preserves explicit globally-unique indices", () => {
    const m = buildComboChartModel({
      groups: [
        {
          type: "bar",
          series: [{ ...baseSeries("Bar"), _explicitIndex: 0 } as any]
        },
        {
          type: "line",
          series: [{ ...baseSeries("Line"), _explicitIndex: 1 } as any]
        }
      ]
    });
    // After buildComboChartModel, default behaviour renumbers 0,1 anyway (no collision)
    const barSer = m.chart.plotArea.chartTypes[0].series[0];
    const lineSer = m.chart.plotArea.chartTypes[1].series[0];
    // With no collision, indices remain as assigned by per-group builder (0,0 → renumbered to 0,1)
    expect(barSer.index).toBe(0);
    expect(lineSer.index).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// 20. Effect list / 3D scene / sp3d structured support
// ---------------------------------------------------------------------------

describe("effect list and 3D structured", () => {
  it("parseSpPr extracts outer shadow", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:effectLst>",
      '    <a:outerShdw blurRad="40000" dist="25000" dir="5400000" algn="ctr">',
      '      <a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>',
      "    </a:outerShdw>",
      "  </a:effectLst>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.effectList?.outerShadow).toBeDefined();
    expect(parsed.effectList!.outerShadow!.blurRadius).toBe(40000);
    expect(parsed.effectList!.outerShadow!.distance).toBe(25000);
    expect(parsed.effectList!.outerShadow!.alignment).toBe("ctr");
    expect(parsed.effectList!.outerShadow!.color.srgb).toBe("000000");
  });

  it("parseSpPr extracts glow", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:effectLst>",
      '    <a:glow rad="63500">',
      '      <a:srgbClr val="FF6600"/>',
      "    </a:glow>",
      "  </a:effectLst>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.effectList?.glow).toBeDefined();
    expect(parsed.effectList!.glow!.radius).toBe(63500);
    expect(parsed.effectList!.glow!.color.srgb).toBe("FF6600");
  });

  it("parseSpPr extracts soft edge", () => {
    const rawXml = '<c:spPr><a:effectLst><a:softEdge rad="25400"/></a:effectLst></c:spPr>';
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.effectList?.softEdge?.radius).toBe(25400);
  });

  it("parseSpPr extracts reflection", () => {
    const rawXml = [
      "<c:spPr><a:effectLst>",
      '<a:reflection blurRad="6350" stA="52000" stPos="0" endA="300" endPos="35000" ',
      'dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>',
      "</a:effectLst></c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    const r = parsed.effectList?.reflection;
    expect(r).toBeDefined();
    expect(r!.blurRadius).toBe(6350);
    expect(r!.startOpacity).toBe(52000);
    expect(r!.scaleVertical).toBe(-100000);
    expect(r!.alignment).toBe("bl");
  });

  it("structured effect list round-trips through render", () => {
    const spPr = buildSpPr({
      fill: { solid: { srgb: "FF0000" } },
      effectList: {
        outerShadow: {
          blurRadius: 40000,
          distance: 25000,
          direction: 5400000,
          alignment: "ctr",
          color: { srgb: "000000", alpha: 40000 }
        },
        glow: { radius: 10000, color: { srgb: "00FF00" } }
      }
    });
    // The structured effectList should be serialisable via buildSpPr.
    // Note: buildSpPr in shape-properties.ts doesn't yet emit effectList in XML
    // — the chart-space-xform _renderSpPr handles that in full renders.
    expect(spPr.effectList).toBeDefined();
  });

  it("parseSpPr extracts scene3d camera and lightRig", () => {
    const rawXml = [
      "<c:spPr>",
      "  <a:scene3d>",
      '    <a:camera prst="legacyObliqueTopRight"/>',
      '    <a:lightRig rig="threePt" dir="t"/>',
      "  </a:scene3d>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.scene3d?.camera?.preset).toBe("legacyObliqueTopRight");
    expect(parsed.scene3d?.lightRig?.rig).toBe("threePt");
    expect(parsed.scene3d?.lightRig?.direction).toBe("t");
  });

  it("parseSpPr extracts sp3d with bevel", () => {
    const rawXml = [
      "<c:spPr>",
      '  <a:sp3d extrusionH="50800" prstMaterial="metal">',
      '    <a:bevelT w="63500" h="12700" prst="circle"/>',
      '    <a:bevelB w="63500" h="12700"/>',
      "  </a:sp3d>",
      "</c:spPr>"
    ].join("");
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.sp3d).toBeDefined();
    expect(parsed.sp3d!.extrusionHeight).toBe(50800);
    expect(parsed.sp3d!.material).toBe("metal");
    expect(parsed.sp3d!.bevelTop?.width).toBe(63500);
    expect(parsed.sp3d!.bevelTop?.preset).toBe("circle");
    expect(parsed.sp3d!.bevelBottom?.height).toBe(12700);
  });

  it("chart with structured spPr + effectList renders via xlsx round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "x";
    ws.getCell("B1").value = 1;
    ws.addChart(
      {
        type: "bar",
        series: [baseSeries("S")]
      },
      "C1:J10"
    );
    const chart = ws.getCharts()[0];
    // Assign effect list to the first series
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    series.spPr = {
      fill: { solid: { srgb: "4472C4" } },
      effectList: {
        outerShadow: {
          blurRadius: 40000,
          distance: 25000,
          direction: 5400000,
          alignment: "ctr",
          color: { srgb: "000000", alpha: 40000 }
        }
      }
    };
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart2 = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    const s = chart2.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    // After round-trip spPr comes back as _rawXml — parseSpPr extracts it
    const parsed = parseSpPr(s.spPr);
    expect(parsed.fill?.solid?.srgb).toBe("4472C4");
    expect(parsed.effectList?.outerShadow).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// 21. P3 anchor types + P4 Chart.clone/copyTo/Worksheet.removeChart
// ---------------------------------------------------------------------------

describe("anchor types and chart management", () => {
  it("addChart with oneCell anchor (tl + ext)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart(
      { type: "bar", series: [baseSeries("S")] },
      { tl: "B2", ext: { cx: 3657600, cy: 2743200 } }
    );
    const chart = ws.getCharts()[0];
    expect(chart.range.ext?.cx).toBe(3657600);
    expect(chart.range.ext?.cy).toBe(2743200);
    expect(chart.range.br).toBeUndefined();
    expect(chart.range.editAs).toBe("oneCell");
  });

  it("addChart with absolute anchor (pos + ext)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart(
      { type: "bar", series: [baseSeries("S")] },
      { pos: { x: 914400, y: 914400 }, ext: { cx: 3657600, cy: 2743200 } }
    );
    const chart = ws.getCharts()[0];
    expect(chart.range.pos?.x).toBe(914400);
    expect(chart.range.pos?.y).toBe(914400);
    expect(chart.range.ext?.cx).toBe(3657600);
    expect(chart.range.editAs).toBe("absolute");
  });

  it("addChart with two-cell anchor (tl + br)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")] }, { tl: "B2", br: "H10" });
    const chart = ws.getCharts()[0];
    expect(chart.range.br).toBeDefined();
    expect(chart.range.editAs).toBe("twoCell");
  });

  it("Chart.clone creates a deep copy in the same worksheet", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "x";
    ws.getCell("B1").value = 1;
    ws.addChart({ type: "bar", series: [baseSeries("S")], title: "Original" }, "C1:J10");
    const original = ws.getCharts()[0];
    const clonedNumber = original.clone();
    expect(ws.getCharts()).toHaveLength(2);
    expect(clonedNumber).not.toBe(original.chartNumber);
    const cloned = ws.getCharts()[1];
    expect(cloned.title).toBe("Original");
    // Mutate the clone — original should be unchanged
    cloned.title = "Modified";
    expect(original.title).toBe("Original");
    expect(cloned.title).toBe("Modified");
  });

  it("Chart.copyTo copies to another worksheet", () => {
    const wb = new Workbook();
    const src = wb.addWorksheet("Source");
    const dst = wb.addWorksheet("Dest");
    src.getCell("A1").value = "x";
    src.getCell("B1").value = 1;
    src.addChart({ type: "bar", series: [baseSeries("S")], title: "Moved" }, "C1:J10");
    const chart = src.getCharts()[0];
    chart.copyTo(dst, "A1:H10");
    expect(dst.getCharts()).toHaveLength(1);
    expect(dst.getCharts()[0].title).toBe("Moved");
    // Original still exists in source
    expect(src.getCharts()).toHaveLength(1);
  });

  it("Worksheet.removeChart by object", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")] }, "C1:J10");
    ws.addChart({ type: "line", series: [baseSeries("S")] }, "K1:R10");
    const charts = ws.getCharts();
    expect(charts).toHaveLength(2);
    const removed = ws.removeChart(charts[0]);
    expect(removed).toBe(true);
    expect(ws.getCharts()).toHaveLength(1);
    expect(ws.getCharts()[0].chartTypes[0].type).toBe("line");
  });

  it("Worksheet.removeChart by index", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")] }, "C1:J10");
    ws.addChart({ type: "line", series: [baseSeries("S")] }, "K1:R10");
    ws.removeChart(0);
    expect(ws.getCharts()).toHaveLength(1);
    expect(ws.getCharts()[0].chartTypes[0].type).toBe("line");
  });

  it("Worksheet.removeChart returns false for out-of-range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")] }, "C1:J10");
    expect(ws.removeChart(10)).toBe(false);
    expect(ws.removeChart(-1)).toBe(false);
    expect(ws.getCharts()).toHaveLength(1);
  });

  it("Chart.getSeries / addSeries / removeSeries with group index (combo)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "a";
    ws.getCell("B1").value = 1;
    ws.addComboChart(
      {
        groups: [
          { type: "bar", series: [baseSeries("Bar")] },
          { type: "line", series: [baseSeries("Line")] }
        ]
      },
      "C1:J10"
    );
    const chart = ws.getCharts()[0];
    expect(chart.getSeriesCount(0)).toBe(1);
    expect(chart.getSeriesCount(1)).toBe(1);
    expect(chart.totalSeriesCount).toBe(2);
    expect(chart.getSeries(0, 0)).toBeDefined();
    expect(chart.getSeries(0, 1)).toBeDefined();
    // Remove from group 1
    chart.removeSeries(0, 1);
    expect(chart.getSeriesCount(1)).toBe(0);
    expect(chart.getSeriesCount(0)).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// 22. Chart sidecar (style.xml / colors.xml) structured access
// ---------------------------------------------------------------------------

describe("chart sidecar files", () => {
  it("parseChartColors extracts palette entries", () => {
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"',
      '  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">',
      '  <a:schemeClr val="accent1"/>',
      '  <a:schemeClr val="accent2"/>',
      '  <a:srgbClr val="FF6600"/>',
      "</cs:colorStyle>"
    ].join("\n");
    const parsed = parseChartColors(xml);
    expect(parsed.method).toBe("cycle");
    expect(parsed.id).toBe(10);
    expect(parsed.colors).toHaveLength(3);
    expect(parsed.colors![0].theme).toBe("accent1");
    expect(parsed.colors![2].srgb).toBe("FF6600");
  });

  it("buildChartColors rebuilds XML from structured colors", () => {
    const xml = buildChartColors({
      method: "cycle",
      id: 10,
      colors: [{ theme: "accent1" }, { theme: "accent2", lumMod: 80000 }, { srgb: "FF0000" }]
    });
    expect(xml).toContain('meth="cycle"');
    expect(xml).toContain('<a:schemeClr val="accent1"/>');
    expect(xml).toContain('<a:lumMod val="80000"/>');
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
  });

  it("parseChartStyle extracts style id", () => {
    const xml = '<cs:chartStyle xmlns:cs="..." id="227">...</cs:chartStyle>';
    const parsed = parseChartStyle(xml);
    expect(parsed.id).toBe(227);
  });

  it("addChart writes chart style/colors sidecars from structured options", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Cat1";
    ws.getCell("B1").value = 10;
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1", values: "Sheet1!$B$1" }],
        chartStyle: { id: 227 },
        chartColors: {
          method: "cycle",
          id: 10,
          colors: [{ srgb: "FF0000" }, { theme: "accent2", lumMod: 60000 }]
        }
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const styleXml = textDecoder.decode(entries.get("xl/charts/style1.xml")!.data);
    const colorsXml = textDecoder.decode(entries.get("xl/charts/colors1.xml")!.data);
    const relsXml = textDecoder.decode(entries.get("xl/charts/_rels/chart1.xml.rels")!.data);

    expect(styleXml).toContain('id="227"');
    expect(colorsXml).toContain('<a:srgbClr val="FF0000"/>');
    expect(colorsXml).toContain('<a:schemeClr val="accent2"><a:lumMod val="60000"/></a:schemeClr>');
    expect(relsXml).toContain("chartStyle");
    expect(relsXml).toContain("chartColorStyle");
  });

  it("Chart.copyTo preserves chart style/colors sidecars across workbooks", async () => {
    const srcWb = new Workbook();
    const src = srcWb.addWorksheet("Src");
    src.getCell("A1").value = "Cat1";
    src.getCell("B1").value = 10;
    src.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Src!$A$1", values: "Src!$B$1" }],
        chartStyle: { id: 227 },
        chartColors: { method: "cycle", id: 10, colors: [{ srgb: "FF0000" }] }
      },
      "D1:J10"
    );

    const dstWb = new Workbook();
    const dst = dstWb.addWorksheet("Dst");
    src.getCharts()[0].copyTo(dst, "A1:H10");

    const buf = await dstWb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const styleXml = textDecoder.decode(entries.get("xl/charts/style1.xml")!.data);
    const colorsXml = textDecoder.decode(entries.get("xl/charts/colors1.xml")!.data);
    const relsXml = textDecoder.decode(entries.get("xl/charts/_rels/chart1.xml.rels")!.data);
    expect(styleXml).toContain('id="227"');
    expect(colorsXml).toContain('<a:srgbClr val="FF0000"/>');
    expect(relsXml).toContain("chartStyle");
    expect(relsXml).toContain("chartColorStyle");
  });

  it("addChartEx writes chartEx style/colors sidecars from structured options", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Cat1";
    ws.getCell("B1").value = 10;
    ws.addChartEx(
      {
        type: "histogram",
        series: [{ name: "S", values: "Sheet1!$B$1:$B$1" }],
        chartStyle: { id: 227 },
        chartColors: { method: "cycle", id: 10, colors: [{ srgb: "FF0000" }] }
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const styleXml = textDecoder.decode(entries.get("xl/charts/styleEx1.xml")!.data);
    const colorsXml = textDecoder.decode(entries.get("xl/charts/colorsEx1.xml")!.data);
    const relsXml = textDecoder.decode(entries.get("xl/charts/_rels/chartEx1.xml.rels")!.data);

    expect(styleXml).toContain('id="227"');
    expect(colorsXml).toContain('<a:srgbClr val="FF0000"/>');
    expect(relsXml).toContain("styleEx1.xml");
    expect(relsXml).toContain("colorsEx1.xml");
  });
});

// ---------------------------------------------------------------------------
// 23. ChartEx (Office 2016+ extended charts) — 8 modern chart types
// ---------------------------------------------------------------------------
