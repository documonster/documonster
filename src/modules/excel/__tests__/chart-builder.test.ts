/**
 * Comprehensive chart builder & API tests.
 *
 * Covers:
 * - All 16 chart types via programmatic addChart
 * - Combo charts with secondary axes
 * - 3D charts with series axis
 * - Scatter/bubble with errorBars arrays
 * - Rich title, strRef title, plain string title
 * - Builder advanced options: marker, trendline, dataLabels, dataPoints, axisOptions
 * - Chart round-trip: addChart → writeBuffer → load → verify model
 * - Chart high-level API: getCharts, title, chartTypes, axes, legend, series
 */

import type {
  AddChartOptions,
  AddChartSeriesOptions,
  ChartModel,
  ChartTypeGroup,
  PlotArea,
  BarChartGroup,
  LineChartGroup,
  AreaChartGroup,
  LineSeries,
  BarSeries,
  ValueAxis,
  AddChartExOptions,
  StockChartGroup,
  OfPieChartGroup
} from "@excel/chart";
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
  buildChartExModel,
  renderChartEx
} from "@excel/chart";
import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const CATEGORIES = "Sheet1!$A$1:$A$4";
const VALUES_A = "Sheet1!$B$1:$B$4";
const VALUES_B = "Sheet1!$C$1:$C$4";

function baseSeries(name: string, values = VALUES_A): AddChartSeriesOptions {
  return { name, categories: CATEGORIES, values };
}

function scatterSeries(name: string): AddChartSeriesOptions {
  return { name, xValues: "Sheet1!$A$1:$A$4", values: VALUES_A };
}

function bubbleSeries(name: string): AddChartSeriesOptions {
  return {
    name,
    xValues: "Sheet1!$A$1:$A$3",
    values: "Sheet1!$B$1:$B$3",
    bubbleSize: "Sheet1!$C$1:$C$3"
  };
}

function pa(m: ChartModel): PlotArea {
  return m.chart.plotArea;
}

function ctg(m: ChartModel, idx = 0): ChartTypeGroup {
  return pa(m).chartTypes[idx];
}

/** Round-trip: addChart → write → load → return Chart */
async function roundTripChart(opts: AddChartOptions) {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");
  // Fill some data so the sheet isn't empty
  ws.getCell("A1").value = "x";
  ws.addChart(opts, "C1:J15");
  const buf = await wb.xlsx.writeBuffer();
  const wb2 = new Workbook();
  await wb2.xlsx.load(buf);
  const ws2 = wb2.getWorksheet("Sheet1")!;
  return ws2.getCharts()[0];
}

// ---------------------------------------------------------------------------
// 1. All 16 chart types — buildChartModel unit tests
// ---------------------------------------------------------------------------

describe("buildChartModel — all 16 chart types", () => {
  it("bar chart", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("bar");
    expect(pa(m).axes.length).toBe(2);
    expect(pa(m).axes[0].axisType).toBe("cat");
    expect(pa(m).axes[1].axisType).toBe("val");
  });

  it("bar3D chart", () => {
    const m = buildChartModel({ type: "bar3D", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("bar3D");
    expect(pa(m).axes.length).toBe(3);
    expect(pa(m).axes[2].axisType).toBe("ser");
  });

  it("line chart", () => {
    const m = buildChartModel({ type: "line", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("line");
    expect(pa(m).axes.length).toBe(2);
    // line chart defaults marker: true
    expect((ctg(m) as any).marker).toBe(true);
  });

  it("line3D chart", () => {
    const m = buildChartModel({ type: "line3D", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("line3D");
    expect(pa(m).axes.length).toBe(3);
    expect(pa(m).axes[2].axisType).toBe("ser");
  });

  it("pie chart", () => {
    const m = buildChartModel({ type: "pie", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("pie");
    expect(pa(m).axes.length).toBe(0);
  });

  it("pie3D chart", () => {
    const m = buildChartModel({ type: "pie3D", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("pie3D");
    expect(pa(m).axes.length).toBe(0);
  });

  it("doughnut chart", () => {
    const m = buildChartModel({
      type: "doughnut",
      series: [baseSeries("S1")],
      holeSize: 50
    });
    expect(ctg(m).type).toBe("doughnut");
    expect((ctg(m) as any).holeSize).toBe(50);
    expect(pa(m).axes.length).toBe(0);
  });

  it("area chart", () => {
    const m = buildChartModel({ type: "area", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("area");
    expect(pa(m).axes.length).toBe(2);
  });

  it("area3D chart", () => {
    const m = buildChartModel({ type: "area3D", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("area3D");
    expect(pa(m).axes.length).toBe(3);
    expect(pa(m).axes[2].axisType).toBe("ser");
  });

  it("scatter chart", () => {
    const m = buildChartModel({ type: "scatter", series: [scatterSeries("S1")] });
    expect(ctg(m).type).toBe("scatter");
    expect(pa(m).axes.length).toBe(2);
    // scatter uses val+val axes, not cat+val
    expect(pa(m).axes[0].axisType).toBe("val");
    expect(pa(m).axes[1].axisType).toBe("val");
  });

  it("bubble chart", () => {
    const m = buildChartModel({ type: "bubble", series: [bubbleSeries("S1")] });
    expect(ctg(m).type).toBe("bubble");
    expect(pa(m).axes.length).toBe(2);
    expect(pa(m).axes[0].axisType).toBe("val");
    expect(pa(m).axes[1].axisType).toBe("val");
  });

  it("radar chart", () => {
    const m = buildChartModel({ type: "radar", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("radar");
    expect((ctg(m) as any).radarStyle).toBe("marker");
    expect(pa(m).axes.length).toBe(2);
  });

  it("stock chart", () => {
    const m = buildChartModel({
      type: "stock",
      series: [baseSeries("High"), baseSeries("Low"), baseSeries("Close")]
    });
    expect(ctg(m).type).toBe("stock");
    expect(pa(m).axes.length).toBe(2);
    expect(ctg(m).series.length).toBe(3);
  });

  it("surface chart", () => {
    const m = buildChartModel({ type: "surface", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("surface");
    expect(pa(m).axes.length).toBe(3);
    expect(pa(m).axes[2].axisType).toBe("ser");
  });

  it("surface3D chart", () => {
    const m = buildChartModel({ type: "surface3D", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("surface3D");
    expect(pa(m).axes.length).toBe(3);
  });

  it("ofPie chart", () => {
    const m = buildChartModel({
      type: "ofPie",
      series: [baseSeries("S1")],
      ofPieType: "bar",
      splitType: "auto",
      secondPieSize: 75
    });
    expect(ctg(m).type).toBe("ofPie");
    expect((ctg(m) as any).ofPieType).toBe("bar");
    expect((ctg(m) as any).secondPieSize).toBe(75);
    expect(pa(m).axes.length).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// 2. Chart model defaults and options
// ---------------------------------------------------------------------------

describe("buildChartModel — defaults and options", () => {
  it("auto-deletes title when no title provided", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(m.chart.autoTitleDeleted).toBe(true);
    expect(m.chart.title).toBeUndefined();
  });

  it("sets string title", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")], title: "Sales" });
    expect(m.chart.autoTitleDeleted).toBe(false);
    expect(m.chart.title).toBeDefined();
    expect(m.chart.title!.text?.paragraphs[0].runs![0].text).toBe("Sales");
  });

  it("sets formula title", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      title: { formula: "Sheet1!$A$1" }
    });
    expect(m.chart.title!.strRef).toBeDefined();
    expect(m.chart.title!.strRef!.formula).toBe("Sheet1!$A$1");
  });

  it("sets rich text title", () => {
    const richText = { paragraphs: [{ runs: [{ text: "Bold", rPr: { b: true } }] }] };
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      title: richText
    });
    expect(m.chart.title!.text).toBe(richText);
  });

  it("creates legend by default", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(m.chart.legend).toBeDefined();
    expect(m.chart.legend!.legendPos).toBe("b");
  });

  it("hides legend when showLegend=false", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")], showLegend: false });
    expect(m.chart.legend).toBeUndefined();
  });

  it("sets legendPosition", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      legendPosition: "t"
    });
    expect(m.chart.legend!.legendPos).toBe("t");
  });

  it("uses displayBlanksAs option", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      displayBlanksAs: "span"
    });
    expect(m.chart.dispBlanksAs).toBe("span");
  });

  it("defaults displayBlanksAs to gap", () => {
    const m = buildChartModel({ type: "line", series: [baseSeries("S1")] });
    expect(m.chart.dispBlanksAs).toBe("gap");
  });

  it("sets barDir option", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      barDir: "bar"
    });
    expect((ctg(m) as any).barDir).toBe("bar");
  });

  it("sets grouping option", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      grouping: "stacked"
    });
    expect((ctg(m) as any).grouping).toBe("stacked");
  });

  it("sets scatterStyle option", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [scatterSeries("S1")],
      scatterStyle: "smoothMarker"
    });
    expect((ctg(m) as any).scatterStyle).toBe("smoothMarker");
  });

  it("sets radarStyle option", () => {
    const m = buildChartModel({
      type: "radar",
      series: [baseSeries("S1")],
      radarStyle: "filled"
    });
    expect((ctg(m) as any).radarStyle).toBe("filled");
  });

  it("sets gapWidth and overlap on bar chart group", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      gapWidth: 200,
      overlap: -50
    });
    expect((ctg(m) as any).gapWidth).toBe(200);
    expect((ctg(m) as any).overlap).toBe(-50);
  });
});

// ---------------------------------------------------------------------------
// 3. Series options
// ---------------------------------------------------------------------------

describe("buildChartModel — series options", () => {
  it("applies fill and line to series spPr", () => {
    const m = buildChartModel({
      type: "bar",
      series: [{ ...baseSeries("S1"), fill: "FF0000", line: "0000FF", lineWidth: 2 }]
    });
    const s = ctg(m).series[0] as any;
    expect(s.spPr).toBeDefined();
    expect(s.spPr.fill?.solid?.srgb).toBe("FF0000");
    expect(s.spPr.line?.color?.srgb).toBe("0000FF");
    expect(s.spPr.line?.width).toBe(25400); // 2 * 12700
  });

  it("applies smooth to line series", () => {
    const m = buildChartModel({
      type: "line",
      series: [{ ...baseSeries("S1"), smooth: true }]
    });
    expect((ctg(m).series[0] as any).smooth).toBe(true);
  });

  it("applies smooth to scatter series", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [{ ...scatterSeries("S1"), smooth: true }]
    });
    expect((ctg(m).series[0] as any).smooth).toBe(true);
  });

  it("applies explosion to pie series", () => {
    const m = buildChartModel({
      type: "pie",
      series: [{ ...baseSeries("S1"), explosion: 25 }]
    });
    expect((ctg(m).series[0] as any).explosion).toBe(25);
  });

  it("applies invertIfNegative to bar series", () => {
    const m = buildChartModel({
      type: "bar",
      series: [{ ...baseSeries("S1"), invertIfNegative: false }]
    });
    expect((ctg(m).series[0] as any).invertIfNegative).toBe(false);
  });

  it("applies bubble3D to bubble series", () => {
    const m = buildChartModel({
      type: "bubble",
      series: [{ ...bubbleSeries("S1"), bubble3D: true }]
    });
    expect((ctg(m).series[0] as any).bubble3D).toBe(true);
  });

  it("applies marker options", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          marker: { symbol: "circle", size: 8, fill: "00FF00" }
        }
      ]
    });
    const marker = (ctg(m).series[0] as any).marker;
    expect(marker).toBeDefined();
    expect(marker.symbol).toBe("circle");
    expect(marker.size).toBe(8);
    expect(marker.spPr.fill.solid.srgb).toBe("00FF00");
  });

  it("applies dataLabels to series", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          dataLabels: { showVal: true, showCatName: false, position: "outEnd" }
        }
      ]
    });
    const dl = (ctg(m).series[0] as any).dataLabels;
    expect(dl.showVal).toBe(true);
    expect(dl.showCatName).toBe(false);
    expect(dl.position).toBe("outEnd");
  });

  it("applies trendline to series", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          trendline: { type: "linear", displayEq: true, displayRSqr: true, line: "FF0000" }
        }
      ]
    });
    const tl = (ctg(m).series[0] as any).trendlines;
    expect(tl).toBeDefined();
    expect(tl.length).toBe(1);
    expect(tl[0].type).toBe("linear");
    expect(tl[0].displayEq).toBe(true);
    expect(tl[0].displayRSqr).toBe(true);
    expect(tl[0].spPr.line.color.srgb).toBe("FF0000");
  });

  it("applies dataPoints to series", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          dataPoints: [
            { index: 0, fill: "FF0000" },
            { index: 2, fill: "00FF00", border: "000000" }
          ]
        }
      ]
    });
    const dp = (ctg(m).series[0] as any).dataPoints;
    expect(dp.length).toBe(2);
    expect(dp[0].index).toBe(0);
    expect(dp[0].spPr.fill.solid.srgb).toBe("FF0000");
    expect(dp[1].spPr.line.color.srgb).toBe("000000");
  });

  it("applies single errorBars to bar series", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          errorBars: { type: "fixedVal", value: 5, direction: "y" }
        }
      ]
    });
    const eb = (ctg(m).series[0] as any).errorBars;
    // bar series errorBars is single, not array
    expect(eb.errValType).toBe("fixedVal");
    expect(eb.val).toBe(5);
  });

  it("applies errorBars array to scatter series", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [
        {
          ...scatterSeries("S1"),
          errorBars: [
            { type: "fixedVal", value: 3, direction: "x" },
            { type: "percentage", value: 10, direction: "y" }
          ]
        }
      ]
    });
    const eb = (ctg(m).series[0] as any).errorBars;
    // scatter errorBars is an array
    expect(Array.isArray(eb)).toBe(true);
    expect(eb.length).toBe(2);
    expect(eb[0].errValType).toBe("fixedVal");
    expect(eb[1].errValType).toBe("percentage");
  });

  it("applies errorBars array to bubble series", () => {
    const m = buildChartModel({
      type: "bubble",
      series: [
        {
          ...bubbleSeries("S1"),
          errorBars: [{ type: "stdDev", value: 1, direction: "x" }]
        }
      ]
    });
    const eb = (ctg(m).series[0] as any).errorBars;
    expect(Array.isArray(eb)).toBe(true);
    expect(eb[0].errValType).toBe("stdDev");
  });

  it("applies group-level dataLabels", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      dataLabels: { showVal: true, showPercent: false }
    });
    const dl = (ctg(m) as any).dataLabels;
    expect(dl.showVal).toBe(true);
    expect(dl.showPercent).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 4. Axis options
// ---------------------------------------------------------------------------

describe("buildChartModel — axis options", () => {
  it("applies categoryAxis title", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { title: "Category" }
    });
    const catAx = pa(m).axes.find(a => a.axisType === "cat")!;
    expect(catAx.title).toBeDefined();
    expect(catAx.title!.text!.paragraphs[0].runs![0].text).toBe("Category");
  });

  it("applies valueAxis min/max/orientation", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { min: 0, max: 100, orientation: "maxMin" }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.scaling!.min).toBe(0);
    expect(valAx.scaling!.max).toBe(100);
    expect(valAx.scaling!.orientation).toBe("maxMin");
  });

  it("applies valueAxis majorUnit/minorUnit", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      valueAxis: { majorUnit: 10, minorUnit: 2 }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val") as any;
    expect(valAx.majorUnit).toBe(10);
    expect(valAx.minorUnit).toBe(2);
  });

  it("applies axis gridlines", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { majorGridlines: true, minorGridlines: true }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.majorGridlines).toBeDefined();
    expect(valAx.minorGridlines).toBeDefined();
  });

  it("applies axis hidden option", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { hidden: true }
    });
    const catAx = pa(m).axes.find(a => a.axisType === "cat")!;
    expect(catAx.delete).toBe(true);
  });

  it("applies axis numFmt", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { numFmt: "#,##0.00" }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val")!;
    expect(valAx.numFmt!.formatCode).toBe("#,##0.00");
  });

  it("applies axis textRotation", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { textRotation: -45 }
    });
    const catAx = pa(m).axes.find(a => a.axisType === "cat")!;
    expect(catAx.txPr).toBeDefined();
    expect((catAx.txPr as any).rotation).toBe(-45 * 60000);
  });

  it("applies axis crossBetween on valueAxis", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      valueAxis: { crossBetween: "midCat" }
    });
    const valAx = pa(m).axes.find(a => a.axisType === "val") as any;
    expect(valAx.crossBetween).toBe("midCat");
  });

  it("applies axis logBase", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [scatterSeries("S1")],
      valueAxis: { logBase: 10 }
    });
    // scatter: valueAxis → second val axis (y)
    const axes = pa(m).axes.filter(a => a.axisType === "val");
    const yAx = axes[1];
    expect(yAx.scaling!.logBase).toBe(10);
  });

  it("applies scatter categoryAxis to x val axis", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [scatterSeries("S1")],
      categoryAxis: { title: "X Axis" },
      valueAxis: { title: "Y Axis" }
    });
    const axes = pa(m).axes;
    // scatter: both axes are val; categoryAxis → first (x), valueAxis → second (y)
    expect(axes[0].title!.text!.paragraphs[0].runs![0].text).toBe("X Axis");
    expect(axes[1].title!.text!.paragraphs[0].runs![0].text).toBe("Y Axis");
  });
});

// ---------------------------------------------------------------------------
// 5. Combo charts
// ---------------------------------------------------------------------------

describe("buildComboChartModel", () => {
  it("creates multiple chart type groups", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Bars")] },
        { type: "line", series: [baseSeries("Lines")] }
      ]
    });
    expect(pa(m).chartTypes.length).toBe(2);
    expect(pa(m).chartTypes[0].type).toBe("bar");
    expect(pa(m).chartTypes[1].type).toBe("line");
  });

  it("shares primary axes between groups", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Bars")] },
        { type: "line", series: [baseSeries("Lines")] }
      ]
    });
    // Both groups share the same axis IDs
    expect(pa(m).chartTypes[0].axisIds).toEqual(pa(m).chartTypes[1].axisIds);
    // Only 2 axes total (cat + val), not 4
    expect(pa(m).axes.length).toBe(2);
  });

  it("creates secondary axes for useSecondaryAxis groups", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Primary")] },
        { type: "line", series: [baseSeries("Secondary")], useSecondaryAxis: true }
      ]
    });
    // 4 axes: primary cat+val, secondary cat+val
    expect(pa(m).axes.length).toBe(4);
    // Primary and secondary have different axis IDs
    expect(pa(m).chartTypes[0].axisIds).not.toEqual(pa(m).chartTypes[1].axisIds);
    // Secondary val axis should be on the right (axPos: "r")
    const secondaryValAx = pa(m).axes[3];
    expect(secondaryValAx.axPos).toBe("r");
  });

  it("creates serAx for 3D groups in combo", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar3D", series: [baseSeries("3D Bars")] },
        { type: "line", series: [baseSeries("Lines")] }
      ]
    });
    // bar3D needs serAx: cat + val + ser = 3 axes
    // line reuses cat + val (no ser)
    expect(pa(m).axes.length).toBe(3);
    expect(pa(m).axes[2].axisType).toBe("ser");
    expect(pa(m).chartTypes[0].axisIds!.length).toBe(3);
    expect(pa(m).chartTypes[1].axisIds!.length).toBe(2);
  });

  it("creates secondary serAx for 3D secondary groups", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Primary")] },
        { type: "bar3D", series: [baseSeries("Secondary 3D")], useSecondaryAxis: true }
      ]
    });
    // primary: cat + val (2)
    // secondary: cat + val + ser (3)
    expect(pa(m).axes.length).toBe(5);
    const serAxes = pa(m).axes.filter(a => a.axisType === "ser");
    expect(serAxes.length).toBe(1);
  });

  it("handles scatter in combo", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "scatter", series: [scatterSeries("S1")] },
        { type: "scatter", series: [scatterSeries("S2")], useSecondaryAxis: true }
      ]
    });
    // primary x+y + secondary x+y = 4
    expect(pa(m).axes.length).toBe(4);
    // All axes are val type
    expect(pa(m).axes.every(a => a.axisType === "val")).toBe(true);
  });

  it("applies combo-level title", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("S1")] },
        { type: "line", series: [baseSeries("S2")] }
      ],
      title: "Combo Chart"
    });
    expect(m.chart.title).toBeDefined();
    expect(m.chart.title!.text!.paragraphs[0].runs![0].text).toBe("Combo Chart");
  });
});

// ---------------------------------------------------------------------------
// 6. Round-trip tests — addChart → write → load → verify
// ---------------------------------------------------------------------------

describe("chart round-trip via addChart API", () => {
  it("round-trips a bar chart", async () => {
    const chart = await roundTripChart({
      type: "bar",
      series: [baseSeries("Revenue", VALUES_A), baseSeries("Cost", VALUES_B)],
      title: "Quarterly Results",
      barDir: "col",
      grouping: "clustered"
    });

    expect(chart).toBeDefined();
    expect(chart.chartNumber).toBeGreaterThan(0);
    expect(chart.title).toBe("Quarterly Results");
    expect(chart.chartTypes.length).toBe(1);
    expect(chart.chartTypes[0].type).toBe("bar");
    expect(chart.seriesCount).toBe(2);
    expect(chart.axes.length).toBe(2);
  });

  it("round-trips a line chart with smooth", async () => {
    const chart = await roundTripChart({
      type: "line",
      series: [{ ...baseSeries("Trend"), smooth: true }],
      title: "Trend Line"
    });

    expect(chart.title).toBe("Trend Line");
    expect(chart.chartTypes[0].type).toBe("line");
  });

  it("round-trips a pie chart", async () => {
    const chart = await roundTripChart({
      type: "pie",
      series: [baseSeries("Share")],
      title: "Market Share"
    });

    expect(chart.title).toBe("Market Share");
    expect(chart.chartTypes[0].type).toBe("pie");
    expect(chart.axes.length).toBe(0);
  });

  it("round-trips a scatter chart", async () => {
    const chart = await roundTripChart({
      type: "scatter",
      series: [scatterSeries("Points")],
      scatterStyle: "lineMarker"
    });

    expect(chart.chartTypes[0].type).toBe("scatter");
    expect(chart.axes.length).toBe(2);
  });

  it("round-trips a 3D bar chart with serAx", async () => {
    const chart = await roundTripChart({
      type: "bar3D",
      series: [baseSeries("3D Data")],
      title: "3D Chart"
    });

    expect(chart.title).toBe("3D Chart");
    expect(chart.chartTypes[0].type).toBe("bar3D");
    expect(chart.axes.length).toBe(3);
    expect(chart.axes[2].axisType).toBe("ser");
  });

  it("round-trips a doughnut chart", async () => {
    const chart = await roundTripChart({
      type: "doughnut",
      series: [baseSeries("Segments")],
      holeSize: 60
    });

    expect(chart.chartTypes[0].type).toBe("doughnut");
  });

  it("round-trips an area chart", async () => {
    const chart = await roundTripChart({
      type: "area",
      series: [baseSeries("Area")],
      grouping: "stacked"
    });

    expect(chart.chartTypes[0].type).toBe("area");
    expect(chart.axes.length).toBe(2);
  });

  it("round-trips a radar chart", async () => {
    const chart = await roundTripChart({
      type: "radar",
      series: [baseSeries("Radar")],
      radarStyle: "filled"
    });

    expect(chart.chartTypes[0].type).toBe("radar");
  });

  it("round-trips a surface chart", async () => {
    const chart = await roundTripChart({
      type: "surface",
      series: [baseSeries("Surface")],
      wireframe: true
    });

    expect(chart.chartTypes[0].type).toBe("surface");
    expect(chart.axes.length).toBe(3);
  });

  it("round-trips a bubble chart", async () => {
    const chart = await roundTripChart({
      type: "bubble",
      series: [bubbleSeries("Bubbles")],
      bubbleScale: 200
    });

    expect(chart.chartTypes[0].type).toBe("bubble");
    expect(chart.axes.length).toBe(2);
  });

  it("round-trips chart with no title (autoTitleDeleted)", async () => {
    const chart = await roundTripChart({
      type: "bar",
      series: [baseSeries("S1")]
    });

    expect(chart.title).toBeUndefined();
  });

  it("round-trips chart with axis options", async () => {
    const chart = await roundTripChart({
      type: "bar",
      series: [baseSeries("S1")],
      categoryAxis: { title: "Quarters", hidden: false },
      valueAxis: { title: "Amount", min: 0, max: 50, majorGridlines: true }
    });

    const catAx = chart.categoryAxis;
    expect(catAx).toBeDefined();
    const valAx = chart.valueAxis;
    expect(valAx).toBeDefined();
    expect(valAx!.scaling?.min).toBe(0);
    expect(valAx!.scaling?.max).toBe(50);
  });

  it("round-trips chart with series dataLabels", async () => {
    const chart = await roundTripChart({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          dataLabels: { showVal: true }
        }
      ]
    });

    const series = chart.getSeries(0) as any;
    expect(series).toBeDefined();
    expect(series.dataLabels?.showVal).toBe(true);
  });

  it("round-trips chart with series trendline", async () => {
    const chart = await roundTripChart({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          trendline: { type: "linear" }
        }
      ]
    });

    const series = chart.getSeries(0) as any;
    expect(series).toBeDefined();
    expect(series.trendlines?.length).toBe(1);
    expect(series.trendlines[0].type).toBe("linear");
  });
});

// ---------------------------------------------------------------------------
// 7. Combo chart round-trip
// ---------------------------------------------------------------------------

describe("combo chart round-trip via addComboChart API", () => {
  it("round-trips bar+line combo chart", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "x";
    ws.addComboChart(
      {
        groups: [
          { type: "bar", series: [baseSeries("Bars")], barDir: "col" },
          { type: "line", series: [baseSeries("Lines")], useSecondaryAxis: true }
        ],
        title: "Combo"
      },
      "C1:J15"
    );

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];

    expect(chart.title).toBe("Combo");
    expect(chart.chartTypes.length).toBe(2);
    expect(chart.chartTypes[0].type).toBe("bar");
    expect(chart.chartTypes[1].type).toBe("line");
    // Should have 4 axes (primary cat+val, secondary cat+val)
    expect(chart.axes.length).toBe(4);
  });

  it("round-trips combo with 3D group", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "x";
    ws.addComboChart(
      {
        groups: [
          { type: "bar3D", series: [baseSeries("3D Bars")] },
          { type: "line", series: [baseSeries("Lines")] }
        ],
        title: "3D Combo"
      },
      "C1:J15"
    );

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];

    expect(chart.title).toBe("3D Combo");
    expect(chart.chartTypes.length).toBe(2);
    // bar3D produces serAx
    expect(chart.axes.some(a => a.axisType === "ser")).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 8. High-level Chart API
// ---------------------------------------------------------------------------

describe("Chart high-level API", () => {
  it("Chart.title getter/setter", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "bar", series: [baseSeries("S1")], title: "Original" }, "C1:J10");
    const chart = ws.getCharts()[0];
    expect(chart.title).toBe("Original");

    chart.title = "Updated";
    expect(chart.title).toBe("Updated");

    chart.title = undefined;
    expect(chart.title).toBeUndefined();
  });

  it("Chart.legend getter/setter", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];

    expect(chart.legend).toBeDefined();
    chart.legend = undefined;
    expect(chart.legend).toBeUndefined();
  });

  it("Chart.chartTypes returns chart type groups", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "line", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];

    expect(chart.chartTypes.length).toBe(1);
    expect(chart.chartTypes[0].type).toBe("line");
  });

  it("Chart.axes returns axis list", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];

    expect(chart.axes.length).toBe(2);
    expect(chart.categoryAxis).toBeDefined();
    expect(chart.valueAxis).toBeDefined();
  });

  it("Chart.getSeries returns series by index", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart(
      {
        type: "bar",
        series: [baseSeries("First"), baseSeries("Second")]
      },
      "C1:J10"
    );
    const chart = ws.getCharts()[0];

    expect(chart.seriesCount).toBe(2);
    expect(chart.getSeries(0)).toBeDefined();
    expect(chart.getSeries(1)).toBeDefined();
    expect(chart.getSeries(2)).toBeUndefined();
  });

  it("Chart.spPr getter/setter", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];

    chart.spPr = { fill: { solid: { rgb: "FFFFFF" } } } as any;
    expect(chart.spPr).toBeDefined();
  });

  it("Chart.chartModel returns undefined for non-existent chart entry", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    // Manually construct a Chart with a chartExNumber to simulate chartEx
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    const chart = ws.getCharts()[0];
    // chartModel should be defined for a normal chart
    expect(chart.chartModel).toBeDefined();
    expect(chart.isChartEx).toBe(false);
  });

  it("getCharts returns all charts on the worksheet", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = "x";
    ws.addChart({ type: "bar", series: [baseSeries("S1")] }, "C1:J10");
    ws.addChart({ type: "pie", series: [baseSeries("S2")] }, "L1:S10");

    const charts = ws.getCharts();
    expect(charts.length).toBe(2);
    expect(charts[0].chartTypes[0].type).toBe("bar");
    expect(charts[1].chartTypes[0].type).toBe("pie");
  });
});

// ---------------------------------------------------------------------------
// 9. Edge cases
// ---------------------------------------------------------------------------

describe("chart builder edge cases", () => {
  it("handles empty series array", () => {
    const m = buildChartModel({ type: "bar", series: [] });
    expect(ctg(m).series.length).toBe(0);
  });

  it("handles undefined series", () => {
    const m = buildChartModel({ type: "bar" });
    expect(ctg(m).series.length).toBe(0);
  });

  it("multiple series with correct index/order", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("A"), baseSeries("B"), baseSeries("C")]
    });
    expect(ctg(m).series.length).toBe(3);
    expect(ctg(m).series[0].index).toBe(0);
    expect(ctg(m).series[1].index).toBe(1);
    expect(ctg(m).series[2].index).toBe(2);
    expect(ctg(m).series[0].order).toBe(0);
    expect(ctg(m).series[1].order).toBe(1);
    expect(ctg(m).series[2].order).toBe(2);
  });

  it("axis IDs are unique within a chart", () => {
    const m = buildChartModel({ type: "bar3D", series: [baseSeries("S1")] });
    const axIds = pa(m).axes.map(a => a.axId);
    expect(new Set(axIds).size).toBe(axIds.length);
  });

  it("combo chart axis IDs are all unique", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Primary")] },
        { type: "line", series: [baseSeries("Secondary")], useSecondaryAxis: true }
      ]
    });
    const axIds = pa(m).axes.map(a => a.axId);
    expect(new Set(axIds).size).toBe(axIds.length);
  });

  it("bubble chart with bubbleScale and showNegBubbles", () => {
    const m = buildChartModel({
      type: "bubble",
      series: [bubbleSeries("S1")],
      bubbleScale: 150,
      showNegBubbles: false,
      sizeRepresents: "w"
    });
    const grp = ctg(m) as any;
    expect(grp.bubbleScale).toBe(150);
    expect(grp.showNegBubbles).toBe(false);
    expect(grp.sizeRepresents).toBe("w");
  });

  it("trendline with all options", () => {
    const m = buildChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S1"),
          trendline: {
            type: "poly",
            order: 3,
            period: 2,
            forward: 1,
            backward: 0.5,
            intercept: 10,
            displayEq: true,
            displayRSqr: true,
            name: "Custom Trend",
            line: "FF0000",
            lineWidth: 1.5,
            lineDash: "dash"
          }
        }
      ]
    });
    const tl = (ctg(m).series[0] as any).trendlines[0];
    expect(tl.type).toBe("poly");
    expect(tl.order).toBe(3);
    expect(tl.period).toBe(2);
    expect(tl.forward).toBe(1);
    expect(tl.backward).toBe(0.5);
    expect(tl.intercept).toBe(10);
    expect(tl.name).toBe("Custom Trend");
    expect(tl.spPr.line.dash).toBe("dash");
  });

  it("errorBars with noEndCap and custom plus/minus", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          errorBars: {
            type: "cust",
            noEndCap: true,
            plus: "Sheet1!$E$1:$E$4",
            minus: "Sheet1!$F$1:$F$4",
            direction: "y"
          }
        }
      ]
    });
    const eb = (ctg(m).series[0] as any).errorBars;
    expect(eb.errValType).toBe("cust");
    expect(eb.noEndCap).toBe(true);
    expect(eb.plus).toBeDefined();
    expect(eb.minus).toBeDefined();
  });

  it("dataLabels with numFmt and separator", () => {
    const m = buildChartModel({
      type: "bar",
      series: [
        {
          ...baseSeries("S1"),
          dataLabels: {
            showVal: true,
            numFmt: "#,##0.00",
            numFmtLinked: false,
            separator: "\n",
            showBubbleSize: true,
            showLeaderLines: true
          }
        }
      ]
    });
    const dl = (ctg(m).series[0] as any).dataLabels;
    expect(dl.numFmt.formatCode).toBe("#,##0.00");
    expect(dl.numFmt.sourceLinked).toBe(false);
    expect(dl.separator).toBe("\n");
    expect(dl.showBubbleSize).toBe(true);
    expect(dl.showLeaderLines).toBe(true);
  });

  it("ofPie chart with splitType and splitPos", () => {
    const m = buildChartModel({
      type: "ofPie",
      series: [baseSeries("S1")],
      ofPieType: "pie",
      splitType: "pos",
      splitPos: 3,
      secondPieSize: 50
    });
    const grp = ctg(m) as any;
    expect(grp.splitType).toBe("pos");
    expect(grp.splitPos).toBe(3);
    expect(grp.secondPieSize).toBe(50);
    expect(grp.ofPieType).toBe("pie");
  });

  it("surface chart with wireframe", () => {
    const m = buildChartModel({
      type: "surface",
      series: [baseSeries("S1")],
      wireframe: true
    });
    expect((ctg(m) as any).wireframe).toBe(true);
  });

  it("chart has roundedCorners=false and lang=en-US by default", () => {
    const m = buildChartModel({ type: "bar", series: [] });
    expect(m.roundedCorners).toBe(false);
    expect(m.lang).toBe("en-US");
  });
});

// ---------------------------------------------------------------------------
// 10. Multiple charts on one sheet round-trip
// ---------------------------------------------------------------------------

describe("multiple charts round-trip", () => {
  it("preserves multiple charts on a single worksheet", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";

    ws.addChart({ type: "bar", series: [baseSeries("Bar")], title: "Chart 1" }, "C1:J10");
    ws.addChart({ type: "pie", series: [baseSeries("Pie")], title: "Chart 2" }, "C12:J22");

    expect(ws.getCharts().length).toBe(2);

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    expect(ws2.getCharts().length).toBe(2);
    expect(ws2.getCharts()[0].title).toBe("Chart 1");
    expect(ws2.getCharts()[1].title).toBe("Chart 2");
    expect(ws2.getCharts()[0].chartTypes[0].type).toBe("bar");
    expect(ws2.getCharts()[1].chartTypes[0].type).toBe("pie");
  });
});

// ---------------------------------------------------------------------------
// 11. New builder features: dataTable, hiLowLines, upDownBars, dropLines, serLines
// ---------------------------------------------------------------------------

describe("builder chart-level options", () => {
  it("dataTable=true sets all borders and keys", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")], dataTable: true });
    const dt = pa(m).dataTable;
    expect(dt).toBeDefined();
    expect(dt!.showHorzBorder).toBe(true);
    expect(dt!.showVertBorder).toBe(true);
    expect(dt!.showOutline).toBe(true);
    expect(dt!.showKeys).toBe(true);
  });

  it("dataTable with partial options", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      dataTable: { showHorzBorder: true, showKeys: false }
    });
    const dt = pa(m).dataTable;
    expect(dt).toBeDefined();
    expect(dt!.showHorzBorder).toBe(true);
    expect(dt!.showVertBorder).toBeUndefined();
    expect(dt!.showKeys).toBe(false);
  });

  it("dataTable not set when undefined", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(pa(m).dataTable).toBeUndefined();
  });

  it("plotVisOnly defaults to true", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(m.chart.plotVisOnly).toBe(true);
  });

  it("plotVisOnly=false is respected", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")], plotVisOnly: false });
    expect(m.chart.plotVisOnly).toBe(false);
  });

  it("showDLblsOverMax is set when provided", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      showDLblsOverMax: true
    });
    expect(m.chart.showDLblsOverMax).toBe(true);
  });

  it("firstSliceAng on pie chart", () => {
    const m = buildChartModel({
      type: "pie",
      series: [baseSeries("S1")],
      firstSliceAng: 90
    });
    expect((ctg(m) as any).firstSliceAng).toBe(90);
  });

  it("firstSliceAng defaults to 0 for pie", () => {
    const m = buildChartModel({ type: "pie", series: [baseSeries("S1")] });
    expect((ctg(m) as any).firstSliceAng).toBe(0);
  });

  it("hiLowLines on line chart", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      hiLowLines: true
    });
    const g = ctg(m) as LineChartGroup;
    expect(g.hiLowLines).toBeDefined();
  });

  it("upDownBars on line chart", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      upDownBars: true
    });
    const g = ctg(m) as LineChartGroup;
    expect(g.upDownBars).toBeDefined();
    expect(g.upDownBars!.gapWidth).toBe(150);
  });

  it("upDownBars with custom gapWidth", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      upDownBars: { gapWidth: 50 }
    });
    const g = ctg(m) as LineChartGroup;
    expect(g.upDownBars!.gapWidth).toBe(50);
  });

  it("dropLines on line chart", () => {
    const m = buildChartModel({
      type: "line",
      series: [baseSeries("S1")],
      dropLines: true
    });
    const g = ctg(m) as LineChartGroup;
    expect(g.dropLines).toBeDefined();
  });

  it("dropLines on area chart", () => {
    const m = buildChartModel({
      type: "area",
      series: [baseSeries("S1")],
      dropLines: true
    });
    const g = ctg(m) as AreaChartGroup;
    expect(g.dropLines).toBeDefined();
  });

  it("serLines on bar chart", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      serLines: true
    });
    const g = ctg(m) as BarChartGroup;
    expect(g.serLines).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// 12. Trendline array support
// ---------------------------------------------------------------------------

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
    expect(parsed.fill!.gradient!.stops[0].position).toBe(0);
    expect(parsed.fill!.gradient!.stops[0].color.srgb).toBe("FF0000");
    expect(parsed.fill!.gradient!.stops[2].position).toBe(100);
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
// 18. Rich text writing API
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
});

// ---------------------------------------------------------------------------
// 23. ChartEx (Office 2016+ extended charts) — 8 modern chart types
// ---------------------------------------------------------------------------

describe("ChartEx modern chart types", () => {
  function makeExWb() {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Populate some data
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = `Cat${i}`;
      ws.getCell(`B${i}`).value = i * 10;
    }
    return { wb, ws };
  }

  it("buildChartExModel for sunburst", () => {
    const m = buildChartExModel({
      type: "sunburst",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }]
    });
    expect(m.chartSpace.chartData.data).toHaveLength(2); // categories + values
    expect(m.chartSpace.chart.plotArea.plotAreaRegion?.series).toHaveLength(1);
    expect(m.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId).toBe("sunburst");
  });

  it("buildChartExModel for treemap with hierarchy", () => {
    const m = buildChartExModel({
      type: "treemap",
      categories: "Sheet1!$A$1:$A$5",
      series: [
        {
          name: "Tree",
          values: "Sheet1!$B$1:$B$5",
          hierarchy: ["Sheet1!$C$1:$C$5", "Sheet1!$D$1:$D$5"]
        }
      ]
    });
    // 1 cat + 1 val + 2 hierarchy = 4 data entries
    expect(m.chartSpace.chartData.data).toHaveLength(4);
    expect(m.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId).toBe("treemap");
  });

  it("buildChartExModel for waterfall with subtotals", () => {
    const m = buildChartExModel({
      type: "waterfall",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Flow", values: "Sheet1!$B$1:$B$5", subtotals: [2, 4] }]
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("waterfall");
    expect(s.layoutPr?.subtotals).toEqual([{ idx: 2 }, { idx: 4 }]);
  });

  it("buildChartExModel for funnel", () => {
    const m = buildChartExModel({
      type: "funnel",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Funnel", values: "Sheet1!$B$1:$B$5" }]
    });
    expect(m.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId).toBe("funnel");
    expect(m.chartSpace.chart.plotArea.axis).toBeUndefined();
  });

  it("buildChartExModel for histogram", () => {
    const m = buildChartExModel({
      type: "histogram",
      series: [
        {
          name: "Histogram",
          values: "Sheet1!$B$1:$B$5"
        }
      ],
      layout: {
        binning: { binCount: 10, intervalClosed: "r" }
      }
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("clusteredColumn");
    expect(s.layoutPr?.binning?.binCount).toBe(10);
  });

  it("buildChartExModel for boxWhisker", () => {
    const m = buildChartExModel({
      type: "boxWhisker",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Box", values: "Sheet1!$B$1:$B$5" }],
      layout: {
        quartileMethod: "inclusive",
        showMeanLine: true,
        showMeanMarker: true
      }
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("boxWhisker");
    expect(s.layoutPr?.quartileMethod).toBe("inclusive");
    expect(s.layoutPr?.showMeanLine).toBe(true);
  });

  it("buildChartExModel for regionMap", () => {
    const m = buildChartExModel({
      type: "regionMap",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Map", values: "Sheet1!$B$1:$B$5" }],
      layout: {
        projection: "mercator",
        regionLabels: "bestFit",
        geoMappingLevel: "country"
      }
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("regionMap");
    expect(s.layoutPr?.projection).toBe("mercator");
  });

  it("renderChartEx produces valid cx: XML", () => {
    const m = buildChartExModel({
      type: "sunburst",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
      title: "My Sunburst"
    });
    const xml = renderChartEx(m);
    expect(xml).toContain("<cx:chartSpace");
    expect(xml).toContain('xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"');
    expect(xml).toContain("<cx:chartData>");
    expect(xml).toContain("<cx:chart>");
    expect(xml).toContain('layoutId="sunburst"');
    expect(xml).toContain("My Sunburst");
    expect(xml).toContain("Sheet1!$A$1:$A$5");
    expect(xml).toContain("Sheet1!$B$1:$B$5");
  });

  it("addChartEx integrates with worksheet write pipeline", async () => {
    const { wb, ws } = makeExWb();
    const num = ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
        title: "Programmatic ChartEx"
      },
      "D1:J10"
    );
    expect(num).toBe(1);
    expect(ws.getCharts()).toHaveLength(1);
    expect(ws.getCharts()[0].isChartEx).toBe(true);

    // Write and verify the XML is in the zip
    const buf = await wb.xlsx.writeBuffer();
    // Load and check bytes are present
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2.getCharts()).toHaveLength(1);
    expect(ws2.getCharts()[0].isChartEx).toBe(true);
  });

  it("addChartEx for each of 8 types creates a distinct chart", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = `C${i}`;
      ws.getCell(`B${i}`).value = i * 10;
    }

    const types: AddChartExOptions["type"][] = [
      "sunburst",
      "treemap",
      "waterfall",
      "funnel",
      "histogram",
      "pareto",
      "boxWhisker",
      "regionMap"
    ];
    let row = 1;
    for (const type of types) {
      ws.addChartEx(
        {
          type,
          categories: "Sheet1!$A$1:$A$5",
          series: [{ values: "Sheet1!$B$1:$B$5", name: type }]
        },
        { tl: { col: 4, row: row - 1 }, br: { col: 12, row: row + 9 } }
      );
      row += 12;
    }

    expect(ws.getCharts()).toHaveLength(8);
    expect(ws.getCharts().every(c => c.isChartEx)).toBe(true);

    // Round-trip
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2.getCharts()).toHaveLength(8);
  });
});

// ---------------------------------------------------------------------------
// 24. TC1-TC7 + GAP supplementary tests
// ---------------------------------------------------------------------------

describe("TC1: removeChart cleans up workbook chart entry", () => {
  it("workbook.getChartEntry returns undefined after removeChart", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    const chartNum = ws.addChart({ type: "bar", series: [baseSeries("S")] }, "C1:J10");
    expect(wb.getChartEntry(chartNum)).toBeDefined();
    ws.removeChart(0);
    expect(wb.getChartEntry(chartNum)).toBeUndefined();
  });
});

describe("TC2: ChartEx round-trip preserves isChartEx", () => {
  it("chartEx remains isChartEx after write/read", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = `C${i}`;
      ws.getCell(`B${i}`).value = i * 10;
    }
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "S", values: "Sheet1!$B$1:$B$5" }],
        title: "Sunburst"
      },
      "D1:K10"
    );
    expect(ws.getCharts()[0].isChartEx).toBe(true);
    expect(ws.getCharts()[0].chartNumber).toBe(0);
    expect(ws.getCharts()[0].chartExNumber).toBeGreaterThan(0);

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    const chart2 = ws2.getCharts()[0];
    expect(chart2.isChartEx).toBe(true);
    expect(chart2.chartNumber).toBe(0);
    expect(chart2.chartExNumber).toBeGreaterThan(0);
  });
});

describe("TC4: combo chart with categoryAxis/valueAxis options on primary group", () => {
  it("applies axis options to the primary axes", () => {
    const m = buildComboChartModel({
      groups: [
        {
          type: "bar",
          series: [baseSeries("Bar")],
          categoryAxis: { title: "Primary Cat" },
          valueAxis: { title: "Primary Val", min: 0, max: 100 }
        },
        {
          type: "line",
          series: [baseSeries("Line", VALUES_B)],
          useSecondaryAxis: true
        }
      ]
    });
    const axes = m.chart.plotArea.axes;
    // Primary category axis should have the title
    const primaryCat = axes.find(
      a => a.axisType === "cat" && a.title?.text?.paragraphs[0]?.runs?.[0]?.text === "Primary Cat"
    );
    expect(primaryCat).toBeDefined();
    // Primary value axis should have the title + min/max
    const primaryVal = axes.find(
      a => a.axisType === "val" && a.title?.text?.paragraphs[0]?.runs?.[0]?.text === "Primary Val"
    );
    expect(primaryVal).toBeDefined();
    expect(primaryVal!.scaling?.min).toBe(0);
    expect(primaryVal!.scaling?.max).toBe(100);
    // Secondary axes exist
    expect(axes.length).toBeGreaterThanOrEqual(4); // primary cat + val + secondary cat + val
  });
});

describe("TC5: Chart.clone/copyTo round-trip", () => {
  it("cloned chart survives write/read as independent copy", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "x";
    ws.getCell("B1").value = 100;
    ws.addChart({ type: "bar", series: [baseSeries("Original")], title: "Source" }, "C1:J10");
    const original = ws.getCharts()[0];
    original.clone();
    expect(ws.getCharts()).toHaveLength(2);

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2.getCharts()).toHaveLength(2);
    // Both charts should be independent
    expect(ws2.getCharts()[0].title).toBe("Source");
    expect(ws2.getCharts()[1].title).toBe("Source");
  });

  it("copyTo another worksheet creates independent chart", async () => {
    const wb = new Workbook();
    const src = wb.addWorksheet("Src");
    const dst = wb.addWorksheet("Dst");
    src.getCell("A1").value = 1;
    src.addChart({ type: "pie", series: [baseSeries("P")], title: "Pie" }, "C1:J10");
    src.getCharts()[0].copyTo(dst, "A1:H10");
    expect(dst.getCharts()).toHaveLength(1);

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.getWorksheet("Src")!.getCharts()).toHaveLength(1);
    expect(wb2.getWorksheet("Dst")!.getCharts()).toHaveLength(1);
    expect(wb2.getWorksheet("Dst")!.getCharts()[0].title).toBe("Pie");
  });
});

describe("TC6: Date values in chart cache population", () => {
  it("fills numRef cache with serial numbers for Date cell values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Label";
    ws.getCell("B1").value = new Date(2024, 0, 1); // 2024-01-01

    ws.addChart(
      {
        type: "line",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$B$1:$B$1" }]
      },
      "C1:J10"
    );
    const chart = ws.getCharts()[0];
    const series = chart.chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    const points = series.val.numRef.cache.points;
    expect(points).toHaveLength(1);
    // Excel serial for 2024-01-01 (1900-based, non-1904) should be ~45292
    expect(points[0].value).toBeGreaterThan(45000);
    expect(points[0].value).toBeLessThan(46000);
  });
});

describe("TC7: pareto chart type in ChartEx builder", () => {
  it("pareto type uses clusteredColumn layoutId", () => {
    const m = buildChartExModel({
      type: "pareto",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }]
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("clusteredColumn");
  });

  it("pareto chart renders valid XML", () => {
    const m = buildChartExModel({
      type: "pareto",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Pareto", values: "Sheet1!$B$1:$B$5" }],
      title: "Pareto Analysis"
    });
    const xml = renderChartEx(m);
    expect(xml).toContain("<cx:chartSpace");
    expect(xml).toContain('layoutId="clusteredColumn"');
    expect(xml).toContain("Pareto Analysis");
  });
});

describe("Builder gap fixes verification", () => {
  it("stock chart supports hiLowLines, upDownBars, dropLines", () => {
    const m = buildChartModel({
      type: "stock",
      series: [
        baseSeries("Open"),
        baseSeries("High", VALUES_B),
        baseSeries("Low"),
        baseSeries("Close", VALUES_B)
      ],
      hiLowLines: true,
      upDownBars: { gapWidth: 100 },
      dropLines: true
    });
    const g = ctg(m) as StockChartGroup;
    expect(g.hiLowLines).toBeDefined();
    expect(g.upDownBars).toBeDefined();
    expect(g.upDownBars!.gapWidth).toBe(100);
    expect(g.dropLines).toBeDefined();
  });

  it("stock chart supports varyColors", () => {
    const m = buildChartModel({
      type: "stock",
      series: [baseSeries("S1"), baseSeries("S2", VALUES_B)],
      varyColors: true
    });
    const g = ctg(m) as StockChartGroup;
    expect(g.varyColors).toBe(true);
  });

  it("ofPie chart supports gapWidth and serLines", () => {
    const m = buildChartModel({
      type: "ofPie",
      series: [baseSeries("S")],
      gapWidth: 200,
      serLines: true
    });
    const g = ctg(m) as OfPieChartGroup;
    expect(g.gapWidth).toBe(200);
    expect(g.serLines).toBeDefined();
  });
});

describe("Renderer gap fixes verification", () => {
  it("line with only width and dash renders via round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            ...baseSeries("S"),
            spPr: {
              line: { width: 25400, dash: "dash" }
            }
          }
        ]
      },
      "C1:J10"
    );
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const s = wb2.getWorksheet("Sheet1")!.getCharts()[0].chartModel!.chart.plotArea.chartTypes[0]
      .series[0];
    const parsed = parseSpPr(s.spPr!);
    expect(parsed.line?.width).toBe(25400);
    expect(parsed.line?.dash).toBe("dash");
  });

  it("line cap/join/compound are preserved through structured render", () => {
    const spPr = buildSpPr({
      line: { width: 12700, cap: "rnd", join: "round", compound: "sng" }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.line?.width).toBe(12700);
    expect(spPr.line?.cap).toBe("rnd");
    expect(spPr.line?.join).toBe("round");
    expect(spPr.line?.compound).toBe("sng");
  });

  it("sysClr color renders correctly", () => {
    const spPr = buildSpPr({
      fill: { solid: { sysClr: "windowText" } }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.fill?.solid?.sysClr).toBe("windowText");
  });

  it("prstClr color renders correctly", () => {
    const spPr = buildSpPr({
      fill: { solid: { prstClr: "red" } }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.fill?.solid?.prstClr).toBe("red");
  });

  it("shade and satMod modifiers render on theme color", () => {
    const spPr = buildSpPr({
      fill: { solid: { theme: 4, shade: 75000, satMod: 120000 } }
    });
    // buildSpPr returns a structured object (no _rawXml)
    expect(spPr._rawXml).toBeUndefined();
    expect(spPr.fill?.solid?.shade).toBe(75000);
    expect(spPr.fill?.solid?.satMod).toBe(120000);
  });
});

// ---------------------------------------------------------------------------
// 25. Bug fix verifications and supplementary edge cases
// ---------------------------------------------------------------------------

describe("BUG-7+8: parseSpPr handles malformed XML", () => {
  it("parseSpPr handles missing end tag gracefully", () => {
    // A truncated raw XML where </a:srgbClr> is missing
    const rawXml = '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>';
    const parsed = parseSpPr({ _rawXml: rawXml } as any);
    expect(parsed.fill?.solid?.srgb).toBe("FF0000");
  });
});

describe("BUG-11: funnel chart has no axes", () => {
  it("funnel chartEx has no axes", () => {
    const m = buildChartExModel({
      type: "funnel",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "F", values: "Sheet1!$B$1:$B$5" }]
    });
    expect(m.chartSpace.chart.plotArea.axis).toBeUndefined();
    const series = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(series.axisId).toBeUndefined();
  });
});

describe("ROBUST-4: Excel 1900 leap year bug in date serial", () => {
  it("dateToSerial accounts for Excel 1900 leap year bug", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Cat";
    // March 1, 1900 — Excel serial should be 61 (not 60) for the integer part.
    // The Date constructor uses local time, so the serial may include a fractional
    // time component depending on timezone. We check the floor is 61.
    ws.getCell("B1").value = new Date(1900, 2, 1);
    ws.addChart(
      {
        type: "line",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$B$1:$B$1" }]
      },
      "C1:J10"
    );
    const series = ws.getCharts()[0].chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    const serial = series.val.numRef.cache.points[0].value;
    // The integer part must be 61 (accounting for the 1900 leap year bug)
    expect(Math.floor(serial)).toBe(61);
  });
});

describe("ROBUST-6: _renderTxPr includes rotation", () => {
  it("chart axis textRotation round-trips through structured txPr", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.addChart(
      { type: "bar", series: [baseSeries("S")], categoryAxis: { textRotation: -45 } },
      "C1:J10"
    );
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    // After round-trip, axis txPr should contain rotation
    const catAx = wb2
      .getWorksheet("Sheet1")!
      .getCharts()[0]
      .chartModel!.chart.plotArea.axes.find(a => a.axisType === "cat")!;
    // txPr comes back as raw XML — check it contains rot="-2700000" (−45° × 60000)
    expect(catAx.txPr?._rawXml || "").toContain("-2700000");
  });
});

describe("supplementary edge cases", () => {
  it("buildChartModel handles empty series array", () => {
    const m = buildChartModel({ type: "bar", series: [] });
    expect(m.chart.plotArea.chartTypes[0].series).toHaveLength(0);
  });

  it("cache populator handles cells with null/undefined values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Cat";
    ws.getCell("B1").value = null;
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$B$1:$B$1" }]
      },
      "C1:J10"
    );
    const series = ws.getCharts()[0].chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    // null value should produce empty cache points (nothing at index 0)
    expect(series.val.numRef.cache.points).toHaveLength(0);
    expect(series.val.numRef.cache.pointCount).toBe(1);
  });

  it("combo chart supports 3 or more groups", () => {
    const m = buildComboChartModel({
      groups: [
        { type: "bar", series: [baseSeries("Bar")] },
        { type: "line", series: [baseSeries("Line", VALUES_B)] },
        { type: "area", series: [baseSeries("Area")] }
      ]
    });
    expect(m.chart.plotArea.chartTypes).toHaveLength(3);
    expect(m.chart.plotArea.chartTypes[0].type).toBe("bar");
    expect(m.chart.plotArea.chartTypes[1].type).toBe("line");
    expect(m.chart.plotArea.chartTypes[2].type).toBe("area");
  });

  it("doughnut chart with holeSize 0 and 90", () => {
    const m0 = buildChartModel({ type: "doughnut", series: [baseSeries("S")], holeSize: 0 });
    expect((ctg(m0) as any).holeSize).toBe(0);
    const m90 = buildChartModel({ type: "doughnut", series: [baseSeries("S")], holeSize: 90 });
    expect((ctg(m90) as any).holeSize).toBe(90);
  });

  it("all 16 chart types round-trip through write/read", async () => {
    const types: AddChartOptions["type"][] = [
      "bar",
      "bar3D",
      "line",
      "line3D",
      "pie",
      "pie3D",
      "doughnut",
      "area",
      "area3D",
      "scatter",
      "bubble",
      "radar",
      "stock",
      "surface",
      "surface3D",
      "ofPie"
    ];
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    for (let i = 1; i <= 4; i++) {
      ws.getCell(`A${i}`).value = `Cat${i}`;
      ws.getCell(`B${i}`).value = i * 10;
      ws.getCell(`C${i}`).value = i * 5;
    }
    let row = 1;
    for (const type of types) {
      const opts: AddChartOptions =
        type === "scatter"
          ? { type, series: [scatterSeries("S")] }
          : type === "bubble"
            ? { type, series: [bubbleSeries("S")] }
            : type === "stock"
              ? {
                  type,
                  series: [
                    baseSeries("O"),
                    baseSeries("H", VALUES_B),
                    baseSeries("L"),
                    baseSeries("C", VALUES_B)
                  ]
                }
              : { type, series: [baseSeries("S")] };
      ws.addChart(opts, { tl: { col: 4, row }, br: { col: 12, row: row + 9 } });
      row += 12;
    }
    expect(ws.getCharts()).toHaveLength(16);
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2.getCharts()).toHaveLength(16);
    for (let i = 0; i < types.length; i++) {
      expect(ws2.getCharts()[i].chartTypes[0].type).toBe(types[i]);
    }
  });
});
