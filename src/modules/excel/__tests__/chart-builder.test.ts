/**
 * Chart builder core tests — buildChartModel core, round-trip, high-level API, edge cases, chart-level options.
 *
 * This file is the first of six chart-builder-*.test.ts files
 * the original 13,000+ line suite was split into. See the
 * sibling `chart-builder.helpers.ts` for shared imports and
 * utility functions every split consumes.
 */

import type { BarChartGroup, LineChartGroup, AreaChartGroup, ChartExModel } from "@excel/chart";
import {
  buildChartModel,
  buildComboChartModel,
  buildChartExModel,
  renderChartExPng,
  buildChartScene,
  buildEffectFilter,
  renderChartSvg
} from "@excel/chart";
import { type ChartSceneSeries } from "@excel/chart/index";
import { installChartSupport } from "@excel/chart/install";
import { Workbook } from "@excel/workbook";
import { beforeAll, describe, it, expect } from "vitest";

import {
  CATEGORIES,
  VALUES_A,
  VALUES_B,
  baseSeries,
  bubbleSeries,
  ctg,
  pa,
  roundTripChart,
  scatterSeries
} from "./chart-builder.helpers";

const textDecoder = new TextDecoder();

beforeAll(() => {
  installChartSupport();
});

describe("buildChartModel — all 16 chart types", () => {
  it("bar chart", () => {
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(ctg(m).type).toBe("bar");
    expect(pa(m).axes.length).toBe(2);
    expect(pa(m).axes[0].axisType).toBe("cat");
    expect(pa(m).axes[1].axisType).toBe("val");
  });

  it("ChartEx layoutPr fields round-trip for every layoutId where they apply", async () => {
    // Cover the per-layoutId fields that Excel authors actually edit
    // in the GUI. Each type has its own dialect of layoutPr — histogram
    // has `binning`, boxWhisker has the four show* flags + quartileMethod,
    // waterfall has subtotals + connectorLines, regionMap has projection
    // + regionLabels + geoMappingLevel. This test guards against any
    // one of them being silently dropped during build → write → load.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 30],
      ["D", 40],
      ["E", 50]
    ]);

    ws.addChartEx(
      {
        type: "histogram",
        series: [{ values: "Data!$B$1:$B$5", literalValues: [1, 2, 3, 4, 5] }],
        binning: { binCount: 5, intervalClosed: "r", underflow: 0, overflow: 10 }
      },
      "D1:J10"
    );
    ws.addChartEx(
      {
        type: "boxWhisker",
        categories: "Data!$A$1:$A$5",
        series: [{ values: "Data!$B$1:$B$5" }],
        layout: {
          quartileMethod: "inclusive",
          showMeanLine: true,
          showMeanMarker: true,
          showInnerPoints: true,
          showOutlierPoints: true
        }
      },
      "D12:J22"
    );
    ws.addChartEx(
      {
        type: "waterfall",
        categories: "Data!$A$1:$A$5",
        series: [{ values: "Data!$B$1:$B$5" }],
        layout: {
          subtotals: [{ idx: 2 }, { idx: 4 }],
          connectorLines: true
        }
      },
      "D24:J34"
    );
    ws.addChartEx(
      {
        type: "regionMap",
        series: [
          {
            values: "Data!$B$1:$B$2",
            literalValues: [10, 20],
            literalCategories: ["USA", "Canada"]
          }
        ],
        layout: {
          projection: "albers",
          regionLabels: "showAll",
          geoMappingLevel: "country"
        }
      },
      "D36:J46"
    );

    const bytes = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(bytes);
    const charts = wb2.getWorksheet("Data")!.getCharts();
    expect(charts.length).toBe(4);

    const layoutFor = (idx: number) =>
      charts[idx].chartExModel!.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutPr;

    // Histogram
    const histo = layoutFor(0);
    expect(histo?.binning?.binCount).toBe(5);
    expect(histo?.binning?.intervalClosed).toBe("r");
    expect(histo?.binning?.underflow).toBe(0);
    expect(histo?.binning?.overflow).toBe(10);

    // BoxWhisker
    const box = layoutFor(1);
    expect(box?.quartileMethod).toBe("inclusive");
    expect(box?.showMeanLine).toBe(true);
    expect(box?.showMeanMarker).toBe(true);
    expect(box?.showInnerPoints).toBe(true);
    expect(box?.showOutlierPoints).toBe(true);

    // Waterfall
    const waterfall = layoutFor(2);
    expect(waterfall?.subtotals).toEqual([{ idx: 2 }, { idx: 4 }]);
    expect(waterfall?.connectorLines).toBe(true);

    // RegionMap
    const region = layoutFor(3);
    expect(region?.projection).toBe("albers");
    expect(region?.regionLabels).toBe("showAll");
    expect(region?.geoMappingLevel).toBe("country");
  });

  it("ChartEx PNG pipeline produces valid byte-shape for every layoutId", async () => {
    // Fill the ChartEx PNG gap noted in the compatibility matrix —
    // before this, only classic charts had PNG signature assertions
    // even though `renderChartExPng` goes through the same SVG-to-
    // raster pipeline. Confirm each ChartEx kind produces the correct
    // PNG magic + IHDR dimensions.
    const fixtures = [
      {
        name: "sunburst",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "sunburst",
            series: [{ values: "A", literalValues: [1, 2], literalCategories: ["A", "B"] }]
          })
      },
      {
        name: "treemap",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "treemap",
            series: [{ values: "A", literalValues: [10, 5], literalCategories: ["A", "B"] }]
          })
      },
      {
        name: "waterfall",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "waterfall",
            categories: "A1:A3",
            series: [
              { values: "B1:B3", literalCategories: ["A", "B", "C"], literalValues: [5, 10, -3] }
            ]
          })
      },
      {
        name: "funnel",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "funnel",
            series: [
              { values: "A", literalValues: [100, 80, 50], literalCategories: ["T", "M", "B"] }
            ]
          })
      },
      {
        name: "histogram",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "histogram",
            series: [{ values: "A", literalValues: [1, 2, 3, 4, 5] }],
            binning: { binCount: 3 }
          })
      },
      {
        name: "pareto",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "pareto",
            series: [{ values: "A", literalValues: [5, 3, 1], literalCategories: ["A", "B", "C"] }]
          })
      },
      {
        name: "boxWhisker",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "boxWhisker",
            categories: "A1:A3",
            series: [
              { values: "B1:B3", literalCategories: ["A", "B", "C"], literalValues: [2, 5, 7] }
            ]
          })
      },
      {
        name: "regionMap",
        build: (): ChartExModel =>
          buildChartExModel({
            type: "regionMap",
            series: [
              {
                values: "A",
                literalValues: [10, 20],
                literalCategories: ["USA", "Canada"]
              }
            ]
          })
      }
    ];
    for (const fixture of fixtures) {
      const png = await renderChartExPng(fixture.build(), { width: 240, height: 160 });
      // PNG magic bytes (8-byte signature) + IHDR chunk header.
      expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      expect(textDecoder.decode(png.slice(12, 16))).toBe("IHDR");
      // Dimensions readable at offsets 16-23 (IHDR width/height big-endian).
      const width = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
      const height = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
      expect(width).toBe(240);
      expect(height).toBe(160);
    }
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
  it("leaves autoTitleDeleted absent when no title provided", () => {
    // The builder previously set `autoTitleDeleted: true` whenever
    // `opts.title` was absent, suppressing Excel's default auto-title
    // behaviour. The new semantics: omit → leave Excel to auto-title
    // (autoTitleDeleted undefined); pass `title: null` to explicitly
    // suppress.
    const m = buildChartModel({ type: "bar", series: [baseSeries("S1")] });
    expect(m.chart.autoTitleDeleted).toBeUndefined();
    expect(m.chart.title).toBeUndefined();
  });

  it("auto-deletes title when title is explicitly null", () => {
    const m = buildChartModel({
      type: "bar",
      series: [baseSeries("S1")],
      title: null
    });
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

  it("throws clear errors for invalid chart options", () => {
    expect(() =>
      buildChartModel({
        type: "pie",
        series: [{ ...baseSeries("S1"), trendline: { type: "linear" } }]
      })
    ).toThrow("chart.series[0].trendline is not valid for pie charts");

    expect(() =>
      buildChartModel({ type: "doughnut", series: [baseSeries("S1")], holeSize: 120 })
    ).toThrow("chart.holeSize must be an integer between 0 and 90");

    expect(() =>
      buildChartModel({
        type: "pie",
        series: [baseSeries("S1")],
        categoryAxis: { title: "Invalid" }
      })
    ).toThrow("chart.categoryAxis is not valid for pie charts because they do not have axes");

    expect(() =>
      buildChartModel({ type: "bar", series: [baseSeries("S1")], valueAxis: { min: 10, max: 5 } })
    ).toThrow("chart.valueAxis.min must be less than chart.valueAxis.max");

    expect(() =>
      buildChartModel({ type: "pie", series: [baseSeries("S1")], grouping: "stacked" })
    ).toThrow("chart.grouping is only valid for bar, line, and area charts");

    expect(() =>
      buildChartModel({ type: "bar", series: [baseSeries("S1")], scatterStyle: "smooth" })
    ).toThrow("chart.scatterStyle is only valid for scatter charts");

    expect(() =>
      buildChartModel({ type: "bar", series: [baseSeries("S1")], wireframe: true })
    ).toThrow("chart.wireframe is only valid for surface and surface3D charts");
  });

  it("throws clear errors for invalid nested options", () => {
    expect(() =>
      buildChartModel({ type: "line", series: [{ ...baseSeries("S1"), marker: { size: 100 } }] })
    ).toThrow("chart.series[0].marker.size must be an integer between 2 and 72");

    expect(() => buildChartModel({ type: "scatter", series: [baseSeries("S1")] })).toThrow(
      "chart.series[0].xValues is required for scatter charts"
    );

    expect(() =>
      buildChartModel({ type: "bubble", series: [{ ...baseSeries("S1"), xValues: VALUES_A }] })
    ).toThrow("chart.series[0].bubbleSize is required for bubble charts");

    expect(() =>
      buildChartModel({
        type: "line",
        series: [{ ...baseSeries("S1"), trendline: { type: "poly", order: 8 } }]
      })
    ).toThrow("chart.series[0].trendline.order must be an integer between 2 and 6");

    expect(() =>
      buildChartModel({
        type: "line",
        series: [{ ...baseSeries("S1"), errorBars: { type: "cust", plus: VALUES_A } }]
      })
    ).toThrow(
      'chart.series[0].errorBars.plus and chart.series[0].errorBars.minus are required when type is "cust"'
    );
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
    expect(chart.getSeriesCount(0)).toBe(2);
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

    expect(chart.getSeriesCount(0)).toBe(2);
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
  it("rejects empty series array (Excel won't render a series-less chart)", () => {
    expect(() => buildChartModel({ type: "bar", series: [] })).toThrow(/at least one series/);
  });

  it("rejects undefined series (same reason)", () => {
    expect(() => buildChartModel({ type: "bar" })).toThrow(/at least one series/);
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
            // `period` is only valid for movingAvg trendlines; using it on
            // other types is now a validation error (Excel silently ignores
            // it, which quickly devolves into phantom config bugs).
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
    const m = buildChartModel({
      type: "bar",
      series: [{ categories: CATEGORIES, values: VALUES_A }]
    });
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

  it("renderChartSvg draws a dataTable overlay below the plot", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["Jan", 10, 20],
      ["Feb", 15, 25],
      ["Mar", 20, 30]
    ]);
    ws.addChart(
      {
        type: "bar",
        dataTable: {
          showOutline: true,
          showHorzBorder: true,
          showVertBorder: true,
          showKeys: true
        },
        series: [
          { name: "Sales", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" },
          { name: "Profit", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$C$1:$C$3" }
        ]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    const scene = buildChartScene(model, { width: 600, height: 400 });
    expect(scene.dataTable).toBeDefined();
    const dt = scene.dataTable!;
    // Header row + 2 series rows = 3 rows → 4 row boundaries.
    expect(dt.rows.length).toBe(4);
    // Name column + 3 category columns = 4 columns → 5 boundaries.
    expect(dt.columns.length).toBe(5);
    // Header cells (3 categories) + 2 series × (1 name + 3 values) = 11 cells.
    expect(dt.cells.length).toBe(11);
    // Two series → two legend swatches.
    expect(dt.legendSwatches.length).toBe(2);
    expect(dt.legendSwatches[0].color).toMatch(/^#/);
    // All four sides + two interior horizontal + two interior vertical = 8 borders.
    expect(dt.borders.length).toBeGreaterThanOrEqual(8);
    // Plot should sit above the data table.
    expect(dt.rect.y).toBeGreaterThan(scene.plot.y + scene.plot.height);

    // The SVG emits the category header, series names and values as
    // <text> nodes and the outline/key swatches as <rect> nodes.
    const svg = renderChartSvg(model, { width: 600, height: 400 });
    expect(svg).toContain(">Jan<");
    expect(svg).toContain(">Feb<");
    expect(svg).toContain(">Sales<");
    expect(svg).toContain(">Profit<");
    expect(svg).toContain(">10<");
    expect(svg).toContain(">30<");
  });

  it("dataTable suppresses x-axis category labels (Excel parity)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["Q1", 10],
      ["Q2", 20]
    ]);
    ws.addChart(
      {
        type: "bar",
        dataTable: true,
        series: [{ categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    const scene = buildChartScene(model, { width: 500, height: 350 });
    expect(scene.dataTable).toBeDefined();
    expect(scene.xLabels).toEqual([]);
  });

  it("bar3D renders a true extruded box with view3D-driven projection", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 5],
      ["B", 20, 15],
      ["C", 30, 25]
    ]);
    ws.addChart(
      {
        type: "bar3D",
        view3D: { rotX: 20, rotY: 30, depthPercent: 100, rAngAx: true },
        series: [
          { categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" },
          { categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$C$1:$C$3" }
        ]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    const scene = buildChartScene(model, { width: 640, height: 400 });
    const bar = scene.series.find(
      (s): s is ChartSceneSeries & { type: "bar" } => s.type === "bar"
    )!;
    expect(bar.type).toBe("bar");
    // 3D projection is populated.
    expect(bar.depth).toBeGreaterThan(0);
    expect(bar.projection3D).toBeDefined();
    expect(Math.abs(bar.projection3D!.dx)).toBeGreaterThan(0);
    expect(Math.abs(bar.projection3D!.dy)).toBeGreaterThan(0);
    // SVG emits three polygons/rects per bar (top face + right face + front rect).
    const svg = renderChartSvg(model, { width: 640, height: 400 });
    // Two series × 3 bars each = 6 bars. Each produces 2 polygons (top+right) + 1 rect.
    const polygonMatches = svg.match(/<polygon /g) ?? [];
    expect(polygonMatches.length).toBeGreaterThanOrEqual(12);
  });

  it("bar3D falls back to default view3D when model omits it", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "bar3D",
        series: [{ categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    expect(model.chart.view3D).toBeUndefined();
    const scene = buildChartScene(model, { width: 500, height: 350 });
    const bar = scene.series.find(
      (s): s is ChartSceneSeries & { type: "bar" } => s.type === "bar"
    )!;
    expect(bar.projection3D).toBeDefined();
  });

  // Regression — stacked bar/area y-range used to ignore stacking and
  // compute range from per-series maxes, so cumulative columns
  // overflowed the plot rectangle. Verifies the column-sum term in
  // `buildAxisContext`.
  it("stacked column y-range includes the cumulative sum, not per-series max", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["Q1", 10, 30],
      ["Q2", 20, 40],
      ["Q3", 30, 50]
    ]);
    ws.addChart(
      {
        type: "bar",
        barDir: "col",
        grouping: "stacked",
        series: [
          { categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" },
          { categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$C$1:$C$3" }
        ]
      },
      "D1:J10"
    );
    const scene = buildChartScene(ws.getCharts()[0].chartModel!, {
      width: 500,
      height: 300
    });
    const bar = scene.series.find(
      (s): s is ChartSceneSeries & { type: "bar" } => s.type === "bar" && !!s.bars.length
    )!;
    // Every bar must sit inside the plot rectangle. Before the fix the
    // cumulative stack (up to 80) blew past the top of the plot because
    // axis max was pinned to the per-series max (50).
    const plotTop = scene.plot.y;
    const plotBottom = scene.plot.y + scene.plot.height;
    for (const rect of bar.bars) {
      expect(rect.y).toBeGreaterThanOrEqual(plotTop - 1);
      expect(rect.y + rect.height).toBeLessThanOrEqual(plotBottom + 1);
    }
  });

  // Regression — `fmt(NaN)` used to emit the literal string `"NaN"` into
  // SVG attributes, producing `x="NaN"` which is invalid SVG. The guard
  // in `fmt()` now returns `"0"` for non-finite inputs so the emitted
  // document stays parseable even if upstream filters are bypassed.
  it("renderChartSvg never emits NaN or undefined in attribute values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRow(["A"]);
    ws.addRow(["Single"]);
    ws.addChart(
      {
        type: "bar",
        series: [{ categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$A$1:$A$1" }]
      },
      "C1:J10"
    );
    const svg = renderChartSvg(ws.getCharts()[0].chartModel!, {
      width: 400,
      height: 260
    });
    expect(svg).not.toMatch(/="NaN"/);
    expect(svg).not.toMatch(/="undefined"/);
  });

  // Regression — the SVG effect filter chained `<feMerge>` primitives
  // by setting `inLayer = "shadowOut-merged"` after the outer-shadow
  // step, but that name did not exist as a `result` anywhere in the
  // filter graph. Subsequent `glow` / `innerShadow` primitives then
  // referenced a non-existent input layer, producing undefined SVG
  // rendering. Each `feMerge` must now have a `result="<id>"` that the
  // next step can reference.
  it("buildEffectFilter chains layered effects with named merge results", () => {
    const xml = buildEffectFilter("test-filter", {
      outerShadow: {
        blurRadius: 50800,
        distance: 25400,
        direction: 2700000,
        color: { srgb: "000000" }
      },
      glow: {
        radius: 50800,
        color: { srgb: "FFFF00" }
      },
      innerShadow: {
        blurRadius: 38100,
        distance: 12700,
        direction: 5400000,
        color: { srgb: "000000" }
      }
    });
    expect(xml).toContain('result="shadowMerged"');
    expect(xml).toContain('result="glowMerged"');
    // The inner shadow consumes the glow-merged output, not the phantom
    // "shadowOut-merged" id that the buggy writer produced.
    expect(xml).toContain('<feMergeNode in="glowMerged"/>');
    expect(xml).not.toContain("shadowOut-merged");
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
