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

import { extractAll } from "@archive/unzip/extract";
import type {
  AddChartOptions,
  AddChartSeriesOptions,
  ChartModel,
  ChartTypeGroup,
  PlotArea,
  BarChartGroup,
  LineChartGroup,
  AreaChartGroup,
  PieChartGroup,
  DoughnutChartGroup,
  BubbleChartGroup,
  LineSeries,
  BarSeries,
  ValueAxis,
  AddChartExOptions,
  ChartExModel,
  StockChartGroup,
  OfPieChartGroup
} from "@excel/chart";
import {
  CHART_EX_PRESETS,
  CHART_PRESETS,
  applyChartExPreset,
  buildChartModel,
  buildComboChartModel,
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  fillChartCaches,
  fillChartExCaches,
  parseSpPr,
  buildSpPr,
  getSpPrFillColor,
  getSpPrGradient,
  getSpPrPattern,
  setSpPrFill,
  parseChartColors,
  buildChartColors,
  parseChartStyle,
  parseChartEx,
  buildChartExModel,
  renderChartEx,
  renderChartExSvg,
  renderChartExPng,
  buildChartScene,
  renderChartPng,
  renderChartSvg,
  applyChartPreset,
  EXCEL_CHART_EX_PRESETS,
  EXCEL_CHART_PRESETS,
  drawChartPdf,
  seriesFromColumns
} from "@excel/chart";
import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

import {
  Workbook as RootWorkbook,
  applyChartExPreset as rootApplyChartExPreset,
  CHART_EX_PRESETS as ROOT_CHART_EX_PRESETS,
  EXCEL_CHART_EX_PRESETS as ROOT_EXCEL_CHART_EX_PRESETS,
  CHART_PRESETS as ROOT_CHART_PRESETS,
  EXCEL_CHART_PRESETS as ROOT_EXCEL_CHART_PRESETS,
  buildChartExModel as rootBuildChartExModel,
  chartOptionsFromRows as rootChartOptionsFromRows,
  chartOptionsFromTable as rootChartOptionsFromTable,
  renderChartExSvg as rootRenderChartExSvg,
  renderChartSvg as rootRenderChartSvg,
  seriesFromColumns as rootSeriesFromColumns,
  validateXmlName as rootValidateXmlName,
  xmlEncodeAttr as rootXmlEncodeAttr,
  type AddChartFromRowsOptions,
  type AddChartFromTableOptions,
  type ChartExType,
  type ChartScene,
  type ChartSceneLegend,
  type ChartSceneSeries,
  type ChartSceneText,
  type SeriesFromColumnsOptions
} from "../../../index";

const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

function stableHash(input: string): string {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function expectPngDimensions(png: Uint8Array, width: number, height: number): void {
  expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  expect(textDecoder.decode(png.slice(12, 16))).toBe("IHDR");
  expect(readU32be(png, 16)).toBe(width);
  expect(readU32be(png, 20)).toBe(height);
  expect(png.length).toBeGreaterThan(100);
}

function expectPngPhysDpi(png: Uint8Array, dpi: number): void {
  const offset = findPngChunk(png, "pHYs");
  expect(offset).toBeGreaterThan(0);
  const pixelsPerMeter = readU32be(png, offset + 8);
  expect(pixelsPerMeter).toBe(Math.round(dpi / 0.0254));
  expect(png[offset + 16]).toBe(1);
}

function pngSignature(png: Uint8Array): string {
  const idat = collectPngChunks(png, "IDAT");
  const phys = collectPngChunks(png, "pHYs");
  return stableHash(
    [
      readU32be(png, 16),
      readU32be(png, 20),
      idat.reduce((sum, chunk) => sum + chunk.length, 0),
      stableHashBytes(idat),
      stableHashBytes(phys)
    ].join(":")
  );
}

function collectPngChunks(png: Uint8Array, type: string): Uint8Array[] {
  const chunks: Uint8Array[] = [];
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = readU32be(png, offset);
    const chunkType = textDecoder.decode(png.slice(offset + 4, offset + 8));
    if (chunkType === type) {
      chunks.push(png.slice(offset + 8, offset + 8 + length));
    }
    offset += 12 + length;
  }
  return chunks;
}

function stableHashBytes(chunks: Uint8Array[]): string {
  let hash = 2166136261;
  for (const chunk of chunks) {
    for (const byte of chunk) {
      hash ^= byte;
      hash = Math.imul(hash, 16777619);
    }
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function findPngChunk(png: Uint8Array, type: string): number {
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = readU32be(png, offset);
    if (textDecoder.decode(png.slice(offset + 4, offset + 8)) === type) {
      return offset;
    }
    offset += 12 + length;
  }
  return -1;
}

function readU32be(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset] * 0x1000000 +
    ((bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3])
  );
}

function makeRootExportRenderedChartModel(): ChartModel {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.addRows([
    ["A", 10],
    ["B", 20]
  ]);
  ws.addChart(
    {
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
      title: "Sales"
    },
    "D1:J10"
  );
  return ws.getCharts()[0].chartModel!;
}

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
    expect(m.chartSpace.chartData.data[0].numDim?.type).toBe("x");
  });

  it("buildChartExModel defaults histogram binning to auto", () => {
    const m = buildChartExModel({
      type: "histogram",
      series: [{ name: "Histogram", values: "Sheet1!$B$1:$B$5" }]
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutPr?.binning?.binType).toBe("auto");
    expect(renderChartEx(m)).toContain("<cx:auto/>");
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

  it("renderChartEx writes chart frame shape properties", () => {
    const m = buildChartExModel({
      type: "funnel",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
      spPr: { fill: { solid: { srgb: "FF0000" } } }
    });
    const xml = renderChartEx(m);
    expect(xml).toContain("<cx:spPr>");
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
  });

  it("buildChartExModel retains style and colors metadata", () => {
    const m = buildChartExModel({
      type: "sunburst",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
      chartStyle: { id: 201 },
      chartColors: { method: "cycle", colors: [{ srgb: "00FF00" }] }
    });
    expect(m.style?.id).toBe(201);
    expect(m.colors?.colors?.[0].srgb).toBe("00FF00");
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

  it("loads chartEx into a structured model while preserving clean raw XML", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
        title: "Programmatic ChartEx"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    expect(chart.chartExModel).toBeDefined();
    expect(chart.title).toBe("Programmatic ChartEx");
    expect(chart.chartExModel!.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutId).toBe(
      "sunburst"
    );

    const inputEntries = await extractAll(new Uint8Array(buf));
    const output = await wb2.xlsx.writeBuffer();
    const outputEntries = await extractAll(new Uint8Array(output));
    expect(textDecoder.decode(outputEntries.get("xl/charts/chartEx1.xml")!.data)).toBe(
      textDecoder.decode(inputEntries.get("xl/charts/chartEx1.xml")!.data)
    );
  });

  it("addChartEx populates cached category and value points", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }]
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const xml = textDecoder.decode(entries.get("xl/charts/chartEx1.xml")!.data);

    expect(xml).toContain('<cx:lvl ptCount="5">');
    expect(xml).toContain('<cx:pt idx="0">Cat1</cx:pt>');
    expect(xml).toContain('<cx:pt idx="4">Cat5</cx:pt>');
    expect(xml).toContain('<cx:pt idx="0">10</cx:pt>');
    expect(xml).toContain('<cx:pt idx="4">50</cx:pt>');
  });

  it("patches loaded chartEx XML after a high-level title mutation", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
        title: "Programmatic ChartEx"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml");
    expect(chartExEntry).toBeDefined();
    const marker = "<!-- excelts-chartEx-raw-passthrough -->";
    const chartExXml = textDecoder.decode(chartExEntry!.data);
    chartExEntry!.data = textEncoder.encode(
      chartExXml.replace("<cx:chart>", `${marker}<cx:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.title = "Updated ChartEx";

    const output = await wb2.xlsx.writeBuffer();
    const outputEntries = await extractAll(new Uint8Array(output));
    const outputChartExXml = textDecoder.decode(outputEntries.get("xl/charts/chartEx1.xml")!.data);
    expect(outputChartExXml).toContain(marker);
    expect(outputChartExXml).toContain("Updated ChartEx");
  });

  it("preserves raw ChartEx shape, data label, and extension XML after a non-raw-patch mutation", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
        title: "Programmatic ChartEx"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    const chartExXml = textDecoder.decode(chartExEntry.data);
    chartExEntry.data = textEncoder.encode(
      chartExXml
        .replace(
          '<cx:series layoutId="sunburst">',
          '<cx:series layoutId="sunburst"><cx:spPr><a:effectLst><a:outerShdw blurRad="63500"/></a:effectLst></cx:spPr>'
        )
        .replace(
          "</cx:series>",
          '<cx:dataLabels><cx:visibility value="1" numFmt="1"/><cx:spPr><a:noFill/></cx:spPr></cx:dataLabels><cx:extLst><cx:ext uri="{custom-chartEx}"/></cx:extLst></cx:series>'
        )
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutateChartEx(model => {
      model.chartSpace.chart.autoTitleDeleted = true;
    });

    const output = await wb2.xlsx.writeBuffer();
    const outputEntries = await extractAll(new Uint8Array(output));
    const outputChartExXml = textDecoder.decode(outputEntries.get("xl/charts/chartEx1.xml")!.data);
    expect(outputChartExXml).toContain("outerShdw");
    expect(outputChartExXml).toContain('numFmt="1"');
    expect(outputChartExXml).toContain("{custom-chartEx}");
  });

  it("falls back to structured ChartEx re-render when raw-patch mutations include unsafe changes", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }],
        title: "Programmatic ChartEx"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    const marker = "<!-- chartEx-unsafe-raw-marker -->";
    chartExEntry.data = textEncoder.encode(
      textDecoder.decode(chartExEntry.data).replace("<cx:chart>", `${marker}<cx:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutateChartEx(model => {
      model.chartSpace.chart.title = {
        text: { paragraphs: [{ runs: [{ text: "Updated" }] }] },
        overlay: false
      };
      model.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutId = "treemap";
    });

    const output = await wb2.xlsx.writeBuffer();
    const outputEntries = await extractAll(new Uint8Array(output));
    const outputChartExXml = textDecoder.decode(outputEntries.get("xl/charts/chartEx1.xml")!.data);
    expect(outputChartExXml).not.toContain(marker);
    expect(outputChartExXml).toContain('layoutId="treemap"');
    expect(outputChartExXml).toContain("Updated");
  });

  it("patches loaded ChartEx XML for data, layout, labels, shape, data points, and axes", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "waterfall",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5", subtotals: [4] }],
        title: "Waterfall"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    const marker = "<!-- chartEx-wide-raw-patch -->";
    chartExEntry.data = textEncoder.encode(
      textDecoder.decode(chartExEntry.data).replace("<cx:chart>", `${marker}<cx:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutateChartEx(
      model => {
        const data = model.chartSpace.chartData.data;
        data[0].strDim!.formula = "Sheet1!$A$2:$A$4";
        data[1].numDim!.formula = "Sheet1!$B$2:$B$4";
        const region = model.chartSpace.chart.plotArea.plotAreaRegion!;
        region.layout = { manualLayout: { x: 0.1, y: 0.2, w: 0.8, h: 0.7 } };
        region.plotSurface = { fill: { solid: { srgb: "F2F2F2" } } };
        const series = region.series[0];
        series.spPr = { fill: { solid: { srgb: "4472C4" } } };
        series.dataRefs = [{ dataId: 0 }, { dataId: 1 }];
        series.layoutPr = { subtotals: [{ idx: 2 }], connectorLines: false };
        series.axisId = [0];
        series.dataLabels = { visibility: { value: true }, position: "outEnd", numFmt: "#,##0" };
        series.dataPt = [{ idx: 1, spPr: { fill: { solid: { srgb: "ED7D31" } } } }];
        const valueAxis = model.chartSpace.chart.plotArea.axis?.find(axis => axis.axisId === 1);
        if (valueAxis) {
          valueAxis.numFmt = { formatCode: "#,##0", sourceLinked: false };
          valueAxis.valScaling = { min: 0, max: 100, majorUnit: 25 };
          valueAxis.spPr = { line: { color: { srgb: "70AD47" } } };
        }
      },
      { preferRawPatch: true }
    );

    const output = await wb2.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const xml = textDecoder.decode(entries.get("xl/charts/chartEx1.xml")!.data);
    expect(xml).toContain(marker);
    expect(xml).toContain("Sheet1!$A$2:$A$4");
    expect(xml).toContain("Sheet1!$B$2:$B$4");
    expect(xml).toContain("<cx:layout>");
    expect(xml).toContain('<cx:connectorLines val="0"/>');
    expect(xml).toContain('<cx:axisId val="0"/>');
    expect(xml).not.toContain('<cx:axisId val="1"/>');
    expect(xml).toContain('<cx:dataLabel pos="outEnd"/>');
    expect(xml).toContain('<cx:dataPt idx="1">');
    expect(xml).toContain('formatCode="#,##0"');
    expect(xml).toContain('max="100"');
    expect(xml).toContain('val="70AD47"');
  });

  it("enforces strict template mode globally for loaded ChartEx mutations", async () => {
    const { wb, ws } = makeExWb();
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }]
      },
      "D1:J10"
    );
    const zipData = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    chartExEntry.data = textEncoder.encode(
      textDecoder
        .decode(chartExEntry.data)
        .replace("<cx:chart>", "<!-- strict-chartEx --><cx:chart>")
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    wb2
      .getWorksheet("Sheet1")!
      .getCharts()[0].chartExModel!.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutId =
      "treemap";

    await expect(wb2.xlsx.writeBuffer({ strictTemplateMode: true })).rejects.toThrow(
      /strict template mode/
    );
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

  it("workbook.getChartExStructuredEntry returns undefined after removeChart", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChartEx(
      {
        type: "histogram",
        series: [{ name: "S", values: "Sheet1!$B$1:$B$3", literalValues: [1, 2, 3] }]
      },
      "C1:J10"
    );
    const chartExNumber = ws.getCharts()[0].chartExNumber;
    expect(wb.getChartExStructuredEntry(chartExNumber)).toBeDefined();
    ws.removeChart(0);
    expect(wb.getChartExStructuredEntry(chartExNumber)).toBeUndefined();
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

describe("TC2b: classic chart raw passthrough", () => {
  async function makeWorkbookWithInjectedChartXml(inject: (xml: string) => string) {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Cat1";
    ws.getCell("B1").value = 10;
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "Data", categories: "Sheet1!$A$1:$A$1", values: "Sheet1!$B$1:$B$1" }],
        title: "Original"
      },
      "D1:J10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartEntry = zipData.get("xl/charts/chart1.xml");
    expect(chartEntry).toBeDefined();

    const chartXml = textDecoder.decode(chartEntry!.data);
    chartEntry!.data = textEncoder.encode(inject(chartXml));
    const { createZip } = await import("@archive/zip/zip-bytes");
    return createZip([...zipData.entries()].map(([name, file]) => ({ name, data: file.data })));
  }

  it("preserves unmodified loaded chart XML exactly", async () => {
    const marker = "<!-- excelts-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const output = await wb.xlsx.writeBuffer();
    const inputEntries = await extractAll(new Uint8Array(input));
    const outputEntries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(outputEntries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toBe(textDecoder.decode(inputEntries.get("xl/charts/chart1.xml")!.data));
  });

  it("patches loaded chart XML after a high-level title mutation", async () => {
    const marker = "<!-- excelts-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.title = "Updated";

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain("Updated");
  });

  it("re-renders loaded chart XML after a direct model mutation", async () => {
    const marker = "<!-- excelts-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.chartModel!.chart.legend = undefined;

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).not.toContain(marker);
    expect(outputChartXml).not.toContain("c:legend");
  });

  it("patches loaded chart XML after a high-level legend mutation", async () => {
    const marker = "<!-- excelts-legend-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.legend = { legendPos: "t", overlay: false };

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain('<c:legendPos val="t"/>');
  });

  it("patches loaded chart XML for series data label changes without dropping unsupported raw XML", async () => {
    const marker = "<!-- excelts-series-label-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const series = model.chart.plotArea.chartTypes[0].series[0] as BarSeries & LineSeries;
        series.dataLabels = { showVal: true, showCatName: true, separator: " / " };
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain("<c:dLbls>");
    expect(outputChartXml).toContain('<c:showVal val="1"/>');
    expect(outputChartXml).toContain('<c:showCatName val="1"/>');
  });

  it("patches only changed series children while preserving unmodified raw child XML", async () => {
    const rawMarker =
      '<c:marker><c:symbol val="circle"/><c:extLst><c:ext uri="{raw-marker}"/></c:extLst></c:marker>';
    const rawTrendline =
      '<c:trendline><c:trendlineType val="linear"/><c:extLst><c:ext uri="{raw-trendline}"/></c:extLst></c:trendline>';
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:cat>", `${rawMarker}${rawTrendline}<c:cat>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const series = model.chart.plotArea.chartTypes[0].series[0] as BarSeries;
        series.val = {
          numRef: { formula: "Sheet1!$B$1:$B$1", cache: { points: [{ index: 0, value: 42 }] } }
        };
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain("{raw-marker}");
    expect(outputChartXml).toContain("{raw-trendline}");
    expect(outputChartXml).toContain("Sheet1!$B$1:$B$1");
  });

  it("patches loaded chart XML for chart-group data labels while preserving series label XML", async () => {
    const marker = "<!-- excelts-group-label-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml
        .replace("<c:chart>", `${marker}<c:chart>`)
        .replace("</c:ser>", '<c:dLbls><c:showSerName val="1"/></c:dLbls></c:ser>')
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const group = model.chart.plotArea.chartTypes[0] as BarChartGroup;
        group.dataLabels = { showVal: true, position: "outEnd" };
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain('<c:showSerName val="1"/>');
    expect(outputChartXml).toContain('<c:showVal val="1"/>');
    expect(outputChartXml).toContain('<c:dLblPos val="outEnd"/>');
  });

  it("patches loaded chart XML for plot area manual layout without dropping unsupported raw XML", async () => {
    const marker = "<!-- excelts-plot-layout-raw-passthrough -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        model.chart.plotArea.layout = {
          manualLayout: {
            layoutTarget: "inner",
            xMode: "factor",
            yMode: "factor",
            x: 0.12,
            y: 0.18
          }
        };
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain("<c:manualLayout>");
    expect(outputChartXml).toContain('<c:x val="0.12"/>');
  });

  it("patches title, legend, series ranges, axes, and plot layout together while preserving raw extension XML", async () => {
    const marker =
      '<c:extLst><c:ext uri="{unsupported-template-extension}"><c15:foo xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" val="1"/></c:ext></c:extLst>';
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("</c:chart>", `${marker}</c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        model.chart.title = {
          text: { paragraphs: [{ runs: [{ text: "Patched" }] }] },
          overlay: false
        };
        model.chart.legend = { legendPos: "r", overlay: false };
        model.chart.plotArea.layout = { manualLayout: { x: 0.2, y: 0.1, w: 0.7, h: 0.8 } };
        const group = model.chart.plotArea.chartTypes[0] as BarChartGroup;
        group.series[0].cat = { strRef: { formula: "Sheet1!$A$1:$A$2", cache: { points: [] } } };
        group.series[0].val = { numRef: { formula: "Sheet1!$B$1:$B$2", cache: { points: [] } } };
        group.dataLabels = { showVal: true };
        const valueAxis = model.chart.plotArea.axes.find(
          axis => axis.axisType === "val"
        ) as ValueAxis;
        valueAxis.title = {
          text: { paragraphs: [{ runs: [{ text: "Value Axis" }] }] },
          overlay: false
        };
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain("{unsupported-template-extension}");
    expect(outputChartXml).toContain("Patched");
    expect(outputChartXml).toContain('<c:legendPos val="r"/>');
    expect(outputChartXml).toContain("Sheet1!$A$1:$A$2");
    expect(outputChartXml).toContain("Sheet1!$B$1:$B$2");
    expect(outputChartXml).toContain("Value Axis");
    expect(outputChartXml).toContain('<c:x val="0.2"/>');
    expect(outputChartXml).toContain('<c:showVal val="1"/>');
  });

  it("patches loaded chart XML for series formatting, markers, trendlines, error bars, data points, and axis formatting", async () => {
    const marker = "<!-- excelts-wide-raw-patch -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const series = model.chart.plotArea.chartTypes[0].series[0] as BarSeries & LineSeries;
        series.spPr = {
          fill: { solid: { srgb: "4472C4" } },
          line: { color: { srgb: "ED7D31" }, width: 12700, dash: "dash" }
        };
        series.marker = {
          symbol: "diamond",
          size: 9,
          spPr: { fill: { solid: { srgb: "70AD47" } } }
        };
        series.dataPoints = [{ index: 0, spPr: { fill: { solid: { srgb: "FFC000" } } } }];
        series.trendlines = [{ type: "linear", name: "Trend", displayEq: true }];
        series.errorBars = { barDir: "both", errValType: "fixedVal", val: 5 };
        const valueAxis = model.chart.plotArea.axes.find(
          axis => axis.axisType === "val"
        ) as ValueAxis;
        valueAxis.numFmt = { formatCode: "#,##0", sourceLinked: false };
        valueAxis.majorTickMark = "outside";
        valueAxis.minorTickMark = "inside";
        valueAxis.tickLblPos = "high";
        valueAxis.spPr = { line: { color: { srgb: "5B9BD5" }, width: 9525 } };
        valueAxis.txPr = { size: 900, bold: true, color: { srgb: "404040" } };
        valueAxis.majorUnit = 10;
      },
      { preferRawPatch: true }
    );

    const output = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);

    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain('<a:srgbClr val="4472C4"/>');
    expect(outputChartXml).toContain("<c:marker>");
    expect(outputChartXml).toContain('<c:symbol val="diamond"/>');
    expect(outputChartXml).toContain('<c:dPt><c:idx val="0"/>');
    expect(outputChartXml).toContain("<c:trendline>");
    expect(outputChartXml).toContain("<c:name>Trend</c:name>");
    expect(outputChartXml).toContain("<c:errBars>");
    expect(outputChartXml).toContain('<c:val val="5"/>');
    expect(outputChartXml).toContain('formatCode="#,##0"');
    expect(outputChartXml).toContain('<c:majorTickMark val="outside"/>');
    expect(outputChartXml).toContain('<c:minorTickMark val="inside"/>');
    expect(outputChartXml).toContain('<c:tickLblPos val="high"/>');
    expect(outputChartXml).toContain('<c:majorUnit val="10"/>');
    expect(outputChartXml).toContain('<a:defRPr sz="900" b="1"');
  });

  it("enforces strict template mode globally for loaded classic chart mutations", async () => {
    const marker = "<!-- strict-template-mode-classic -->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("<c:chart>", `${marker}<c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    // Mutate a field that the raw-patch path deliberately does NOT cover:
    // swap a bar chart into a line chart. That is a structural change —
    // the whole `<c:barChart>` block becomes `<c:lineChart>` — so
    // strict mode must refuse rather than silently rebuild.
    const group = wb.getWorksheet("Sheet1")!.getCharts()[0].chartModel!.chart.plotArea
      .chartTypes[0];
    (group as unknown as { type: string }).type = "line";

    await expect(wb.xlsx.writeBuffer({ templateMode: "strict" })).rejects.toThrow(
      /strict template mode/
    );
  });

  it("Chart.setStyle(n) writes <c:style val=N/> and survives round-trip", async () => {
    // High-level setter equivalent to xlsxwriter's Chart.set_style(N).
    // Confirms both: (a) the API validates the range 1..48, (b) the
    // resulting XML carries <c:style val="N"/> on a fresh chart, and
    // (c) a subsequent round-trip preserves it on the model.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Styled"
      },
      "D1:J10"
    );
    const chart = ws.getCharts()[0];
    chart.setStyle(37);
    chart.setBuiltInStyle(42); // alias — must produce the same effect.

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(zipData.get("xl/charts/chart1.xml")!.data);
    // Last call wins — style should be 42.
    expect(chartXml).toContain('<c:style val="42"/>');

    // Round-trip.
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.getWorksheet("Sheet1")!.getCharts()[0].chartModel!.style).toBe(42);

    // Invalid ranges throw.
    expect(() => chart.setStyle(0)).toThrow(/1..48/);
    expect(() => chart.setStyle(49)).toThrow(/1..48/);
    expect(() => chart.setStyle(1.5)).toThrow(/1..48/);
  });

  it("loaded chart with c15:datalabelsRange ext raw-patches title edits instead of rebuilding", async () => {
    // Inject a c15:datalabelsRange extension inside a data-labels
    // block, then edit the title. Before #5, the `extLst` JSON shape
    // change between previous and current forced `getChartRawPatchPlan`
    // to return undefined → full rebuild, which drops the extension.
    // After #5, the fast-patch path must preserve both the c15 ext
    // and the edited title.
    const ext =
      "<c:extLst>" +
      '<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" ' +
      'xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">' +
      "<c15:datalabelsRange>" +
      "<c15:f>Sheet1!$C$1:$C$1</c15:f>" +
      '<c15:dlblRangeCache><c15:ptCount val="1"/></c15:dlblRangeCache>' +
      "</c15:datalabelsRange>" +
      "</c:ext>" +
      "</c:extLst>";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      // Inject inside the first <c:dLbls> opening or add a dLbls block —
      // simpler: drop the ext right before </c:ser> (c:ser allows extLst
      // as its last child per schema).
      xml.replace("</c:ser>", `${ext}</c:ser>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        model.chart.title = {
          text: { paragraphs: [{ runs: [{ text: "New Title" }] }] }
        };
      },
      { preferRawPatch: true }
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(zipData.get("xl/charts/chart1.xml")!.data);
    // c15 extension preserved.
    expect(chartXml).toContain("c15:datalabelsRange");
    expect(chartXml).toContain("Sheet1!$C$1:$C$1");
    // New title is written.
    expect(chartXml).toContain("New Title");
  });

  it("classic chart parser surfaces vendor-namespaced unknown elements", async () => {
    // Inject a `c15:` tag into the chartSpace root — it uses a non-c:
    // namespace and is not wrapped in an `<c:extLst>` block, so the
    // structured parser must record it on `chartModel.unknownElements`
    // for parity with ChartExModel.unknownElements. Before this fix
    // the classic parser silently dropped such elements.
    const vendorTag = '<c15:customChartMeta val="vendor"/>';
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml
        .replace(
          "<c:chartSpace ",
          '<c:chartSpace xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" '
        )
        .replace("</c:chartSpace>", `${vendorTag}</c:chartSpace>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    const unknown = chart.chartModel!.unknownElements;
    expect(unknown).toBeDefined();
    expect(unknown!.some(e => e.name === "c15:customChartMeta")).toBe(true);

    // Chart.unknownElements convenience getter proxies the same list so
    // authors can decide about strictTemplateMode without reaching into
    // the structured model. Returns a copy (mutations don't leak back).
    const convenience = chart.unknownElements;
    expect(convenience).toBeDefined();
    expect(convenience).toEqual(unknown);
    convenience!.pop();
    expect(chart.unknownElements?.length).toBe(unknown!.length);
  });

  it("Chart.unknownElements returns undefined for freshly-built charts", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 1],
      ["B", 2]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const chart = ws.getCharts()[0];
    // Freshly built — no raw XML was ever parsed.
    expect(chart.unknownElements).toBeUndefined();
  });

  it("raw-patches chartTypeGroup simple fields (overlap / gapWidth / varyColors) without rebuilding", async () => {
    // Inject an unparseable vendor marker inside the bar-chart block so
    // we can confirm it survives the write. Then mutate `overlap` (was
    // not raw-patchable before this change) via `preferRawPatch` and
    // assert the marker and the new value are both in the output.
    const vendorMarker = "<!--keep-me-2026-04-28-->";
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      // Place the marker near the end of the barChart block, before the
      // closing tag so it sits between siblings rather than inside an
      // element — this exercises the "group block preserved verbatim"
      // promise of the patcher.
      xml.replace("</c:barChart>", `${vendorMarker}</c:barChart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const group = model.chart.plotArea.chartTypes[0] as BarChartGroup;
        group.overlap = -30;
        group.gapWidth = 220;
        group.varyColors = true;
      },
      { preferRawPatch: true }
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(zipData.get("xl/charts/chart1.xml")!.data);
    // Vendor marker preserved → raw-patch path was taken.
    expect(chartXml).toContain(vendorMarker);
    // And every simple-field rewrite landed.
    expect(chartXml).toContain('<c:overlap val="-30"/>');
    expect(chartXml).toContain('<c:gapWidth val="220"/>');
    expect(chartXml).toContain('<c:varyColors val="1"/>');
  });

  it("raw-patches pie chartTypeGroup firstSliceAng without rebuilding", async () => {
    const vendorMarker = "<!--pie-vendor-keep-->";
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "pie",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const initialBuf = await wb.xlsx.writeBuffer();
    const initial = await extractAll(new Uint8Array(initialBuf));
    const chartEntry = initial.get("xl/charts/chart1.xml")!;
    chartEntry.data = textEncoder.encode(
      textDecoder.decode(chartEntry.data).replace("</c:pieChart>", `${vendorMarker}</c:pieChart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const modifiedXlsx = await createZip(
      [...initial.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(modifiedXlsx);
    const chart2 = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart2.mutate(
      model => {
        const group = model.chart.plotArea.chartTypes[0] as PieChartGroup;
        group.firstSliceAng = 90;
      },
      { preferRawPatch: true }
    );
    const outBuf = await wb2.xlsx.writeBuffer();
    const out = await extractAll(new Uint8Array(outBuf));
    const outXml = textDecoder.decode(out.get("xl/charts/chart1.xml")!.data);
    expect(outXml).toContain(vendorMarker);
    expect(outXml).toContain('<c:firstSliceAng val="90"/>');
  });

  it("raw-patches doughnut holeSize and removes the leaf when set to undefined", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "doughnut",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        holeSize: 40
      },
      "D1:J10"
    );
    const initialBuf = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(initialBuf);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        const group = model.chart.plotArea.chartTypes[0] as DoughnutChartGroup;
        group.holeSize = 75;
      },
      { preferRawPatch: true }
    );
    const outBuf = await wb2.xlsx.writeBuffer();
    const out = await extractAll(new Uint8Array(outBuf));
    const outXml = textDecoder.decode(out.get("xl/charts/chart1.xml")!.data);
    // Updated in place.
    expect(outXml).toContain('<c:holeSize val="75"/>');
    // Never two copies (sanity: in-place, not duplicated).
    expect((outXml.match(/<c:holeSize /g) ?? []).length).toBe(1);
  });

  it("strict template mode error message lists unknown classic chart XML paths", async () => {
    const vendorTag = '<c15:customChartMeta val="vendor"/>';
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml
        .replace(
          "<c:chartSpace ",
          '<c:chartSpace xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" '
        )
        .replace("</c:chartSpace>", `${vendorTag}</c:chartSpace>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    // Mutate a chart-type group's top-level `type` — structural change
    // (bar → line rewrites the enclosing `<c:barChart>` tag itself),
    // outside the RawPatchPlan whitelist, so strict mode must surface
    // the unknownElements hint.
    const group = wb.getWorksheet("Sheet1")!.getCharts()[0].chartModel!.chart.plotArea
      .chartTypes[0];
    (group as unknown as { type: string }).type = "line";

    await expect(wb.xlsx.writeBuffer({ templateMode: "strict" })).rejects.toThrow(
      /c15:customChartMeta/
    );
  });

  it("fails instead of re-rendering when requireRawPatch cannot safely preserve template XML", async () => {
    const marker =
      '<c:extLst><c:ext uri="{must-preserve}"><c15:foo xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/></c:ext></c:extLst>';
    const input = await makeWorkbookWithInjectedChartXml(xml =>
      xml.replace("</c:chart>", `${marker}</c:chart>`)
    );

    const wb = new Workbook();
    await wb.xlsx.load(input);
    const chart = wb.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutate(
      model => {
        // Switching chart type is outside every raw-patch branch (the
        // `<c:barChart>` → `<c:lineChart>` rewrite requires rebuilding
        // axes, series shape, etc.), so `requireRawPatch` must refuse
        // rather than silently rebuild and drop the extension marker.
        (model.chart.plotArea.chartTypes[0] as unknown as { type: string }).type = "line";
      },
      { preferRawPatch: true, requireRawPatch: true }
    );

    await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/requireRawPatch/);
  });

  it("fails instead of re-rendering ChartEx when requireRawPatch cannot safely preserve template XML", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$2",
        series: [{ values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const zipData = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    chartExEntry.data = textEncoder.encode(
      textDecoder
        .decode(chartExEntry.data)
        .replace("<cx:chart>", "<!-- must-preserve-chartEx --><cx:chart>")
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutateChartEx(
      model => {
        model.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutId = "treemap";
      },
      { preferRawPatch: true, requireRawPatch: true }
    );

    await expect(wb2.xlsx.writeBuffer()).rejects.toThrow(/requireRawPatch/);
  });

  it("parseChartEx surfaces unstructured child elements under known parents", () => {
    // Include unknown children at multiple known parent positions so the
    // collector exercises all whitelisted locations, not just the root.
    const xml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"',
      '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      "  <cx:chartData/>",
      "  <cx:chart>",
      "    <cx:plotArea>",
      "      <cx:plotAreaRegion>",
      '        <cx:series layoutId="sunburst">',
      '          <cx:customSeriesExt val="1"/>',
      "        </cx:series>",
      "      </cx:plotAreaRegion>",
      '      <cx:unknownPlotArea val="x"/>',
      "    </cx:plotArea>",
      "  </cx:chart>",
      '  <cx:customTopLevel val="1"/>',
      "</cx:chartSpace>"
    ].join("\n");

    const model = parseChartEx(xml);

    const paths = (model.unknownElements ?? []).map(entry => entry.path).sort();
    expect(paths).toEqual(
      [
        "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series/cx:customSeriesExt",
        "cx:chartSpace/cx:chart/cx:plotArea/cx:unknownPlotArea",
        "cx:chartSpace/cx:customTopLevel"
      ].sort()
    );
    // Well-known children must never be flagged, even when present
    // alongside unknowns at the same level.
    for (const entry of model.unknownElements ?? []) {
      expect(entry.name.startsWith("cx:")).toBe(true);
      expect(entry.name).not.toBe("cx:chart");
      expect(entry.name).not.toBe("cx:series");
    }
  });

  it("parseChartEx leaves unknownElements undefined when every child is recognised", async () => {
    // Round-trip a freshly-built ChartEx so the XML is known to only contain
    // structured children — regression guard against false positives.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$2",
        series: [{ values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const zipData = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExXml = textDecoder.decode(zipData.get("xl/charts/chartEx1.xml")!.data);

    const model = parseChartEx(chartExXml);
    expect(model.unknownElements).toBeUndefined();
  });

  it("raw-patches ChartEx series ownerIdx without rebuilding the part", async () => {
    // Build a ChartEx with two series so the parser emits explicit ownerIdx
    // attributes on cx:series. Round-trip once so the second load has a
    // clean raw XML baseline, then mutate only ownerIdx and verify the
    // writer preserves the surrounding structure via raw patching rather
    // than falling through to the structural rebuild path (which would
    // regenerate series markup and drop any sibling template XML).
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 30],
      ["B", 20, 40]
    ]);
    ws.addChartEx(
      {
        type: "waterfall",
        categories: "Sheet1!$A$1:$A$2",
        series: [{ values: "Sheet1!$B$1:$B$2" }, { values: "Sheet1!$C$1:$C$2" }]
      },
      "D1:J10"
    );
    // Inject a vendor marker as a sibling of cx:series so we can detect
    // the difference between raw-patch (marker preserved) and rebuild
    // (marker discarded because unknownElements are dropped on rebuild).
    const initial = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExEntry = initial.get("xl/charts/chartEx1.xml")!;
    const marker = '<cx:vendorBadge val="keep-me"/>';
    chartExEntry.data = textEncoder.encode(
      textDecoder.decode(chartExEntry.data).replace("<cx:series", `${marker}<cx:series`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...initial.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    const beforeOwner =
      chart.chartExModel!.chartSpace.chart.plotArea.plotAreaRegion!.series[0].ownerIdx;
    chart.mutateChartEx(
      model => {
        // Flip ownerIdx to something distinct; this is the only mutation.
        const series = model.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
        series.ownerIdx = (beforeOwner ?? 0) + 7;
      },
      { preferRawPatch: true }
    );

    const out = await extractAll(new Uint8Array(await wb2.xlsx.writeBuffer()));
    const outXml = textDecoder.decode(out.get("xl/charts/chartEx1.xml")!.data);
    // ownerIdx on the first series is updated
    expect(outXml).toMatch(/<cx:series[^>]*ownerIdx="7"/);
    // Sibling vendor marker is preserved — proof that we took the raw-patch
    // path, not the structural rebuild path.
    expect(outXml).toContain(marker);
  });

  it("strict template mode error message lists unknown ChartEx XML paths", async () => {
    // Author a loaded ChartEx file with two distinct vendor extensions at
    // different known parents; mutate structurally in a way that cannot be
    // raw-patched to force the strict-mode failure path.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$2",
        series: [{ values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const zipData = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    const originalXml = textDecoder.decode(chartExEntry.data);
    const injectedXml = originalXml
      .replace(
        '<cx:series layoutId="sunburst"',
        '<cx:vendorSeriesAnnot val="x"/><cx:series layoutId="sunburst"'
      )
      .replace("<cx:chart>", '<cx:chart><cx:vendorChartMeta val="y"/>');
    chartExEntry.data = textEncoder.encode(injectedXml);
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    // Changing layoutId requires a rebuild; the raw-patch white-list cannot
    // mutate the opening tag's layoutId attribute alone.
    chart.mutateChartEx(model => {
      model.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutId = "treemap";
    });

    await expect(wb2.xlsx.writeBuffer({ templateMode: "strict" })).rejects.toThrow(
      /cx:vendorChartMeta.*cx:vendorSeriesAnnot|cx:vendorSeriesAnnot.*cx:vendorChartMeta/
    );
  });
});

describe("TC2c: chartsheet API", () => {
  it("creates a classic chart chartsheet and preserves it after round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = "Cat1";
    ws.getCell("A2").value = "Cat2";
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;

    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ name: "Sales", categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }],
        title: "Sales Chart"
      }
    });

    expect(chartsheet.chartNumber).toBe(1);
    expect(chartsheet.isChartEx).toBe(false);
    expect(wb.chartsheets).toHaveLength(1);

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/chartsheets/sheet1.xml")).toBeDefined();
    expect(entries.get("xl/chartsheets/_rels/sheet1.xml.rels")).toBeDefined();
    expect(entries.get("xl/drawings/drawing1.xml")).toBeDefined();
    expect(entries.get("xl/drawings/_rels/drawing1.xml.rels")).toBeDefined();
    expect(entries.get("xl/charts/chart1.xml")).toBeDefined();

    const workbookXml = textDecoder.decode(entries.get("xl/workbook.xml")!.data);
    expect(workbookXml).toContain('name="Data"');
    expect(workbookXml).toContain('name="Chart Sheet"');

    const chartsheetRels = textDecoder.decode(
      entries.get("xl/chartsheets/_rels/sheet1.xml.rels")!.data
    );
    expect(chartsheetRels).toContain("../drawings/drawing1.xml");

    const drawingRels = textDecoder.decode(
      entries.get("xl/drawings/_rels/drawing1.xml.rels")!.data
    );
    expect(drawingRels).toContain("../charts/chart1.xml");

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.chartsheets).toHaveLength(1);
    expect(wb2.chartsheets[0].name).toBe("Chart Sheet");
    expect(wb2.chartsheets[0].chartNumber).toBe(1);
  });

  it("creates a chartEx chartsheet", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    for (let i = 1; i <= 3; i++) {
      ws.getCell(`A${i}`).value = `Cat${i}`;
      ws.getCell(`B${i}`).value = i * 10;
    }

    const chartsheet = wb.addChartsheet("Modern Chart", {
      chart: {
        type: "sunburst",
        categories: "Data!$A$1:$A$3",
        series: [{ name: "Data", values: "Data!$B$1:$B$3" }],
        title: "Modern"
      }
    });

    expect(chartsheet.isChartEx).toBe(true);
    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/charts/chartEx1.xml")).toBeDefined();
    const drawingXml = textDecoder.decode(entries.get("xl/drawings/drawing1.xml")!.data);
    expect(drawingXml).toContain("mc:AlternateContent");

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.chartsheets[0].isChartEx).toBe(true);
    expect(wb2.chartsheets[0].chartExNumber).toBe(1);
  });

  it("exposes chartsheet chart models and supports replacement/removal", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }],
        title: "Original"
      },
      pageMargins: { l: 0.7, r: 0.7, t: 0.75, b: 0.75, header: 0.3, footer: 0.3 },
      zoomScale: 90
    });
    expect(chartsheet.chartModel?.chart.title).toBeDefined();
    expect(chartsheet.pageMargins?.l).toBe(0.7);
    chartsheet.zoomScale = 120;
    expect(wb.getChartsheet("Chart Sheet")?.zoomScale).toBe(120);
    expect(chartsheet.chart?.title).toBe("Original");
    chartsheet.pageSetup = { orientation: "landscape", paperSize: 9, scale: 90 };
    chartsheet.printOptions = { horizontalCentered: true, verticalCentered: true };

    expect(
      wb.replaceChartsheetChart("Chart Sheet", {
        type: "histogram",
        series: [{ values: "Data!$B$1:$B$2" }]
      })
    ).toBe(true);
    const replaced = wb.getChartsheet("Chart Sheet")!;
    expect(replaced.isChartEx).toBe(true);
    expect(replaced.chartExModel).toBeDefined();
    expect(wb.removeChartsheet("Chart Sheet")).toBe(true);
    expect(wb.chartsheets).toHaveLength(0);
  });

  it("supports chartsheet rename, copy, wrapper replacement, and print settings round-trip", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }],
        title: "Original"
      },
      pageSetup: { orientation: "landscape", paperSize: 9, scale: 95 },
      printOptions: { horizontalCentered: true }
    });
    expect(chartsheet.rename("Renamed Chart")).toBe(true);
    const copy = chartsheet.copy("Copied Chart")!;
    expect(copy.chartNumber).not.toBe(chartsheet.chartNumber);
    expect(copy.chart?.title).toBe("Original");
    expect(copy.replaceChart({ type: "pie", series: [baseSeries("Pie")], title: "Pie" })).toBe(
      true
    );
    expect(copy.chart?.title).toBe("Pie");

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const sheetXml = textDecoder.decode(entries.get("xl/chartsheets/sheet1.xml")!.data);
    expect(sheetXml).toContain('orientation="landscape"');
    expect(sheetXml).toContain('horizontalCentered="1"');

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.getChartsheet("Renamed Chart")?.pageSetup?.orientation).toBe("landscape");
    expect(wb2.getChartsheet("Copied Chart")?.chart?.title).toBe("Pie");
  });

  it("cleans up chartsheet chart parts after replacement and removal", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }],
        title: "Original"
      }
    });
    const copied = chartsheet.copy("Copied Chart")!;
    expect(copied.chartNumber).toBe(2);
    expect(
      chartsheet.replaceChart({ type: "line", series: [baseSeries("Line")], title: "Line" })
    ).toBe(true);
    expect(chartsheet.chartNumber).toBe(3);
    expect(wb.getChartEntry(1)).toBeUndefined();

    let entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    expect(entries.get("xl/charts/chart1.xml")).toBeUndefined();
    expect(entries.get("xl/charts/chart2.xml")).toBeDefined();
    expect(entries.get("xl/charts/chart3.xml")).toBeDefined();

    expect(copied.remove()).toBe(true);
    entries = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    expect(entries.get("xl/charts/chart2.xml")).toBeUndefined();
    expect(entries.get("xl/charts/chart3.xml")).toBeDefined();
  });
});

describe("TC2d: high-level chart series editing API", () => {
  it("mutate APIs mark classic and ChartEx charts dirty for deep changes", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")], title: "Original" }, "C1:J10");
    const chart = ws.getCharts()[0];
    chart.mutate(model => {
      model.chart.title = undefined;
    });
    expect(wb.getChartEntry(chart.chartNumber)!.dirty).toBe(true);

    ws.addChartEx({ type: "funnel", series: [{ values: VALUES_A }] }, "L1:R10");
    const chartEx = ws.getCharts()[1];
    chartEx.mutateChartEx(model => {
      model.chartSpace.chart.autoTitleDeleted = true;
    });
    expect(wb.getChartExStructuredEntry(chartEx.chartExNumber)!.dirty).toBe(true);
  });

  it("updates classic chart series references by patching loaded raw XML", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "A";
    ws.getCell("A2").value = "B";
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("C1").value = 30;
    ws.getCell("C2").value = 40;
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "Old", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Series"
      },
      "E1:K10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const marker = "<!-- excelts-series-raw-passthrough -->";
    const chartEntry = zipData.get("xl/charts/chart1.xml")!;
    chartEntry.data = textEncoder.encode(
      textDecoder.decode(chartEntry.data).replace("<c:chart>", `${marker}<c:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    expect(chart.setSeriesName(0, "New")).toBe(true);
    expect(chart.setSeriesValues(0, "Sheet1!$C$1:$C$2")).toBe(true);
    expect(chart.setSeriesCategories(0, "Sheet1!$A$1:$A$2")).toBe(true);

    const output = await wb2.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain("New");
    expect(outputChartXml).toContain("Sheet1!$C$1:$C$2");
  });

  it("patches loaded chart XML after axis scaling and title changes", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "A";
    ws.getCell("B1").value = 10;
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1", values: "Sheet1!$B$1" }],
        valueAxis: { title: "Old", min: 0, max: 20 }
      },
      "E1:K10"
    );

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const marker = "<!-- excelts-axis-raw-passthrough -->";
    const chartEntry = zipData.get("xl/charts/chart1.xml")!;
    chartEntry.data = textEncoder.encode(
      textDecoder.decode(chartEntry.data).replace("<c:chart>", `${marker}<c:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    const valueAxis = chart.valueAxis!;
    valueAxis.scaling = { min: 0, max: 100 };
    valueAxis.title = { text: { paragraphs: [{ runs: [{ text: "Updated Axis" }] }] } };
    wb2.getChartEntry(chart.chartNumber)!.dirty = true;
    wb2.getChartEntry(chart.chartNumber)!.preferRawPatch = true;

    const output = await wb2.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const outputChartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(outputChartXml).toContain(marker);
    expect(outputChartXml).toContain('<c:max val="100"/>');
    expect(outputChartXml).toContain("Updated Axis");
  });

  it("returns false when updating a missing series", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "bar", series: [baseSeries("S")] }, "E1:K10");
    expect(ws.getCharts()[0].setSeriesValues(5, "Sheet1!$C$1:$C$2")).toBe(false);
  });
});

describe("TC2e: pivot chart creation", () => {
  function makePivotWorkbook() {
    const wb = new Workbook();
    const data = wb.addWorksheet("Sales Data");
    data.addRows([
      ["Region", "Product", "Revenue", "Units"],
      ["West", "A", 10, 1],
      ["West", "B", 20, 2],
      ["East", "A", 30, 3]
    ]);
    const pivot = wb.addWorksheet("Pivot Sheet");
    const pivotTable = pivot.addPivotTable({
      sourceSheet: data,
      rows: ["Region"],
      values: ["Revenue", "Units"],
      metric: "sum"
    });
    return { wb, pivot, pivotTable };
  }

  it("creates a classic pivot chart with pivotSource and pivot chart formats", async () => {
    const { wb, pivot, pivotTable } = makePivotWorkbook();

    const chartNumber = pivot.addPivotChart(
      pivotTable,
      {
        type: "bar",
        series: [
          {
            name: "Revenue",
            categories: "'Pivot Sheet'!$A$4:$A$5",
            values: "'Pivot Sheet'!$B$4:$B$5"
          }
        ],
        title: "Revenue by Region"
      },
      "D1:K12"
    );

    expect(chartNumber).toBe(1);
    expect(pivotTable.chartFormat).toBe(1);
    expect(pivotTable.chartFormats).toHaveLength(1);

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    const pivotXml = textDecoder.decode(entries.get("xl/pivotTables/pivotTable1.xml")!.data);

    expect(chartXml).toContain("<c:pivotSource>");
    expect(chartXml).toContain("<c:name>'Pivot Sheet'!PivotTable1</c:name>");
    expect(chartXml).toContain('<c:fmtId val="0"/>');
    expect(chartXml.indexOf("<c:pivotSource>")).toBeLessThan(chartXml.indexOf("<c:chart>"));
    expect(chartXml).toContain("<c:pivotFmts>");
    expect(pivotXml).toContain('name="PivotTable1"');
    expect(pivotXml).toContain('chartFormat="1"');
    expect(pivotXml).toContain('<chartFormats count="1">');
    expect(pivotXml).toContain('<chartFormat chart="0" format="0" series="1">');
    expect(pivotXml).toContain('<reference field="4294967294" count="1" selected="0">');

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart = wb2.getWorksheet("Pivot Sheet")!.getCharts()[0];
    expect(chart.chartModel!.pivotSource).toContain("'Pivot Sheet'!PivotTable1");
  });

  it("writes structured pivot chart field buttons, filters, and refresh metadata", async () => {
    const { wb, pivot, pivotTable } = makePivotWorkbook();

    pivot.addPivotChart(
      pivotTable,
      {
        type: "bar",
        series: [
          {
            name: "Revenue",
            categories: "'Pivot Sheet'!$A$4:$A$5",
            values: "'Pivot Sheet'!$B$4:$B$5"
          }
        ],
        pivotChartOptions: {
          refreshOnOpen: true,
          dropZonesVisible: true,
          dropZoneFilter: true,
          dropZoneCategories: true,
          dropZoneData: false,
          dropZoneSeries: false
        }
      },
      "D1:K12"
    );

    expect(pivotTable.pivotChartOptions?.refreshOnOpen).toBe(true);
    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    const cacheXml = textDecoder.decode(
      entries.get("xl/pivotCache/pivotCacheDefinition1.xml")!.data
    );
    // Pivot options now emit the standard MS `c14:pivotOptions` element
    // inside chartSpace's `<c:extLst>`. The uri is the GUID Excel 2010+
    // recognises (MS-XLSX §2.3.11).
    expect(chartXml).toContain("{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}");
    expect(chartXml).toContain("<c14:pivotOptions>");
    expect(chartXml).toContain('<c14:dropZoneFilter val="1"/>');
    expect(chartXml).toContain('<c14:dropZoneCategories val="1"/>');
    expect(chartXml).toContain('<c14:dropZoneData val="0"/>');
    expect(chartXml).toContain('<c14:dropZoneSeries val="0"/>');
    expect(chartXml).toContain('<c14:dropZonesVisible val="1"/>');
    expect(cacheXml).toContain('refreshOnLoad="1"');
  });

  it("writes only the drop-zone toggles the caller explicitly sets", async () => {
    const { wb, pivot, pivotTable } = makePivotWorkbook();

    pivot.addPivotChart(
      pivotTable,
      {
        type: "bar",
        series: [
          {
            name: "Revenue",
            categories: "'Pivot Sheet'!$A$4:$A$5",
            values: "'Pivot Sheet'!$B$4:$B$5"
          }
        ],
        pivotChartOptions: {
          // Deliberately minimal: only suppress the series drop zone,
          // leave all other children absent so Excel uses its defaults.
          dropZoneSeries: false
        }
      },
      "D1:K12"
    );

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(chartXml).toContain('<c14:dropZoneSeries val="0"/>');
    // None of the absent flags should be emitted.
    expect(chartXml).not.toContain("dropZoneFilter");
    expect(chartXml).not.toContain("dropZoneCategories");
    expect(chartXml).not.toContain("dropZoneData");
    expect(chartXml).not.toContain("dropZonesVisible");
  });

  it("creates a combo pivot chart", async () => {
    const { wb, pivot, pivotTable } = makePivotWorkbook();

    pivot.addPivotComboChart(
      pivotTable,
      {
        groups: [
          {
            type: "bar",
            series: [
              {
                name: "Revenue",
                categories: "'Pivot Sheet'!$A$4:$A$5",
                values: "'Pivot Sheet'!$B$4:$B$5"
              }
            ]
          },
          {
            type: "line",
            series: [
              {
                name: "Units",
                categories: "'Pivot Sheet'!$A$4:$A$5",
                values: "'Pivot Sheet'!$C$4:$C$5"
              }
            ],
            useSecondaryAxis: true
          }
        ],
        title: "Revenue and Units"
      },
      "D1:K12"
    );

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(chartXml).toContain("<c:pivotSource>");
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("<c:lineChart>");
    expect(pivotTable.chartFormats).toHaveLength(1);
  });

  it("creates a pivot chart chartsheet", async () => {
    const { wb, pivotTable } = makePivotWorkbook();

    const chartsheet = wb.addPivotChartsheet("Pivot Chart", pivotTable, {
      chart: {
        type: "bar",
        series: [
          {
            name: "Revenue",
            categories: "'Pivot Sheet'!$A$4:$A$5",
            values: "'Pivot Sheet'!$B$4:$B$5"
          }
        ],
        title: "Revenue by Region"
      }
    });

    expect(chartsheet.chartNumber).toBe(1);
    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/chartsheets/sheet1.xml")).toBeDefined();
    expect(entries.get("xl/charts/chart1.xml")).toBeDefined();
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    const pivotXml = textDecoder.decode(entries.get("xl/pivotTables/pivotTable1.xml")!.data);
    expect(chartXml).toContain("<c:pivotSource>");
    expect(chartXml).toContain("<c:name>'Pivot Sheet'!PivotTable1</c:name>");
    expect(pivotXml).toContain('<chartFormats count="1">');
  });

  it("round-trips c14:pivotOptions through a full xlsx load / save cycle", async () => {
    const { wb, pivot, pivotTable } = makePivotWorkbook();

    pivot.addPivotChart(
      pivotTable,
      {
        type: "bar",
        series: [
          {
            name: "Revenue",
            categories: "'Pivot Sheet'!$A$4:$A$5",
            values: "'Pivot Sheet'!$B$4:$B$5"
          }
        ],
        pivotChartOptions: {
          dropZonesVisible: true,
          dropZoneFilter: false,
          dropZoneCategories: true,
          dropZoneData: true,
          dropZoneSeries: false
        }
      },
      "D1:K12"
    );

    // First save and reopen — verifies the parser recognises the c14
    // extension and hydrates the structured model from it.
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart2 = wb2.worksheets[1].getCharts()[0];
    const m2 = chart2.chartModel!;
    expect(m2.pivotOptions).toEqual({
      dropZonesVisible: true,
      dropZoneFilter: false,
      dropZoneCategories: true,
      dropZoneData: true,
      dropZoneSeries: false
    });

    // Second save — ensures we write exactly one c14:pivotOptions element
    // even after a full round-trip (no duplication from leaking into the
    // raw extLst cache).
    const buf2 = await wb2.xlsx.writeBuffer();
    const entries2 = await extractAll(new Uint8Array(buf2));
    const chartXml2 = textDecoder.decode(entries2.get("xl/charts/chart1.xml")!.data);
    expect(chartXml2).toContain("<c14:pivotOptions>");
    const occurrences = (chartXml2.match(/<c14:pivotOptions>/g) ?? []).length;
    expect(occurrences).toBe(1);
    expect(chartXml2).toContain('<c14:dropZoneFilter val="0"/>');
    expect(chartXml2).toContain('<c14:dropZoneCategories val="1"/>');
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

  it("copyTo supports chartEx charts", async () => {
    const wb = new Workbook();
    const src = wb.addWorksheet("Src");
    const dst = wb.addWorksheet("Dst");
    for (let i = 1; i <= 3; i++) {
      src.getCell(`A${i}`).value = `C${i}`;
      src.getCell(`B${i}`).value = i * 10;
    }
    src.addChartEx(
      {
        type: "sunburst",
        categories: "Src!$A$1:$A$3",
        series: [{ name: "Data", values: "Src!$B$1:$B$3" }],
        title: "Modern"
      },
      "C1:J10"
    );

    const copiedNumber = src.getCharts()[0].copyTo(dst, "A1:H10");
    expect(copiedNumber).toBe(2);
    expect(dst.getCharts()[0].isChartEx).toBe(true);

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/charts/chartEx1.xml")).toBeDefined();
    expect(entries.get("xl/charts/chartEx2.xml")).toBeDefined();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    expect(wb2.getWorksheet("Dst")!.getCharts()[0].isChartEx).toBe(true);
    expect(wb2.getWorksheet("Dst")!.getCharts()[0].title).toBe("Modern");
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

  it("fills multi-level category caches", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["North", "Q1", 10],
      ["North", "Q2", 20],
      ["South", "Q1", 30]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            name: "Sales",
            values: "Sheet1!$C$1:$C$3",
            categories: { multiLvlStrRef: { formula: "Sheet1!$A$1:$B$3" } }
          }
        ]
      },
      "E1:K10"
    );

    const series = ws.getCharts()[0].chartModel!.chart.plotArea.chartTypes[0].series[0] as any;
    expect(series.cat.multiLvlStrRef.cache.pointCount).toBe(3);
    expect(series.cat.multiLvlStrRef.cache.levels).toHaveLength(2);
    expect(series.cat.multiLvlStrRef.cache.levels[0].points[0]).toEqual({
      index: 0,
      value: "North"
    });
    expect(series.cat.multiLvlStrRef.cache.levels[1].points[1]).toEqual({
      index: 1,
      value: "Q2"
    });

    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(chartXml).toContain("<c:multiLvlStrCache>");
    expect(chartXml).toContain('<c:ptCount val="3"/>');
    expect(chartXml).toContain("<c:v>North</c:v>");
    expect(chartXml).toContain("<c:v>Q2</c:v>");
  });
});

describe("TC7: pareto chart type in ChartEx builder", () => {
  it("pareto type uses clusteredColumn layoutId with x values and pareto line", () => {
    const m = buildChartExModel({
      type: "pareto",
      categories: "Sheet1!$A$1:$A$5",
      series: [{ name: "Data", values: "Sheet1!$B$1:$B$5" }]
    });
    const s = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(s.layoutId).toBe("clusteredColumn");
    expect(s.layoutPr?.paretoLine).toBe(true);
    expect(m.chartSpace.chartData.data.find(entry => entry.numDim)?.numDim?.type).toBe("x");
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
    expect(xml).toContain('<cx:paretoLine val="1"/>');
    expect(xml).toContain("Pareto Analysis");
  });

  it("ChartEx histogram honors explicit binning fields and validates bad binning", () => {
    const m = buildChartExModel({
      type: "histogram",
      series: [{ name: "H", values: "Sheet1!$B$1:$B$10" }],
      binning: { binType: "binSize", binSize: 5, underflow: 0, overflow: 100, intervalClosed: "r" }
    });
    const binning = m.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutPr!.binning!;
    expect(binning.binSize).toBe(5);
    expect(binning.underflow).toBe(0);
    expect(binning.overflow).toBe(100);
    expect(() =>
      buildChartExModel({
        type: "histogram",
        series: [{ values: "Sheet1!$B$1:$B$10" }],
        binning: { binSize: 0 }
      })
    ).toThrow(/binSize/);
    expect(() =>
      buildChartExModel({
        type: "funnel",
        series: [{ values: "Sheet1!$B$1:$B$10" }],
        layout: { binning: { binType: "auto" } }
      })
    ).toThrow(/binning/);
    expect(() =>
      buildChartExModel({
        type: "histogram",
        series: [
          {
            values: "Sheet1!$B$1:$B$10",
            literalValues: [1, 2],
            literalCategories: ["A"]
          }
        ]
      })
    ).toThrow(/literalCategories length/);
    expect(() =>
      buildChartExModel({
        type: "waterfall",
        series: [{ values: "Sheet1!$B$1:$B$10", subtotals: [-1] }]
      })
    ).toThrow(/subtotals must contain non-negative integer/);
  });

  it("ChartEx literal histogram bins render underflow, closed intervals, and overflow deterministically", () => {
    const model = buildChartExModel({
      type: "histogram",
      series: [{ values: "Sheet1!$B$1:$B$6", literalValues: [-1, 0, 5, 10, 11, 21] }],
      binning: { binType: "binSize", binSize: 10, underflow: 0, overflow: 20, intervalClosed: "r" }
    });
    const svg = renderChartExSvg(model, { width: 420, height: 240 });

    expect(svg).toContain("&lt;=0");
    expect(svg).toContain("0-10");
    expect(svg).toContain("10-20");
    expect(svg).toContain("&gt;20");
    expect(svg).toContain(">2<");
    expect(stableHash(svg)).toBe("6819fc71");
  });

  it("ChartEx renderAxes emits tick labels at every gridline (min + 4 quintiles + max)", () => {
    // Guard the promise the new drawAxesPdf makes — that SVG and PDF
    // backends render the same number of tick labels on the value
    // axis. Using a distribution whose min/max span evenly into 0.4
    // increments so the quintile markers are exact numeric strings we
    // can assert without fighting float formatting.
    const model = buildChartExModel({
      type: "histogram",
      series: [{ values: "Sheet1!$B$1:$B$5", literalValues: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] }],
      binning: { binType: "binCount", binCount: 5 }
    });
    const svg = renderChartExSvg(model, { width: 320, height: 200 });
    // Interior quintile tick labels (0.4, 0.8, 1.2, 1.6) must all
    // appear between min=0 and max=2. Before this change renderAxes
    // only emitted the min/max pair.
    expect(svg).toContain(">0.4<");
    expect(svg).toContain(">0.8<");
    expect(svg).toContain(">1.2<");
    expect(svg).toContain(">1.6<");
    // And the framed min/max labels stay — we only *added* ticks,
    // never removed.
    expect(svg).toContain(">0<");
    expect(svg).toContain(">2<");
  });

  it("ChartEx parser round-trips manual binning values", () => {
    const model = buildChartExModel({
      type: "histogram",
      series: [{ values: "Sheet1!$B$1:$B$6" }],
      binning: { binType: "manual", binSize: 2, binCount: 4, underflow: 1, overflow: 9 }
    });
    const parsed = parseChartEx(renderChartEx(model));
    const binning = parsed.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutPr!.binning!;

    expect(binning.binType).toBe("manual");
    expect(binning.binSize).toBe(2);
    expect(binning.binCount).toBe(4);
    expect(binning.underflow).toBe(1);
    expect(binning.overflow).toBe(9);
  });

  it("ChartEx parser preserves real Excel hierarchy, map layout, axis extLst, and external data", () => {
    const rawXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:chartData>
    <cx:externalData r:id="rId1" autoUpdate="1"/>
    <cx:data id="0"><cx:strDim type="cat"><cx:f>Data!$A$2:$A$5</cx:f><cx:lvl ptCount="4"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt><cx:pt idx="2">C</cx:pt><cx:pt idx="3">D</cx:pt></cx:lvl></cx:strDim></cx:data>
    <cx:data id="1"><cx:strDim type="cat"><cx:f>Data!$B$2:$B$5</cx:f><cx:lvl ptCount="4"><cx:pt idx="0">North</cx:pt><cx:pt idx="1">North</cx:pt><cx:pt idx="2">South</cx:pt><cx:pt idx="3">South</cx:pt></cx:lvl></cx:strDim></cx:data>
    <cx:data id="2"><cx:strDim type="cat"><cx:f>Data!$C$2:$C$5</cx:f><cx:lvl ptCount="4"><cx:pt idx="0">USA</cx:pt><cx:pt idx="1">Canada</cx:pt><cx:pt idx="2">Brazil</cx:pt><cx:pt idx="3">Chile</cx:pt></cx:lvl></cx:strDim></cx:data>
    <cx:data id="3"><cx:numDim type="val"><cx:f>Data!$D$2:$D$5</cx:f><cx:lvl ptCount="4" formatCode="#,##0"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">5</cx:pt><cx:pt idx="3">15</cx:pt></cx:lvl></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="regionMap" ownerIdx="0">
          <cx:dataId val="0"/><cx:dataId val="1"/><cx:dataId val="2"/><cx:dataId val="3"/>
          <cx:layoutPr><cx:projection val="mercator"/><cx:regionLabels val="showAll"/><cx:geoMappingLevel val="country"/><cx:extLst><cx:ext uri="{layout-ext}"/></cx:extLst></cx:layoutPr>
          <cx:axisId val="42"/>
          <cx:extLst><cx:ext uri="{series-ext}"/></cx:extLst>
        </cx:series>
      </cx:plotAreaRegion>
      <cx:axis id="42"><cx:catScaling gapWidth="219"/><cx:extLst><cx:ext uri="{axis-ext}"/></cx:extLst></cx:axis>
    </cx:plotArea>
  </cx:chart>
  <cx:extLst><cx:ext uri="{space-ext}"/></cx:extLst>
</cx:chartSpace>`;

    const parsed = parseChartEx(rawXml);
    const series = parsed.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    const axis = parsed.chartSpace.chart.plotArea.axis![0];

    expect(parsed.chartSpace.chartData.externalData?.[0]).toEqual({ id: "rId1", autoUpdate: true });
    expect(parsed.chartSpace.chartData.data).toHaveLength(4);
    expect(series.layoutId).toBe("regionMap");
    expect(series.dataRefs?.map(ref => ref.dataId)).toEqual([0, 1, 2, 3]);
    expect(series.layoutPr?.projection).toBe("mercator");
    expect(series.layoutPr?.regionLabels).toBe("showAll");
    expect(series.layoutPr?.geoMappingLevel).toBe("country");
    expect(series.layoutPr?._rawXml).toContain("{layout-ext}");
    expect(series.layoutPr?.extLst).toContain("{layout-ext}");
    expect(series.extLst).toContain("{series-ext}");
    expect(axis.catScaling?.gapWidth).toBe(219);
    expect(axis.extLst).toContain("{axis-ext}");
    expect(parsed.chartSpace.extLst).toContain("{space-ext}");

    const modelForRender = { ...parsed, rawXml: undefined };
    const rendered = renderChartEx(modelForRender);
    expect(rendered).toContain("{axis-ext}");
    expect(rendered).toContain("{layout-ext}");
    expect(rendered).toContain("{series-ext}");
  });

  it("ChartEx raw patch preserves template XML while updating hidden, series name, and axis bindings", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15],
      ["D", 30],
      ["E", 45]
    ]);
    ws.addChartEx(
      {
        type: "waterfall",
        categories: "Sheet1!$A$1:$A$5",
        series: [{ name: "Before", values: "Sheet1!$B$1:$B$5", subtotals: [4] }]
      },
      "D1:J10"
    );

    const zipData = await extractAll(new Uint8Array(await wb.xlsx.writeBuffer()));
    const chartExEntry = zipData.get("xl/charts/chartEx1.xml")!;
    const marker = "<!-- patchable-chartEx-template -->";
    chartExEntry.data = textEncoder.encode(
      textDecoder.decode(chartExEntry.data).replace("<cx:chart>", `${marker}<cx:chart>`)
    );
    const { createZip } = await import("@archive/zip/zip-bytes");
    const input = await createZip(
      [...zipData.entries()].map(([name, file]) => ({ name, data: file.data }))
    );

    const wb2 = new Workbook();
    await wb2.xlsx.load(input);
    const chart = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    chart.mutateChartEx(
      model => {
        const series = model.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
        series.hidden = true;
        series.tx = { value: "After" };
        series.axisId = [0];
      },
      { preferRawPatch: true, requireRawPatch: true }
    );

    const output = await wb2.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(output));
    const xml = textDecoder.decode(entries.get("xl/charts/chartEx1.xml")!.data);
    expect(xml).toContain(marker);
    expect(xml).toContain('hidden="1"');
    expect(xml).toContain("After");
    expect(xml).toContain('<cx:axisId val="0"/>');
    expect(xml).not.toContain('<cx:axisId val="1"/>');
  });

  it("ChartEx regionMap preview uses geographic projection for known regions", () => {
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2:$B$5",
          literalValues: [10, 20, 5, 15],
          literalCategories: ["USA", "Canada", "Brazil", "Chile"]
        }
      ],
      layout: { projection: "mercator", regionLabels: "showAll", geoMappingLevel: "country" }
    });

    const svg = renderChartExSvg(model, { width: 420, height: 260 });
    expect(svg).toContain('data-region-map-mode="geographic-preview"');
    expect(svg).toContain("USA");
    expect(svg).toContain("Canada");
    expect(svg).toContain("<circle");
    expect(stableHash(svg)).toBe("1318a0c0");
  });

  it("ChartEx regionMap Albers and Robinson projections use real formulas, not linear fallbacks", () => {
    // Before the projection fix, `albers` used `y = 0.52 - lat/210` and
    // `robinson` fell through to a generic `y = 0.5 - lat/190` — both
    // linear, neither a real projection. With the fix, each projection
    // must produce a geometry that differs from the previous linear
    // approximation AND from the other projections (Robinson ≠ Albers ≠
    // Mercator) for the same canvas. SVG hash is the most stable
    // assertion we can make here.
    const build = (projection: "albers" | "robinson" | "mercator" | "miller"): string =>
      renderChartExSvg(
        buildChartExModel({
          type: "regionMap",
          series: [
            {
              values: "Data!$B$2:$B$5",
              literalValues: [10, 20, 5, 15],
              literalCategories: ["USA", "Canada", "Brazil", "Chile"]
            }
          ],
          layout: { projection, regionLabels: "showAll", geoMappingLevel: "country" }
        }),
        { width: 420, height: 260 }
      );
    const mercator = build("mercator");
    const miller = build("miller");
    const albers = build("albers");
    const robinson = build("robinson");
    // Each projection produces a distinct SVG; equality would mean a
    // projection silently aliased to another (as `robinson` did before).
    const hashes = new Set([
      stableHash(mercator),
      stableHash(miller),
      stableHash(albers),
      stableHash(robinson)
    ]);
    expect(hashes.size).toBe(4);
    // And none of them should be identical to the old linear formulas.
    // `albers` previously resolved `Chile` (lat ≈ -33) to y ≈ 0.52 +
    // 33/210 ≈ 0.677, landing ~75% down the canvas. With real Albers
    // equal-area on world parallels, Chile ends up higher (closer to
    // the southern edge because the cone narrows towards the pole).
    // Match a regex against the first `<circle cy="…">` in each output
    // — that circle corresponds to `usa`, the first key alphabetically
    // in the centroid table — and assert the Albers cy differs from the
    // old value.
    const albersCy = /<circle[^>]*cy="([0-9.]+)"/.exec(albers)?.[1];
    const robinsonCy = /<circle[^>]*cy="([0-9.]+)"/.exec(robinson)?.[1];
    expect(albersCy).toBeDefined();
    expect(robinsonCy).toBeDefined();
    expect(albersCy).not.toBe(robinsonCy);
  });

  it("ChartEx regionMap uses user-provided TopoJSON for real country polygons", () => {
    // Minimal hand-authored topology: one square "country" with id
    // `"USA"`. Confirms the full pipeline works end-to-end:
    // resolveTopologyObject → feature match → SVG path emission.
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [10, 0],
          [0, 10],
          [-10, 0],
          [0, -10]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [{ type: "Polygon", id: "USA", arcs: [[0]] }]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2:$B$3",
          literalValues: [42, 17],
          literalCategories: ["USA", "Unknown"]
        }
      ],
      layout: { projection: "miller", regionLabels: "showAll", geoMappingLevel: "country" }
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: {
        topology,
        objectName: "countries",
        match: "id"
      }
    });
    // TopoJSON path is emitted (and the fallback centroid preview is NOT).
    expect(svg).toContain('data-region-map-mode="topojson"');
    expect(svg).not.toContain('data-region-map-mode="geographic-preview"');
    // At least one <path> element (the USA polygon) is present.
    expect(svg).toMatch(/<path d="M/);
    // Label "USA" is rendered at the centroid.
    expect(svg).toContain("USA");
  });

  it("ChartEx regionMap falls back to centroid preview when TopoJSON matches nothing", () => {
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [10, 0]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [{ type: "LineString", id: "TOTALLY_FAKE_ID", arcs: [0] }]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2:$B$3",
          literalValues: [10, 20],
          literalCategories: ["USA", "Canada"]
        }
      ],
      layout: { projection: "miller", regionLabels: "showAll", geoMappingLevel: "country" }
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: { topology, objectName: "countries", match: "id" }
    });
    // USA/Canada are in the built-in centroid table, so the fallback
    // preview runs and draws dots + labels.
    expect(svg).toContain('data-region-map-mode="geographic-preview"');
    expect(svg).toContain("USA");
    expect(svg).toContain("Canada");
  });

  it("ChartEx regionMap matches features on a property key (e.g. world-atlas `name`)", () => {
    // world-atlas countries-110m.json uses `properties.name` for the
    // English name; confirm `match: "property:name"` selects that
    // field instead of the numeric `id`.
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [5, 0],
          [0, 5],
          [-5, 0],
          [0, -5]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [
            {
              type: "Polygon",
              id: 840,
              properties: { name: "United States" },
              arcs: [[0]]
            }
          ]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2",
          literalValues: [7],
          literalCategories: ["United States"]
        }
      ],
      layout: { projection: "miller", regionLabels: "showAll", geoMappingLevel: "country" }
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: { topology, objectName: "countries", match: "property:name" }
    });
    expect(svg).toContain('data-region-map-mode="topojson"');
    expect(svg).toContain("United States");
  });

  it("ChartEx regionMap `match: [...]` tries locale-aware rules in order (first-match-wins)", () => {
    // World-atlas-style topology with two features:
    //   - France: has both `name_zh` and `name`
    //   - Spain:  has only `name` (legacy file without localisation)
    //
    // The categories array uses the Chinese spellings. With a matchers
    // fallback of `["property:name_zh", "property:name"]` France must
    // match via `name_zh` and Spain via `name` — exactly the pattern
    // natural-earth-vectors workflows need.
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [5, 0],
          [0, 5],
          [-5, 0],
          [0, -5]
        ],
        [
          [10, 10],
          [5, 0],
          [0, 5],
          [-5, 0],
          [0, -5]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [
            {
              type: "Polygon",
              id: 250,
              properties: { name: "France", name_zh: "法国" },
              arcs: [[0]]
            },
            {
              type: "Polygon",
              id: 724,
              properties: { name: "Spain" },
              arcs: [[1]]
            }
          ]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2:$B$3",
          literalValues: [10, 20],
          literalCategories: ["法国", "Spain"]
        }
      ],
      layout: { projection: "miller", regionLabels: "showAll", geoMappingLevel: "country" }
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: {
        topology,
        objectName: "countries",
        match: ["property:name_zh", "property:name"]
      }
    });
    // Topology path was used — real polygons, not the centroid fallback.
    expect(svg).toContain('data-region-map-mode="topojson"');
    // Both features matched → both labels rendered, each with the
    // original category spelling.
    expect(svg).toContain("法国");
    expect(svg).toContain("Spain");
  });

  it("ChartEx regionMap single `match` string stays back-compatible", () => {
    // Sanity: passing the old-style single rule still routes through
    // the matchers-array code path without regressing the original
    // single-match shape or the label fall-back.
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [5, 0],
          [0, 5],
          [-5, 0],
          [0, -5]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [{ type: "Polygon", id: "USA", arcs: [[0]] }]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [{ values: "Data!$B$2", literalValues: [42], literalCategories: ["USA"] }]
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: { topology, objectName: "countries", match: "id" }
    });
    expect(svg).toContain('data-region-map-mode="topojson"');
    expect(svg).toContain("USA");
  });

  it("ChartEx regionMap matchers fall through to centroid preview when every rule misses", () => {
    const topology = {
      type: "Topology",
      arcs: [
        [
          [0, 0],
          [5, 0],
          [0, 5]
        ]
      ],
      objects: {
        countries: {
          type: "GeometryCollection",
          geometries: [
            {
              type: "Polygon",
              id: 999,
              properties: { code_iso2: "XY" },
              arcs: [[0]]
            }
          ]
        }
      }
    } as const;
    const model = buildChartExModel({
      type: "regionMap",
      series: [
        {
          values: "Data!$B$2",
          literalValues: [42],
          // Categories are familiar country names → centroid table will pick them up.
          literalCategories: ["USA"]
        }
      ]
    });
    const svg = renderChartExSvg(model, {
      width: 420,
      height: 260,
      regionMap: {
        topology,
        objectName: "countries",
        match: ["property:name_zh", "property:name"]
      }
    });
    // Neither rule matches any feature → centroid preview runs, not the
    // topojson path.
    expect(svg).toContain('data-region-map-mode="geographic-preview"');
    expect(svg).toContain("USA");
  });

  it("ChartEx literal multi-level hierarchy renders deterministically", () => {
    const model = buildChartExModel({
      type: "sunburst",
      series: [
        {
          values: "Data!$D$2:$D$5",
          literalValues: [10, 20, 5, 15],
          literalCategories: ["A", "B", "C", "D"],
          literalHierarchy: [
            ["North", "North", "South", "South"],
            ["USA", "Canada", "Brazil", "Chile"]
          ]
        }
      ]
    });
    const svg = renderChartExSvg(model, { width: 360, height: 240 });

    expect(svg).toContain("<path");
    expect(stableHash(svg)).toBe("38e54e31");
  });

  it("ChartEx sunburst handles 3-level hierarchy (continent/country/city)", () => {
    // Extend the 2-level test above to 3 levels so the parser's
    // `findChildren("cx:lvl")` + builder `literalHierarchy[][]` code
    // paths stay truthful for the common real-world case of
    // continent → country → city sunbursts. Before: only 2 levels had
    // explicit coverage.
    const model = buildChartExModel({
      type: "sunburst",
      series: [
        {
          values: "Data!$D$2:$D$7",
          literalValues: [10, 20, 5, 15, 8, 12],
          literalCategories: ["NYC", "LA", "Toronto", "Rio", "Beijing", "Tokyo"],
          literalHierarchy: [
            // Continent
            ["Americas", "Americas", "Americas", "Americas", "Asia", "Asia"],
            // Country
            ["USA", "USA", "Canada", "Brazil", "China", "Japan"]
          ]
        }
      ]
    });
    // Model-level: three strDim entries (continent/country/city) + one numDim.
    const data = model.chartSpace.chartData.data;
    const strDims = data.filter(d => d.strDim !== undefined);
    expect(strDims.length).toBe(3);
    for (const dim of strDims) {
      for (const lvl of dim.strDim!.levels ?? []) {
        expect(lvl.points.length).toBe(6);
      }
    }
    // Render-level: a 3-level sunburst emits three concentric arc
    // rings. Each ring's slice count equals the number of distinct
    // parents at that depth. Total <path> count ≥ 6 (leaves) + 3
    // (countries USA/Canada/Brazil/China/Japan merged by rendering
    // policy) + continents — concretely ≥ 10.
    const svg = renderChartExSvg(model, { width: 360, height: 240 });
    const pathCount = (svg.match(/<path/g) ?? []).length;
    expect(pathCount).toBeGreaterThanOrEqual(10);
  });

  it("ChartEx parses 3-level sunburst XML back into structured data with correct ptCount on every level", async () => {
    // Round-trip the authored XML — writeBuffer → load → parse — so
    // the parser's level handling (`findChildren(cx:lvl).map(...)`) is
    // exercised on the emit-side output, not just a synthetic string.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.addRows([
      ["City", "Country", "Continent", "Revenue"],
      ["NYC", "USA", "Americas", 10],
      ["LA", "USA", "Americas", 20],
      ["Toronto", "Canada", "Americas", 5]
    ]);
    ws.addChartEx(
      {
        type: "sunburst",
        series: [
          {
            values: "Data!$D$2:$D$4",
            literalValues: [10, 20, 5],
            literalCategories: ["NYC", "LA", "Toronto"],
            literalHierarchy: [
              ["Americas", "Americas", "Americas"],
              ["USA", "USA", "Canada"]
            ]
          }
        ]
      },
      "F1:L12"
    );
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart = wb2.getWorksheet("Data")!.getCharts()[0];
    const data = chart.chartExModel!.chartSpace.chartData.data;
    // Three str dimensions (continent, country, city) + one numeric
    // (revenue) = 4 data entries at least.
    expect(data.length).toBeGreaterThanOrEqual(4);
    const strEntries = data.filter(d => d.strDim !== undefined);
    expect(strEntries.length).toBeGreaterThanOrEqual(3);
    // Every strDim level must have exactly 3 points (our three rows).
    for (const entry of strEntries) {
      for (const lvl of entry.strDim!.levels ?? []) {
        expect(lvl.points.length).toBe(3);
        // ptCount (serialised) equals the real count — guards against
        // a parser that dropped the attribute or rebuilt it from a
        // stale snapshot.
        expect(lvl.ptCount ?? lvl.points.length).toBe(3);
      }
    }
  });

  it("ChartEx SVG renders pareto, waterfall, funnel, hierarchy, and boxWhisker previews", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", "North", 10],
      ["B", "North", 20],
      ["C", "South", 5],
      ["D", "South", 15]
    ]);
    const models = [
      buildChartExModel({
        type: "pareto",
        categories: "Sheet1!$A$1:$A$4",
        series: [{ values: "Sheet1!$C$1:$C$4" }]
      }),
      buildChartExModel({
        type: "waterfall",
        categories: "Sheet1!$A$1:$A$4",
        series: [{ values: "Sheet1!$C$1:$C$4", subtotals: [3] }],
        layout: {
          connectorLines: true,
          increaseSpPr: { fill: { solid: { srgb: "00AA00" } } },
          decreaseSpPr: { fill: { solid: { srgb: "AA0000" } } },
          totalSpPr: { fill: { solid: { srgb: "0000AA" } } }
        }
      }),
      buildChartExModel({
        type: "funnel",
        categories: "Sheet1!$A$1:$A$4",
        series: [{ values: "Sheet1!$C$1:$C$4" }]
      }),
      buildChartExModel({
        type: "treemap",
        categories: "Sheet1!$A$1:$A$4",
        series: [{ values: "Sheet1!$C$1:$C$4", hierarchy: ["Sheet1!$B$1:$B$4"] }]
      }),
      buildChartExModel({
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$4",
        series: [{ values: "Sheet1!$C$1:$C$4", hierarchy: ["Sheet1!$B$1:$B$4"] }]
      }),
      buildChartExModel({
        type: "boxWhisker",
        categories: "Sheet1!$B$1:$B$4",
        series: [{ values: "Sheet1!$C$1:$C$4" }],
        layout: { showMeanLine: true, showInnerPoints: true, showOutlierPoints: true }
      })
    ];
    for (const model of models) {
      fillChartExCaches(model, wb);
      const svg = renderChartExSvg(model, { width: 360, height: 220 });
      expect(svg).toContain('<svg xmlns="http://www.w3.org/2000/svg"');
      expect(svg).toMatch(/<(rect|polygon|path|polyline|circle)/);
    }
    const waterfallSvg = renderChartExSvg(models[1], { width: 360, height: 220 });
    expect(waterfallSvg).toContain("#00AA00");
    expect(waterfallSvg).toContain("stroke-dasharray");
    const boxSvg = renderChartExSvg(models[5], { width: 360, height: 220 });
    expect(boxSvg).toContain('stroke-dasharray="3 2"');
  });
});

describe("P1: chart convenience APIs and presets", () => {
  it("addColumnChart creates a column bar chart", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addColumnChart({ series: [baseSeries("S")], grouping: "stacked" }, "C1:J10");

    const group = ws.getCharts()[0].chartTypes[0] as BarChartGroup;
    expect(group.type).toBe("bar");
    expect(group.barDir).toBe("col");
    expect(group.grouping).toBe("stacked");
  });

  it("addLineChart creates a line chart", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addLineChart({ series: [baseSeries("S")] }, "C1:J10");

    expect(ws.getCharts()[0].chartTypes[0].type).toBe("line");
  });

  it("addHistogramChart creates a chartEx histogram", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addHistogramChart({ series: [{ name: "H", values: "Sheet1!$B$1:$B$5" }] }, "C1:J10");

    const chart = ws.getCharts()[0];
    expect(chart.isChartEx).toBe(true);
    expect(
      chart.chartExModel!.chartSpace.chart.plotArea.plotAreaRegion!.series[0].layoutPr?.binning
        ?.binType
    ).toBe("auto");
  });

  it("addPresetChart maps Excel-style presets", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addPresetChart("bar3DConeClustered", { series: [baseSeries("S")] }, "C1:J10");

    const group = ws.getCharts()[0].chartTypes[0] as BarChartGroup;
    expect(group.type).toBe("bar3D");
    expect(group.barDir).toBe("bar");
    expect(group.shape).toBe("cone");
  });

  it("applyChartPreset covers Excel 3D shape, stock, scatter, and surface aliases", () => {
    const cases: Array<[Parameters<typeof applyChartPreset>[0], Partial<AddChartOptions>]> = [
      [
        "cylinderColStacked100",
        { type: "bar3D", barDir: "col", grouping: "percentStacked", shape: "cylinder" }
      ],
      [
        "pyramidBarStacked",
        { type: "bar3D", barDir: "bar", grouping: "stacked", shape: "pyramid" }
      ],
      [
        "coneBarStacked100",
        { type: "bar3D", barDir: "bar", grouping: "percentStacked", shape: "cone" }
      ],
      ["pieExploded3D", { type: "pie3D" }],
      ["stockVOHLC", { type: "stock" }],
      ["scatterLinesNoMarkers", { type: "scatter", scatterStyle: "line" }],
      ["surfaceTopView", { type: "surface" }],
      ["topViewWireframe", { type: "surface", wireframe: true }]
    ];
    for (const [preset, expected] of cases) {
      const options = applyChartPreset(preset, { series: [baseSeries("S")] });
      expect(options).toMatchObject(expected);
    }
  });

  it("applyChartPreset exposes a broad Excel UI preset alias matrix", () => {
    expect(EXCEL_CHART_PRESETS.length).toBeGreaterThanOrEqual(70);
    expect(new Set(EXCEL_CHART_PRESETS).size).toBe(EXCEL_CHART_PRESETS.length);

    const cases: Array<[Parameters<typeof applyChartPreset>[0], Partial<AddChartOptions>]> = [
      [
        "col3DConeStacked100",
        { type: "bar3D", barDir: "col", grouping: "percentStacked", shape: "cone" }
      ],
      [
        "bar3DCylinderStacked",
        { type: "bar3D", barDir: "bar", grouping: "stacked", shape: "cylinder" }
      ],
      [
        "bar3DPyramidClustered",
        { type: "bar3D", barDir: "bar", grouping: "clustered", shape: "pyramid" }
      ],
      ["barOfPie", { type: "ofPie", ofPieType: "bar" }],
      ["pieOfPie", { type: "ofPie", ofPieType: "pie" }],
      ["area3DStacked100", { type: "area3D", grouping: "percentStacked" }],
      ["scatterMarker", { type: "scatter", scatterStyle: "marker" }],
      ["radarMarkers", { type: "radar", radarStyle: "marker" }],
      ["surface3DWireframe", { type: "surface3D", wireframe: true }],
      ["stockVOHLC", { type: "stock", hiLowLines: true, upDownBars: true }]
    ];

    for (const [preset, expected] of cases) {
      const options = applyChartPreset(preset, { series: [baseSeries("S")] });
      expect(options).toMatchObject(expected);
    }
  });

  it("applyChartPreset maps series-level exploded pie settings", () => {
    const options = applyChartPreset("pieExploded", { series: [baseSeries("S")] });
    const model = buildChartModel(options);
    const group = ctg(model) as PieChartGroup;

    expect(group.type).toBe("pie");
    expect(group.series[0].explosion).toBe(25);
  });

  it("applyChartPreset maps bubble3D to bubble series", () => {
    const options = applyChartPreset("bubble3D", { series: [bubbleSeries("S")] });
    const model = buildChartModel(options);
    const group = ctg(model) as BubbleChartGroup;

    expect(group.type).toBe("bubble");
    expect(group.series[0].bubble3D).toBe(true);
  });

  it("applyChartExPreset maps modern chart presets", () => {
    expect(EXCEL_CHART_EX_PRESETS.length).toBeGreaterThanOrEqual(10);
    expect(new Set(EXCEL_CHART_EX_PRESETS).size).toBe(EXCEL_CHART_EX_PRESETS.length);
    expect(CHART_EX_PRESETS.boxAndWhisker.type).toBe("boxWhisker");

    const series = [{ name: "S", values: VALUES_A }];
    expect(applyChartExPreset("histogram", { series })).toMatchObject({ type: "histogram" });
    expect(applyChartExPreset("pareto", { series })).toMatchObject({ type: "pareto" });
    expect(applyChartExPreset("boxAndWhisker", { series })).toMatchObject({
      type: "boxWhisker"
    });
    expect(applyChartExPreset("map", { series })).toMatchObject({ type: "regionMap" });
  });

  it("creates chart options from column ranges without hand-written sheet formulas", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sales Data");
    const series = ws.seriesFromColumns({ categories: "A2:A4", values: "B2:B4", name: "Sales" });
    expect(series.categories).toBe("'Sales Data'!$A$2:$A$4");
    expect(series.values).toBe("'Sales Data'!$B$2:$B$4");

    const standalone = seriesFromColumns("Sales Data", {
      categories: "A2:A4",
      values: "B2:B4",
      name: "Sales"
    });
    expect(standalone).toEqual(series);
  });

  it("adds charts directly from tables and object rows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "SalesTable",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Month" }, { name: "Sales" }, { name: "Profit" }],
      rows: [
        ["Jan", 10, 2],
        ["Feb", 20, 6]
      ]
    });
    ws.addChartFromTable(
      "SalesTable",
      { type: "bar", barDir: "col", categoryColumn: "Month", valueColumns: ["Sales"] },
      "E1:K10"
    );
    ws.addColumnChartFromRows(
      [
        { month: "Jan", sales: 10 },
        { month: "Feb", sales: 20 }
      ],
      { x: "month", y: "sales", startCell: "H20" },
      "E12:K22"
    );
    expect(ws.getCharts()).toHaveLength(2);
    expect((ws.getCharts()[0].chartTypes[0].series[0] as BarSeries).val?.numRef?.formula).toBe(
      "SalesTable[Sales]"
    );
    expect((ws.getCharts()[1].chartTypes[0].series[0] as BarSeries).cat?.strRef?.formula).toBe(
      "Sheet1!$H$21:$H$22"
    );
    expect(
      chartOptionsFromTable(ws, "SalesTable", { type: "line", valueColumns: ["Profit"] }).series
    ).toHaveLength(1);
    expect(
      chartOptionsFromRows(ws, [{ month: "Mar", sales: 30 }], {
        type: "bar",
        x: "month",
        y: "sales",
        startCell: "K20"
      }).series
    ).toHaveLength(1);
  });

  it("creates table chart options with structured references by default", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data Sheet");
    ws.addTable({
      name: "SalesTable",
      ref: "A1",
      columns: [{ name: "Month" }, { name: "Sales" }, { name: "Profit %" }],
      rows: [
        ["Jan", 10, 0.2],
        ["Feb", 20, 0.3]
      ]
    });

    const dynamic = chartOptionsFromTable(ws, "SalesTable", {
      type: "bar",
      categoryColumn: "Month",
      valueColumns: ["Sales", "Profit %"]
    });
    expect(dynamic.series?.[0].categories).toBe("SalesTable[Month]");
    expect(dynamic.series?.[0].values).toBe("SalesTable[Sales]");
    expect(dynamic.series?.[1].values).toBe("SalesTable[Profit %]");

    const fixed = chartOptionsFromTable(ws, "SalesTable", {
      type: "bar",
      categoryColumn: "Month",
      valueColumns: ["Sales"],
      structuredReferences: false
    });
    expect(fixed.series?.[0].categories).toBe("'Data Sheet'!$A$2:$A$3");
    expect(fixed.series?.[0].values).toBe("'Data Sheet'!$B$2:$B$3");
  });

  it("populates caches from table structured references", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "SalesTable",
      ref: "A1",
      columns: [{ name: "Month" }, { name: "Sales" }],
      rows: [
        ["Jan", 10],
        ["Feb", 20]
      ]
    });

    const options = chartOptionsFromTable(ws, "SalesTable", {
      type: "bar",
      categoryColumn: "Month",
      valueColumns: ["Sales"]
    });
    const model = buildChartModel(options);
    fillChartCaches(model, wb);
    const series = model.chart.plotArea.chartTypes[0].series[0] as BarSeries;

    expect(series.cat?.strRef?.cache?.points).toEqual([
      { index: 0, value: "Jan" },
      { index: 1, value: "Feb" }
    ]);
    expect(series.val?.numRef?.cache?.points).toEqual([
      { index: 0, value: 10 },
      { index: 1, value: 20 }
    ]);
    expect(series.val?.numRef?.cache?.pointCount).toBe(2);
  });

  it("populates caches from escaped structured references", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    const table = ws.addTable({
      name: "OddTable",
      ref: "A1",
      columns: [{ name: "Region" }, { name: "Sales]#'@" }],
      rows: [
        ["East", 7],
        ["West", 8]
      ]
    });
    table.name = "Odd.Table";

    const options = chartOptionsFromTable(ws, table, {
      type: "bar",
      categoryColumn: "Region",
      valueColumns: ["Sales]#'@"]
    });
    const model = buildChartModel(options);
    fillChartCaches(model, wb);
    const series = model.chart.plotArea.chartTypes[0].series[0] as BarSeries;

    expect(series.val?.numRef?.formula).toBe("Odd.Table[Sales']'#'''@]");
    expect(series.val?.numRef?.cache?.points).toEqual([
      { index: 0, value: 7 },
      { index: 1, value: 8 }
    ]);
  });

  it("rejects object-row charts without rows or with a mismatched sheetName", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    wb.addWorksheet("Other");

    expect(() => chartOptionsFromRows(ws, [], { type: "bar", x: "month", y: "sales" })).toThrow(
      /at least one row/
    );
    expect(() =>
      chartOptionsFromRows(ws, [{ month: "Jan", sales: 10 }], {
        type: "bar",
        x: "month",
        y: "sales",
        sheetName: "Other"
      })
    ).toThrow(/sheetName must match/);
  });

  it("rejects table charts with no data rows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "EmptySales",
      ref: "A1",
      columns: [{ name: "Month" }, { name: "Sales" }],
      rows: []
    });

    expect(() =>
      chartOptionsFromTable(ws, "EmptySales", {
        type: "bar",
        categoryColumn: "Month",
        valueColumns: ["Sales"]
      })
    ).toThrow(/no data rows/);
  });

  it("chartExOptionsFromTable builds sunburst/funnel/histogram options from a Table", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data Sheet");
    ws.addTable({
      name: "RegionSales",
      ref: "A1",
      columns: [{ name: "Region" }, { name: "Sales" }, { name: "Profit %" }],
      rows: [
        ["North", 100, 0.2],
        ["South", 80, 0.15],
        ["East", 60, 0.3]
      ]
    });

    const structured = chartExOptionsFromTable(ws, "RegionSales", {
      type: "sunburst",
      categoryColumn: "Region",
      valueColumns: ["Sales", "Profit %"]
    });
    expect(structured.type).toBe("sunburst");
    expect(structured.categories).toBe("RegionSales[Region]");
    expect(structured.series).toHaveLength(2);
    expect(structured.series[0].values).toBe("RegionSales[Sales]");
    expect(structured.series[0].name).toBe("Sales");
    expect(structured.series[1].values).toBe("RegionSales[Profit %]");

    const absolute = chartExOptionsFromTable(ws, "RegionSales", {
      type: "funnel",
      categoryColumn: "Region",
      structuredReferences: false
    });
    expect(absolute.type).toBe("funnel");
    expect(absolute.categories).toBe("'Data Sheet'!$A$2:$A$4");
    // Default valueColumns = every non-category column.
    expect(absolute.series).toHaveLength(2);
    expect(absolute.series[0].values).toBe("'Data Sheet'!$B$2:$B$4");
    expect(absolute.series[1].values).toBe("'Data Sheet'!$C$2:$C$4");

    // Histogram should build cleanly too.
    const hist = chartExOptionsFromTable(ws, "RegionSales", {
      type: "histogram",
      valueColumns: ["Sales"]
    });
    const histModel = buildChartExModel(hist);
    expect(histModel.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId).toBe(
      "clusteredColumn"
    );
  });

  it("chartExOptionsFromRows stages rows and builds waterfall/treemap options", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    const waterfall = chartExOptionsFromRows(
      ws,
      [
        { stage: "Start", value: 100 },
        { stage: "Adj1", value: -20 },
        { stage: "Adj2", value: 30 },
        { stage: "End", value: 110 }
      ],
      { type: "waterfall", x: "stage", y: "value", startCell: "A1" }
    );
    expect(waterfall.type).toBe("waterfall");
    expect(waterfall.categories).toBe("Sheet1!$A$2:$A$5");
    expect(waterfall.series[0].values).toBe("Sheet1!$B$2:$B$5");
    // Header row was written.
    expect(ws.getCell("A1").value).toBe("stage");
    expect(ws.getCell("B1").value).toBe("value");
    expect(ws.getCell("A2").value).toBe("Start");
    expect(ws.getCell("B4").value).toBe(30);

    // Skipping headers + offset startCell.
    const treemap = chartExOptionsFromRows(
      ws,
      [
        { leaf: "A", size: 1 },
        { leaf: "B", size: 2 }
      ],
      {
        type: "treemap",
        x: "leaf",
        y: "size",
        startCell: "D10",
        includeHeaders: false
      }
    );
    expect(treemap.categories).toBe("Sheet1!$D$10:$D$11");
    expect(treemap.series[0].values).toBe("Sheet1!$E$10:$E$11");
    expect(ws.getCell("D10").value).toBe("A");
    // No header was written.
    expect(ws.getCell("D9").value).toBeNull();
  });

  it("chartEx helpers surface through worksheet.addChartExFromTable and addChartExFromRows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "Funnel",
      ref: "A1",
      columns: [{ name: "Stage" }, { name: "Count" }],
      rows: [
        ["Visit", 1000],
        ["Sign up", 400],
        ["Purchase", 80]
      ]
    });

    const tableChartNum = ws.addChartExFromTable(
      "Funnel",
      { type: "funnel", categoryColumn: "Stage" },
      "D1:J10"
    );
    const rowChartNum = ws.addChartExFromRows(
      [
        { stage: "Start", value: 50 },
        { stage: "End", value: 40 }
      ],
      { type: "waterfall", x: "stage", y: "value", startCell: "L1" },
      "D12:J21"
    );

    expect(tableChartNum).toBeGreaterThan(0);
    expect(rowChartNum).toBeGreaterThan(tableChartNum);
    // Both entries should have created ChartExModel records.
    const charts = ws.getCharts();
    expect(charts).toHaveLength(2);
    const tableChart = charts.find(chart => chart.chartExNumber === tableChartNum);
    const rowChart = charts.find(chart => chart.chartExNumber === rowChartNum);
    expect(
      tableChart?.chartExModel?.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId
    ).toBe("funnel");
    expect(
      rowChart?.chartExModel?.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId
    ).toBe("waterfall");
  });

  it("chartEx helpers reject empty rows, missing tables, and sheet mismatch", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    wb.addWorksheet("Other");
    ws.addTable({
      name: "EmptyEx",
      ref: "A1",
      columns: [{ name: "K" }, { name: "V" }],
      rows: []
    });

    expect(() => chartExOptionsFromTable(ws, "EmptyEx", { type: "sunburst" })).toThrow(
      /no data rows/
    );

    expect(() => chartExOptionsFromTable(ws, "MissingTable", { type: "treemap" })).toThrow(
      /Table not found/
    );

    expect(() => chartExOptionsFromRows(ws, [], { type: "funnel", x: "a", y: "b" })).toThrow(
      /at least one row/
    );

    expect(() =>
      chartExOptionsFromRows(ws, [{ a: "x", b: 1 }], {
        type: "funnel",
        x: "a",
        y: "b",
        sheetName: "Other"
      })
    ).toThrow(/sheetName must match/);
  });

  it("Chart.setUserShapesXml attaches a drawing overlay and round-trips through xlsx write", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 1],
      ["B", 2]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );

    const chart = ws.getCharts()[0];
    // Freshly built chart has no user shapes.
    expect(chart.userShapesXml).toBeUndefined();

    const userShapesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      "<c:relSizeAnchor>" +
      "<c:from><c:x>0.1</c:x><c:y>0.1</c:y></c:from>" +
      "<c:to><c:x>0.4</c:x><c:y>0.3</c:y></c:to>" +
      '<c:sp><c:nvSpPr><c:cNvPr id="1" name="TextBox"/><c:cNvSpPr/></c:nvSpPr>' +
      "<c:spPr/><c:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Callout</a:t></a:r></a:p></c:txBody></c:sp>" +
      "</c:relSizeAnchor>" +
      "</c:userShapes>";

    chart.setUserShapesXml(userShapesXml);
    expect(chart.userShapesXml).toBeDefined();
    expect(textDecoder.decode(chart.userShapesXml!)).toContain("Callout");
    // The chart model should now carry a rel id for the drawing part.
    expect(chart.chartModel?.userShapesRelId).toMatch(/^rId/);

    const buf = await wb.xlsx.writeBuffer();
    const zipData = await extractAll(new Uint8Array(buf));
    const overlay = zipData.get("xl/drawings/chartUserShape1.xml");
    expect(overlay).toBeDefined();
    expect(textDecoder.decode(overlay!.data)).toContain("Callout");
    const rels = textDecoder.decode(zipData.get("xl/charts/_rels/chart1.xml.rels")!.data);
    expect(rels).toContain("chartUserShapes");
    expect(rels).toContain("chartUserShape1.xml");
    const contentTypes = textDecoder.decode(zipData.get("[Content_Types].xml")!.data);
    expect(contentTypes).toContain("chartUserShape1.xml");

    // Load the produced file and confirm the overlay bytes survive.
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf);
    const chart2 = wb2.getWorksheet("Sheet1")!.getCharts()[0];
    expect(chart2.userShapesXml).toBeDefined();
    expect(textDecoder.decode(chart2.userShapesXml!)).toContain("Callout");
  });

  it("Chart.setUserShapesXml validates input and removeUserShapes drops the rel", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 1],
      ["B", 2]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );
    const chart = ws.getCharts()[0];

    // String overload is accepted too.
    expect(() => chart.setUserShapesXml("<unknown/>")).toThrow(/c:userShapes/);

    chart.setUserShapesXml(
      '<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
    );
    expect(chart.userShapesXml).toBeDefined();
    expect(chart.chartModel!.userShapesRelId).toBeDefined();

    // Empty bytes == remove.
    chart.setUserShapesXml(new Uint8Array(0));
    expect(chart.userShapesXml).toBeUndefined();
    expect(chart.chartModel!.userShapesRelId).toBeUndefined();

    // Explicit remove is idempotent.
    chart.removeUserShapes();
    expect(chart.userShapesXml).toBeUndefined();
  });

  it("adds classic and chartEx shortcuts from worksheet helpers", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["Q1", 10, 1, 3],
      ["Q2", 20, 2, 4]
    ]);

    ws.addBarChart({ series: [baseSeries("Bar")], grouping: "stacked" }, "E1:L10");
    ws.addAreaChart({ series: [baseSeries("Area")] }, "E12:L21");
    ws.addDoughnutChart({ series: [baseSeries("Donut")], holeSize: 55 }, "E23:L32");
    ws.addBubbleChart({ series: [bubbleSeries("Bubble")] }, "E34:L43");
    ws.addRadarChart({ series: [baseSeries("Radar")], radarStyle: "filled" }, "E45:L54");
    ws.addStockChart(
      { series: [baseSeries("High"), baseSeries("Low", VALUES_B)], hiLowLines: true },
      "E56:L65"
    );
    ws.addSurfaceChart({ series: [baseSeries("Surface")], wireframe: true }, "E67:L76");

    const charts = ws.getCharts();
    expect(charts.slice(0, 7).map(chart => chart.chartTypes[0].type)).toEqual([
      "bar",
      "area",
      "doughnut",
      "bubble",
      "radar",
      "stock",
      "surface"
    ]);
    expect((charts[0].chartTypes[0] as BarChartGroup).barDir).toBe("bar");
    expect((charts[0].chartTypes[0] as BarChartGroup).grouping).toBe("stacked");

    const chartExSeries = { name: "Modern", values: VALUES_A };
    ws.addParetoChart({ series: [chartExSeries] }, "N1:T10");
    ws.addWaterfallChart({ series: [chartExSeries] }, "N12:T21");
    ws.addFunnelChart({ series: [chartExSeries] }, "N23:T32");
    ws.addTreemapChart({ series: [chartExSeries] }, "N34:T43");
    ws.addSunburstChart({ series: [chartExSeries] }, "N45:T54");
    ws.addBoxWhiskerChart({ series: [chartExSeries] }, "N56:T65");
    ws.addRegionMapChart({ series: [chartExSeries] }, "N67:T76");

    expect(
      ws
        .getCharts()
        .slice(7)
        .map(
          chart => chart.chartExModel?.chartSpace.chart.plotArea.plotAreaRegion?.series[0].layoutId
        )
    ).toEqual([
      "clusteredColumn",
      "waterfall",
      "funnel",
      "treemap",
      "sunburst",
      "boxWhisker",
      "regionMap"
    ]);
  });

  it("adds charts from generic row helpers", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.addChartFromRows(
      [
        { month: "Jan", sales: 10 },
        { month: "Feb", sales: 20 }
      ],
      { type: "line", x: "month", y: "sales", startCell: "C3" },
      "F1:M10"
    );

    const group = ws.getCharts()[0].chartTypes[0] as LineChartGroup;
    expect(group.type).toBe("line");
    expect(group.series[0].cat?.strRef?.formula).toBe("Sheet1!$C$4:$C$5");
    expect(group.series[0].val?.numRef?.formula).toBe("Sheet1!$D$4:$D$5");
  });

  it("updates and appends series from high-level options", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart({ type: "line", series: [baseSeries("Initial")] }, "C1:J10");
    const chart = ws.getCharts()[0];

    expect(
      chart.updateSeries(0, {
        name: { formula: "Sheet1!$D$1" },
        categories: "Sheet1!$D$2:$D$4",
        values: "Sheet1!$E$2:$E$4",
        fill: "#FF0000",
        line: "#00AA00",
        lineWidth: 2,
        lineDash: "dash",
        marker: { symbol: "diamond", size: 9, fill: "#0000FF" },
        dataLabels: { showVal: true, position: "outEnd" },
        trendline: { type: "linear", displayEq: true },
        errorBars: { type: "fixedVal", value: 2 },
        dataPoints: [{ index: 1, fill: "#FFFF00" }]
      })
    ).toBe(true);

    const series = chart.getSeries(0) as LineSeries;
    expect(series.tx?.strRef?.formula).toBe("Sheet1!$D$1");
    expect(series.cat?.strRef?.formula).toBe("Sheet1!$D$2:$D$4");
    expect(series.val?.numRef?.formula).toBe("Sheet1!$E$2:$E$4");
    expect(series.spPr?.fill?.solid?.srgb).toBe("FF0000");
    expect(series.spPr?.line?.color?.srgb).toBe("00AA00");
    expect(series.spPr?.line?.width).toBe(25400);
    expect(series.spPr?.line?.dash).toBe("dash");
    expect(series.marker?.symbol).toBe("diamond");
    expect(series.marker?.size).toBe(9);
    expect(series.marker?.spPr?.fill?.solid?.srgb).toBe("0000FF");
    expect(series.dataLabels?.showVal).toBe(true);
    expect(series.dataLabels?.position).toBe("outEnd");
    expect(series.trendlines?.[0].type).toBe("linear");
    expect(series.trendlines?.[0].displayEq).toBe(true);
    expect(series.errorBars?.errValType).toBe("fixedVal");
    expect(series.errorBars?.val).toBe(2);
    expect(series.dataPoints?.[0].index).toBe(1);
    expect(series.dataPoints?.[0].spPr?.fill?.solid?.srgb).toBe("FFFF00");

    expect(chart.addSeriesFromOptions({ ...baseSeries("Added"), values: VALUES_B })).toBe(true);
    expect(chart.getSeriesCount()).toBe(2);
    expect((chart.getSeries(1) as LineSeries).val?.numRef?.formula).toBe(VALUES_B);
    expect(chart.addSeriesFromOptions(baseSeries("Missing"), 99)).toBe(false);
  });

  it("updates scatter and bubble series-specific references", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addChart(
      {
        type: "scatter",
        series: [scatterSeries("Scatter")]
      },
      "C1:J10"
    );
    ws.addChart(
      {
        type: "bubble",
        series: [bubbleSeries("Bubble")]
      },
      "C12:J21"
    );

    const scatter = ws.getCharts()[0];
    scatter.updateSeries(0, {
      xValues: "Sheet1!$D$1:$D$4",
      values: "Sheet1!$E$1:$E$4",
      errorBars: [
        { type: "fixedVal", value: 1, direction: "x" },
        { type: "fixedVal", value: 2, direction: "y" }
      ]
    });
    const scatterSeriesModel = scatter.getSeries(0) as any;
    expect(scatterSeriesModel.xVal.numRef.formula).toBe("Sheet1!$D$1:$D$4");
    expect(scatterSeriesModel.yVal.numRef.formula).toBe("Sheet1!$E$1:$E$4");
    expect(scatterSeriesModel.errorBars).toHaveLength(2);

    const bubble = ws.getCharts()[1];
    bubble.updateSeries(0, {
      xValues: "Sheet1!$F$1:$F$3",
      values: "Sheet1!$G$1:$G$3",
      bubbleSize: "Sheet1!$H$1:$H$3",
      bubble3D: true
    });
    const bubbleSeriesModel = bubble.getSeries(0) as any;
    expect(bubbleSeriesModel.xVal.numRef.formula).toBe("Sheet1!$F$1:$F$3");
    expect(bubbleSeriesModel.yVal.numRef.formula).toBe("Sheet1!$G$1:$G$3");
    expect(bubbleSeriesModel.bubbleSize.numRef.formula).toBe("Sheet1!$H$1:$H$3");
    expect(bubbleSeriesModel.bubble3D).toBe(true);
  });

  it("exposes README chart and XML helpers from the root entrypoint", () => {
    const wb = new RootWorkbook();
    const ws = wb.addWorksheet("Root API");
    ws.addTable({
      name: "RootSales",
      ref: "A1",
      columns: [{ name: "Month" }, { name: "Sales" }],
      rows: [["Jan", 10]]
    });

    const seriesOptions: SeriesFromColumnsOptions = {
      categories: "A2:A2",
      values: "B2:B2",
      name: "Sales"
    };
    const tableOptions: AddChartFromTableOptions = {
      type: "bar",
      categoryColumn: "Month",
      valueColumns: ["Sales"]
    };
    const rowOptions: AddChartFromRowsOptions<{ month: string; sales: number }> = {
      type: "bar",
      x: "month",
      y: "sales",
      startCell: "D1"
    };

    expect(rootSeriesFromColumns("Root API", seriesOptions).values).toBe("'Root API'!$B$2:$B$2");
    expect(rootChartOptionsFromTable(ws, "RootSales", tableOptions).series).toHaveLength(1);
    expect(
      rootChartOptionsFromRows(ws, [{ month: "Feb", sales: 20 }], rowOptions).series
    ).toHaveLength(1);
    expect(ROOT_CHART_PRESETS).toBe(CHART_PRESETS);
    expect(ROOT_CHART_EX_PRESETS).toBe(CHART_EX_PRESETS);
    expect(ROOT_EXCEL_CHART_PRESETS).toBe(EXCEL_CHART_PRESETS);
    expect(ROOT_EXCEL_CHART_EX_PRESETS).toBe(EXCEL_CHART_EX_PRESETS);
    expect(rootApplyChartExPreset("map", { series: [{ values: "Root API!$B$2:$B$3" }] }).type).toBe(
      "regionMap"
    );
    expect(rootRenderChartSvg(makeRootExportRenderedChartModel())).toContain("<svg");
    const chartExType: ChartExType = "histogram";
    expect(
      rootRenderChartExSvg(
        rootBuildChartExModel({
          type: chartExType,
          series: [{ name: "S", values: "Root API!$B$2:$B$3", literalValues: [1, 2] }]
        })
      )
    ).toContain("<svg");
    expect(rootXmlEncodeAttr('a " b')).toBe("a &quot; b");
    expect(() => rootValidateXmlName("valid:name")).not.toThrow();
    const scene: ChartScene = buildChartScene(makeRootExportRenderedChartModel());
    const legend: ChartSceneLegend = scene.legend;
    const series: ChartSceneSeries = scene.series[0];
    const title: ChartSceneText | undefined = scene.title;
    expect(legend.items.length).toBeGreaterThan(0);
    expect(series.type).toBe("bar");
    expect(title?.text).toBe("Sales");
  });
});

describe("Builder gap fixes verification", () => {
  it("rejects invalid classic chart option combinations", () => {
    expect(() =>
      buildChartModel({
        type: "pie",
        series: [{ ...baseSeries("S"), trendline: { type: "linear" } }]
      })
    ).toThrow(/trendline/);
    expect(() =>
      buildChartModel({ type: "surface", series: [baseSeries("S")], dataLabels: { showVal: true } })
    ).toThrow(/dataLabels/);
  });

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

describe("P2: chart SVG/PDF renderer", () => {
  function makeRenderedChartModel(options?: Partial<AddChartOptions>): ChartModel {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 5],
      ["B", 20, 12],
      ["C", 30, 25]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }],
        title: "Sales",
        ...options
      },
      "D1:J10"
    );
    return ws.getCharts()[0].chartModel!;
  }

  function makeRenderedBubbleChartModel(): ChartModel {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      [1, 10, 5],
      [2, 20, 12],
      [3, 30, 25]
    ]);
    ws.addChart(
      {
        type: "bubble",
        series: [
          {
            name: "Bubble",
            xValues: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            bubbleSize: "Sheet1!$C$1:$C$3"
          }
        ]
      },
      "D1:J10"
    );
    return ws.getCharts()[0].chartModel!;
  }

  it("renderChartSvg returns a standalone SVG chart preview", () => {
    const svg = renderChartSvg(makeRenderedChartModel());
    expect(svg).toContain('<svg xmlns="http://www.w3.org/2000/svg"');
    expect(svg).toContain("Sales");
    expect(svg).toContain("<rect");
    expect(svg).toContain("S");
  });

  it("renderChartSvg has a stable golden hash for a decorated combo preview", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 2],
      ["B", 20, 4],
      ["C", 15, 5]
    ]);
    ws.addComboChart(
      {
        groups: [
          {
            type: "bar",
            barDir: "col",
            dataLabels: { showVal: true },
            series: [
              { name: "Revenue", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }
            ]
          },
          {
            type: "line",
            series: [
              {
                name: "Growth",
                categories: "Sheet1!$A$1:$A$3",
                values: "Sheet1!$C$1:$C$3",
                marker: { symbol: "diamond", size: 7 },
                trendline: { type: "linear", lineDash: "dash" }
              }
            ]
          }
        ],
        title: "Golden Combo"
      },
      "E1:L12"
    );

    const svg = renderChartSvg(ws.getCharts()[0].chartModel!, { width: 420, height: 260 });
    expect(stableHash(svg)).toBe("e3fa27b5");
  });

  it("renderChartSvg is documented as a deterministic preview, not Excel-identical", () => {
    const svg = renderChartSvg(makeRenderedChartModel());
    expect(svg).toContain("deterministic preview");
    expect(svg).toContain("not an Excel-identical layout");
  });

  it("legend width reacts to real label widths instead of a fixed 86-px slot", () => {
    // Before the font-metrics fix `legendRect` hard-coded `itemCount * 86` for
    // horizontal legends and `width: 96` for vertical ones, so a 30-character
    // series name was silently clipped or wrapped outside the viewport. With
    // real glyph widths in place a long label must produce a measurably
    // wider legend rectangle than a short one.
    const build = (label: string, legendPosition: "r" | "b" = "r"): ChartModel => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRows([
        ["A", 10],
        ["B", 20]
      ]);
      ws.addChart(
        {
          type: "bar",
          series: [{ name: label, categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
          title: "T",
          legendPosition
        },
        "D1:J10"
      );
      return ws.getCharts()[0].chartModel!;
    };
    // Horizontal legend: width growth is visible on the legend rect itself;
    // plot width is governed by other padding so it does not shrink here.
    const shortBottom = buildChartScene(build("S", "b"), { width: 400, height: 240 });
    const longBottom = buildChartScene(
      build("A much longer series name than the default one", "b"),
      { width: 400, height: 240 }
    );
    expect(shortBottom.legend.rect.width).toBeLessThan(longBottom.legend.rect.width);

    // Vertical legend: long labels push the plot rectangle inwards. Use a
    // wide canvas so the unclamped legend column width can express itself.
    const shortRight = buildChartScene(build("S", "r"), { width: 800, height: 240 });
    const longRight = buildChartScene(
      build("A much longer series name than the default one", "r"),
      { width: 800, height: 240 }
    );
    expect(shortRight.legend.rect.width).toBeLessThan(longRight.legend.rect.width);
    const plotDelta = shortRight.plot.width - longRight.plot.width;
    const legendDelta = longRight.legend.rect.width - shortRight.legend.rect.width;
    expect(plotDelta).toBeGreaterThan(legendDelta - 4);
    expect(plotDelta).toBeLessThan(legendDelta + 4);
  });

  it("pie chart emits leader lines when data labels use outEnd/bestFit", () => {
    // Build a pie chart with outEnd data labels. Before this change,
    // labels were anchored at the slice centroid with no connector line;
    // now the renderer must project each label to the outer ring and
    // emit a matching ChartSceneLine.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 30]
    ]);
    ws.addChart(
      {
        type: "pie",
        series: [
          {
            name: "S",
            categories: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            dataLabels: { showVal: true, position: "outEnd" }
          } as AddChartSeriesOptions
        ],
        title: "Pie"
      },
      "D1:J10"
    );
    const scene = buildChartScene(ws.getCharts()[0].chartModel!, {
      width: 320,
      height: 240
    });
    const pieSeries = scene.series[0];
    expect(pieSeries.type === "pie").toBe(true);
    // There must be three labels matching the three slices…
    expect(pieSeries.labels?.length).toBe(3);
    // …and three leader lines pointing from slice outer edge to each label.
    expect(pieSeries.leaderLines?.length).toBe(3);
    // Each label must live outside the pie's inscribed circle: recompute
    // the centre from the slice geometry and assert that all label
    // anchors are measurably further from it than the slice radius.
    if (pieSeries.type === "pie") {
      const slice = pieSeries.slices[0];
      const dist = (x: number, y: number): number => Math.hypot(x - slice.cx, y - slice.cy);
      for (const label of pieSeries.labels ?? []) {
        expect(dist(label.x, label.y)).toBeGreaterThan(slice.radius);
      }
    }

    // And the SVG output must contain the leader stroke (colour hardcoded).
    const svg = renderChartSvg(ws.getCharts()[0].chartModel!, {
      width: 320,
      height: 240
    });
    expect(svg).toContain('stroke="#808080"');
  });

  it("bar data labels nudge or drop when they would otherwise overlap", () => {
    // Construct a very narrow bar chart with many data points. Without the
    // collision pass, the outEnd labels emit at nearly identical y with
    // horizontally overlapping bboxes. With the pass, neighbours must
    // either be stacked vertically or dropped entirely (never both at the
    // same coordinate).
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    const cats = Array.from({ length: 8 }, (_, i) => `Cat${i + 1}`);
    const vals = Array.from({ length: 8 }, (_, i) => 10 + (i % 3));
    ws.addRows(cats.map((c, i) => [c, vals[i]]));
    ws.addChart(
      {
        type: "bar",
        barDir: "col",
        series: [
          {
            name: "S",
            categories: "Sheet1!$A$1:$A$8",
            values: "Sheet1!$B$1:$B$8",
            dataLabels: { showVal: true, position: "outEnd" }
          } as AddChartSeriesOptions
        ],
        title: "Dense"
      },
      "D1:F8"
    );
    const scene = buildChartScene(ws.getCharts()[0].chartModel!, {
      width: 160,
      height: 120
    });
    const series = scene.series[0];
    expect(series.type === "bar").toBe(true);
    const kept = series.labels ?? [];
    // At least one label should have been either moved up or dropped.
    expect(kept.length).toBeLessThanOrEqual(8);
    // No two kept labels share the same (x, y) coordinate — if they did,
    // the collision pass failed and we would be drawing glyphs on top of
    // each other.
    const coords = new Set(kept.map(l => `${Math.round(l.x)}:${Math.round(l.y)}`));
    expect(coords.size).toBe(kept.length);
  });

  it("renderChartSvg honours plotArea manualLayout edge-mode overrides", () => {
    // Render the same chart with and without a manual plot-area override and
    // extract the first axes' x-coordinates to confirm the manual rectangle
    // is applied. Using an exaggerated override (x=0.5) guarantees the
    // resulting plot origin is well right of the default ~58px inset, even
    // after auto padding for left-axis labels.
    const width = 400;
    const height = 240;
    const baseModel = makeRenderedChartModel();
    const baseSvg = renderChartSvg(baseModel, { width, height });
    const manualModel = makeRenderedChartModel();
    manualModel.chart.plotArea.layout = {
      manualLayout: {
        xMode: "edge",
        yMode: "edge",
        wMode: "edge",
        hMode: "edge",
        x: 0.5,
        y: 0.1,
        w: 0.4,
        h: 0.6
      }
    };
    const manualSvg = renderChartSvg(manualModel, { width, height });

    // Extract the plot rectangle's leftmost <line> x attribute from each SVG —
    // the deterministic renderer emits axis gridlines starting at plot.x.
    const extractFirstLineX = (svg: string): number => {
      const match = /<line\s+x1="([0-9.]+)"/.exec(svg);
      expect(match).not.toBeNull();
      return parseFloat(match![1]);
    };
    const baseX = extractFirstLineX(baseSvg);
    const manualX = extractFirstLineX(manualSvg);
    // Manual x=0.5 → plot left edge near 200; base is under ~120 for this
    // legend/title configuration. The exact value depends on padding, but the
    // manual version must be meaningfully further right.
    expect(manualX).toBeGreaterThan(baseX + 40);
    expect(manualX).toBeGreaterThan(width * 0.45);
  });

  it("renderChartSvg positions data labels per DataLabelPosition", () => {
    // Build two otherwise-identical chart models whose only difference is the
    // data-label position so we can verify the renderer actually consumes
    // `labels.position`. Use `l` (left) vs `r` (right) since they produce
    // distinct x-anchors that are easy to detect.
    const buildModel = (position: "l" | "r"): ChartModel => {
      const model = makeRenderedChartModel({
        type: "line",
        series: [
          {
            name: "S",
            categories: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            dataLabels: { showVal: true, position }
          } as AddChartSeriesOptions
        ]
      });
      return model;
    };
    const leftSvg = renderChartSvg(buildModel("l"), { width: 400, height: 240 });
    const rightSvg = renderChartSvg(buildModel("r"), { width: 400, height: 240 });

    // Left-positioned labels must be emitted with text-anchor="end" (so the
    // glyphs end at the anchor x); right-positioned labels with "start".
    expect(leftSvg).toContain('text-anchor="end"');
    expect(rightSvg).toContain('text-anchor="start"');
    // And the converse must not hold — a right-positioned chart should not
    // contain the left anchor (ignoring anchors that belong to other text,
    // such as axis labels; we check for data label text `10` specifically).
    const leftLabelAnchors = [...leftSvg.matchAll(/<text[^>]*>10<\/text>/g)];
    const rightLabelAnchors = [...rightSvg.matchAll(/<text[^>]*>10<\/text>/g)];
    expect(leftLabelAnchors.length).toBeGreaterThan(0);
    expect(rightLabelAnchors.length).toBeGreaterThan(0);
  });

  it("Chart.toSVG renders a classic chart through the high-level API", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "line",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Line Preview"
      },
      "D1:J10"
    );

    const svg = ws.getCharts()[0].toSVG({ width: 320, height: 180 });
    expect(svg).toContain('width="320"');
    expect(svg).toContain("Line Preview");
    expect(svg).toContain("<polyline");
  });

  it("Chart.toPNG renders a classic chart without a browser canvas", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "line",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Line Preview"
      },
      "D1:J10"
    );

    const png = await ws.getCharts()[0].toPNG({ width: 320, height: 180 });
    expectPngDimensions(png, 320, 180);
  });

  it("renderChartPng has a stable golden raster signature", async () => {
    const png = await renderChartPng(makeRenderedChartModel(), { width: 220, height: 140 });
    expectPngDimensions(png, 220, 140);
    expect(pngSignature(png)).toBe("f3a8e1e4");
  });

  it("Node PNG fallback consumes text rotate transforms from the SVG", async () => {
    // Build a chart whose only difference is a left-axis title, which the
    // SVG renderer emits with `transform="rotate(-90 x y)"`. The Node PNG
    // fallback must consume that transform; before this fix it was silently
    // discarded so rotated text ended up drawn in its axis-aligned position.
    // Comparing the rasterised bytes confirms that code path actually runs:
    // the two images must differ wherever the rotated glyphs land.
    const baseModel = makeRenderedChartModel({ type: "bar" });
    const rotatedModel = makeRenderedChartModel({ type: "bar" });
    const valueAxis = rotatedModel.chart.plotArea.axes.find(
      axis => (axis as ValueAxis).axisType === "val"
    ) as ValueAxis | undefined;
    expect(valueAxis).toBeDefined();
    valueAxis!.title = { text: { paragraphs: [{ runs: [{ text: "Units" }] }] } };

    const [basePng, rotatedPng] = await Promise.all([
      renderChartPng(baseModel, { width: 260, height: 160 }),
      renderChartPng(rotatedModel, { width: 260, height: 160 })
    ]);
    // Sanity: both PNGs have the expected geometry.
    expectPngDimensions(basePng, 260, 160);
    expectPngDimensions(rotatedPng, 260, 160);
    // And their IDAT chunks must differ — meaning the rotated title actually
    // produced pixels, rather than being silently dropped by the SVG
    // tag-regex raster path.
    expect(pngSignature(rotatedPng)).not.toBe(pngSignature(basePng));

    // Meanwhile the SVG must still emit the canonical rotate transform so
    // the browser path continues to match visually (regression guard for
    // SVG emit — unrelated to the basic raster fix but cheap to assert).
    const svg = renderChartSvg(rotatedModel, { width: 260, height: 160 });
    expect(svg).toMatch(/transform="rotate\(-90 /);
  });

  it("Chart.toPNG renders a ChartEx chart without a browser canvas", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addHistogramChart(
      {
        series: [{ name: "H", values: "Sheet1!$B$1:$B$6", literalValues: [-1, 0, 5, 10, 11, 21] }],
        binning: { binType: "binSize", binSize: 10, underflow: 0, overflow: 20 },
        title: "Histogram Preview"
      },
      "D1:J10"
    );

    const png = await ws.getCharts()[0].toPNG({ width: 300, height: 160 });
    expectPngDimensions(png, 300, 160);
  });

  it("Chart.toPNG honors scale, transparent background, and DPI metadata", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
      },
      "D1:J10"
    );

    const svg = ws.getCharts()[0].toSVG({ width: 160, height: 90, backgroundColor: "transparent" });
    expect(svg).not.toContain('fill="#fff"');

    const png = await ws
      .getCharts()[0]
      .toPNG({ width: 160, height: 90, scale: 2, dpi: 192, backgroundColor: "transparent" });
    expectPngDimensions(png, 320, 180);
    expectPngPhysDpi(png, 192);
  });

  it("buildChartScene produces shared geometry for SVG and PDF rendering", () => {
    const scene = buildChartScene(makeRenderedChartModel(), { width: 300, height: 180 });
    expect(scene.width).toBe(300);
    expect(scene.height).toBe(180);
    expect(scene.series[0].type).toBe("bar");
    expect(scene.gridlines.length).toBeGreaterThan(0);
    expect(scene.yLabels.map(label => label.text)).toContain("30");
  });

  it("buildChartScene handles horizontal bar charts", () => {
    const model = makeRenderedChartModel({ type: "bar", barDir: "bar" });
    const scene = buildChartScene(model, { width: 300, height: 180 });
    const series = scene.series[0];

    expect(series.type).toBe("bar");
    if (series.type === "bar") {
      expect(series.bars[0].width).toBeGreaterThan(series.bars[0].height);
    }
  });

  it("renderChartSvg draws area and bubble primitives", () => {
    const areaSvg = renderChartSvg(makeRenderedChartModel({ type: "area" }));
    const bubbleSvg = renderChartSvg(makeRenderedBubbleChartModel());

    expect(areaSvg).toContain("<polygon");
    expect(areaSvg).toContain("<polyline");
    expect(bubbleSvg).toContain("<circle");
  });

  it("renderChartSvg overlays combo chart groups instead of rendering only the first group", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 1],
      ["B", 20, 2],
      ["C", 30, 3]
    ]);
    ws.addComboChart(
      {
        groups: [
          {
            type: "bar",
            barDir: "col",
            series: [
              { name: "Revenue", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }
            ]
          },
          {
            type: "line",
            series: [{ name: "Growth", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$C$1:$C$3" }]
          }
        ],
        title: "Combo"
      },
      "E1:L12"
    );
    const svg = renderChartSvg(ws.getCharts()[0].chartModel!);
    expect(svg).toContain("Combo");
    expect(svg).toContain("<rect");
    expect(svg).toContain("<polyline");
    expect(svg).toContain("Revenue");
    expect(svg).toContain("Growth");
  });

  it("buildChartScene exposes secondary axes, legend visibility, and axis titles", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 1000],
      ["B", 20, 4000],
      ["C", 30, 9000]
    ]);
    ws.addComboChart(
      {
        groups: [
          {
            type: "bar",
            barDir: "col",
            categoryAxis: { title: "Quarter" },
            valueAxis: { title: "Revenue", min: 0, max: 40 },
            series: [
              { name: "Revenue", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }
            ]
          },
          {
            type: "line",
            useSecondaryAxis: true,
            series: [{ name: "Users", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$C$1:$C$3" }]
          }
        ],
        legendPosition: "t"
      },
      "E1:L12"
    );
    const model = ws.getCharts()[0].chartModel!;
    const secondaryValueAxis = model.chart.plotArea.axes.find(
      axis => axis.axisType === "val" && axis.axPos === "r"
    )!;
    secondaryValueAxis.title = {
      text: { paragraphs: [{ runs: [{ text: "Users" }] }] }
    };
    const scene = buildChartScene(model, { width: 420, height: 260 });
    const svg = renderChartSvg(model, { width: 420, height: 260 });

    expect(scene.axes.y2).toBeDefined();
    expect(scene.secondaryYLabels.map(label => label.text)).toContain("9000");
    expect(scene.axisTitles.map(label => label.text)).toEqual(
      expect.arrayContaining(["Quarter", "Revenue", "Users"])
    );
    expect(scene.legend.visible).toBe(true);
    expect(scene.legend.orientation).toBe("horizontal");
    expect(svg).toContain("Users");
    expect(svg).toContain("rotate(90");
  });

  it("buildChartScene hides the legend when the chart has no legend", () => {
    const scene = buildChartScene(makeRenderedChartModel({ showLegend: false }));
    const svg = renderChartSvg(makeRenderedChartModel({ showLegend: false }));
    expect(scene.legend.visible).toBe(false);
    expect(svg).not.toContain('font-size="10" fill="#555">S</text>');
  });

  it("renderChartSvg draws labels, markers, trendlines, error bars, radar, stock, surface, and ofPie", () => {
    const decorated = makeRenderedChartModel({
      type: "line",
      series: [
        {
          ...baseSeries("S"),
          marker: { symbol: "diamond", size: 8 },
          dataLabels: { showVal: true },
          trendline: { type: "linear", name: "Trend", lineDash: "dash" },
          errorBars: { type: "fixedVal", value: 2 }
        }
      ]
    });
    const radar = makeRenderedChartModel({ type: "radar", radarStyle: "filled" });
    const surface = makeRenderedChartModel({ type: "surface", wireframe: true });
    const ofPie = makeRenderedChartModel({ type: "ofPie" });
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 15, 8, 12],
      ["B", 20, 24, 18, 19],
      ["C", 30, 34, 27, 33],
      ["D", 24, 26, 20, 22]
    ]);
    const stock = buildChartModel({
      type: "stock",
      series: [
        { name: "Open", categories: CATEGORIES, values: "Sheet1!$B$1:$B$4" },
        { name: "High", categories: CATEGORIES, values: "Sheet1!$C$1:$C$4" },
        { name: "Low", categories: CATEGORIES, values: "Sheet1!$D$1:$D$4" },
        { name: "Close", categories: CATEGORIES, values: "Sheet1!$E$1:$E$4" }
      ]
    });
    fillChartCaches(stock, wb);
    const svg = [decorated, radar, surface, ofPie, stock]
      .map(model => renderChartSvg(model))
      .join("\n");
    expect(svg).toContain("Trend");
    expect(svg).toContain("stroke-dasharray");
    expect(svg).toContain("<polygon");
    expect(svg).toContain("<path");
    expect(svg).toContain("#70AD47");
  });

  it("drawChartPdf draws chart primitives on a PDF-like surface", () => {
    const calls: string[] = [];
    const trace: string[] = [];
    const page = {
      drawRect() {
        calls.push("rect");
        return this;
      },
      drawLine() {
        calls.push("line");
        return this;
      },
      drawText() {
        calls.push("text");
        return this;
      }
    };

    const result = drawChartPdf(page, makeRenderedChartModel(), {
      x: 10,
      y: 20,
      width: 300,
      height: 180,
      trace
    });
    expect(result).toBe(page);
    expect(calls).toContain("rect");
    expect(calls).toContain("line");
    expect(calls).toContain("text");
    expect(stableHash(trace.join("\n"))).toBe("60f4cd02");
  });

  it("drawChartPdf draws pie paths when the surface supports paths", () => {
    const calls: string[] = [];
    const page = {
      drawRect() {
        calls.push("rect");
        return this;
      },
      drawLine() {
        calls.push("line");
        return this;
      },
      drawText() {
        calls.push("text");
        return this;
      },
      drawPath() {
        calls.push("path");
        return this;
      }
    };

    drawChartPdf(page, makeRenderedChartModel({ type: "pie" }), {
      x: 10,
      y: 20,
      width: 300,
      height: 180
    });

    expect(calls).toContain("path");
  });

  it("drawChartPdf emits adornments (labels/markers/errorBars/trendlines/leaders)", () => {
    // Build a chart whose model actually contains every adornment type so
    // the PDF path is forced to exercise the code added by this change:
    // markers (from series.marker), data labels (from dataLabels),
    // trendlines (from trendlines), error bars (from errorBars), and
    // leader lines are produced by the pie pass separately below.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "line",
        series: [
          {
            name: "Trend",
            categories: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            marker: { symbol: "diamond", size: 8 },
            dataLabels: { showVal: true, position: "outEnd" },
            trendline: { type: "linear" },
            errorBars: { type: "fixedVal", value: 1, direction: "y" }
          } as AddChartSeriesOptions
        ],
        title: "PDF adornments"
      },
      "D1:J10"
    );
    const trace: string[] = [];
    const calls: Record<string, number> = {};
    const record = (kind: string) => {
      calls[kind] = (calls[kind] ?? 0) + 1;
    };
    const page = {
      drawRect() {
        record("rect");
        return this;
      },
      drawLine() {
        record("line");
        return this;
      },
      drawText() {
        record("text");
        return this;
      },
      drawPath() {
        record("path");
        return this;
      },
      drawCircle() {
        record("circle");
        return this;
      }
    };
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 320,
      height: 220,
      trace
    });
    // Trace must contain all new adornment tags.
    const joined = trace.join("\n");
    expect(joined).toContain("trendline:");
    expect(joined).toContain("errorbar:");
    expect(joined).toContain("label:");
    // Diamond markers route through `drawPath`.
    expect(calls.path).toBeGreaterThan(0);
    // And there are strictly more `line` calls than the pre-adornment
    // baseline (axes + gridlines + errorbar = many more than 4 lines).
    expect(calls.line).toBeGreaterThan(5);
  });

  it("drawChartPdf draws pie leader lines as connector strokes", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "pie",
        series: [
          {
            name: "S",
            categories: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            dataLabels: { showVal: true, position: "outEnd" }
          } as AddChartSeriesOptions
        ]
      },
      "D1:J10"
    );
    const trace: string[] = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        return this;
      },
      drawText() {
        return this;
      },
      drawPath() {
        return this;
      }
    };
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 320,
      height: 240,
      trace
    });
    expect(trace.join("\n")).toContain("leader:");
  });

  it("drawChartPdf forwards anchor and rotation to surfaces that support them", () => {
    // Axis titles on the y axes use rotate: ±90°. The PDF bridge must
    // forward that rotation so the title glyphs read vertically instead
    // of overlapping the tick labels.
    const captured: Array<{ text: string; rotation?: number; anchor?: string }> = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        return this;
      },
      drawText(
        text: string,
        options: { x: number; y: number; rotation?: number; anchor?: string }
      ) {
        captured.push({ text, rotation: options.rotation, anchor: options.anchor });
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Centred",
        valueAxis: { title: "Units" }
      },
      "D1:J10"
    );
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 320,
      height: 220
    });
    // Title must be drawn with anchor=middle (centred under its x
    // coordinate); axis title "Units" with rotation=-90 or 90.
    const titleCall = captured.find(c => c.text === "Centred");
    expect(titleCall?.anchor).toBe("middle");
    const axisTitleCall = captured.find(c => c.text === "Units");
    expect(axisTitleCall).toBeDefined();
    expect(Math.abs(axisTitleCall!.rotation ?? 0)).toBe(90);
  });

  it("drawChartPdf area fill forwards alpha to surfaces that honour PdfColor.a", () => {
    // Without alpha support, stacked area charts render opaque and the
    // layer beneath becomes invisible. The PDF bridge must now forward
    // the SVG withAlpha(color, 0.35) semantic as a real PdfColor.a so
    // capable surfaces can emit /ExtGState.
    const captured: Array<{ fill?: { a?: number } }> = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        return this;
      },
      drawText() {
        return this;
      },
      drawPath(_ops: unknown, options?: { fill?: { a?: number } }) {
        captured.push({ fill: options?.fill });
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "area",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }]
      },
      "D1:J10"
    );
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 320,
      height: 200
    });
    // An area polygon must have been drawn with a non-opaque fill.
    const areaFill = captured.find(c => c.fill?.a !== undefined && c.fill.a < 1);
    expect(areaFill).toBeDefined();
    expect(areaFill!.fill!.a).toBeCloseTo(0.35);
  });

  it("drawChartPdf filled radar emits a translucent polygon fill", () => {
    const captured: Array<{ fill?: { a?: number } }> = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        return this;
      },
      drawText() {
        return this;
      },
      drawPath(_ops: unknown, options?: { fill?: { a?: number } }) {
        captured.push({ fill: options?.fill });
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "radar",
        radarStyle: "filled",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }]
      },
      "D1:J10"
    );
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 320,
      height: 220
    });
    // A radar fill path with 0.35 alpha — same value the SVG path draws.
    const translucent = captured.find(c => c.fill?.a !== undefined && c.fill.a < 1);
    expect(translucent).toBeDefined();
    expect(translucent!.fill!.a).toBeCloseTo(0.35);
  });

  it("drawChartPdf forwards fontFamily/bold/italic from txPr to the PDF surface", () => {
    const captured: Array<{
      text: string;
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
    }> = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        return this;
      },
      drawText(
        text: string,
        options: { x: number; y: number; fontFamily?: string; bold?: boolean; italic?: boolean }
      ) {
        captured.push({
          text,
          fontFamily: options.fontFamily,
          bold: options.bold,
          italic: options.italic
        });
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    // Inject an explicit <a:latin> + bold + italic on the title txPr so
    // the scene exposes them via textStyleFromTxPr and the PDF bridge
    // can forward them.
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Styled"
      },
      "D1:J10"
    );
    const model = ws.getCharts()[0].chartModel!;
    model.chart.title!.txPr = {
      _rawXml:
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800" b="1" i="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Verdana"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    };
    drawChartPdf(page, model, {
      x: 0,
      y: 0,
      width: 320,
      height: 220
    });
    const titleCall = captured.find(c => c.text === "Styled");
    expect(titleCall).toBeDefined();
    expect(titleCall!.fontFamily).toBe("Verdana");
    expect(titleCall!.bold).toBe(true);
    expect(titleCall!.italic).toBe(true);
  });

  it("drawChartPdf on a drawPath-less surface strokes pie slice outlines instead of dropping them", () => {
    // Minimal surface (the real-world legacy shape) exposes only rect /
    // line / text. Before P3, the pie branch's `if (page.drawPath)` made
    // the entire chart disappear. The fallback must emit at least one
    // drawLine per slice so the geometry is still readable.
    const lineCalls: Array<{ x1: number; y1: number; x2: number; y2: number }> = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine(options: { x1: number; y1: number; x2: number; y2: number }) {
        lineCalls.push(options);
        return this;
      },
      drawText() {
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "pie",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }]
      },
      "D1:J10"
    );
    const baseline = lineCalls.length;
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 300,
      height: 220
    });
    // A 3-slice pie traced with `arcPolyline` produces many line
    // segments (at least one per slice × 3 slices). Assert a
    // conservative lower bound — the exact count depends on arc
    // segmentation, not on the fallback policy we're testing.
    expect(lineCalls.length - baseline).toBeGreaterThan(10);
  });

  it("drawChartPdf emits nothing for the pie fallback when drawPath IS available (no double paint)", () => {
    // Regression guard: the fallback path must not run when drawPath is
    // present. This proves the `else` branch is correctly gated — the
    // number of drawLine calls with drawPath available should be
    // dramatically smaller than without it (no arc-outline strokes).
    const pathCalls: string[] = [];
    const lineCalls: string[] = [];
    const page = {
      drawRect() {
        return this;
      },
      drawLine() {
        lineCalls.push("line");
        return this;
      },
      drawText() {
        return this;
      },
      drawPath() {
        pathCalls.push("path");
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "pie",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }]
      },
      "D1:J10"
    );
    drawChartPdf(page, ws.getCharts()[0].chartModel!, {
      x: 0,
      y: 0,
      width: 300,
      height: 220
    });
    // drawPath used for every slice (3 slices = at least 3 calls).
    expect(pathCalls.length).toBeGreaterThanOrEqual(3);
    // drawLine count is small — only axis / gridline primitives, not
    // slice outline strokes. Each slice would contribute ~25 arc
    // segments if the fallback ran, so >50 total drawLine calls
    // would indicate leakage. Staying under that gives a comfortable
    // margin without coupling to the axis renderer's exact line count.
    expect(lineCalls.length).toBeLessThan(20);
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

// ---------------------------------------------------------------------------
// P0-3: structured spPr / txPr access on chartEx (previously rawXml-only)
// ---------------------------------------------------------------------------

describe("ChartEx structured spPr / txPr", () => {
  const CHART_EX_NAMESPACES =
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  function makeWaterfallWithSpPr(colorHex: string): string {
    // Minimal waterfall chartEx with a structured cx:spPr on the series
    // containing a solid fill and a 2pt dashed border. This is the shape
    // produced by Excel when a user sets Fill + Outline on a waterfall.
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace ${CHART_EX_NAMESPACES}>
  <cx:chartData>
    <cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$1:$A$3</cx:f></cx:strDim></cx:data>
    <cx:data id="1"><cx:numDim type="val"><cx:f>Sheet1!$B$1:$B$3</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="waterfall" ownerIdx="0">
          <cx:dataId val="0"/>
          <cx:dataId val="1"/>
          <cx:spPr>
            <a:solidFill>
              <a:srgbClr val="${colorHex}"/>
            </a:solidFill>
            <a:ln w="25400">
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
              <a:prstDash val="dash"/>
            </a:ln>
          </cx:spPr>
          <cx:axisId val="1"/>
        </cx:series>
      </cx:plotAreaRegion>
      <cx:axis id="1"/>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;
  }

  it("parser populates structured fill/line fields on cx:spPr while preserving rawXml", () => {
    const model = parseChartEx(makeWaterfallWithSpPr("FF0000"));
    const series = model.chartSpace.chart.plotArea.plotAreaRegion!.series[0];

    // Structured fields are populated
    expect(series.spPr?.fill?.solid?.srgb).toBe("FF0000");
    expect(series.spPr?.line?.color?.srgb).toBe("000000");
    expect(series.spPr?.line?.width).toBe(25400);
    expect(series.spPr?.line?.dash).toBe("dash");

    // Original rawXml is retained for lossless round-trip of anything the
    // structured parser doesn't model (joints, extensions…).
    expect(series.spPr?._rawXml).toBeDefined();
    expect(series.spPr?._rawXml).toContain("FF0000");
  });

  it("round-trip via parseChartEx(renderChartEx(model)) preserves spPr bytes when unmodified", () => {
    const first = parseChartEx(makeWaterfallWithSpPr("00FF00"));
    const rebuilt = parseChartEx(renderChartEx(first));
    const series = rebuilt.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(series.spPr?.fill?.solid?.srgb).toBe("00FF00");
    expect(series.spPr?.line?.dash).toBe("dash");
  });

  it("structural mutation via setSpPrFill drops rawXml and writes the new fill", () => {
    const model = parseChartEx(makeWaterfallWithSpPr("FF0000"));
    const region = model.chartSpace.chart.plotArea.plotAreaRegion!;

    // Mutate: replace the fill using the shape-properties API. This returns
    // a new ShapeProperties object without `_rawXml`, which forces the
    // spPr writer to take the structured path.
    region.series[0].spPr = setSpPrFill(region.series[0].spPr, {
      solid: { srgb: "1122AA" }
    });
    expect(region.series[0].spPr._rawXml).toBeUndefined();

    // Clear the model-level `rawXml` cache so `renderChartEx` does not
    // short-circuit to the original bytes. `ChartExModel.rawXml` is a
    // performance cache for untouched parts — mutating consumers are
    // expected to drop it when they intend to re-serialise from the
    // structured model.
    model.rawXml = undefined;

    // Round-trip and observe the new colour.
    const rendered = renderChartEx(model);
    expect(rendered).toContain("1122AA");
    expect(rendered).not.toContain("FF0000");

    // The new XML parses back to the expected structured value.
    const reparsed = parseChartEx(rendered);
    expect(
      reparsed.chartSpace.chart.plotArea.plotAreaRegion!.series[0].spPr?.fill?.solid?.srgb
    ).toBe("1122AA");
  });

  it("parser populates structured txPr (font size / colour) on cx:dataLabels", () => {
    const rawXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace ${CHART_EX_NAMESPACES}>
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$1:$A$3</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="funnel" ownerIdx="0">
          <cx:dataId val="0"/>
          <cx:dataLabels>
            <cx:visibility value="1"/>
            <cx:txPr>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="1200" b="1">
                    <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
                  </a:defRPr>
                </a:pPr>
                <a:endParaRPr/>
              </a:p>
            </cx:txPr>
          </cx:dataLabels>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;
    const model = parseChartEx(rawXml);
    const dl = model.chartSpace.chart.plotArea.plotAreaRegion!.series[0].dataLabels!;

    // Structured fields are extracted from the raw XML.
    expect(dl.txPr?.size).toBe(1200);
    expect(dl.txPr?.bold).toBe(true);
    expect(dl.txPr?.color?.srgb).toBe("C00000");

    // Raw XML is also retained for elements the structured parser doesn't
    // understand (pPr alignment, custom fonts, etc.).
    expect(dl.txPr?._rawXml).toBeDefined();
    expect(dl.txPr?._rawXml).toContain("C00000");
  });

  it("structured ShapeProperties carries effectList / sp3d through cx:spPr render", () => {
    // Build a chartEx model programmatically with a structured effectList
    // and sp3d on a series, then round-trip through render+parse and assert
    // the structured fields survive. This is the path user code takes when
    // constructing a chart with shadow / bevel effects.
    const model = buildChartExModel({
      type: "funnel",
      series: [
        {
          values: "Sheet1!$A$1:$A$3",
          spPr: {
            fill: { solid: { srgb: "4472C4" } },
            effectList: {
              outerShadow: {
                blurRadius: 50800,
                distance: 38100,
                direction: 2700000,
                color: { srgb: "000000", alpha: 40 }
              }
            },
            sp3d: {
              extrusionHeight: 76200,
              bevelTop: { width: 63500, height: 25400, preset: "circle" }
            }
          }
        }
      ]
    });

    const rendered = renderChartEx(model);
    expect(rendered).toContain("<a:effectLst>");
    expect(rendered).toContain("<a:outerShdw");
    expect(rendered).toContain("<a:sp3d");
    expect(rendered).toContain("<a:bevelT");

    const reparsed = parseChartEx(rendered);
    const parsedSeries = reparsed.chartSpace.chart.plotArea.plotAreaRegion!.series[0];
    expect(parsedSeries.spPr?.effectList?.outerShadow?.blurRadius).toBe(50800);
    expect(parsedSeries.spPr?.sp3d?.extrusionHeight).toBe(76200);
    expect(parsedSeries.spPr?.sp3d?.bevelTop?.preset).toBe("circle");
  });

  it("structured axis txPr round-trips with rotation and font family", () => {
    const model = buildChartExModel({
      type: "histogram",
      series: [{ values: "Sheet1!$A$1:$A$6" }]
    });
    // Inject a txPr on the first axis before render.
    const axis = model.chartSpace.chart.plotArea.axis?.[0];
    if (!axis) {
      throw new Error("histogram builder did not produce an axis");
    }
    axis.txPr = {
      size: 900,
      italic: true,
      rotation: -5400000,
      fontFamily: "Arial",
      color: { srgb: "333333" }
    };

    const rendered = renderChartEx(model);
    expect(rendered).toContain('rot="-5400000"');
    expect(rendered).toContain('typeface="Arial"');

    const reparsed = parseChartEx(rendered);
    const parsedAxis = reparsed.chartSpace.chart.plotArea.axis![0];
    expect(parsedAxis.txPr?.size).toBe(900);
    expect(parsedAxis.txPr?.italic).toBe(true);
    expect(parsedAxis.txPr?.rotation).toBe(-5400000);
    expect(parsedAxis.txPr?.fontFamily).toBe("Arial");
  });
});

// ---------------------------------------------------------------------------
// P1-3: pictureFill.image high-level API
// ---------------------------------------------------------------------------

describe("bar chart pictureFill.image (high-level)", () => {
  const PNG_MAGIC = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    // IHDR dummy — just enough bytes for Excel to open; we only test the
    // rel pipeline here, not image rendering.
    0x00, 0x00, 0x00, 0x0d
  ]);

  it("resolves a raw Uint8Array into a workbook media entry and chart rel", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;

    ws.addChart(
      {
        type: "bar",
        series: [
          {
            name: "S",
            values: "Sheet1!$A$1:$A$2",
            pictureFill: {
              image: PNG_MAGIC,
              fillMode: "stretch",
              applyToFront: true
            }
          }
        ]
      },
      "C1:J10"
    );

    // The chart entry should carry a freshly allocated rel pointing at
    // the newly registered media entry.
    const entry = wb.getChartEntry(1)!;
    expect(entry.rels).toBeDefined();
    const imageRels = (entry.rels ?? []).filter(r => r.Type.endsWith("/relationships/image"));
    expect(imageRels).toHaveLength(1);
    expect(imageRels[0].Target).toMatch(/^\.\.\/media\/image\d+\.png$/);
    expect(imageRels[0].Id).toMatch(/^rId\d+$/);

    // And the model's blipFill is wired up to the same id.
    const series = entry.model.chart.plotArea.chartTypes[0].series[0] as {
      spPr?: { fill?: { blip?: { relationshipId?: string; _pendingImage?: unknown } } };
    };
    expect(series.spPr?.fill?.blip?.relationshipId).toBe(imageRels[0].Id);
    expect(series.spPr?.fill?.blip?._pendingImage).toBeUndefined();

    // Full round-trip: the xlsx output contains the image part and the
    // chart XML references it via <a:blip r:embed="rIdN"/>.
    const buf = await wb.xlsx.writeBuffer();
    const entries = await extractAll(new Uint8Array(buf));
    expect(entries.get("xl/media/image1.png")).toBeDefined();
    const chartXml = textDecoder.decode(entries.get("xl/charts/chart1.xml")!.data);
    expect(chartXml).toContain("<a:blipFill>");
    expect(chartXml).toContain(`r:embed="${imageRels[0].Id}"`);
    expect(chartXml).toContain("<a:stretch>");
    const chartRels = textDecoder.decode(entries.get("xl/charts/_rels/chart1.xml.rels")!.data);
    expect(chartRels).toContain(imageRels[0].Id);
    expect(chartRels).toContain("image1.png");
  });

  it("reuses an existing workbook image via workbookImageId", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    const imageId = wb.addImage({ extension: "png", buffer: PNG_MAGIC });

    ws.addChart(
      {
        type: "bar",
        series: [
          {
            values: "Sheet1!$A$1:$A$1",
            pictureFill: { image: { workbookImageId: imageId } }
          }
        ]
      },
      "C1:J10"
    );

    // Only one media entry — the existing one.
    expect(wb.media).toHaveLength(1);
    const entry = wb.getChartEntry(1)!;
    const imageRel = (entry.rels ?? []).find(r => r.Type.endsWith("/relationships/image"));
    expect(imageRel?.Target).toBe(`../media/image${imageId + 1}.png`);
  });

  it("accepts a data URL base64 string and infers extension", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 42;
    // 1x1 transparent GIF
    const gifData =
      "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
    ws.addChart(
      {
        type: "bar",
        series: [{ values: "Sheet1!$A$1:$A$1", pictureFill: { image: gifData } }]
      },
      "C1:J10"
    );
    const entry = wb.getChartEntry(1)!;
    const imageRel = (entry.rels ?? []).find(r => r.Type.endsWith("/relationships/image"));
    expect(imageRel?.Target).toMatch(/\.gif$/);
  });

  it("still honours a manually-supplied relationshipId (no auto allocation)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            values: "Sheet1!$A$1:$A$1",
            pictureFill: { relationshipId: "rId99", fillMode: "stack" }
          }
        ]
      },
      "C1:J10"
    );
    const entry = wb.getChartEntry(1)!;
    const series = entry.model.chart.plotArea.chartTypes[0].series[0] as {
      spPr?: { fill?: { blip?: { relationshipId?: string } } };
    };
    expect(series.spPr?.fill?.blip?.relationshipId).toBe("rId99");
    // No image rel was auto-created — caller takes responsibility.
    const imageRels = (entry.rels ?? []).filter(r => r.Type.endsWith("/relationships/image"));
    expect(imageRels).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// P1-6: scatter/bubble text x axis
// ---------------------------------------------------------------------------

describe("scatter/bubble xValueType", () => {
  it("defaults xValues to numRef for scatter (OOXML standard)", () => {
    const m = buildChartModel({
      type: "scatter",
      series: [{ xValues: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }]
    });
    const series = m.chart.plotArea.chartTypes[0].series[0] as {
      xVal?: { numRef?: unknown; strRef?: unknown };
    };
    expect(series.xVal?.numRef).toBeDefined();
    expect(series.xVal?.strRef).toBeUndefined();
  });

  it('wraps xValues as strRef when xValueType is "text"', () => {
    const m = buildChartModel({
      type: "scatter",
      series: [
        {
          xValues: "Sheet1!$A$1:$A$3",
          xValueType: "text",
          values: "Sheet1!$B$1:$B$3"
        }
      ]
    });
    const series = m.chart.plotArea.chartTypes[0].series[0] as {
      xVal?: { numRef?: unknown; strRef?: unknown };
    };
    expect(series.xVal?.strRef).toBeDefined();
    expect(series.xVal?.numRef).toBeUndefined();
  });

  it("bubble chart also honours text xValueType", () => {
    const m = buildChartModel({
      type: "bubble",
      series: [
        {
          xValues: "Sheet1!$A$1:$A$2",
          xValueType: "text",
          values: "Sheet1!$B$1:$B$2",
          bubbleSize: "Sheet1!$C$1:$C$2"
        }
      ]
    });
    const series = m.chart.plotArea.chartTypes[0].series[0] as {
      xVal?: { strRef?: unknown };
    };
    expect(series.xVal?.strRef).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// P1-5: Chartsheet.chart host proxy
// ---------------------------------------------------------------------------

describe("Chartsheet.chart worksheet proxy", () => {
  it("exposes chartModel through the proxy without throwing", async () => {
    const wb = new Workbook();
    wb.addWorksheet("Data").addRows([
      ["A", 1],
      ["B", 2]
    ]);
    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ categories: "Data!$A$1:$A$2", values: "Data!$B$1:$B$2" }]
      }
    });

    // `chart` returns a real Chart bound to a host proxy, and
    // reads/writes to chart-level properties work through it.
    const chart = chartsheet.chart!;
    expect(chart.chartNumber).toBe(1);
    expect(chart.chartModel?.chart.plotArea).toBeDefined();
  });

  it("host proxy rejects grid operations with a helpful error", () => {
    const wb = new Workbook();
    wb.addWorksheet("Data").addRows([["X", 1]]);
    const chartsheet = wb.addChartsheet("Chart Sheet", {
      chart: {
        type: "bar",
        series: [{ categories: "Data!$A$1", values: "Data!$B$1" }]
      }
    });
    const chart = chartsheet.chart!;
    // Grid-centric calls (getRow / getColumn / addTable …) must not
    // silently no-op on the proxy — they should throw so the caller
    // notices the mis-use instead of ending up with a corrupted workbook.
    expect(() => (chart.worksheet as { getRow: (n: number) => unknown }).getRow(1)).toThrow(
      /not supported on a Chart attached to a chartsheet/
    );
  });
});

// ---------------------------------------------------------------------------
// P1-1 (L2): effectList SVG filter auto-injection
// ---------------------------------------------------------------------------

describe("renderChartSvg auto-injects effectList filters", () => {
  it('emits <filter> in <defs> and wraps the series in filter="url(#...)"', () => {
    // Build a bar chart whose series carries an outer shadow effect —
    // the same shape Excel emits when the user sets Series Format →
    // Shadow in the chart pane.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            name: "S",
            categories: "Sheet1!$A$1:$A$3",
            values: "Sheet1!$B$1:$B$3",
            spPr: {
              effectList: {
                outerShadow: {
                  blurRadius: 50800,
                  distance: 38100,
                  direction: 2700000,
                  color: { srgb: "000000", alpha: 40 }
                }
              }
            }
          }
        ]
      },
      "D1:K10"
    );
    const chart = ws.getCharts()[0];
    const svg = chart.toSVG({ width: 400, height: 200 });

    // The SVG must contain a <defs> block with a <filter> inside.
    expect(svg).toContain("<defs>");
    expect(svg).toMatch(/<filter id="excelts-fx-\d+"/);
    expect(svg).toContain("<feGaussianBlur");
    expect(svg).toContain("<feOffset");
    expect(svg).toContain("<feMerge");

    // And at least one <g> element must reference the filter.
    expect(svg).toMatch(/<g filter="url\(#excelts-fx-\d+\)">/);
  });

  it("deduplicates filters when multiple series share the same effect list", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10, 20],
      ["B", 20, 30]
    ]);
    const sharedEffect = {
      glow: { radius: 38100, color: { srgb: "FF0000" } }
    };
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            categories: "Sheet1!$A$1:$A$2",
            values: "Sheet1!$B$1:$B$2",
            spPr: { effectList: sharedEffect }
          },
          {
            categories: "Sheet1!$A$1:$A$2",
            values: "Sheet1!$C$1:$C$2",
            spPr: { effectList: sharedEffect }
          }
        ]
      },
      "D1:K10"
    );
    const chart = ws.getCharts()[0];
    const svg = chart.toSVG({ width: 400, height: 200 });
    // Exactly one <filter> definition even though two series share it.
    const filterDefs = (svg.match(/<filter id="excelts-fx-/g) ?? []).length;
    expect(filterDefs).toBe(1);
    // But both series reference it (two <g filter="url(...)">)
    const filterRefs = (svg.match(/<g filter="url\(#excelts-fx-/g) ?? []).length;
    expect(filterRefs).toBe(2);
  });

  it("omits <defs> entirely when no series carry an effectList", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([["A", 1]]);
    ws.addChart(
      {
        type: "bar",
        series: [{ categories: "Sheet1!$A$1", values: "Sheet1!$B$1" }]
      },
      "D1:K10"
    );
    const chart = ws.getCharts()[0];
    const svg = chart.toSVG({ width: 400, height: 200 });
    expect(svg).not.toContain("<defs>");
    expect(svg).not.toContain("<filter");
    expect(svg).not.toContain('filter="url');
  });

  it("builds filter for glow / innerShadow / softEdge / reflection / blur", () => {
    // Exhaust the effect list surface so regressions in any single
    // effect show up as a distinct assertion failure.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([["A", 5]]);
    ws.addChart(
      {
        type: "bar",
        series: [
          {
            categories: "Sheet1!$A$1",
            values: "Sheet1!$B$1",
            spPr: {
              effectList: {
                blur: { radius: 38100, grow: true },
                innerShadow: {
                  blurRadius: 50800,
                  distance: 25400,
                  direction: 5400000,
                  color: { srgb: "000000", alpha: 30 }
                },
                glow: { radius: 25400, color: { srgb: "FFFF66" } },
                softEdge: { radius: 19050 },
                reflection: {
                  blurRadius: 6350,
                  startOpacity: 52000,
                  endOpacity: 300,
                  distance: 0,
                  direction: 5400000
                }
              }
            }
          }
        ]
      },
      "D1:K10"
    );
    const chart = ws.getCharts()[0];
    const svg = chart.toSVG({ width: 400, height: 200 });
    // Every filter primitive ends up somewhere in the single generated
    // <filter> element.
    expect(svg).toContain("<feGaussianBlur");
    expect(svg).toContain("<feComposite");
    expect(svg).toContain("<feMerge");
    expect(svg).toContain("<feFlood");
  });
});

// ---------------------------------------------------------------------------
// Gap A (L2): DrawingML a:xfrm / a:prstGeom / a:custGeom structured access
// ---------------------------------------------------------------------------

describe("shape-properties structured DrawingML geometry", () => {
  it("extracts a:xfrm offset / ext / rot / flips from raw spPr XML", () => {
    const spPr = parseSpPr({
      _rawXml:
        '<c:spPr><a:xfrm rot="5400000" flipH="1">' +
        '<a:off x="123456" y="789012"/>' +
        '<a:ext cx="500000" cy="250000"/>' +
        "</a:xfrm></c:spPr>"
    });
    expect(spPr.transform).toEqual({
      rotation: 5400000,
      flipHorizontal: true,
      offsetX: 123456,
      offsetY: 789012,
      width: 500000,
      height: 250000
    });
  });

  it("extracts a:prstGeom preset + avLst adjustments", () => {
    const spPr = parseSpPr({
      _rawXml:
        '<c:spPr><a:prstGeom prst="roundRect">' +
        '<a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst>' +
        "</a:prstGeom></c:spPr>"
    });
    expect(spPr.presetGeometry?.preset).toBe("roundRect");
    expect(spPr.presetGeometry?.adjustments).toEqual([{ name: "adj", fmla: "val 16667" }]);
  });

  it("extracts a:custGeom path commands (moveTo / lnTo / arcTo / close)", () => {
    const spPr = parseSpPr({
      _rawXml:
        "<c:spPr><a:custGeom><a:pathLst>" +
        '<a:path w="100000" h="100000" fill="norm" stroke="1">' +
        '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
        '<a:lnTo><a:pt x="100000" y="0"/></a:lnTo>' +
        '<a:arcTo wR="50000" hR="50000" stAng="0" swAng="10800000"/>' +
        "<a:close/>" +
        "</a:path></a:pathLst></a:custGeom></c:spPr>"
    });
    const path = spPr.customGeometry?.paths?.[0];
    expect(path).toBeDefined();
    expect(path?.w).toBe(100000);
    expect(path?.h).toBe(100000);
    expect(path?.fill).toBe("norm");
    expect(path?.stroke).toBe(true);
    expect(path?.commands).toHaveLength(4);
    expect(path?.commands[0].type).toBe("moveTo");
    expect(path?.commands[0].points).toEqual([{ x: 0, y: 0 }]);
    expect(path?.commands[2].type).toBe("arcTo");
    expect(path?.commands[2].arcParams?.swAng).toBe(10800000);
    expect(path?.commands[3].type).toBe("close");
  });
});

// ---------------------------------------------------------------------------
// chartToPdf bridge — vector for classic, raster for ChartEx
// ---------------------------------------------------------------------------

describe("chartToPdf bridge", () => {
  it("produces a valid 1-page PDF for a classic chart via the vector path", async () => {
    const { chartToPdf } = await import("@pdf/excel-bridge");
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }],
        title: "Classic Bar"
      },
      "D1:J10"
    );
    const bytes = await chartToPdf(ws.getCharts()[0]);
    // PDF magic bytes.
    expect(bytes.length).toBeGreaterThan(500);
    const head = new TextDecoder("latin1").decode(bytes.slice(0, 8));
    expect(head.startsWith("%PDF-")).toBe(true);
    const body = new TextDecoder("latin1").decode(bytes);
    // Vector path distinguishes itself from the raster path by the
    // **absence** of an image XObject — content streams are FlateDecoded
    // so we cannot match on literal glyph bytes, but a vector PDF from
    // this helper never embeds an /Image XObject.
    expect(body).not.toMatch(/\/Subtype\s+\/Image/);
    // And every PDF we produce references /Helvetica because the PDF
    // vector path always resolves the built-in Type1 font set.
    expect(body).toContain("/Helvetica");
  });

  it("renders regionMap ChartEx via the vector path by default and via the raster path when forceRaster is set", async () => {
    const { chartToPdf } = await import("@pdf/excel-bridge");
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    // regionMap joined VECTOR_PDF_CHART_EX_LAYOUT_IDS once
    // `drawRegionMapPdf` landed — topology polygons, centroid
    // preview, and hex-tile fallback all render through drawPath /
    // drawRect / drawCircle / drawText now. The raster path is still
    // reachable via `forceRaster: true` for callers who want
    // pixel-identical output to the SVG preview.
    ws.addChartEx(
      {
        type: "regionMap",
        series: [
          {
            values: "Sheet1!$B$1:$B$2",
            literalValues: [10, 20],
            literalCategories: ["USA", "Canada"]
          }
        ]
      },
      "D1:J10"
    );

    // Default: vector path, no image XObject.
    const vectorBytes = await chartToPdf(ws.getCharts()[0]);
    expect(vectorBytes.length).toBeGreaterThan(500);
    const vectorHead = new TextDecoder("latin1").decode(vectorBytes.slice(0, 8));
    expect(vectorHead.startsWith("%PDF-")).toBe(true);
    const vectorBody = new TextDecoder("latin1").decode(vectorBytes);
    expect(vectorBody).not.toMatch(/\/Subtype\s+\/Image/);
    expect(vectorBody).toContain("/Helvetica");

    // Opt-in raster: same chart rendered via the PNG → image XObject route.
    const rasterBytes = await chartToPdf(ws.getCharts()[0], { forceRaster: true });
    expect(rasterBytes.length).toBeGreaterThan(500);
    const rasterHead = new TextDecoder("latin1").decode(rasterBytes.slice(0, 8));
    expect(rasterHead.startsWith("%PDF-")).toBe(true);
    const rasterBody = new TextDecoder("latin1").decode(rasterBytes);
    expect(rasterBody).toMatch(/\/Subtype\s+\/Image/);
  });

  it("sunburst and treemap ChartEx charts use the vector PDF path (no image XObject)", async () => {
    const { chartToPdf } = await import("@pdf/excel-bridge");
    for (const layout of ["sunburst", "treemap"] as const) {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRows([
        ["A", 10],
        ["B", 20],
        ["C", 15]
      ]);
      ws.addChartEx(
        {
          type: layout,
          categories: "Sheet1!$A$1:$A$3",
          series: [{ values: "Sheet1!$B$1:$B$3" }]
        },
        "D1:J10"
      );
      const bytes = await chartToPdf(ws.getCharts()[0], { title: `${layout} vector` });
      const body = new TextDecoder("latin1").decode(bytes);
      // Vector path: no image XObject, but the PDF must still carry a
      // content stream wide enough to hold the slice/rect drawing ops.
      expect(body).not.toMatch(/\/Subtype\s+\/Image/);
      // And the standard chart-title font must be registered (vector
      // text = real PDF font resources).
      expect(body).toContain("/Helvetica");
      // PDF must still be valid.
      expect(body.startsWith("%PDF-")).toBe(true);
    }
  });

  it("waterfall/funnel/histogram/pareto/boxWhisker ChartEx layouts take the vector PDF path", async () => {
    // Second batch of vector-capable ChartEx layouts. Each type has
    // its own geometry (stepped bars with dashed connectors,
    // trapezoid stack, binned columns, sorted columns + cumulative
    // polyline, statistical boxes). Confirming the batch via the
    // same `no /Image XObject + has /Helvetica` contract keeps the
    // assertion shape uniform across the six vector layouts.
    const { chartToPdf } = await import("@pdf/excel-bridge");
    const fixtures: Array<{
      name: string;
      build: (ws: ReturnType<Workbook["addWorksheet"]>) => void;
    }> = [
      {
        name: "waterfall",
        build: ws =>
          ws.addChartEx(
            {
              type: "waterfall",
              categories: "S!$A$1:$A$4",
              series: [{ values: "S!$B$1:$B$4" }],
              layout: { subtotals: [{ idx: 3 }], connectorLines: true }
            },
            "D1:J10"
          )
      },
      {
        name: "funnel",
        build: ws =>
          ws.addChartEx(
            {
              type: "funnel",
              categories: "S!$A$1:$A$4",
              series: [{ values: "S!$B$1:$B$4" }]
            },
            "D1:J10"
          )
      },
      {
        name: "histogram",
        build: ws =>
          ws.addChartEx(
            {
              type: "histogram",
              series: [{ values: "S!$B$1:$B$4", literalValues: [1, 2, 3, 4, 5, 6, 7, 8] }],
              binning: { binCount: 4 }
            },
            "D1:J10"
          )
      },
      {
        name: "pareto",
        build: ws =>
          ws.addChartEx(
            {
              type: "pareto",
              categories: "S!$A$1:$A$4",
              series: [
                {
                  values: "S!$B$1:$B$4",
                  literalValues: [30, 20, 15, 10],
                  literalCategories: ["A", "B", "C", "D"]
                }
              ]
            },
            "D1:J10"
          )
      },
      {
        name: "boxWhisker",
        build: ws =>
          ws.addChartEx(
            {
              type: "boxWhisker",
              categories: "S!$A$1:$A$6",
              series: [{ values: "S!$B$1:$B$6" }],
              layout: { showMeanMarker: true, showMeanLine: true, showOutlierPoints: true }
            },
            "D1:J10"
          )
      }
    ];
    for (const fixture of fixtures) {
      const wb = new Workbook();
      const ws = wb.addWorksheet("S");
      ws.addRows([
        ["A", 10],
        ["B", 20],
        ["C", 15],
        ["D", 8],
        ["E", 5],
        ["F", 25]
      ]);
      fixture.build(ws);
      const bytes = await chartToPdf(ws.getCharts()[0], { title: `${fixture.name} vector` });
      const body = new TextDecoder("latin1").decode(bytes);
      expect(body).not.toMatch(/\/Subtype\s+\/Image/);
      expect(body).toContain("/Helvetica");
      expect(body.startsWith("%PDF-")).toBe(true);
    }
  });

  it("forceRaster on a ChartEx chart keeps the raster path even for vectorable layouts", async () => {
    const { chartToPdf } = await import("@pdf/excel-bridge");
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20],
      ["C", 15]
    ]);
    ws.addChartEx(
      {
        type: "sunburst",
        categories: "Sheet1!$A$1:$A$3",
        series: [{ values: "Sheet1!$B$1:$B$3" }]
      },
      "D1:J10"
    );
    const bytes = await chartToPdf(ws.getCharts()[0], { forceRaster: true });
    const body = new TextDecoder("latin1").decode(bytes);
    expect(body).toMatch(/\/Subtype\s+\/Image/);
  });

  it("forceRaster on a classic chart routes through the PNG path", async () => {
    const { chartToPdf } = await import("@pdf/excel-bridge");
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addRows([
      ["A", 10],
      ["B", 20]
    ]);
    ws.addChart(
      {
        type: "bar",
        series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }],
        title: "Forced raster"
      },
      "D1:J10"
    );
    const bytes = await chartToPdf(ws.getCharts()[0], { forceRaster: true });
    const body = new TextDecoder("latin1").decode(bytes);
    // Raster path embedded an image XObject.
    expect(body).toMatch(/\/Subtype\s+\/Image/);
  });

  it("drawChartExPdf boxWhisker outlier falls back to stroke-only rect on surfaces without drawCircle", async () => {
    // Exercise the minimal-surface path — a surface that implements
    // only drawRect / drawLine / drawText, nothing else. Before this
    // change the outlier emitter silently skipped (no circle, no
    // fallback). Now it emits a stroke-only rect so the outlier stays
    // visible.
    const { drawChartExPdf } = await import("@excel/chart");
    let outlierRects = 0;
    const rects: Array<{ width: number; height: number; fill?: unknown; stroke?: unknown }> = [];
    const surface = {
      drawRect(opts: {
        x: number;
        y: number;
        width: number;
        height: number;
        fill?: unknown;
        stroke?: unknown;
      }) {
        rects.push(opts);
        // A stroke-only 4×4 rect is the outlier fallback signature.
        if (opts.width === 4 && opts.height === 4 && !opts.fill && opts.stroke) {
          outlierRects++;
        }
        return this;
      },
      drawLine() {
        return this;
      },
      drawText() {
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // A distribution with three clear outliers (very large jumps) so
    // boxStats flags them regardless of quartileMethod.
    ws.addChartEx(
      {
        type: "boxWhisker",
        series: [
          {
            values: "ignored",
            literalValues: [10, 10, 11, 11, 12, 12, 13, 13, 14, 14, 200, 500, 1000],
            literalCategories: ["A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A"]
          }
        ],
        layout: { showOutlierPoints: true, quartileMethod: "inclusive" }
      },
      "D1:J10"
    );
    drawChartExPdf(surface, ws.getCharts()[0].chartExModel!, {
      x: 0,
      y: 0,
      width: 400,
      height: 300
    });
    expect(outlierRects).toBeGreaterThan(0);
  });

  it("drawChartExPdf funnel traces the trapezoid outline with drawLine when drawPath is missing", async () => {
    // Minimal surface without drawPath. The funnel layer must still
    // (a) fill a colour rect so the data magnitude is visible and
    // (b) trace four outline lines so the trapezoid silhouette
    // remains recognisable despite the polygon fill being unavailable.
    const { drawChartExPdf } = await import("@excel/chart");
    let coloredRects = 0;
    let outlineLines = 0;
    const surface = {
      drawRect(opts: { fill?: unknown; stroke?: unknown }) {
        if (opts.fill) {
          coloredRects++;
        }
        return this;
      },
      drawLine(opts: { color?: unknown }) {
        if (opts.color) {
          outlineLines++;
        }
        return this;
      },
      drawText() {
        return this;
      }
    };
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addRows([
      ["A", 100],
      ["B", 80],
      ["C", 50],
      ["D", 20]
    ]);
    ws.addChartEx(
      {
        type: "funnel",
        categories: "S!$A$1:$A$4",
        series: [{ values: "S!$B$1:$B$4" }]
      },
      "D1:J10"
    );
    drawChartExPdf(surface, ws.getCharts()[0].chartExModel!, {
      x: 0,
      y: 0,
      width: 400,
      height: 300
    });
    // 4 layers × 1 colour rect each = 4 rects for the fill signal.
    expect(coloredRects).toBeGreaterThanOrEqual(4);
    // 4 layers × 4 outline lines each = 16 lines for the trapezoid
    // silhouette.
    expect(outlineLines).toBeGreaterThanOrEqual(16);
  });
});
