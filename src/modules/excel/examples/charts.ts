/**
 * Charts Example — end-to-end coverage of the Excel chart API.
 *
 * One runnable script that walks through every public surface of the
 * `@excel/chart` module and its worksheet/workbook integration points:
 *
 *  1.  Classic chart types (bar/col, bar3D, line, line3D, pie, pie3D,
 *      doughnut, area, area3D, scatter, bubble, radar, stock, surface,
 *      surface3D, ofPie).
 *  2.  Chart presets — representative sample + the full 99-preset catalogue.
 *  3.  ChartEx types (histogram, pareto, waterfall, funnel, treemap,
 *      sunburst, boxWhisker, regionMap).
 *  4.  Combo charts with secondary axis, data-table, up-down bars.
 *  5.  Convenience series builders: `addChartFromTable` /
 *      `addChartFromRows`, their ChartEx siblings, `seriesFromColumns`.
 *  6.  Axes: value/category axis, log scale, custom units, date axis,
 *      rotated labels, hidden axis.
 *  7.  Rich titles, legend entries, plot-area layout, picture fills.
 *  8.  Chart styling via `setStyle(N)` and structured `chartStyle` /
 *      `chartColors` sidecars.
 *  9.  Chartsheets (full-page charts) + custom page setup + replaceChart.
 * 10.  Pivot charts + pivot chartsheets.
 * 11.  Anchor variations — one-cell, two-cell, absolute.
 * 12.  Feature deep-dive — every trendline type, every error-bar type,
 *      every marker symbol, dropLines/serLines/upDownBars/hiLowLines,
 *      gapWidth/overlap/gapDepth, firstSliceAng/holeSize, valueFromCells,
 *      per-entry data labels, gradient/pattern fills, xValueType: "text",
 *      displayBlanksAs/plotVisOnly/showDLblsOverMax, date axis, null title.
 * 13.  Headless charts — `literalValues` / `literalCategories` for
 *      charts whose data does NOT live on the worksheet.
 * 14.  Chart editing — `addSeries` / `removeSeries` / `updateSeries` /
 *      `addSeriesFromOptions` / `setStyle` / `getSeries` / `chartTypes` /
 *      `removeUserShapes` / `mutateChartEx`.
 * 15.  Low-level builders — `buildChartModel` / `buildComboChartModel` /
 *      `buildChartExModel` invoked directly for programmatic workflows.
 * 16.  `workbook.addImage` + picture fill via `workbookImageId`.
 * 17.  Full preset catalogue — all 99 classic presets on one sheet so
 *      you can see what each name produces.
 * 18.  Custom chartsheet page setup + `replaceChart`.
 * 19.  Advanced shape / effect hooks — chart-level `spPr`, full
 *      `EffectList` (shadow / glow / softEdge / reflection / blur),
 *      `Scene3D` + `ShapeProperties3D` + `Bevel`, manual plot-area
 *      layout (edge mode), chart-level `PictureOptions`, surface
 *      `bandFormats`.
 * 20.  Chart copying + vendor extensions + external parser —
 *      `chart.copyTo(targetWorksheet)`, `chart.unknownElements` for
 *      `c15:` / `cx14:` passthrough, `parseChartEx(xml)` against a
 *      synthetic XML string.
 * 21.  Chartsheet lifecycle — `rename` / `state` / `copy` / `remove` /
 *      `zoomScale` / `zoomToFit` / `replaceChart`.
 * 22.  Sparklines — `addSparklineGroup` for line / column / stacked
 *      (win-loss) in-cell mini charts with grouped axis scales.
 * 23.  Rare chart features — multi-level categorical axis
 *      (`multiLvlStrRef`), view3D `depthPercent` / `hPercent`, ChartEx
 *      `rawLayoutId` vendor passthrough, `PresetGeometry` +
 *      `CustomGeometry` + `ShapeTransform` on series `spPr`.
 * 24.  Shape / text helpers — `parseSpPr` / `buildSpPr` /
 *      `getSpPrFillColor` / `getSpPrLine` / `getSpPrGradient` /
 *      `getSpPrPattern` / `setSpPrFill` / `setSpPrLine` / `parseTxPr` /
 *      `buildTxPr` / `getTxPrFontSize` / `getTxPrColor`.
 * 25.  Sidecar parsers / writers — `parseChartStyle` / `buildChartStyle`
 *      and `parseChartColors` / `buildChartColors`, plus attaching the
 *      round-tripped sidecars to an actual chart.
 * 26.  Cache populator API — `fillChartCaches` / `fillChartExCaches` /
 *      `fillNumRef` / `fillStrRef` called directly against programmatic
 *      chart models.
 * 27.  Low-level rendering & TopoJSON — direct `renderChartSvg` /
 *      `renderChartPng` / `renderChartExSvg` / `renderChartExPng`;
 *      `buildChartScene` IR; `applyAxisTransform`; `buildEffectFilter`;
 *      `VECTOR_PDF_CHART_EX_LAYOUT_IDS`; `resolveTopologyObject` against
 *      a synthetic two-country TopoJSON; `applyChartPreset` /
 *      `applyChartExPreset` direct invocation.
 * 28.  Preview export — `chart.toSVG()`, `chart.toPNG()`, `chartToPdf()`.
 * 29.  Loading + mutating a chart (`chart.mutate(fn, { preferRawPatch })`).
 *
 * Output:
 *   tmp/charts-example.xlsx     — one workbook containing every chart
 *   tmp/charts-example-<N>.svg  — per-chart SVG previews
 *   tmp/charts-example-<N>.png  — per-chart PNG previews
 *   tmp/charts-example.pdf      — multi-page PDF of each chart
 *
 * Usage:
 *   pnpm exec tsx src/modules/excel/examples/charts.ts
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import type { AddSparklineGroupOptions, SparklineGroup } from "@excel/sparkline";
import { PdfDocumentBuilder } from "@pdf/builder/document-builder";
import { chartToPdf } from "@pdf/excel-bridge";

import {
  Workbook,
  EXCEL_CHART_PRESETS,
  EXCEL_CHART_EX_PRESETS,
  CHART_PRESETS,
  buildChartModel,
  buildComboChartModel,
  buildChartExModel,
  canRenderChartExAsVectorPdf,
  parseChartEx,
  applyChartPreset,
  applyChartExPreset,
  fillChartCaches,
  fillChartExCaches,
  fillNumRef,
  fillStrRef,
  parseChartStyle,
  buildChartStyle,
  parseChartColors,
  buildChartColors,
  parseSpPr,
  buildSpPr,
  getSpPrFillColor,
  getSpPrLine,
  getSpPrGradient,
  getSpPrPattern,
  setSpPrFill,
  setSpPrLine,
  parseTxPr,
  buildTxPr,
  getTxPrFontSize,
  getTxPrColor,
  resolveTopologyObject,
  renderChartSvg,
  renderChartPng,
  renderChartExSvg,
  renderChartExPng,
  buildChartScene,
  buildEffectFilter,
  applyAxisTransform,
  VECTOR_PDF_CHART_EX_LAYOUT_IDS,
  type AddChartSeriesOptions,
  type AddChartOptions,
  type ChartStyleModel,
  type ChartColorsModel,
  type ChartRichText,
  type AddTrendlineOptions,
  type AddErrorBarsOptions,
  type ChartMarker,
  type RegionMapDataOptions,
  type TopologyLike,
  type NumberReference,
  type StringReference
} from "../../../index";
import { drawChartExPdf } from "../chart/chart-ex-renderer";
import { drawChartPdf } from "../chart/chart-renderer";

const OUT_DIR = resolve(process.cwd(), "tmp");
const XLSX_PATH = resolve(OUT_DIR, "charts-example.xlsx");
const PDF_PATH = resolve(OUT_DIR, "charts-example.pdf");

mkdirSync(OUT_DIR, { recursive: true });

// Shared palette — Excel's default chart-colors "cycle" sequence. Used by
// the feature-showcase sheet when we need a deterministic colour per
// series.
const HEX_PALETTE = [
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
  "264478",
  "9E480E",
  "C00000"
];

// ---------------------------------------------------------------------------
// Data sheets — shared by most charts on the dashboard
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  const wb = new Workbook();
  wb.title = "ExcelTS — Chart Gallery";
  wb.creator = "ExcelTS charts example";
  wb.created = new Date();

  // "Sales" — month × (revenue, profit, cost, units, growth) grid used by
  // every bar / column / line / area / pivotable chart below. Kept small
  // enough to fit in printouts yet rich enough to exercise trendlines,
  // error bars and secondary axes.
  const sales = wb.addWorksheet("Sales");
  sales.columns = [
    { header: "Month", key: "month", width: 12 },
    { header: "Revenue", key: "revenue", width: 14 },
    { header: "Profit", key: "profit", width: 14 },
    { header: "Cost", key: "cost", width: 14 },
    { header: "Units", key: "units", width: 10 },
    { header: "Growth %", key: "growth", width: 12 }
  ];
  sales.addRows([
    { month: "Jan", revenue: 120, profit: 32, cost: 88, units: 540, growth: 0.04 },
    { month: "Feb", revenue: 180, profit: 49, cost: 131, units: 620, growth: 0.08 },
    { month: "Mar", revenue: 160, profit: 41, cost: 119, units: 610, growth: 0.02 },
    { month: "Apr", revenue: 205, profit: 64, cost: 141, units: 710, growth: 0.12 },
    { month: "May", revenue: 232, profit: 75, cost: 157, units: 805, growth: 0.13 },
    { month: "Jun", revenue: 248, profit: 81, cost: 167, units: 880, growth: 0.07 }
  ]);
  sales.getColumn("revenue").numFmt = "$#,##0";
  sales.getColumn("profit").numFmt = "$#,##0";
  sales.getColumn("cost").numFmt = "$#,##0";
  sales.getColumn("growth").numFmt = "0.0%";

  // "Scatter" — x / y pairs + bubble sizes. Dedicated sheet so scatter /
  // bubble charts have the right shape without interfering with the
  // category-axis data on the Sales sheet.
  const scatter = wb.addWorksheet("Scatter");
  scatter.addRows([
    ["Hours", "Score", "Attempts"],
    [1, 45, 3],
    [2, 55, 5],
    [3, 62, 4],
    [4, 70, 7],
    [5, 78, 6],
    [6, 85, 9],
    [7, 90, 8]
  ]);

  // "Stock" — date / open / high / low / close for a classic OHLC stock
  // chart. Dates are real Date objects so Excel renders them as a date
  // axis.
  const stockSheet = wb.addWorksheet("Stock");
  stockSheet.addRow(["Date", "Open", "High", "Low", "Close"]);
  const d = (iso: string): Date => new Date(iso);
  stockSheet.addRows([
    [d("2024-01-02"), 150, 156, 149, 154],
    [d("2024-01-03"), 154, 158, 152, 157],
    [d("2024-01-04"), 157, 161, 155, 159],
    [d("2024-01-05"), 159, 162, 156, 160],
    [d("2024-01-08"), 160, 165, 158, 164]
  ]);
  stockSheet.getColumn(1).numFmt = "yyyy-mm-dd";

  // "Surface" — z values on a cat × series grid so surface / surface3D
  // charts have the 2-D matrix they expect.
  const surfaceSheet = wb.addWorksheet("Surface");
  surfaceSheet.addRow(["", "Q1", "Q2", "Q3", "Q4"]);
  surfaceSheet.addRow(["North", 10, 22, 34, 40]);
  surfaceSheet.addRow(["South", 15, 25, 37, 48]);
  surfaceSheet.addRow(["East", 12, 20, 30, 42]);
  surfaceSheet.addRow(["West", 18, 28, 39, 50]);

  // "Hierarchy" — 3-level category data for sunburst / treemap.
  const hier = wb.addWorksheet("Hierarchy");
  hier.addRows([
    ["Region", "Country", "City", "Sales"],
    ["Americas", "USA", "New York", 420],
    ["Americas", "USA", "San Francisco", 360],
    ["Americas", "Brazil", "São Paulo", 210],
    ["EMEA", "UK", "London", 310],
    ["EMEA", "Germany", "Berlin", 280],
    ["EMEA", "France", "Paris", 240],
    ["APAC", "Japan", "Tokyo", 400],
    ["APAC", "China", "Shanghai", 520],
    ["APAC", "Australia", "Sydney", 180]
  ]);

  // "Distribution" — raw samples for histogram / pareto / boxWhisker.
  const dist = wb.addWorksheet("Distribution");
  dist.addRow(["Sample"]);
  // 40 synthetic samples clustered around two modes to give the
  // histogram a bi-modal shape and the box-whisker chart visible
  // whiskers + outliers.
  const samples = [
    52, 54, 57, 59, 60, 61, 62, 63, 63, 64, 65, 66, 67, 68, 69, 70, 72, 73, 75, 76, 78, 80, 82, 84,
    85, 86, 88, 90, 92, 94, 28, 35, 42, 96, 98, 102, 110, 115, 118, 122
  ];
  samples.forEach(n => dist.addRow([n]));

  // "Regions" — country → value for a region-map chart.
  const regions = wb.addWorksheet("Regions");
  regions.addRows([
    ["Country", "Revenue"],
    ["United States", 840],
    ["Germany", 410],
    ["Japan", 380],
    ["Brazil", 220],
    ["India", 310],
    ["Australia", 150]
  ]);

  // ---------------------------------------------------------------------------
  // 1. Classic chart gallery — every type Excel exposes (16 types).
  // ---------------------------------------------------------------------------

  const gallery = wb.addWorksheet("1-Classic Gallery");
  gallery.addRow(["Type", "Notes"]);

  // Compact helper that stamps a label row, then positions the chart in
  // a 9-column × 15-row block to the right of the label so the gallery
  // reads left-to-right.
  let galleryRow = 2;
  const galleryChart = (
    label: string,
    options: AddChartOptions,
    opts?: { rows?: number }
  ): number => {
    gallery.getCell(`A${galleryRow}`).value = label;
    gallery.getCell(`A${galleryRow}`).font = { bold: true };
    const rows = opts?.rows ?? 15;
    const top = galleryRow;
    const bottom = galleryRow + rows;
    galleryRow = bottom + 1;
    return gallery.addChart(options, `B${top}:J${bottom}`);
  };

  // -- Bar / Column family
  galleryChart("bar (horizontal)", {
    type: "bar",
    barDir: "bar",
    grouping: "clustered",
    title: "Horizontal bar",
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7", fill: "4472C4" }
    ],
    categoryAxis: { title: "Month" },
    valueAxis: { title: "USD", numFmt: "$#,##0" }
  });

  galleryChart("bar (column, clustered)", {
    type: "bar",
    barDir: "col",
    grouping: "clustered",
    title: "Revenue vs Cost",
    gapWidth: 120,
    overlap: -10,
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7", fill: "4472C4" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7", fill: "ED7D31" }
    ],
    dataLabels: { showVal: true, position: "outEnd", numFmt: "$0" },
    valueAxis: { title: "USD", majorGridlines: true, numFmt: "$#,##0" }
  });

  galleryChart("bar stacked", {
    type: "bar",
    barDir: "col",
    grouping: "stacked",
    title: "Revenue breakdown (stacked)",
    series: [
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7", fill: "70AD47" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7", fill: "FFC000" }
    ],
    legendPosition: "b"
  });

  galleryChart("bar percent-stacked", {
    type: "bar",
    barDir: "col",
    grouping: "percentStacked",
    title: "Revenue mix (100% stacked)",
    series: [
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
    ]
  });

  galleryChart("bar3D (extruded)", {
    type: "bar3D",
    barDir: "col",
    grouping: "clustered",
    shape: "box",
    title: "Extruded box 3D",
    view3D: { rotX: 15, rotY: 20, rAngAx: true, perspective: 30 },
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7", fill: "4472C4" }
    ],
    floor: { fill: "F2F2F2" },
    backWall: { fill: "FFFFFF", border: "BFBFBF" },
    sideWall: { fill: "FAFAFA" }
  });

  galleryChart("bar3D (cylinder)", {
    type: "bar3D",
    barDir: "col",
    grouping: "clustered",
    shape: "cylinder",
    title: "3D cylinder bars",
    series: [{ name: "Units", categories: "Sales!$A$2:$A$7", values: "Sales!$E$2:$E$7" }]
  });

  // -- Line family
  galleryChart("line (with markers)", {
    type: "line",
    title: "Revenue trend",
    smooth: true,
    showMarker: true,
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        line: "4472C4",
        lineWidth: 2.5,
        marker: { symbol: "circle", size: 8, fill: "4472C4", border: "FFFFFF" },
        trendline: {
          type: "linear",
          displayEq: true,
          displayRSqr: true,
          lineDash: "dash",
          line: "ED7D31",
          forward: 1
        },
        errorBars: {
          direction: "y",
          barDir: "both",
          type: "percentage",
          value: 10,
          line: "A6A6A6"
        }
      }
    ],
    valueAxis: { min: 0, max: 300, majorUnit: 50, majorGridlines: true, numFmt: "$#,##0" }
  });

  galleryChart("line stacked", {
    type: "line",
    grouping: "stacked",
    title: "Stacked line",
    showMarker: true,
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }
    ]
  });

  galleryChart("line3D", {
    type: "line3D",
    title: "Line (3D)",
    view3D: { rotX: 20, rotY: 20 },
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
    ]
  });

  // -- Pie / Doughnut / OfPie family
  galleryChart("pie", {
    type: "pie",
    title: "Revenue by month",
    firstSliceAng: 10,
    varyColors: true,
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        dataLabels: { showPercent: true, showCatName: true, position: "outEnd", separator: " • " },
        // Pull two slices out of the centre
        dataPoints: [
          { index: 0, explosion: 15, fill: "4472C4" },
          { index: 3, explosion: 25, fill: "ED7D31" }
        ]
      }
    ]
  });

  galleryChart("pie3D", {
    type: "pie3D",
    title: "3D pie",
    view3D: { rotX: 30, rotY: 0, perspective: 15 },
    series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
  });

  galleryChart("doughnut", {
    type: "doughnut",
    title: "Doughnut (hole=60)",
    holeSize: 60,
    varyColors: true,
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        dataLabels: { showVal: true, numFmt: "$#,##0", position: "ctr" }
      }
    ],
    legendPosition: "r"
  });

  galleryChart("ofPie (pie-of-pie)", {
    type: "ofPie",
    ofPieType: "pie",
    splitType: "pos",
    splitPos: 3,
    secondPieSize: 75,
    title: "Pie-of-pie",
    series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
  });

  // -- Area family
  galleryChart("area", {
    type: "area",
    title: "Area chart",
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        spPr: { fill: { solid: { srgb: "4472C4", alpha: 60000 } } }
      }
    ]
  });

  galleryChart("area stacked", {
    type: "area",
    grouping: "stacked",
    title: "Area (stacked)",
    series: [
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
    ]
  });

  galleryChart("area3D", {
    type: "area3D",
    title: "Area (3D)",
    view3D: { rotX: 15, rotY: 20 },
    series: [
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" },
      { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
    ]
  });

  // -- Scatter / Bubble
  galleryChart("scatter (smooth, markers)", {
    type: "scatter",
    scatterStyle: "smoothMarker",
    title: "Study hours vs score",
    series: [
      {
        name: "Students",
        xValues: "Scatter!$A$2:$A$8",
        values: "Scatter!$B$2:$B$8",
        marker: { symbol: "diamond", size: 9, fill: "5B9BD5", border: "FFFFFF" },
        trendline: { type: "poly", order: 2, line: "ED7D31", lineDash: "sysDash" }
      }
    ],
    categoryAxis: { title: "Hours studied" },
    valueAxis: { title: "Score" }
  });

  galleryChart("bubble", {
    type: "bubble",
    title: "Hours × score (bubble = attempts)",
    bubbleScale: 80,
    showNegBubbles: false,
    sizeRepresents: "area",
    series: [
      {
        name: "Students",
        xValues: "Scatter!$A$2:$A$8",
        values: "Scatter!$B$2:$B$8",
        bubbleSize: "Scatter!$C$2:$C$8",
        fill: "70AD47"
      }
    ]
  });

  // -- Radar family
  galleryChart("radar (standard)", {
    type: "radar",
    radarStyle: "standard",
    title: "Radar",
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
      { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }
    ]
  });

  galleryChart("radar (filled)", {
    type: "radar",
    radarStyle: "filled",
    title: "Filled radar",
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7",
        spPr: { fill: { solid: { srgb: "4472C4", alpha: 40000 } } }
      }
    ]
  });

  // -- Stock (OHLC)
  galleryChart("stock (OHLC)", {
    type: "stock",
    title: "Open-High-Low-Close",
    hiLowLines: true,
    upDownBars: {
      gapWidth: 150,
      upBars: { fill: "FFFFFF", border: "70AD47" },
      downBars: { fill: "C00000", border: "C00000" }
    },
    series: [
      {
        name: "Open",
        categories: "Stock!$A$2:$A$6",
        values: "Stock!$B$2:$B$6",
        line: "70AD47",
        marker: { symbol: "none" }
      },
      {
        name: "High",
        categories: "Stock!$A$2:$A$6",
        values: "Stock!$C$2:$C$6",
        line: "70AD47",
        marker: { symbol: "none" }
      },
      {
        name: "Low",
        categories: "Stock!$A$2:$A$6",
        values: "Stock!$D$2:$D$6",
        line: "C00000",
        marker: { symbol: "none" }
      },
      {
        name: "Close",
        categories: "Stock!$A$2:$A$6",
        values: "Stock!$E$2:$E$6",
        line: "000000",
        marker: { symbol: "none" }
      }
    ],
    categoryAxis: { numFmt: "yyyy-mm-dd" }
  });

  // -- Surface family
  galleryChart("surface (contour)", {
    type: "surface",
    title: "Contour (surface top-down)",
    series: [
      { name: "North", categories: "Surface!$B$1:$E$1", values: "Surface!$B$2:$E$2" },
      { name: "South", categories: "Surface!$B$1:$E$1", values: "Surface!$B$3:$E$3" },
      { name: "East", categories: "Surface!$B$1:$E$1", values: "Surface!$B$4:$E$4" },
      { name: "West", categories: "Surface!$B$1:$E$1", values: "Surface!$B$5:$E$5" }
    ]
  });

  galleryChart("surface3D (wireframe)", {
    type: "surface3D",
    wireframe: true,
    title: "Wireframe surface",
    view3D: { rotX: 20, rotY: 30, perspective: 30 },
    series: [
      { name: "North", categories: "Surface!$B$1:$E$1", values: "Surface!$B$2:$E$2" },
      { name: "South", categories: "Surface!$B$1:$E$1", values: "Surface!$B$3:$E$3" },
      { name: "East", categories: "Surface!$B$1:$E$1", values: "Surface!$B$4:$E$4" },
      { name: "West", categories: "Surface!$B$1:$E$1", values: "Surface!$B$5:$E$5" }
    ]
  });

  // ---------------------------------------------------------------------------
  // 2. Chart presets — 70+ Excel UI aliases exposed through `addPresetChart`.
  // ---------------------------------------------------------------------------

  const presets = wb.addWorksheet("2-Presets");
  presets.addRow(["Preset", "Notes"]);
  presets.getCell("A1").font = { bold: true };

  // Demonstrate a representative sample from every preset family.
  const presetCatalog: Array<{ preset: string; title: string }> = [
    { preset: "columnClustered", title: "Column, clustered" },
    { preset: "columnStacked", title: "Column, stacked" },
    { preset: "columnStacked100", title: "Column, 100% stacked" },
    { preset: "col3DConeStacked100", title: "3D cone, 100% stacked" },
    { preset: "col3DCylinderClustered", title: "3D cylinder, clustered" },
    { preset: "col3DPyramidStacked", title: "3D pyramid, stacked" },
    { preset: "barStacked", title: "Bar, stacked" },
    { preset: "pieOfPie", title: "Pie-of-pie" },
    { preset: "barOfPie", title: "Bar-of-pie" },
    { preset: "lineMarkers", title: "Line with markers" },
    { preset: "pieExploded", title: "Exploded pie" },
    { preset: "doughnutExploded", title: "Exploded doughnut" },
    { preset: "scatterSmoothMarker", title: "Scatter smooth + markers" },
    { preset: "bubble3D", title: "Bubble with 3D effect" },
    { preset: "radarFilled", title: "Radar filled" },
    { preset: "stockOHLC", title: "Stock OHLC" },
    { preset: "surface3D", title: "Surface 3D" },
    { preset: "surfaceTopView", title: "Surface top view" },
    { preset: "wireframeSurface", title: "Wireframe surface" }
  ];

  presetCatalog.forEach((entry, i) => {
    const row = 2 + i * 16;
    presets.getCell(`A${row}`).value = entry.preset;
    presets.getCell(`A${row}`).font = { bold: true };
    presets.getCell(`A${row + 1}`).value = entry.title;

    // `addPresetChart` auto-fills the `type` from the preset. The caller
    // only has to supply series + any chart-specific overrides.
    const series: AddChartSeriesOptions[] =
      entry.preset === "scatterSmoothMarker"
        ? [{ xValues: "Scatter!$A$2:$A$8", values: "Scatter!$B$2:$B$8", name: "Scatter" }]
        : entry.preset === "bubble3D"
          ? [
              {
                xValues: "Scatter!$A$2:$A$8",
                values: "Scatter!$B$2:$B$8",
                bubbleSize: "Scatter!$C$2:$C$8",
                name: "Bubble",
                bubble3D: true
              }
            ]
          : entry.preset === "stockOHLC"
            ? [
                { name: "Open", categories: "Stock!$A$2:$A$6", values: "Stock!$B$2:$B$6" },
                { name: "High", categories: "Stock!$A$2:$A$6", values: "Stock!$C$2:$C$6" },
                { name: "Low", categories: "Stock!$A$2:$A$6", values: "Stock!$D$2:$D$6" },
                { name: "Close", categories: "Stock!$A$2:$A$6", values: "Stock!$E$2:$E$6" }
              ]
            : entry.preset.startsWith("surface")
              ? [
                  { name: "N", categories: "Surface!$B$1:$E$1", values: "Surface!$B$2:$E$2" },
                  { name: "S", categories: "Surface!$B$1:$E$1", values: "Surface!$B$3:$E$3" },
                  { name: "E", categories: "Surface!$B$1:$E$1", values: "Surface!$B$4:$E$4" },
                  { name: "W", categories: "Surface!$B$1:$E$1", values: "Surface!$B$5:$E$5" }
                ]
              : [
                  {
                    name: "Revenue",
                    categories: "Sales!$A$2:$A$7",
                    values: "Sales!$B$2:$B$7"
                  },
                  {
                    name: "Cost",
                    categories: "Sales!$A$2:$A$7",
                    values: "Sales!$D$2:$D$7"
                  }
                ];

    presets.addPresetChart(
      entry.preset as (typeof EXCEL_CHART_PRESETS)[number],
      { title: entry.title, series },
      `B${row}:J${row + 14}`
    );
  });

  // ---------------------------------------------------------------------------
  // 3. ChartEx gallery — Office 2016+ modern chart types.
  // ---------------------------------------------------------------------------

  const ex = wb.addWorksheet("3-ChartEx Gallery");
  ex.addRow(["Type"]);
  ex.getCell("A1").font = { bold: true };
  let exRow = 2;
  const addExChart = (label: string, anchorHeight: number, cb: (range: string) => void): void => {
    ex.getCell(`A${exRow}`).value = label;
    ex.getCell(`A${exRow}`).font = { bold: true };
    const top = exRow;
    const bottom = exRow + anchorHeight;
    cb(`B${top}:J${bottom}`);
    exRow = bottom + 1;
  };

  addExChart("histogram", 15, range => {
    ex.addHistogramChart(
      {
        title: "Histogram",
        series: [{ name: "Samples", values: "Distribution!$A$2:$A$41" }],
        binning: { binType: "auto" }
      },
      range
    );
  });

  addExChart("pareto", 15, range => {
    ex.addParetoChart(
      {
        title: "Pareto (histogram + cumulative)",
        series: [{ name: "Samples", values: "Distribution!$A$2:$A$41" }],
        binning: { binType: "binCount", binCount: 8 }
      },
      range
    );
  });

  addExChart("waterfall", 15, range => {
    ex.addWaterfallChart(
      {
        title: "Revenue waterfall",
        categories: "Sales!$A$2:$A$7",
        series: [
          {
            name: "Delta",
            values: "Sales!$C$2:$C$7",
            subtotals: [0, 5] // Mark first & last data points as subtotals
          }
        ],
        layout: { connectorLines: true }
      },
      range
    );
  });

  addExChart("funnel", 15, range => {
    ex.addFunnelChart(
      {
        title: "Sales funnel",
        categories: "Sales!$A$2:$A$7",
        series: [{ name: "Stages", values: "Sales!$B$2:$B$7" }]
      },
      range
    );
  });

  addExChart("treemap", 15, range => {
    ex.addTreemapChart(
      {
        title: "Sales treemap",
        categories: "Hierarchy!$C$2:$C$10",
        series: [
          {
            name: "Sales",
            values: "Hierarchy!$D$2:$D$10",
            hierarchy: ["Hierarchy!$A$2:$A$10", "Hierarchy!$B$2:$B$10"]
          }
        ],
        layout: { parentLabelLayout: "banner" }
      },
      range
    );
  });

  addExChart("sunburst", 15, range => {
    ex.addSunburstChart(
      {
        title: "Sales sunburst",
        categories: "Hierarchy!$C$2:$C$10",
        series: [
          {
            name: "Sales",
            values: "Hierarchy!$D$2:$D$10",
            hierarchy: ["Hierarchy!$A$2:$A$10", "Hierarchy!$B$2:$B$10"]
          }
        ]
      },
      range
    );
  });

  addExChart("boxWhisker", 15, range => {
    ex.addBoxWhiskerChart(
      {
        title: "Distribution box-whisker",
        categories: "Distribution!$A$1:$A$1",
        series: [
          {
            name: "Samples",
            values: "Distribution!$A$2:$A$41"
          }
        ],
        layout: {
          quartileMethod: "inclusive",
          showMeanLine: true,
          showMeanMarker: true,
          showInnerPoints: true,
          showOutlierPoints: true
        }
      },
      range
    );
  });

  addExChart("regionMap", 15, range => {
    ex.addRegionMapChart(
      {
        title: "Revenue by country",
        categories: "Regions!$A$2:$A$7",
        series: [{ name: "Revenue", values: "Regions!$B$2:$B$7" }],
        layout: {
          projection: "mercator",
          regionLabels: "bestFit",
          geoMappingLevel: "country"
        }
      },
      range
    );
  });

  // ChartEx presets — same shape as classic presets but using
  // `addPresetChartEx` / `EXCEL_CHART_EX_PRESETS`.
  addExChart("preset: boxAndWhisker alias", 15, range => {
    ex.addPresetChartEx(
      "boxAndWhisker",
      {
        title: "Preset alias",
        categories: "Distribution!$A$1:$A$1",
        series: [{ name: "Samples", values: "Distribution!$A$2:$A$41" }]
      },
      range
    );
  });
  console.log(`ChartEx presets known: ${EXCEL_CHART_EX_PRESETS.join(", ")}`);

  // ---------------------------------------------------------------------------
  // 4. Combo charts — multiple chart-type groups with secondary axes,
  //    up-down bars, data tables, and 3D wall styling.
  // ---------------------------------------------------------------------------

  const combo = wb.addWorksheet("4-Combo");
  combo.addRow(["Combo charts"]);
  combo.getCell("A1").font = { bold: true };

  // Column + line on secondary axis — the canonical combo layout.
  combo.addComboChart(
    {
      title: "Revenue (bars) vs Growth (line, secondary axis)",
      groups: [
        {
          type: "bar",
          barDir: "col",
          grouping: "clustered",
          series: [
            {
              name: "Revenue",
              categories: "Sales!$A$2:$A$7",
              values: "Sales!$B$2:$B$7",
              fill: "4472C4"
            }
          ]
        },
        {
          type: "line",
          useSecondaryAxis: true,
          series: [
            {
              name: "Growth",
              categories: "Sales!$A$2:$A$7",
              values: "Sales!$F$2:$F$7",
              line: "ED7D31",
              lineWidth: 2,
              marker: { symbol: "circle", size: 7, fill: "ED7D31", border: "FFFFFF" }
            }
          ]
        }
      ],
      showLegend: true,
      legendPosition: "b",
      dataTable: { showHorzBorder: true, showVertBorder: true, showOutline: true, showKeys: true }
    },
    "A3:L22"
  );

  // Stacked bar + line trend combo
  combo.addComboChart(
    {
      title: "Stacked revenue + total trend",
      groups: [
        {
          type: "bar",
          barDir: "col",
          grouping: "stacked",
          series: [
            { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" },
            { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
          ]
        },
        {
          type: "line",
          series: [
            {
              name: "Revenue",
              categories: "Sales!$A$2:$A$7",
              values: "Sales!$B$2:$B$7",
              smooth: true,
              line: "C00000",
              lineWidth: 2.5,
              marker: { symbol: "square", size: 7 }
            }
          ]
        }
      ]
    },
    "A24:L43"
  );

  // Volume-HLC: bar for volume on primary axis + stock on secondary
  // axis — the exact combo that used to be named `stockVHLC` in the
  // preset registry; see `chart-presets.ts` for the removal note.
  combo.addComboChart(
    {
      title: "Volume + HLC (combo)",
      groups: [
        {
          type: "bar",
          barDir: "col",
          series: [
            {
              name: "Volume",
              categories: "Stock!$A$2:$A$6",
              values: "Stock!$E$2:$E$6",
              fill: "A6A6A6"
            }
          ]
        },
        {
          type: "stock",
          useSecondaryAxis: true,
          hiLowLines: true,
          series: [
            { name: "High", categories: "Stock!$A$2:$A$6", values: "Stock!$C$2:$C$6" },
            { name: "Low", categories: "Stock!$A$2:$A$6", values: "Stock!$D$2:$D$6" },
            { name: "Close", categories: "Stock!$A$2:$A$6", values: "Stock!$E$2:$E$6" }
          ]
        }
      ]
    },
    "A45:L64"
  );

  // ---------------------------------------------------------------------------
  // 5. Helper APIs — `addChartFromRows`, `addChartFromTable`,
  //    `addColumnChartFromRows`, `addChartExFromRows`, `seriesFromColumns`.
  // ---------------------------------------------------------------------------

  const helpers = wb.addWorksheet("5-Helper APIs");
  helpers.addRow(["Helper APIs"]);
  helpers.getCell("A1").font = { bold: true };

  // -- `addChartFromRows`: write rows + build chart from the JS array.
  // Returns a regular chart, so you can chain formatting later.
  helpers.getCell("A3").value = "addChartFromRows";
  helpers.getCell("A3").font = { bold: true };
  helpers.addChartFromRows(
    [
      { day: "Mon", visits: 312 },
      { day: "Tue", visits: 400 },
      { day: "Wed", visits: 280 },
      { day: "Thu", visits: 530 },
      { day: "Fri", visits: 470 }
    ],
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Visitors this week",
      startCell: "A4",
      x: "day",
      y: "visits"
    },
    "C3:K17"
  );

  // -- `addColumnChartFromRows`: shortcut — same as above but `type` /
  // `barDir` are implied.
  helpers.getCell("A20").value = "addColumnChartFromRows";
  helpers.getCell("A20").font = { bold: true };
  helpers.addColumnChartFromRows(
    [
      { quarter: "Q1", revenue: 1.2 },
      { quarter: "Q2", revenue: 1.8 },
      { quarter: "Q3", revenue: 1.5 },
      { quarter: "Q4", revenue: 2.1 }
    ],
    {
      title: "Quarterly revenue",
      startCell: "A21",
      x: "quarter",
      y: "revenue"
    },
    "C20:K34"
  );

  // -- `addChartFromTable`: pull series out of a structured Excel Table
  // so the chart expands when the table grows.
  const tblSheet = wb.addWorksheet("5b-Table-backed");
  const table = tblSheet.addTable({
    name: "MonthlyKpis",
    ref: "A1",
    headerRow: true,
    totalsRow: true,
    columns: [
      { name: "Month" },
      { name: "Revenue", totalsRowFunction: "sum" },
      { name: "Profit", totalsRowFunction: "sum" },
      { name: "Units", totalsRowFunction: "sum" }
    ],
    rows: [
      ["Jan", 120, 32, 540],
      ["Feb", 180, 49, 620],
      ["Mar", 160, 41, 610],
      ["Apr", 205, 64, 710],
      ["May", 232, 75, 805],
      ["Jun", 248, 81, 880]
    ]
  });

  tblSheet.addChartFromTable(
    table,
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Table-backed — expands with data",
      categoryColumn: "Month",
      valueColumns: ["Revenue", "Profit"]
    },
    "F1:N18"
  );

  // -- `seriesFromColumns`: compose series from raw column ranges.
  helpers.getCell("A38").value = "seriesFromColumns";
  helpers.getCell("A38").font = { bold: true };
  const seriesA = helpers.seriesFromColumns({
    categories: "Sales!$A$2:$A$7",
    values: "Sales!$B$2:$B$7",
    name: "Revenue"
  });
  const seriesB = helpers.seriesFromColumns({
    categories: "Sales!$A$2:$A$7",
    values: "Sales!$C$2:$C$7",
    name: "Profit"
  });
  helpers.addChart(
    {
      type: "line",
      title: "Built via seriesFromColumns",
      showMarker: true,
      series: [seriesA, seriesB]
    },
    "C38:K52"
  );

  // -- `addChartExFromTable`: same idea, for modern ChartEx types.
  tblSheet.addChartExFromTable(
    table,
    {
      type: "funnel",
      title: "Funnel from table",
      categoryColumn: "Month",
      valueColumns: ["Revenue"]
    },
    "F20:N37"
  );

  // ---------------------------------------------------------------------------
  // 6. Axis features — secondary axis, log scale, rotated labels, custom
  //    units, date axis, hidden axis.
  // ---------------------------------------------------------------------------

  const axes = wb.addWorksheet("6-Axes");
  axes.addRow(["Axis showcase"]);
  axes.getCell("A1").font = { bold: true };

  axes.addChart(
    {
      type: "line",
      title: "Log scale axis, rotated labels",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          smooth: true
        }
      ],
      categoryAxis: {
        title: "Month",
        textRotation: -45,
        majorTickMark: "outside",
        minorTickMark: "none"
      },
      valueAxis: {
        title: "Revenue (USD, log10)",
        logBase: 10,
        min: 10,
        max: 1000,
        majorGridlines: true,
        minorGridlines: true,
        numFmt: "$#,##0"
      }
    },
    "A3:I22"
  );

  axes.addChart(
    {
      type: "line",
      title: "Display units (thousands)",
      series: [{ name: "Units", categories: "Sales!$A$2:$A$7", values: "Sales!$E$2:$E$7" }],
      valueAxis: {
        title: "Units sold",
        displayUnits: "thousands",
        displayUnitsLabel: "× 1 000"
      }
    },
    "A24:I43"
  );

  axes.addChart(
    {
      type: "line",
      title: "Hidden value axis, minimalist axes",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          line: "4472C4",
          lineWidth: 3,
          marker: { symbol: "none" }
        }
      ],
      valueAxis: { hidden: true },
      categoryAxis: { majorTickMark: "none", minorTickMark: "none" },
      showLegend: false
    },
    "A45:I64"
  );

  // ---------------------------------------------------------------------------
  // 7. Rich titles, legend entries, plot-area layout, picture fills.
  // ---------------------------------------------------------------------------

  const rich = wb.addWorksheet("7-Rich Formatting");
  rich.addRow(["Rich formatting"]);
  rich.getCell("A1").font = { bold: true };

  // Structured rich-text title: two runs, different formatting.
  const richTitle: ChartRichText = {
    paragraphs: [
      {
        runs: [
          { text: "Q2 ", properties: { bold: true, size: 1600, color: { srgb: "4472C4" } } },
          { text: "Performance", properties: { italic: true, size: 1400 } }
        ]
      }
    ]
  };

  // Build a single programmatic PNG for the picture-fill demo. Every
  // byte is generated here — no external dependency.
  const demoPng = makeSolidColorPng(256, 160, 0x4472c4);

  rich.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: richTitle,
      titleOptions: { overlay: false, spPr: { border: "BFBFBF" } },
      legendOptions: {
        overlay: false,
        entries: [
          // Hide the second series' legend entry without hiding the series.
          { index: 1, hidden: true }
        ],
        txPr: { size: 900, color: { srgb: "595959" } }
      },
      plotAreaOptions: {
        spPr: { fill: "FAFAFA", border: "D9D9D9" }
      },
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          pictureFill: { image: demoPng, fillMode: "stretch" }
        },
        {
          name: "Cost",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$D$2:$D$7",
          fill: "ED7D31"
        }
      ]
    },
    "A3:I22"
  );

  // Title from a cell formula — stays live when the cell changes.
  rich.getCell("A25").value = "Live chart title →";
  rich.getCell("A25").font = { bold: true };
  rich.getCell("B25").value = "Automatic dashboard";
  rich.addChart(
    {
      type: "line",
      title: { formula: "'7-Rich Formatting'!$B$25" },
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A27:I46"
  );

  // ---------------------------------------------------------------------------
  // 8. Chart styling — built-in 1-48 catalogue + structured sidecars.
  // ---------------------------------------------------------------------------

  const styled = wb.addWorksheet("8-Styles");
  styled.addRow(["Chart styling"]);
  styled.getCell("A1").font = { bold: true };

  const styledChartNum = styled.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Built-in style",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A3:I22"
  );
  // `setStyle(N)` == `setBuiltInStyle(N)` — lightweight `<c:style/>` knob
  // mapped to the 2007/2010 catalogue (1-48).
  styled.getCharts()[0].setStyle(42);
  console.log(`Built-in-style chart is #${styledChartNum}`);

  // Modern styleN.xml / colorsN.xml sidecars — full Office 2013+ styling.
  const chartStyle: ChartStyleModel = {
    id: 201,
    elements: {
      chartArea: { fillRefIdx: 1, lnRefIdx: 1, effectRefIdx: 0, fontRefIdx: "minor" },
      title: { fillRefIdx: 0, lnRefIdx: 0, effectRefIdx: 0, fontRefIdx: "major" },
      dataPoint: { fillRefIdx: 2, lnRefIdx: 0, effectRefIdx: 0, fontRefIdx: "minor" }
    }
  };
  const chartColors: ChartColorsModel = {
    method: "cycle",
    id: 10,
    colors: [{ srgb: "4472C4" }, { srgb: "ED7D31" }, { srgb: "A5A5A5" }, { srgb: "FFC000" }]
  };
  styled.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Structured style sidecar",
      series: [
        { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
        { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }
      ],
      chartStyle,
      chartColors
    },
    "A25:I44"
  );

  // ---------------------------------------------------------------------------
  // 9. Chartsheet — full-page chart that takes its own tab.
  // ---------------------------------------------------------------------------

  wb.addChartsheet("Dashboard", {
    tabSelected: false,
    zoomToFit: true,
    chart: {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Full-page dashboard",
      series: [
        { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
        { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
      ],
      showLegend: true,
      legendPosition: "b",
      dataTable: true
    }
  });

  // ChartEx chartsheet — same interface, modern layout.
  wb.addChartsheet("Map", {
    chart: {
      type: "regionMap",
      title: "Revenue heat-map",
      categories: "Regions!$A$2:$A$7",
      series: [{ name: "Revenue", values: "Regions!$B$2:$B$7" }],
      layout: { projection: "robinson", regionLabels: "bestFit" }
    }
  });

  // ---------------------------------------------------------------------------
  // 10. Pivot chart — classic pivot-backed chart + pivot chartsheet.
  // ---------------------------------------------------------------------------

  const pivotSource = wb.addWorksheet("10-Pivot Source");
  pivotSource.addTable({
    name: "PivotSrc",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "Region" }, { name: "Category" }, { name: "Quarter" }, { name: "Revenue" }],
    rows: [
      ["East", "A", "Q1", 120],
      ["East", "B", "Q1", 95],
      ["East", "A", "Q2", 140],
      ["West", "A", "Q1", 180],
      ["West", "B", "Q1", 110],
      ["West", "A", "Q2", 200],
      ["South", "A", "Q1", 80],
      ["South", "B", "Q1", 70]
    ]
  });
  const pivotSheet = wb.addWorksheet("10-Pivot");
  const pivot = pivotSheet.addPivotTable({
    sourceTable: pivotSource.getTable("PivotSrc"),
    rows: ["Region"],
    columns: ["Category"],
    values: ["Revenue"],
    metric: "sum"
  });

  pivotSheet.addPivotChart(
    pivot,
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Revenue by region / category (pivot)",
      series: [
        {
          name: "Revenue",
          categories: "'10-Pivot Source'!$A$2:$A$9",
          values: "'10-Pivot Source'!$D$2:$D$9"
        }
      ],
      pivotChartOptions: {
        dropZonesVisible: true,
        dropZoneFilter: true,
        dropZoneCategories: true,
        dropZoneData: true,
        dropZoneSeries: true,
        refreshOnOpen: true,
        showExpandCollapseFieldButtons: true
      }
    },
    "F1:N20"
  );

  wb.addPivotChartsheet("Pivot Dashboard", pivot, {
    chart: {
      type: "line",
      title: "Pivot (chartsheet)",
      showMarker: true,
      series: [
        {
          name: "Revenue",
          categories: "'10-Pivot Source'!$A$2:$A$9",
          values: "'10-Pivot Source'!$D$2:$D$9"
        }
      ]
    }
  });

  // ---------------------------------------------------------------------------
  // 11. Anchor variations — one-cell, two-cell and absolute EMU anchors.
  // ---------------------------------------------------------------------------

  const anchors = wb.addWorksheet("11-Anchors");
  anchors.addRow(["Anchor forms"]);
  anchors.getCell("A1").font = { bold: true };

  // Two-cell anchor (default — tl to br).
  anchors.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Two-cell anchor",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    { tl: { col: 1, row: 2 }, br: { col: 8, row: 17 } }
  );

  // One-cell anchor — pinned to a cell, fixed EMU extent (5×3 inches).
  anchors.addChart(
    {
      type: "line",
      title: "One-cell anchor (5×3 in)",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    {
      tl: { col: 1, row: 19 },
      ext: { cx: 5 * 914400, cy: 3 * 914400 },
      editAs: "oneCell"
    }
  );

  // Absolute anchor — fixed EMU position + size; ignores rows/columns.
  // 0.5 inch top margin, 1 inch left margin, 5×3 in size.
  anchors.addChart(
    {
      type: "pie",
      title: "Absolute anchor",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    {
      pos: { x: 914400, y: 36 * 914400 },
      ext: { cx: 5 * 914400, cy: 3 * 914400 },
      editAs: "absolute"
    }
  );

  // ---------------------------------------------------------------------------
  // 12. Feature deep-dive — every trendline / error-bar / marker variant,
  //     dropLines / serLines / hiLowLines, gapWidth / overlap / gapDepth,
  //     firstSliceAng / holeSize, valueFromCells, per-entry data labels,
  //     gradient / pattern fills, xValueType: "text", displayBlanksAs /
  //     plotVisOnly / showDLblsOverMax, date axis, null title.
  // ---------------------------------------------------------------------------

  const features = wb.addWorksheet("12-Feature Deep-Dive");
  features.addRow(["Feature showcase"]);
  features.getCell("A1").font = { bold: true };

  // --- 12.1 every trendline type on one chart
  const trendlines: AddTrendlineOptions[] = [
    { type: "linear", name: "Linear", line: "4472C4", lineDash: "solid" },
    { type: "exp", name: "Exponential", line: "ED7D31", lineDash: "dash" },
    { type: "log", name: "Logarithmic", line: "70AD47", lineDash: "sysDash" },
    { type: "poly", name: "Polynomial (3)", order: 3, line: "FFC000", lineDash: "dot" },
    { type: "movingAvg", name: "Moving avg (2)", period: 2, line: "5B9BD5", lineDash: "lgDash" },
    { type: "power", name: "Power", line: "A5A5A5", lineDash: "lgDashDot" }
  ];
  features.getCell("A3").value =
    "12.1 — every trendline type (linear/exp/log/poly/movingAvg/power)";
  features.getCell("A3").font = { bold: true };
  features.addChart(
    {
      type: "line",
      title: "All six trendline types",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          line: "000000",
          lineWidth: 1,
          marker: { symbol: "circle", size: 6 },
          trendline: trendlines
        }
      ]
    },
    "A4:L22"
  );

  // --- 12.2 every error-bar value type
  const errorBars: AddErrorBarsOptions[] = [
    { type: "fixedVal", value: 10, direction: "y", barDir: "both", line: "4472C4" },
    { type: "percentage", value: 15, direction: "y", barDir: "plus", line: "ED7D31" },
    { type: "stdDev", value: 1, direction: "y", barDir: "minus", line: "70AD47" },
    { type: "stdErr", direction: "y", barDir: "both", line: "FFC000" },
    {
      // Custom error bars — explicit per-point plus/minus ranges.
      type: "cust",
      direction: "y",
      barDir: "both",
      plus: "Sales!$C$2:$C$7",
      minus: "Sales!$C$2:$C$7",
      line: "5B9BD5",
      noEndCap: false
    }
  ];
  features.getCell("A25").value = "12.2 — every error-bar type";
  features.getCell("A25").font = { bold: true };
  features.addChart(
    {
      type: "scatter",
      scatterStyle: "marker",
      title: "All five error-bar types",
      series: [
        {
          name: "Points",
          xValues: "Scatter!$A$2:$A$8",
          values: "Scatter!$B$2:$B$8",
          marker: { symbol: "circle", size: 8, fill: "4472C4" },
          errorBars
        }
      ]
    },
    "A26:L44"
  );

  // --- 12.3 every marker symbol
  const markerSymbols: NonNullable<ChartMarker["symbol"]>[] = [
    "circle",
    "square",
    "diamond",
    "triangle",
    "x",
    "star",
    "plus",
    "dash",
    "dot"
  ];
  features.getCell("A47").value = "12.3 — every marker symbol";
  features.getCell("A47").font = { bold: true };
  features.addChart(
    {
      type: "line",
      title: "Marker catalogue",
      showMarker: true,
      series: markerSymbols.map((sym, i) => ({
        name: sym,
        categories: "Sales!$A$2:$A$7",
        // Fan each series out by an offset so they don't all overlap
        values: "Sales!$B$2:$B$7",
        line: "BFBFBF",
        lineWidth: 1,
        marker: { symbol: sym, size: 10, fill: HEX_PALETTE[i % HEX_PALETTE.length] }
      }))
    },
    "A48:L67"
  );

  // --- 12.4 dropLines + serLines + hiLowLines + upDownBars on one chart
  features.getCell("A70").value = "12.4 — dropLines / serLines / hiLowLines / upDownBars";
  features.getCell("A70").font = { bold: true };
  features.addChart(
    {
      type: "line",
      title: "Drop lines + hi-low lines + up-down bars",
      series: [
        { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
        { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }
      ],
      dropLines: true,
      hiLowLines: true,
      upDownBars: {
        gapWidth: 150,
        upBars: { fill: "70AD47" },
        downBars: { fill: "C00000" }
      }
    },
    "A71:L90"
  );

  features.getCell("A93").value =
    "12.5 — pie with serLines (leader) + firstSliceAng + exploded point";
  features.getCell("A93").font = { bold: true };
  features.addChart(
    {
      type: "ofPie",
      ofPieType: "bar",
      title: "Bar-of-pie — serLines + 45° start",
      firstSliceAng: 45,
      splitType: "percent",
      splitPos: 10,
      serLines: true,
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          dataLabels: { showPercent: true, showCatName: true, position: "bestFit" }
        }
      ]
    },
    "A94:L113"
  );

  // --- 12.6 gapWidth / overlap / gapDepth demo — `overlap` is 2-D only
  // (CT_Bar3DChart has no overlap), so we show gapWidth + gapDepth on a
  // bar3D, plus overlap on a 2-D bar alongside.
  features.getCell("A116").value = "12.6 — gapWidth + gapDepth (bar3D)";
  features.getCell("A116").font = { bold: true };
  features.addChart(
    {
      type: "bar3D",
      barDir: "col",
      grouping: "clustered",
      title: "3D bar — gapWidth=50 gapDepth=300",
      gapWidth: 50,
      gapDepth: 300,
      view3D: { rotX: 20, rotY: 25, rAngAx: true, perspective: 30 },
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "4472C4"
        },
        { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7", fill: "ED7D31" }
      ]
    },
    "A117:L136"
  );

  // --- 12.7 doughnut with several holeSize values (via dataPoints)
  features.getCell("A139").value =
    "12.7 — overlap (2-D bar, bar3D has no overlap in CT_Bar3DChart)";
  features.getCell("A139").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Overlap=-25 (bars visually separated)",
      overlap: -25,
      gapWidth: 200,
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "4472C4"
        },
        { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7", fill: "ED7D31" }
      ]
    },
    "A140:L159"
  );

  features.getCell("A162").value = "12.8 — doughnut with holeSize=75";
  features.getCell("A162").font = { bold: true };
  features.addChart(
    {
      type: "doughnut",
      holeSize: 75,
      firstSliceAng: 0,
      varyColors: true,
      title: "Thin doughnut",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          dataLabels: { showVal: true, position: "ctr" }
        }
      ]
    },
    "A163:L182"
  );

  // --- 12.9 data labels with valueFromCells (Excel 2013+)
  const labelsSheet = wb.addWorksheet("12b-Value-From-Cells");
  labelsSheet.addRows([
    ["Label", "Units"],
    ["🔥 Hot", 540],
    ["🌤 Warm", 620],
    ["☁ Cool", 610],
    ["❄ Cold", 710],
    ["💧 Rainy", 805],
    ["⚡ Storm", 880]
  ]);
  features.getCell("A185").value =
    "12.9 — data labels valueFromCells (Excel 2013+ 'Value From Cells')";
  features.getCell("A185").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Labels pulled from a separate range",
      series: [
        {
          name: "Units",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$E$2:$E$7",
          dataLabels: {
            position: "outEnd",
            showVal: false,
            valueFromCells: "'12b-Value-From-Cells'!$A$2:$A$7"
          }
        }
      ],
      showLegend: false
    },
    "A186:L205"
  );

  // --- 12.10 per-entry data label overrides
  features.getCell("A208").value =
    "12.10 — per-entry data label overrides (hide / custom text / recolour)";
  features.getCell("A208").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Per-entry overrides",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          dataLabels: {
            showVal: true,
            position: "outEnd",
            numFmt: "$#,##0",
            entries: [
              // Hide the first label.
              { index: 0, delete: true },
              // Replace the second label with a constant string.
              {
                index: 1,
                text: {
                  paragraphs: [
                    {
                      runs: [
                        {
                          text: "★ PEAK",
                          properties: { bold: true, color: { srgb: "C00000" }, size: 1200 }
                        }
                      ]
                    }
                  ]
                },
                position: "t"
              },
              // Move the last label into the bar, large.
              {
                index: 5,
                position: "inEnd",
                txPr: { size: 1400, bold: true, color: { srgb: "FFFFFF" } }
              }
            ]
          }
        }
      ]
    },
    "A209:L228"
  );

  // --- 12.11 gradient fill
  features.getCell("A231").value = "12.11 — series filled with a linear gradient";
  features.getCell("A231").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Linear gradient fill",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          spPr: {
            fill: {
              gradient: {
                type: "linear",
                angle: 90,
                stops: [
                  { position: 0, color: { srgb: "4472C4" } },
                  { position: 1, color: { srgb: "ED7D31" } }
                ]
              }
            }
          }
        }
      ]
    },
    "A232:L251"
  );

  // --- 12.12 pattern fill
  features.getCell("A254").value = "12.12 — pattern fill (preset 'dashDnDiag')";
  features.getCell("A254").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "DrawingML pattern fill",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          spPr: {
            fill: {
              pattern: {
                preset: "dashDnDiag",
                foreground: { srgb: "4472C4" },
                background: { srgb: "FFFFFF" }
              }
            }
          }
        }
      ]
    },
    "A255:L274"
  );

  // --- 12.13 scatter with xValueType: "text"
  const xTextSheet = wb.addWorksheet("12c-Text-X-Axis");
  xTextSheet.addRows([
    ["Phase", "Value"],
    ["Alpha", 22],
    ["Beta", 41],
    ["Gamma", 38],
    ["Delta", 52],
    ["Epsilon", 60]
  ]);
  features.getCell("A277").value = "12.13 — scatter with text x-axis (xValueType='text')";
  features.getCell("A277").font = { bold: true };
  features.addChart(
    {
      type: "scatter",
      scatterStyle: "lineMarker",
      title: "Scatter over named phases",
      series: [
        {
          name: "Value",
          xValues: "'12c-Text-X-Axis'!$A$2:$A$6",
          values: "'12c-Text-X-Axis'!$B$2:$B$6",
          xValueType: "text",
          marker: { symbol: "diamond", size: 9, fill: "4472C4" }
        }
      ]
    },
    "A278:L297"
  );

  // --- 12.14 date axis with custom time units
  const timeSheet = wb.addWorksheet("12d-Date-Axis");
  timeSheet.addRow(["Date", "Incidents"]);
  for (let i = 0; i < 12; i++) {
    const date = new Date(Date.UTC(2024, i, 15));
    timeSheet.addRow([date, 20 + Math.round(Math.sin(i) * 10) + i]);
  }
  timeSheet.getColumn(1).numFmt = "yyyy-mm-dd";
  features.getCell("A300").value = "12.14 — date axis (baseTimeUnit=months)";
  features.getCell("A300").font = { bold: true };
  features.addChart(
    {
      type: "line",
      title: "Incidents per month",
      series: [
        {
          name: "Incidents",
          categories: "'12d-Date-Axis'!$A$2:$A$13",
          values: "'12d-Date-Axis'!$B$2:$B$13",
          smooth: true
        }
      ],
      categoryAxis: {
        title: "Date",
        baseTimeUnit: "months",
        majorTimeUnit: "months",
        minorTimeUnit: "days",
        numFmt: "mmm-yy"
      }
    },
    "A301:L320"
  );

  // --- 12.15 null title (autoTitleDeleted)
  features.getCell("A323").value = "12.15 — title: null — Excel will NOT auto-generate a title";
  features.getCell("A323").font = { bold: true };
  features.addChart(
    {
      type: "bar",
      barDir: "col",
      title: null, // <-- explicitly suppressed
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }],
      showLegend: false
    },
    "A324:L343"
  );

  // --- 12.16 plotVisOnly / displayBlanksAs / showDLblsOverMax
  features.getCell("A346").value = "12.16 — plotVisOnly + displayBlanksAs + showDLblsOverMax";
  features.getCell("A346").font = { bold: true };
  features.addChart(
    {
      type: "line",
      title: "Interpolated across blanks",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }],
      plotVisOnly: true,
      displayBlanksAs: "span",
      showDLblsOverMax: true
    },
    "A347:L366"
  );

  // ---------------------------------------------------------------------------
  // 13. Headless charts — series whose cached values do NOT come from any
  //     worksheet range. Useful when generating a PDF preview from
  //     pre-aggregated numbers without bothering to stage them into a
  //     sheet first.
  // ---------------------------------------------------------------------------

  const headless = wb.addWorksheet("13-Headless Literal Data");
  headless.addRow(["Literal (no worksheet ranges)"]);
  headless.getCell("A1").font = { bold: true };

  // Literal classic chart — `literalValues` on a regular chart. The
  // builder emits a `c:numLit` for values + `c:strLit` for categories
  // so Excel can still render the chart on load without pointing at
  // any cell. (Writers that also want live data should additionally
  // set `categories` / `values` to a real range; the literal caches
  // act as a seed.)
  headless.getCell("A3").value =
    "13.1 — classic chart with literalCategories + literalValues (pie)";
  headless.getCell("A3").font = { bold: true };
  // The classic chart literal path writes the cache via a real range
  // pointing at cells we pre-populate here. Excel renders from the
  // cache so the referenced cells never have to be present, but an
  // addressable range keeps the XML valid when the file is re-opened.
  headless.addRows([
    ["Alpha", 30],
    ["Beta", 45],
    ["Gamma", 15],
    ["Delta", 10]
  ]);
  headless.addChart(
    {
      type: "pie",
      title: "Literal pie",
      varyColors: true,
      series: [
        {
          name: "Share",
          categories: "'13-Headless Literal Data'!$A$2:$A$5",
          values: "'13-Headless Literal Data'!$B$2:$B$5",
          dataLabels: { showPercent: true, showCatName: true, position: "bestFit" }
        }
      ]
    },
    "D1:L20"
  );

  // ChartEx treemap with `literalHierarchy` — the modern path that
  // genuinely takes literal arrays and produces a self-contained part.
  headless.getCell("A23").value = "13.2 — ChartEx treemap with literalCategories + literalValues";
  headless.getCell("A23").font = { bold: true };
  headless.addChartEx(
    {
      type: "treemap",
      title: "Literal treemap",
      // `values` is a required field. When literal arrays drive the
      // cache we still need a syntactically-valid reference (Excel
      // rebuilds the numRef shell from it on open).
      categories: "A1",
      series: [
        {
          name: "Sales",
          values: "A1",
          literalValues: [120, 80, 60, 40, 30, 25, 15, 10],
          literalCategories: ["NA", "EU", "APAC", "LATAM", "ME", "AF", "IN", "CN"],
          literalHierarchy: [
            ["Americas", "Americas", "APAC", "Americas", "EMEA", "EMEA", "APAC", "APAC"],
            ["USA", "Canada", "Japan", "Mexico", "Germany", "France", "India", "China"]
          ]
        }
      ],
      layout: { parentLabelLayout: "overlapping" }
    },
    "D23:L42"
  );

  // ---------------------------------------------------------------------------
  // 14. Chart editing — getter/setter + mutation APIs on the Chart object.
  // ---------------------------------------------------------------------------

  const edit = wb.addWorksheet("14-Chart Editing");
  edit.addRow(["Chart editing"]);
  edit.getCell("A1").font = { bold: true };

  edit.addChart(
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Start with one series",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "4472C4"
        }
      ]
    },
    "A3:L22"
  );

  const editing = edit.getCharts()[0];

  // `chartTypes` — inspect the plot-area groups. Classic charts have one
  // entry per chart-type group; combo charts return one per group.
  console.log("Editing chart groups:", editing.chartTypes.length, editing.chartTypes[0]?.type);
  console.log("Initial series count:", editing.getSeriesCount());

  // Mutate the title via the setter — triggers a raw-patch.
  editing.title = "Retitled via setter";

  // Pull an existing series, modify it with `updateSeries`.
  editing.updateSeries(0, { line: "C00000", lineWidth: 2, fill: "C00000" });

  // `addSeriesFromOptions` — append a second series after construction.
  editing.addSeriesFromOptions({
    name: "Profit",
    categories: "Sales!$A$2:$A$7",
    values: "Sales!$C$2:$C$7",
    fill: "70AD47",
    dataLabels: { showVal: true }
  });

  // `removeSeries(index)` — drop a series by index (returns the removed
  // object). Here we append a throwaway series then remove it.
  editing.addSeriesFromOptions({
    name: "Throwaway",
    categories: "Sales!$A$2:$A$7",
    values: "Sales!$E$2:$E$7"
  });
  editing.removeSeries(editing.getSeriesCount() - 1);

  // Low-level `setStyle(N)` toggle — applies one of the 48 built-in styles.
  editing.setStyle(26);

  // `setBuiltInStyle(N)` — xlsxwriter-compatible alias for `setStyle`.
  edit.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "setBuiltInStyle(42)",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A25:L44"
  );
  edit.getCharts()[1].setBuiltInStyle(42);

  // Narrow mutation on an already-built chart via `mutate(fn)`. The
  // callback receives the structured `ChartModel`; any fields touched
  // are re-emitted and the rest of the chart XML is preserved. Pass
  // `preferRawPatch: true` to keep the bytes for untouched blocks,
  // and `requireRawPatch: true` to FAIL if the edit would force a
  // structural rebuild.
  edit.addChart(
    {
      type: "line",
      title: "Mutated via chart.mutate()",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A47:L66"
  );
  const toMutate = edit.getCharts()[2];
  toMutate.mutate(
    model => {
      // Flip the first series into a heavy red dashed line.
      const firstGroup = model.chart.plotArea.chartTypes[0];
      if (firstGroup && firstGroup.series[0]) {
        const s = firstGroup.series[0];
        s.spPr = {
          line: { width: 25400, color: { srgb: "C00000" }, dash: "dash" }
        };
      }
    },
    { preferRawPatch: true }
  );

  // `mutateChartEx(fn)` — same idea for a ChartEx chart.
  edit.getCell("A69").value = "chart.mutateChartEx() — edit a treemap parentLabelLayout";
  edit.getCell("A69").font = { bold: true };
  edit.addChartEx(
    {
      type: "treemap",
      title: "Treemap — parentLabelLayout='banner'",
      categories: "Hierarchy!$C$2:$C$10",
      series: [
        {
          name: "Sales",
          values: "Hierarchy!$D$2:$D$10",
          hierarchy: ["Hierarchy!$A$2:$A$10", "Hierarchy!$B$2:$B$10"]
        }
      ]
    },
    "A70:L89"
  );
  const exChart = edit.getCharts()[3];
  exChart.mutateChartEx(
    model => {
      const series = model.chartSpace.chart.plotArea.plotAreaRegion?.series?.[0];
      if (series) {
        series.layoutPr = { ...series.layoutPr, parentLabelLayout: "banner" };
      }
    },
    { preferRawPatch: true }
  );

  // `removeUserShapes()` — no-op on freshly-built charts (they carry no
  // user shapes); demonstrated here so the code path is exercised.
  editing.removeUserShapes();

  // ---------------------------------------------------------------------------
  // 15. Low-level builders — call `buildChartModel` /
  //     `buildComboChartModel` / `buildChartExModel` directly. Useful
  //     when a caller wants a `ChartModel` in hand for headless rendering
  //     or for transforming programmatically before writing.
  // ---------------------------------------------------------------------------

  const standaloneClassic = buildChartModel({
    type: "bar",
    barDir: "col",
    title: "Built via buildChartModel()",
    series: [
      {
        name: "Literal",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$B$2:$B$7"
      }
    ]
  });
  console.log(
    "Standalone classic model — chartTypes:",
    standaloneClassic.chart.plotArea.chartTypes[0].type
  );

  const standaloneCombo = buildComboChartModel({
    title: "Built via buildComboChartModel()",
    groups: [
      {
        type: "bar",
        barDir: "col",
        series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
      },
      {
        type: "line",
        useSecondaryAxis: true,
        series: [{ name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }]
      }
    ]
  });
  console.log("Standalone combo model — groups:", standaloneCombo.chart.plotArea.chartTypes.length);

  const standaloneEx = buildChartExModel({
    type: "waterfall",
    title: "Built via buildChartExModel()",
    categories: "Sales!$A$2:$A$7",
    series: [{ name: "Delta", values: "Sales!$C$2:$C$7", subtotals: [0, 5] }]
  });
  console.log("Standalone ChartEx model — vectorable:", canRenderChartExAsVectorPdf(standaloneEx));

  // ---------------------------------------------------------------------------
  // 16. Picture fill via `workbook.addImage` — referring to a shared
  //     workbook media entry by id so multiple charts can reuse the same
  //     image without re-embedding it.
  // ---------------------------------------------------------------------------

  const logo = makeSolidColorPng(128, 64, 0x70ad47);
  const logoId = wb.addImage({ buffer: logo, extension: "png" });

  const pic = wb.addWorksheet("16-Picture Fill");
  pic.addRow(["Picture-fill via workbookImageId"]);
  pic.getCell("A1").font = { bold: true };
  pic.addChart(
    {
      type: "bar",
      barDir: "col",
      title: `Shared image #${logoId}`,
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          pictureFill: { image: { workbookImageId: logoId }, fillMode: "stack" }
        }
      ]
    },
    "A3:L22"
  );

  // ---------------------------------------------------------------------------
  // 17. Full preset catalogue — one chart per classic preset name
  //     (99 total), laid out in a grid so the caller can compare them
  //     side by side. Each chart uses the same series data so the
  //     differences are purely visual.
  // ---------------------------------------------------------------------------

  const catalogue = wb.addWorksheet("17-Full Preset Catalogue");
  catalogue.addRow(["All 99 classic presets"]);
  catalogue.getCell("A1").font = { bold: true };

  const presetNames = Object.keys(CHART_PRESETS);
  console.log(`Writing ${presetNames.length} presets on one sheet …`);
  const chartsPerRow = 3;
  const chartW = 9; // columns
  const chartH = 15; // rows
  const gap = 1;
  presetNames.forEach((name, idx) => {
    const gridX = idx % chartsPerRow;
    const gridY = Math.floor(idx / chartsPerRow);
    const startCol = 1 + gridX * (chartW + gap); // B, L, V, …
    const startRow = 3 + gridY * (chartH + gap);
    const endCol = startCol + chartW - 1;
    const endRow = startRow + chartH - 1;

    // Write the preset name above the chart.
    catalogue.getCell(startRow - 1, startCol + 1).value = name;
    catalogue.getCell(startRow - 1, startCol + 1).font = { bold: true, size: 10 };

    // Pick series inputs compatible with every preset family — surface
    // needs a 2-D matrix, stock needs OHLC, scatter/bubble need xValues.
    let series: AddChartSeriesOptions[];
    if (name.toLowerCase().includes("stock")) {
      series = [
        { name: "High", categories: "Stock!$A$2:$A$6", values: "Stock!$C$2:$C$6" },
        { name: "Low", categories: "Stock!$A$2:$A$6", values: "Stock!$D$2:$D$6" },
        { name: "Close", categories: "Stock!$A$2:$A$6", values: "Stock!$E$2:$E$6" }
      ];
    } else if (name.toLowerCase().includes("scatter")) {
      series = [{ xValues: "Scatter!$A$2:$A$8", values: "Scatter!$B$2:$B$8", name: "Series" }];
    } else if (name.toLowerCase().includes("bubble")) {
      series = [
        {
          xValues: "Scatter!$A$2:$A$8",
          values: "Scatter!$B$2:$B$8",
          bubbleSize: "Scatter!$C$2:$C$8",
          name: "Series"
        }
      ];
    } else if (name.toLowerCase().includes("surface") || name.toLowerCase().includes("contour")) {
      series = [
        { name: "N", categories: "Surface!$B$1:$E$1", values: "Surface!$B$2:$E$2" },
        { name: "S", categories: "Surface!$B$1:$E$1", values: "Surface!$B$3:$E$3" },
        { name: "E", categories: "Surface!$B$1:$E$1", values: "Surface!$B$4:$E$4" },
        { name: "W", categories: "Surface!$B$1:$E$1", values: "Surface!$B$5:$E$5" }
      ];
    } else {
      series = [
        { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
        { name: "Cost", categories: "Sales!$A$2:$A$7", values: "Sales!$D$2:$D$7" }
      ];
    }

    catalogue.addPresetChart(
      name as (typeof EXCEL_CHART_PRESETS)[number],
      { title: name, series, showLegend: false },
      {
        tl: { col: startCol, row: startRow },
        br: { col: endCol, row: endRow }
      }
    );
  });

  // ---------------------------------------------------------------------------
  // 18. Chartsheet with custom page setup + replaceChart demo
  // ---------------------------------------------------------------------------

  const cs = wb.addChartsheet("Custom Chartsheet", {
    tabSelected: false,
    zoomScale: 90,
    zoomToFit: false,
    pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    pageSetup: { orientation: "landscape", paperSize: 9, horizontalDpi: 300, verticalDpi: 300 },
    chart: {
      type: "bar",
      barDir: "col",
      title: "Original (will be replaced)",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    }
  });

  // `replaceChart` — swap the chart out in-place, preserving the tab
  // settings + page setup we just configured.
  cs.replaceChart({
    type: "line",
    title: "Replaced after creation",
    showMarker: true,
    series: [
      {
        name: "Profit",
        categories: "Sales!$A$2:$A$7",
        values: "Sales!$C$2:$C$7",
        line: "70AD47",
        lineWidth: 2.5
      }
    ]
  });

  // ---------------------------------------------------------------------------
  // 19. Advanced shape properties & effects — things that sit on the
  //     chart shape tree rather than on individual series.
  // ---------------------------------------------------------------------------
  //
  //   19.1  chart-level `spPr` (chart frame background + border + radius)
  //   19.2  `EffectList` — outer shadow + glow + softEdge + reflection + blur
  //   19.3  `Scene3D` + `ShapeProperties3D` + `Bevel` — 3D metadata
  //   19.4  Manual plot-area layout (edge mode — pins the plot area to
  //         exact fractional coordinates)
  //   19.5  Chart-level `PictureOptions` for bar3D walls
  //   19.6  `bandFormats` on surface chart (per-level colour bands)
  //
  // All six modify the chart model directly after `addChart` returns
  // (the high-level options don't always expose every leaf of the
  // OOXML schema — this shows how to reach in when you need to).

  const advanced = wb.addWorksheet("19-Advanced Shape & Effects");
  advanced.addRow(["Advanced shape / effects"]);
  advanced.getCell("A1").font = { bold: true };

  // --- 19.1 chart-level frame — rounded corners + gradient background
  advanced.getCell("A3").value = "19.1 — chart-level spPr (background gradient + border)";
  advanced.getCell("A3").font = { bold: true };
  advanced.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Chart frame with gradient",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A4:L23"
  );
  const chartFrame = advanced.getCharts()[0];
  chartFrame.mutate(model => {
    model.spPr = {
      fill: {
        gradient: {
          type: "linear",
          angle: 45,
          stops: [
            { position: 0, color: { srgb: "F2F9FF" } },
            { position: 1, color: { srgb: "D9E7F5" } }
          ]
        }
      },
      line: { color: { srgb: "5B9BD5" }, width: 19050 }
    };
    model.roundedCorners = true;
  });

  // --- 19.2 every effect kind on one chart
  advanced.getCell("A26").value =
    "19.2 — EffectList: outerShadow + glow + softEdge + reflection + blur";
  advanced.getCell("A26").font = { bold: true };
  advanced.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Effects showcase",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "4472C4"
        }
      ]
    },
    "A27:L46"
  );
  const withEffects = advanced.getCharts()[1];
  withEffects.mutate(model => {
    const firstGroup = model.chart.plotArea.chartTypes[0];
    const firstSeries = firstGroup?.series[0];
    if (firstSeries) {
      firstSeries.spPr = {
        ...firstSeries.spPr,
        fill: { solid: { srgb: "4472C4" } },
        effectList: {
          // Outer shadow — drop shadow below each bar
          outerShadow: {
            blurRadius: 50800, // 4 pt
            distance: 38100, // 3 pt
            direction: 5400000, // 90° × 60000
            color: { srgb: "000000", alpha: 40000 }
          },
          // Glow — soft coloured halo around the shape
          glow: {
            radius: 63500, // 5 pt
            color: { srgb: "4472C4", alpha: 60000 }
          },
          // Soft edges — feathers the bar edges
          softEdge: { radius: 12700 },
          // Reflection — mirrors the shape below itself with decreasing opacity
          reflection: {
            blurRadius: 6350,
            startOpacity: 50000,
            endOpacity: 0,
            distance: 0,
            direction: 5400000,
            fadeDirection: 5400000,
            scaleHorizontal: 100000,
            scaleVertical: -100000
          },
          // Blur — softens the whole shape
          blur: { radius: 6350, grow: false }
        }
      };
    }
  });

  // --- 19.3 Scene3D + ShapeProperties3D + Bevel on a bar3D chart
  advanced.getCell("A49").value = "19.3 — Scene3D + sp3d bevel/extrusion/contour/material";
  advanced.getCell("A49").font = { bold: true };
  advanced.addChart(
    {
      type: "bar3D",
      barDir: "col",
      title: "3D bar with bevel + extrusion",
      view3D: { rotX: 20, rotY: 30, rAngAx: true, perspective: 30 },
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "ED7D31"
        }
      ]
    },
    "A50:L69"
  );
  const with3D = advanced.getCharts()[2];
  with3D.mutate(model => {
    const firstGroup = model.chart.plotArea.chartTypes[0];
    const firstSeries = firstGroup?.series[0];
    if (firstSeries) {
      firstSeries.spPr = {
        ...firstSeries.spPr,
        fill: { solid: { srgb: "ED7D31" } },
        // Scene3D — the camera + light rig Excel renders bevels under
        scene3d: {
          camera: {
            preset: "orthographicFront",
            fov: 30,
            zoom: 100000
          },
          lightRig: {
            rig: "threePt",
            direction: "t",
            rotation: { lat: 0, lon: 0, rev: 1200000 }
          }
        },
        // Extruded 3D shape with bevelled top and bottom
        sp3d: {
          z: 0,
          extrusionHeight: 50800, // 4 pt extrusion
          contourWidth: 6350, // 0.5 pt contour
          material: "metal",
          bevelTop: { width: 76200, height: 76200, preset: "circle" },
          bevelBottom: { width: 76200, height: 38100, preset: "angle" },
          extrusionColor: { srgb: "9E480E" },
          contourColor: { srgb: "4A1F00" }
        }
      };
    }
  });

  // --- 19.4 manual plot-area layout in edge mode
  advanced.getCell("A72").value = "19.4 — manual plot-area layout (edge mode pins coordinates)";
  advanced.getCell("A72").font = { bold: true };
  advanced.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Plot area pinned to the right half",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }],
      plotAreaOptions: {
        layout: {
          manualLayout: {
            layoutTarget: "inner",
            xMode: "edge",
            yMode: "edge",
            wMode: "edge",
            hMode: "edge",
            x: 0.5, // 50% from the left of the chart area
            y: 0.15,
            w: 0.45,
            h: 0.7
          }
        },
        spPr: { fill: "F2F9FF", border: "5B9BD5" }
      }
    },
    "A73:L92"
  );

  // --- 19.5 chart-level PictureOptions (bar3D wall picture fills)
  advanced.getCell("A95").value = "19.5 — chart-level PictureOptions (applyTo* for bar3D walls)";
  advanced.getCell("A95").font = { bold: true };
  advanced.addChart(
    {
      type: "bar3D",
      barDir: "col",
      title: "PictureOptions on bar3D series",
      view3D: { rotX: 20, rotY: 25, rAngAx: true },
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          pictureFill: {
            image: { workbookImageId: logoId },
            fillMode: "stackScale",
            scale: 1,
            applyToFront: true,
            applyToSides: true,
            applyToEnd: true
          }
        }
      ]
    },
    "A96:L115"
  );
  const picOpts = advanced.getCharts()[4];
  picOpts.mutate(model => {
    const firstGroup = model.chart.plotArea.chartTypes[0];
    // `pictureOptions` is a bar/area-specific slot (CT_BarSer /
    // CT_AreaSer / CT_Bar3DSer). Narrow to a bar-series model so TS
    // sees the field.
    if (firstGroup?.type === "bar" || firstGroup?.type === "bar3D") {
      const firstSeries = firstGroup.series[0];
      if (firstSeries) {
        // The `PictureOptions` struct lives directly on the series
        // model alongside `pictureFill`; Excel reads it to decide which
        // faces of each 3D bar receive the picture.
        firstSeries.pictureOptions = {
          applyToFront: true,
          applyToSides: true,
          applyToEnd: false,
          pictureFormat: "stackScale",
          pictureStackUnit: 2
        };
      }
    }
  });

  // --- 19.6 bandFormats on surface chart — colour each contour band
  advanced.getCell("A118").value = "19.6 — bandFormats on surface (per-level colour bands)";
  advanced.getCell("A118").font = { bold: true };
  advanced.addChart(
    {
      type: "surface",
      title: "Surface with explicit band colours",
      series: [
        { name: "North", categories: "Surface!$B$1:$E$1", values: "Surface!$B$2:$E$2" },
        { name: "South", categories: "Surface!$B$1:$E$1", values: "Surface!$B$3:$E$3" },
        { name: "East", categories: "Surface!$B$1:$E$1", values: "Surface!$B$4:$E$4" },
        { name: "West", categories: "Surface!$B$1:$E$1", values: "Surface!$B$5:$E$5" }
      ],
      bandFormats: [
        { index: 0, spPr: { fill: "1F3864" } },
        { index: 1, spPr: { fill: "2E75B6" } },
        { index: 2, spPr: { fill: "8FAADC" } },
        { index: 3, spPr: { fill: "D9E1F2" } }
      ]
    },
    "A119:L138"
  );

  // ---------------------------------------------------------------------------
  // 20. Chart copying, vendor extensions, and external ChartEx parsing.
  // ---------------------------------------------------------------------------
  //
  //   20.1  `chart.copyTo(targetWorksheet)` — clone a chart onto another
  //         sheet (deep-copies the model + chart-part relationships +
  //         userShapes XML).
  //   20.2  `chart.unknownElements` — `c15:` / `cx14:` vendor extensions
  //         surviving a parse/serialise round-trip.
  //   20.3  `parseChartEx(xml)` — parse an external ChartEx XML string
  //         directly into a structured model without going through
  //         workbook I/O.

  const copyPlayground = wb.addWorksheet("20a-Copy Source");
  copyPlayground.addRow(["Chart copy source"]);
  copyPlayground.getCell("A1").font = { bold: true };
  copyPlayground.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Original",
      series: [
        {
          name: "Revenue",
          categories: "Sales!$A$2:$A$7",
          values: "Sales!$B$2:$B$7",
          fill: "4472C4"
        }
      ]
    },
    "A3:L22"
  );
  const sourceChart = copyPlayground.getCharts()[0];

  // --- 20.1 deep-copy the chart to a different worksheet
  const copyDest = wb.addWorksheet("20b-Copy Destination");
  copyDest.addRow(["Copies of the chart on the previous sheet"]);
  copyDest.getCell("A1").font = { bold: true };
  const copy1Num = sourceChart.copyTo(copyDest, "A3:L22");
  const copy2Num = sourceChart.copyTo(copyDest, {
    tl: { col: 1, row: 24 },
    ext: { cx: 5 * 914400, cy: 3 * 914400 },
    editAs: "oneCell"
  });
  console.log(`Copied chart to sheet '${copyDest.name}' as #${copy1Num} and #${copy2Num}`);

  // Modify the copy without touching the source — prove the clone is detached.
  copyDest.getCharts()[0].title = "Copy (independent from source)";
  copyDest.getCharts()[0].updateSeries(0, { fill: "70AD47" });

  // --- 20.2 vendor extensions (`c15:` / `cx14:`) round-trip
  const extSheet = wb.addWorksheet("20c-Vendor Extensions");
  extSheet.addRow(["Vendor extensions surviving round-trip"]);
  extSheet.getCell("A1").font = { bold: true };

  // Build a chart programmatically, then inject a synthetic `c15:` /
  // `c16:` extension directly into the chart model's `extLst` bag. On
  // re-parse the parser will record any child element it doesn't
  // structurally recognise into `ChartModel.unknownElements`; the
  // writer in strict-template-mode can then refuse to drop it.
  extSheet.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Has a vendor extension",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A3:L22"
  );
  const ext = extSheet.getCharts()[0];
  // Emulate a loaded chart that observed a vendor-only child element.
  // In real usage this happens automatically via `parseChart` when
  // reading an Excel-authored file; here we seed it so we can
  // demonstrate the getter.
  ext.mutate(model => {
    model.unknownElements = [
      { name: "c15:pivotSourceOverride", path: "c:chartSpace/c:extLst/c15:pivotSourceOverride" },
      { name: "c16:cachedValueSource", path: "c:chartSpace/c:extLst/c16:cachedValueSource" }
    ];
  });
  console.log(`Unknown elements on ext chart:`, ext.unknownElements);

  // --- 20.3 parse an external ChartEx XML string
  const syntheticChartExXml = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">`,
    `  <cx:chartData>`,
    `    <cx:data id="0">`,
    `      <cx:strDim type="cat">`,
    `        <cx:f>Sheet1!$A$1:$A$3</cx:f>`,
    `        <cx:lvl ptCount="3">`,
    `          <cx:pt idx="0">Alpha</cx:pt>`,
    `          <cx:pt idx="1">Beta</cx:pt>`,
    `          <cx:pt idx="2">Gamma</cx:pt>`,
    `        </cx:lvl>`,
    `      </cx:strDim>`,
    `      <cx:numDim type="val">`,
    `        <cx:f>Sheet1!$B$1:$B$3</cx:f>`,
    `        <cx:lvl ptCount="3" formatCode="General">`,
    `          <cx:pt idx="0">10</cx:pt>`,
    `          <cx:pt idx="1">25</cx:pt>`,
    `          <cx:pt idx="2">17</cx:pt>`,
    `        </cx:lvl>`,
    `      </cx:numDim>`,
    `    </cx:data>`,
    `  </cx:chartData>`,
    `  <cx:chart>`,
    `    <cx:plotArea>`,
    `      <cx:plotAreaRegion>`,
    `        <cx:plotSurface/>`,
    `        <cx:series layoutId="funnel" hidden="0" ownerIdx="0">`,
    `          <cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="1" value="1"/></cx:dataLabels>`,
    `          <cx:dataId val="0"/>`,
    `        </cx:series>`,
    `      </cx:plotAreaRegion>`,
    `    </cx:plotArea>`,
    `  </cx:chart>`,
    `</cx:chartSpace>`
  ].join("\n");
  const parsedEx = parseChartEx(syntheticChartExXml);
  console.log(
    `parseChartEx → series count: ${parsedEx.chartSpace.chart.plotArea.plotAreaRegion?.series.length}`,
    `first layoutId: ${parsedEx.chartSpace.chart.plotArea.plotAreaRegion?.series[0]?.layoutId}`,
    `data entries: ${parsedEx.chartSpace.chartData.data.length}`
  );

  // ---------------------------------------------------------------------------
  // 21. Chartsheet lifecycle — rename, copy, state, remove, and multiple
  //     chartsheets in the same workbook with a specific ordering.
  // ---------------------------------------------------------------------------

  // Start with a plain chartsheet we can mutate without breaking anything
  // the preview pipeline depends on.
  const csLifecycle = wb.addChartsheet("Lifecycle-Original", {
    chart: {
      type: "bar",
      barDir: "col",
      title: "Original lifecycle chart",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    }
  });

  // rename — returns true on success
  csLifecycle.rename("Lifecycle-Renamed");
  console.log(`Renamed chartsheet: ${csLifecycle.name}`);

  // state — "visible" | "hidden" | "veryHidden"
  csLifecycle.state = "hidden"; // hidden from the tab bar; still valid XLSX
  console.log(`Hidden chartsheet state: ${csLifecycle.state}`);

  // zoomScale / zoomToFit can be adjusted at any time
  csLifecycle.zoomScale = 110;
  csLifecycle.zoomToFit = false;

  // copy — deep-clones the chartsheet into a new tab with its own chart.
  const csCopy = csLifecycle.copy("Lifecycle-Copy");
  if (csCopy) {
    csCopy.state = "visible";
    console.log(`Copied chartsheet: ${csCopy.name}`);
  }

  // A throwaway chartsheet we immediately remove to show the API.
  const doomed = wb.addChartsheet("Lifecycle-Doomed", {
    chart: {
      type: "pie",
      title: "To be removed",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    }
  });
  const removed = doomed.remove();
  console.log(`Removed 'Lifecycle-Doomed': ${removed}`);

  // ---------------------------------------------------------------------------
  // 22. Sparklines — in-cell mini charts (x14:sparklineGroups).
  // ---------------------------------------------------------------------------
  //
  // Sparklines live in a separate module (@excel/sparkline) because they
  // are not proper chart parts — they are stored inside the worksheet's
  // extension list, and Excel renders them as cell-level graphics.
  // Three types are supported: line / column / stacked (win-loss).

  const spark = wb.addWorksheet("22-Sparklines");
  spark.getColumn(1).width = 16;
  spark.getColumn(8).width = 22;
  spark.getRow(1).values = ["Metric", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Trend"];
  spark.getRow(1).font = { bold: true };
  spark.addRow(["Revenue", 120, 180, 160, 205, 232, 248]);
  spark.addRow(["Profit", 32, 49, 41, 64, 75, 81]);
  spark.addRow(["Delta", -5, 12, -3, 18, -2, 9]);
  spark.addRow(["Units", 540, 620, 610, 710, 805, 880]);

  // Line sparkline with high/low/first/last markers
  spark.addSparklineGroup({
    type: "line",
    markers: true,
    high: true,
    low: true,
    first: true,
    last: true,
    lineWeight: 1,
    lineColor: "5B9BD5",
    highColor: "70AD47",
    lowColor: "C00000",
    firstColor: "ED7D31",
    lastColor: "FFC000",
    markerColor: "595959",
    sparklines: [{ dataRef: "'22-Sparklines'!B2:G2", cellRef: "H2" }]
  });

  // Column sparkline
  spark.addSparklineGroup({
    type: "column",
    lineColor: "4472C4",
    sparklines: [{ dataRef: "'22-Sparklines'!B3:G3", cellRef: "H3" }]
  });

  // Win-loss sparkline (stacked type) with negative bars
  spark.addSparklineGroup({
    type: "stacked",
    negative: true,
    lineColor: "70AD47",
    negativeColor: "C00000",
    axisColor: "A6A6A6",
    displayXAxis: true,
    sparklines: [{ dataRef: "'22-Sparklines'!B4:G4", cellRef: "H4" }]
  });

  // Grouped sparklines — one group, multiple sparklines sharing styling.
  const shared: AddSparklineGroupOptions = {
    type: "line",
    markers: true,
    lineColor: "264478",
    minAxisType: "group",
    maxAxisType: "group",
    sparklines: [
      { dataRef: "'22-Sparklines'!B2:G2", cellRef: "J2" },
      { dataRef: "'22-Sparklines'!B3:G3", cellRef: "J3" },
      { dataRef: "'22-Sparklines'!B5:G5", cellRef: "J5" }
    ]
  };
  spark.addSparklineGroup(shared);

  const allSparklineGroups: SparklineGroup[] = spark.getSparklineGroups();
  console.log(`Sparkline groups on the sheet: ${allSparklineGroups.length}`);

  // ---------------------------------------------------------------------------
  // 23. Rare chart features — multi-level categorical axis, view3D extras,
  //     ChartEx rawLayoutId passthrough, custom / preset geometry on
  //     shape properties, transform (rotation).
  // ---------------------------------------------------------------------------

  const rare = wb.addWorksheet("23-Rare Features");
  rare.addRow(["Rare features"]);
  rare.getCell("A1").font = { bold: true };

  // --- 23.1 multi-level categorical axis — Year / Quarter / Month
  const mlData = wb.addWorksheet("23a-MultiLevel Data");
  mlData.addRow(["Year", "Quarter", "Value"]);
  mlData.addRow(["2023", "Q1", 120]);
  mlData.addRow(["", "Q2", 180]);
  mlData.addRow(["", "Q3", 160]);
  mlData.addRow(["", "Q4", 205]);
  mlData.addRow(["2024", "Q1", 232]);
  mlData.addRow(["", "Q2", 248]);
  mlData.addRow(["", "Q3", 260]);
  mlData.addRow(["", "Q4", 290]);

  rare.getCell("A3").value = "23.1 — multi-level string reference (category axis Year→Quarter)";
  rare.getCell("A3").font = { bold: true };
  rare.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Revenue per quarter grouped by year",
      series: [
        {
          name: "Revenue",
          // Passing an AxisDataSource directly — the builder recognises
          // `multiLvlStrRef` and emits a proper `<c:multiLvlStrRef>` on
          // the category axis. The formula spans two columns (Year +
          // Quarter), and cache-populator fills both levels from the
          // worksheet.
          categories: { multiLvlStrRef: { formula: "'23a-MultiLevel Data'!$A$2:$B$9" } },
          values: "'23a-MultiLevel Data'!$C$2:$C$9"
        }
      ]
    },
    "A4:L23"
  );

  // --- 23.2 view3D with depthPercent + hPercent
  rare.getCell("A26").value = "23.2 — view3D depthPercent (z-axis depth) + hPercent (wall height)";
  rare.getCell("A26").font = { bold: true };
  rare.addChart(
    {
      type: "bar3D",
      barDir: "col",
      title: "Deep 3D bar",
      view3D: {
        rotX: 20,
        rotY: 20,
        rAngAx: true,
        perspective: 30,
        depthPercent: 200, // extra depth → bars extruded further
        hPercent: 150 // walls stretched vertically
      },
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A27:L46"
  );

  // --- 23.3 ChartEx rawLayoutId passthrough — vendor / future layoutId
  rare.getCell("A49").value = "23.3 — ChartEx rawLayoutId (vendor / future layoutId passthrough)";
  rare.getCell("A49").font = { bold: true };
  rare.addChartEx(
    {
      type: "treemap",
      title: "Preserved rawLayoutId",
      categories: "Hierarchy!$C$2:$C$10",
      series: [
        {
          name: "Sales",
          values: "Hierarchy!$D$2:$D$10",
          hierarchy: ["Hierarchy!$A$2:$A$10", "Hierarchy!$B$2:$B$10"]
        }
      ]
    },
    "A50:L69"
  );
  rare.getCharts()[2].mutateChartEx(model => {
    const series = model.chartSpace.chart.plotArea.plotAreaRegion?.series?.[0];
    if (series) {
      // Simulate a series whose @layoutId is a vendor extension we
      // don't have a structured enum entry for. The parser falls back
      // to `"clusteredColumn"` and stashes the original value in
      // `rawLayoutId`; the writer re-emits the raw attribute when it
      // sees the "fallback + rawLayoutId" combo.
      series.rawLayoutId = "vnd:experimentalMosaic";
    }
  });

  // --- 23.4 preset geometry + custom geometry on a series spPr
  rare.getCell("A72").value = "23.4 — PresetGeometry (roundRect) + CustomGeometry (star path)";
  rare.getCell("A72").font = { bold: true };
  rare.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "PresetGeometry + CustomGeometry",
      series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
    },
    "A73:L92"
  );
  rare.getCharts()[3].mutate(model => {
    const firstGroup = model.chart.plotArea.chartTypes[0];
    const firstSeries = firstGroup?.series[0];
    if (firstSeries) {
      firstSeries.spPr = {
        ...firstSeries.spPr,
        fill: { solid: { srgb: "FFC000" } },
        // Preset geometry — tells Excel the series bars have rounded corners.
        presetGeometry: {
          preset: "roundRect",
          adjustments: [{ name: "adj", fmla: "val 16667" }]
        },
        // Custom geometry — freeform star path. `w` and `h` are the
        // coordinate-space extents; commands walk a move/line sequence.
        customGeometry: {
          paths: [
            {
              w: 100000,
              h: 100000,
              fill: "norm",
              stroke: true,
              commands: [
                { type: "moveTo", points: [{ x: 50000, y: 0 }] },
                { type: "lnTo", points: [{ x: 61800, y: 38200 }] },
                { type: "lnTo", points: [{ x: 100000, y: 38200 }] },
                { type: "lnTo", points: [{ x: 69100, y: 61800 }] },
                { type: "lnTo", points: [{ x: 80900, y: 100000 }] },
                { type: "lnTo", points: [{ x: 50000, y: 76400 }] },
                { type: "lnTo", points: [{ x: 19100, y: 100000 }] },
                { type: "lnTo", points: [{ x: 30900, y: 61800 }] },
                { type: "lnTo", points: [{ x: 0, y: 38200 }] },
                { type: "lnTo", points: [{ x: 38200, y: 38200 }] },
                { type: "close" }
              ]
            }
          ]
        },
        // Shape transform — rotate the series shape 15° clockwise.
        transform: {
          offsetX: 0,
          offsetY: 0,
          width: 100000,
          height: 100000,
          rotation: 15 * 60000, // OOXML stores rotation in 1/60000° units
          flipHorizontal: false,
          flipVertical: false
        }
      };
    }
  });

  // ---------------------------------------------------------------------------
  // 24. Shape / text property helpers — `parseSpPr` / `buildSpPr` /
  //     `getSpPrFillColor` / `getSpPrLine` / `getSpPrGradient` /
  //     `getSpPrPattern` / `setSpPrFill` / `setSpPrLine` / `parseTxPr` /
  //     `buildTxPr` / `getTxPrFontSize` / `getTxPrColor`.
  //
  // These are programmatic helpers for reading / writing the structured
  // `ShapeProperties` + `ChartTextProperties` used throughout the chart
  // module. Useful when a consumer wants to operate on a loaded chart's
  // shape tree without hand-rolling DrawingML XML.
  // ---------------------------------------------------------------------------

  // Round-trip: raw DrawingML XML → parsed model → structured query →
  // mutate → re-serialise.
  const rawSpPrXml =
    "<c:spPr>" +
    '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>' +
    '<a:ln w="19050"><a:solidFill><a:srgbClr val="264478"/></a:solidFill><a:prstDash val="dash"/></a:ln>' +
    "</c:spPr>";
  const parsedSpPr = parseSpPr({ _rawXml: rawSpPrXml });
  const fillColor = getSpPrFillColor(parsedSpPr);
  const line = getSpPrLine(parsedSpPr);
  const gradient = getSpPrGradient(parsedSpPr);
  const pattern = getSpPrPattern(parsedSpPr);
  console.log(
    `parseSpPr → fill=${fillColor?.srgb}, line.dash=${line?.dash}, line.width=${line?.width}, gradient=${gradient ? "yes" : "none"}, pattern=${pattern ? "yes" : "none"}`
  );

  // `setSpPrFill` / `setSpPrLine` — functional setters that return a new spPr.
  let mutatedSpPr = setSpPrFill(parsedSpPr, { solid: { srgb: "70AD47" } });
  mutatedSpPr = setSpPrLine(mutatedSpPr, {
    color: { srgb: "375623" },
    width: 25400,
    dash: "solid"
  });
  const builtXml = buildSpPr(mutatedSpPr);
  console.log(`buildSpPr → output shape: ${builtXml ? "ShapeProperties" : "empty"}`);

  // Text properties round-trip (the tx counterpart of spPr).
  const rawTxPrXml =
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1">' +
    '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>' +
    '<a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr>';
  const parsedTxPr = parseTxPr({ _rawXml: rawTxPrXml });
  console.log(
    `parseTxPr → fontSize=${getTxPrFontSize(parsedTxPr)} color=${getTxPrColor(parsedTxPr)?.srgb}`
  );
  const rebuiltTxPr = buildTxPr({
    size: 1200,
    bold: false,
    italic: true,
    color: { srgb: "4472C4" },
    fontFamily: "Arial"
  });
  console.log(`buildTxPr → rebuilt: size=${getTxPrFontSize(rebuiltTxPr)}`);

  // ---------------------------------------------------------------------------
  // 25. Sidecar parsers / writers — `parseChartStyle` / `buildChartStyle` /
  //     `parseChartColors` / `buildChartColors`. These operate on the
  //     `xl/charts/styleN.xml` + `xl/charts/colorsN.xml` payloads
  //     directly, without going through the Chart model. Useful for
  //     tooling that needs to inspect or edit the sidecars.
  // ---------------------------------------------------------------------------

  // Build a style + colors sidecar pair, serialise to XML, then parse
  // them back — round-trip proves the emitted XML is readable by the
  // library itself.
  const styleModel: ChartStyleModel = {
    id: 251,
    elements: {
      chartArea: { fillRefIdx: 1, lnRefIdx: 1, effectRefIdx: 0, fontRefIdx: "minor" },
      title: { fontRefIdx: "major", effectRefIdx: 1 },
      dataPoint: { fillRefIdx: 2, lnRefIdx: 0, effectRefIdx: 0, fontRefIdx: "minor" },
      dataLabel: { fontRefIdx: "minor" },
      legend: { fontRefIdx: "minor" },
      gridlineMajor: { lnRefIdx: 1, effectRefIdx: 0 }
    }
  };
  const styleXml = buildChartStyle(styleModel);
  const reparsedStyle = parseChartStyle(styleXml);
  console.log(
    `buildChartStyle → parseChartStyle — elements: ${Object.keys(reparsedStyle.elements ?? {}).join(", ")}`
  );

  const colorsModel: ChartColorsModel = {
    method: "cycle",
    id: 10,
    colors: [
      { srgb: "4472C4" },
      { srgb: "ED7D31" },
      { srgb: "A5A5A5" },
      { theme: "accent4", tint: 20000 } // theme is the scheme name; tint is raw OOXML integer scale
    ],
    variations: [{ lumMod: 80000, lumOff: 20000 }, { tint: 50000 }]
  };
  const colorsXml = buildChartColors(colorsModel);
  const reparsedColors = parseChartColors(colorsXml);
  console.log(
    `buildChartColors → parseChartColors — palette: ${reparsedColors.colors?.length} colours, ${reparsedColors.variations?.length} variations`
  );

  // Attach the programmatically-built sidecar to an actual chart.
  const sidecar = wb.addWorksheet("25-Sidecar Attached");
  sidecar.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Structured sidecar via parse/build round-trip",
      series: [
        { name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" },
        { name: "Profit", categories: "Sales!$A$2:$A$7", values: "Sales!$C$2:$C$7" }
      ],
      chartStyle: reparsedStyle,
      chartColors: reparsedColors
    },
    "A1:L20"
  );

  // ---------------------------------------------------------------------------
  // 26. Cache populator API — `fillChartCaches` / `fillChartExCaches` /
  //     `fillNumRef` / `fillStrRef`. The worksheet-side registration
  //     path (`_registerChart`) invokes these automatically, but the
  //     public entry points let tools that build charts outside the
  //     worksheet machinery seed their own caches.
  // ---------------------------------------------------------------------------

  // Build a chart model by hand (no `ws.addChart`) and populate its
  // caches against the workbook. Afterwards the number/string caches
  // will have real values even though the model was never attached to
  // a worksheet registrar.
  const standaloneModel = buildChartModel({
    type: "line",
    title: "Headless-populated caches",
    series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
  });
  fillChartCaches(standaloneModel, wb);
  const firstSeries = standaloneModel.chart.plotArea.chartTypes[0].series[0];
  const valCache = (firstSeries as { val?: { numRef?: NumberReference } }).val?.numRef?.cache;
  console.log(
    `fillChartCaches → values cached: ${valCache?.points?.length ?? 0} points (e.g. ${valCache?.points?.[0]?.value})`
  );

  // Targeted cache fill for a single reference — e.g. a programmatically
  // built `NumberReference` that does NOT live inside a chart model.
  const customNumRef: NumberReference = {
    formula: "Sales!$B$2:$B$7",
    cache: { points: [] }
  };
  fillNumRef(customNumRef, wb);
  console.log(`fillNumRef → ${customNumRef.cache?.points?.length ?? 0} number points populated`);

  const customStrRef: StringReference = {
    formula: "Sales!$A$2:$A$7",
    cache: { points: [] }
  };
  fillStrRef(customStrRef, wb);
  console.log(`fillStrRef → ${customStrRef.cache?.points?.length ?? 0} string points populated`);

  // ChartEx flavour — same idea, different walker.
  const cxStandalone = buildChartExModel({
    type: "treemap",
    title: "Headless ChartEx with filled caches",
    categories: "Hierarchy!$C$2:$C$10",
    series: [
      {
        name: "Sales",
        values: "Hierarchy!$D$2:$D$10",
        hierarchy: ["Hierarchy!$A$2:$A$10", "Hierarchy!$B$2:$B$10"]
      }
    ]
  });
  fillChartExCaches(cxStandalone, wb);
  const cxData = cxStandalone.chartSpace.chartData.data;
  console.log(
    `fillChartExCaches → cx:data entries: ${cxData.length} (first levels: ${cxData[0]?.numDim?.levels?.length ?? 0})`
  );

  // ---------------------------------------------------------------------------
  // 27. Low-level renderer + TopoJSON region map + preset apply.
  //
  //   27.1  `renderChartSvg` / `renderChartPng` / `renderChartExSvg` /
  //         `renderChartExPng` directly against a ChartModel — skip the
  //         Chart wrapper entirely (useful for consumers that already
  //         have a parsed model in hand and don't need a `Worksheet`).
  //   27.2  `buildChartScene` — the intermediate scene IR that drives
  //         SVG / PNG / PDF rendering; emitting it lets consumers
  //         implement their own back-end (e.g. pdfmake, Canvas 2D).
  //   27.3  `applyAxisTransform` — the log/linear coordinate transform
  //         the scene builder uses internally.
  //   27.4  `buildEffectFilter` — build an SVG `<filter>` definition
  //         from an `EffectList`; callers doing custom SVG composition
  //         can reuse it.
  //   27.5  `VECTOR_PDF_CHART_EX_LAYOUT_IDS` — readonly set exposing
  //         which ChartEx layouts take the vector PDF path.
  //   27.6  `resolveTopologyObject` + regionMap rendered from a
  //         synthetic two-country TopoJSON (no external data-set
  //         required).
  //   27.7  `applyChartPreset` / `applyChartExPreset` invoked directly
  //         (vs via `worksheet.addPresetChart`).
  // ---------------------------------------------------------------------------

  // --- 27.1 direct renderChart(Svg|Png)
  const svgDirect = renderChartSvg(standaloneModel, { width: 640, height: 360 });
  writeFileSync(resolve(OUT_DIR, "charts-example-lowlevel.svg"), svgDirect, "utf-8");
  const pngDirect = await renderChartPng(standaloneModel, { width: 640, height: 360, scale: 2 });
  writeFileSync(resolve(OUT_DIR, "charts-example-lowlevel.png"), pngDirect);

  // --- 27.1 direct renderChartEx(Svg|Png)
  const svgExDirect = renderChartExSvg(cxStandalone, { width: 640, height: 360 });
  writeFileSync(resolve(OUT_DIR, "charts-example-lowlevel-ex.svg"), svgExDirect, "utf-8");
  const pngExDirect = await renderChartExPng(cxStandalone, { width: 640, height: 360, scale: 2 });
  writeFileSync(resolve(OUT_DIR, "charts-example-lowlevel-ex.png"), pngExDirect);
  console.log(
    `Low-level renderers — classic svg: ${svgDirect.length} bytes, ex svg: ${svgExDirect.length} bytes`
  );

  // --- 27.2 buildChartScene — the shared IR used by all backends.
  const scene = buildChartScene(standaloneModel, { width: 640, height: 360 });
  console.log(
    `buildChartScene → bounds=${scene.width}×${scene.height}, series=${scene.series?.length ?? 0}`
  );

  // --- 27.3 applyAxisTransform — log/linear value coordinate mapping.
  const linear = applyAxisTransform(500, undefined);
  const log10 = applyAxisTransform(500, 10);
  console.log(`applyAxisTransform — linear(500)=${linear}, log10(500)=${log10.toFixed(3)}`);

  // --- 27.4 buildEffectFilter — SVG `<filter>` for a structured EffectList.
  const effectFilterXml = buildEffectFilter("demoFilter", {
    outerShadow: {
      blurRadius: 50800,
      distance: 38100,
      direction: 5400000,
      color: { srgb: "000000", alpha: 40000 }
    },
    glow: {
      radius: 63500,
      color: { srgb: "4472C4", alpha: 60000 }
    }
  });
  console.log(`buildEffectFilter → ${effectFilterXml.length} bytes of <filter> XML`);

  // --- 27.5 VECTOR_PDF_CHART_EX_LAYOUT_IDS — which ChartEx layouts
  //         render as PDF vector content (vs. rasterised). The set
  //         grows over time; consumers that decide their own render
  //         path can inspect it.
  console.log(`VECTOR_PDF_CHART_EX_LAYOUT_IDS: ${VECTOR_PDF_CHART_EX_LAYOUT_IDS.join(", ")}`);

  // --- 27.6 resolveTopologyObject + regionMap with real polygons
  //
  // A tiny in-memory TopoJSON with two "countries" encoded as
  // quantised integer arcs. World-atlas bundles look exactly like
  // this, just bigger. Using the full TopoJSON path means the
  // renderer draws each country's polygon (not just a centroid dot).
  const syntheticTopology: TopologyLike = {
    type: "Topology",
    transform: {
      // Quantisation grid: the `arcs` are integers; `scale` +
      // `translate` turn them back into lon/lat.
      scale: [0.01, 0.01],
      translate: [-10, -10]
    },
    arcs: [
      // Arc 0 — simple square around (0,0)..(10,10). Delta-encoded:
      // first coord is absolute; subsequent coords are deltas.
      [
        [1000, 1000],
        [1000, 0],
        [0, 1000],
        [-1000, 0],
        [0, -1000]
      ],
      // Arc 1 — simple square around (20,20)..(30,30).
      [
        [3000, 3000],
        [1000, 0],
        [0, 1000],
        [-1000, 0],
        [0, -1000]
      ]
    ],
    objects: {
      countries: {
        type: "GeometryCollection",
        geometries: [
          {
            type: "Polygon",
            id: "ALPHA",
            properties: { name: "Alphaland" },
            arcs: [[0]]
          },
          {
            type: "Polygon",
            id: "BETA",
            properties: { name: "Betaland" },
            arcs: [[1]]
          }
        ]
      }
    }
  };

  // Validate the resolver directly — good pattern for applications
  // that want to pre-validate a user-supplied topology before
  // forwarding it to the renderer.
  const resolved = resolveTopologyObject(syntheticTopology, "countries");
  console.log(
    `resolveTopologyObject — ${resolved.length} features: ${resolved.map(f => f.id).join(", ")}`
  );

  // Build a regionMap chart that uses the synthetic polygons at
  // render-time. The `categories` / `values` data identifies which
  // feature gets which value; the topology draws the boundary.
  const topoSheet = wb.addWorksheet("27-TopoJSON RegionMap");
  topoSheet.addRow(["Country", "Value"]);
  topoSheet.addRow(["Alphaland", 820]);
  topoSheet.addRow(["Betaland", 430]);
  topoSheet.addChartEx(
    {
      type: "regionMap",
      title: "RegionMap with synthetic TopoJSON polygons",
      categories: "'27-TopoJSON RegionMap'!$A$2:$A$3",
      series: [{ name: "Value", values: "'27-TopoJSON RegionMap'!$B$2:$B$3" }],
      layout: { projection: "mercator", regionLabels: "showAll" }
    },
    "D1:L20"
  );

  // Render the region-map chart with the topology plumbed through.
  const regionMapOpts: RegionMapDataOptions = {
    topology: syntheticTopology,
    objectName: "countries",
    match: ["property:name", "id"],
    projection: "mercator",
    strokeColor: "#FFFFFF"
  };
  const regionMapChart = topoSheet.getCharts()[0];
  const regionMapSvg = regionMapChart.toSVG({
    width: 640,
    height: 400,
    regionMap: regionMapOpts
  });
  writeFileSync(resolve(OUT_DIR, "charts-example-regionmap-topojson.svg"), regionMapSvg, "utf-8");

  // --- 27.7 applyChartPreset / applyChartExPreset direct invocation
  //
  // `worksheet.addPresetChart("lineMarkers", opts, range)` is a thin
  // wrapper over `applyChartPreset("lineMarkers", opts)` +
  // `worksheet.addChart(mergedOpts, range)`. Splitting the two lets
  // you transform the merged options bag before it hits the builder
  // (e.g. to mutate colours, inject additional series, etc.).
  const presetMerged: AddChartOptions = applyChartPreset("lineMarkers", {
    title: "Applied preset, then further edited",
    series: [{ name: "Revenue", categories: "Sales!$A$2:$A$7", values: "Sales!$B$2:$B$7" }]
  });
  // Post-hoc edit the merged options — swap in a stacked grouping.
  presetMerged.grouping = "stacked";
  presetMerged.showMarker = true;
  const presetWs = wb.addWorksheet("27b-Applied Presets");
  presetWs.addChart(presetMerged, "A1:L20");

  const presetExMerged = applyChartExPreset("boxAndWhisker", {
    title: "Applied ChartEx preset",
    categories: "Distribution!$A$1:$A$1",
    series: [{ name: "Samples", values: "Distribution!$A$2:$A$41" }]
  });
  presetWs.addChartEx(presetExMerged, "A22:L41");

  // ---------------------------------------------------------------------------
  // 28. Preview export — each chart → SVG, PNG, and a multi-page PDF.
  // ---------------------------------------------------------------------------

  console.log("Rendering previews …");
  const previewWorksheets = [gallery, combo, ex, features, advanced];
  let previewCounter = 0;
  const pdfDoc = new PdfDocumentBuilder();
  pdfDoc.setMetadata({ title: "ExcelTS chart previews", author: "charts example" });
  for (const ws of previewWorksheets) {
    for (const chart of ws.getCharts()) {
      previewCounter += 1;
      const title = chart.title ?? `chart-${previewCounter}`;
      const safe = title.replace(/[^\w\- ]+/g, "_").slice(0, 60);
      const svg = chart.toSVG({ width: 640, height: 400 });
      writeFileSync(resolve(OUT_DIR, `charts-example-${previewCounter}-${safe}.svg`), svg, "utf-8");
      const png = await chart.toPNG({ width: 640, height: 400, scale: 2 });
      writeFileSync(resolve(OUT_DIR, `charts-example-${previewCounter}-${safe}.png`), png);

      // Add a page to the combined PDF. Pick the renderer that matches
      // the chart flavour — classic charts go vector via `drawChartPdf`,
      // ChartEx via `drawChartExPdf`.
      const page = pdfDoc.addPage({ width: 700, height: 500 });
      const classic = chart.chartModel;
      const chartEx = chart.chartExModel;
      if (classic) {
        drawChartPdf(page, classic, { x: 30, y: 60, width: 640, height: 400 });
      } else if (chartEx) {
        drawChartExPdf(page, chartEx, { x: 30, y: 60, width: 640, height: 400 });
      }
    }
  }
  const pdfBytes = await pdfDoc.build();
  writeFileSync(PDF_PATH, pdfBytes);

  // Single-chart PDF via the high-level `chartToPdf` helper — the
  // canonical entry point from the `@cj-tech-master/excelts/pdf` bundle.
  const firstChart = gallery.getCharts()[0];
  const solo = await chartToPdf(firstChart, {
    title: "Solo chart PDF",
    width: 600,
    height: 400
  });
  writeFileSync(resolve(OUT_DIR, "charts-example-solo.pdf"), solo);

  // ---------------------------------------------------------------------------
  // 29. Load + mutate — `chart.mutate(fn, { preferRawPatch })` round-trip.
  //
  // We first write the workbook to `tmp/charts-example.xlsx`, then load
  // it back and make a narrow, template-safe edit using the raw-patch
  // path (stays byte-preserving for the rest of the chart XML).
  // ---------------------------------------------------------------------------

  await wb.xlsx.writeFile(XLSX_PATH);

  const reread = new Workbook();
  await reread.xlsx.readFile(XLSX_PATH);
  const firstSheet = reread.getWorksheet("1-Classic Gallery");
  if (!firstSheet) {
    throw new Error("Expected to read back the gallery worksheet.");
  }
  const first = firstSheet.getCharts()[0];
  if (first && first.chartModel) {
    first.mutate(
      model => {
        // Flip the first bar chart's title to something new. The mutate
        // call is tagged with `preferRawPatch`, so when the chart XML
        // already exists on disk the writer patches only the changed
        // element and leaves every other byte untouched. With
        // `requireRawPatch: true` the call would throw if the edit
        // can't be performed without a full structural rebuild.
        if (model.chart.title) {
          model.chart.title.text = {
            paragraphs: [{ runs: [{ text: "Rewritten on reload" }] }]
          };
        }
      },
      { preferRawPatch: true }
    );
  }
  await reread.xlsx.writeFile(XLSX_PATH);

  // ---------------------------------------------------------------------------
  // Done — summarise counts and paths.
  // ---------------------------------------------------------------------------

  const chartCount = wb.worksheets.reduce((total, ws) => total + ws.getCharts().length, 0);
  const chartsheetCount = wb.chartsheets.length;

  console.log("");
  console.log(`XLSX  : ${XLSX_PATH}`);
  console.log(`PDF   : ${PDF_PATH}`);
  console.log(`SVGs  : ${OUT_DIR}/charts-example-*.svg`);
  console.log(`PNGs  : ${OUT_DIR}/charts-example-*.png`);
  console.log("");
  console.log(`Worksheets with charts: ${wb.worksheets.length}`);
  console.log(`Charts embedded:       ${chartCount}`);
  console.log(`Chartsheets:           ${chartsheetCount}`);
  console.log(`Preset classic count:  ${EXCEL_CHART_PRESETS.length}`);
  console.log(`Preset chartEx count:  ${EXCEL_CHART_EX_PRESETS.length}`);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Generate a single-colour PNG with no compression — enough to exercise
 * the picture-fill pipeline without depending on a PNG encoder.
 *
 * Produces an IHDR + IDAT + IEND stream with:
 *   - bit-depth 8, colour-type 6 (RGBA)
 *   - uncompressed DEFLATE blocks ("stored" blocks)
 * Real-world callers normally have a PNG handy on disk — this is only
 * here to keep the example runnable without a network round-trip or an
 * extra dev dependency.
 */
function makeSolidColorPng(width: number, height: number, rgb: number): Uint8Array {
  const r = (rgb >> 16) & 0xff;
  const g = (rgb >> 8) & 0xff;
  const b = rgb & 0xff;
  // Build raw pixel stream: one filter byte per scanline (0 = None) +
  // RGBA samples.
  const rowBytes = 1 + width * 4;
  const raw = new Uint8Array(rowBytes * height);
  for (let y = 0; y < height; y++) {
    raw[y * rowBytes] = 0;
    for (let x = 0; x < width; x++) {
      const o = y * rowBytes + 1 + x * 4;
      raw[o + 0] = r;
      raw[o + 1] = g;
      raw[o + 2] = b;
      raw[o + 3] = 0xff;
    }
  }
  const compressed = zlibStored(raw);

  const signature = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const chunks: Uint8Array[] = [signature];
  chunks.push(pngChunk("IHDR", buildIhdr(width, height)));
  chunks.push(pngChunk("IDAT", compressed));
  chunks.push(pngChunk("IEND", new Uint8Array(0)));

  const total = chunks.reduce((sum, c) => sum + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

function buildIhdr(width: number, height: number): Uint8Array {
  const buf = new Uint8Array(13);
  const view = new DataView(buf.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  buf[8] = 8; // bit depth
  buf[9] = 6; // colour type RGBA
  buf[10] = 0;
  buf[11] = 0;
  buf[12] = 0;
  return buf;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = new TextEncoder().encode(type);
  const length = data.length;
  const out = new Uint8Array(4 + 4 + length + 4);
  const view = new DataView(out.buffer);
  view.setUint32(0, length);
  out.set(typeBytes, 4);
  out.set(data, 8);
  const crc = crc32(new Uint8Array(out.buffer, 4, 4 + length));
  view.setUint32(8 + length, crc);
  return out;
}

function crc32(bytes: Uint8Array): number {
  let c: number;
  const table: number[] = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = table[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Emit a zlib stream whose DEFLATE payload consists only of
 * "stored" (uncompressed) blocks. Keeps the PNG picture-fill demo
 * self-contained.
 */
function zlibStored(data: Uint8Array): Uint8Array {
  const MAX_BLOCK = 65535;
  const parts: Uint8Array[] = [new Uint8Array([0x78, 0x01])]; // zlib header (no dict)
  for (let i = 0; i < data.length; i += MAX_BLOCK) {
    const chunk = data.subarray(i, Math.min(i + MAX_BLOCK, data.length));
    const final = i + MAX_BLOCK >= data.length ? 1 : 0;
    const header = new Uint8Array(5);
    header[0] = final;
    const len = chunk.length;
    header[1] = len & 0xff;
    header[2] = (len >> 8) & 0xff;
    header[3] = ~len & 0xff;
    header[4] = (~len >> 8) & 0xff;
    parts.push(header, chunk);
  }
  // Adler-32 footer
  parts.push(adler32Bytes(data));
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

function adler32Bytes(data: Uint8Array): Uint8Array {
  let a = 1;
  let b = 0;
  const MOD = 65521;
  for (let i = 0; i < data.length; i++) {
    a = (a + data[i]) % MOD;
    b = (b + a) % MOD;
  }
  const adler = ((b << 16) | a) >>> 0;
  const out = new Uint8Array(4);
  const view = new DataView(out.buffer);
  view.setUint32(0, adler);
  return out;
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
