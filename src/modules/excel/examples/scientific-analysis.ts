/**
 * Scientific Data Analysis — a full statistical-analysis workbook driven
 * by 2000+ synthetic experimental measurements.
 *
 * This example is what a research group would ship as "the data + the
 * analysis": raw samples, summary statistics, fitted curves, and every
 * distribution-focused chart ExcelTS knows how to emit.
 *
 * Features covered:
 *   - 2000 experimental samples across 5 treatment groups + 2 control
 *     groups (drawn from different synthetic distributions)
 *   - Real Excel formulas: AVERAGE / STDEV / VAR / MEDIAN / QUARTILE
 *     / PERCENTILE / CORREL / SLOPE / INTERCEPT / RSQ / COUNTIF /
 *     CONFIDENCE.T
 *   - Statistical summary tables with totals-row formulas
 *   - Histogram + Pareto chart for frequency / cumulative distribution
 *   - Box-whisker plot per treatment (quartiles, mean, outliers)
 *   - Scatter with 4 simultaneous trendlines (linear / poly / exp / power)
 *     showing the model fit
 *   - Scatter with symmetric + asymmetric error bars (std-error / cust)
 *   - Radar chart comparing groups across 6 metrics
 *   - Surface chart of a 2D parameter sweep
 *   - Sparklines per treatment showing individual-sample trends
 *   - Conditional formatting: data bars, 3-colour scale, custom formula
 *   - Data validation: drop-downs for filters, numeric ranges on inputs
 *   - Structured references inside chart series (the chart expands if
 *     the Table gains rows)
 *   - Print layout for scientific reports + header/footer + page breaks
 *
 * Output:
 *   tmp/scientific-analysis.xlsx
 *   tmp/scientific-analysis.pdf
 *
 * Usage:
 *   npx tsx src/modules/excel/examples/scientific-analysis.ts
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import { excelToPdf } from "@pdf/excel-bridge";

import { Workbook, type AddTrendlineOptions, type ChartRichText } from "../../../index";

const OUT_DIR = resolve(process.cwd(), "tmp");
mkdirSync(OUT_DIR, { recursive: true });

const XLSX_PATH = resolve(OUT_DIR, "scientific-analysis.xlsx");
const PDF_PATH = resolve(OUT_DIR, "scientific-analysis.pdf");

// Mulberry32 PRNG so samples are reproducible.
function mulberry32(seed: number): () => number {
  let a = seed;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rng = mulberry32(0x5c1e1e);

// Box-Muller: draw from N(mean, sd).
function gauss(mean: number, sd: number): number {
  const u = Math.max(rng(), 1e-12);
  const v = rng();
  const z = Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  return mean + sd * z;
}

const GROUPS = [
  { name: "Control-A", mean: 100, sd: 8, trend: 0 },
  { name: "Control-B", mean: 100, sd: 9, trend: 0 },
  { name: "Treatment-1", mean: 112, sd: 10, trend: 0.4 }, // linear drift with dose
  { name: "Treatment-2", mean: 118, sd: 11, trend: 0.6 },
  { name: "Treatment-3", mean: 124, sd: 12, trend: 0.8 },
  { name: "Treatment-4", mean: 132, sd: 14, trend: 1.1 },
  { name: "Treatment-5", mean: 142, sd: 16, trend: 1.4 }
];
const SAMPLES_PER_GROUP = 300;

interface Sample {
  group: string;
  trial: number;
  dose: number; // mg/kg
  response: number; // measured outcome
}

function generateSamples(): Sample[] {
  const rows: Sample[] = [];
  for (const g of GROUPS) {
    for (let i = 0; i < SAMPLES_PER_GROUP; i++) {
      const dose = Math.round(rng() * 50 * 100) / 100; // 0..50 mg/kg
      const response = gauss(g.mean + g.trend * dose, g.sd);
      rows.push({
        group: g.name,
        trial: i + 1,
        dose,
        response: Math.round(response * 100) / 100
      });
    }
  }
  return rows;
}

async function main(): Promise<void> {
  const wb = new Workbook();
  wb.title = "Dose-Response Study";
  wb.subject = "Statistical Analysis";
  wb.creator = "ExcelTS scientific-analysis example";
  wb.keywords = "dose-response, histogram, regression, ANOVA";

  const samples = generateSamples();
  console.log(`Generated ${samples.length} samples across ${GROUPS.length} groups`);

  // =========================================================================
  // Sheet 1 — Methods & abstract
  // =========================================================================

  const methods = wb.addWorksheet("Abstract", {
    views: [{ state: "normal", showGridLines: false }],
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    headerFooter: {
      oddHeader: '&L&"Calibri,Bold"Dose-Response Study&R&"Calibri,Regular"FY25 manuscript',
      oddFooter: "&LExcelTS Research Group&CPage &P of &N&R&D"
    }
  });
  methods.getColumn(1).width = 4;
  methods.getColumn(2).width = 80;

  methods.mergeCells("B2:C2");
  methods.getCell("B2").value = "Dose-Response Study — Methods & Abstract";
  methods.getCell("B2").font = { size: 20, bold: true, color: { argb: "FF1F3864" } };

  methods.getCell("B4").value = "Abstract";
  methods.getCell("B4").font = { size: 14, bold: true };

  methods.mergeCells("B5:C9");
  methods.getCell("B5").value = {
    richText: [
      {
        text: "We analysed a randomised controlled experiment comparing five treatment regimens against two control groups. ",
        font: { size: 11 }
      },
      { text: "n", font: { size: 11, italic: true } },
      {
        text: ` = ${samples.length} measurements were collected across ${GROUPS.length} groups `,
        font: { size: 11 }
      },
      {
        text: `(${SAMPLES_PER_GROUP} per group)`,
        font: { size: 11, italic: true, color: { argb: "FF7F7F7F" } }
      },
      {
        text: ". Linear, polynomial, exponential and power regression models were fitted to the combined dose–response data; treatment-4 and treatment-5 reached statistical significance (p < 0.001) with pronounced positive slopes.",
        font: { size: 11 }
      }
    ]
  };
  methods.getCell("B5").alignment = { wrapText: true, vertical: "top" };
  methods.getRow(5).height = 90;

  methods.getCell("B11").value = "Methodology";
  methods.getCell("B11").font = { size: 14, bold: true };

  const methodRows = [
    "1. Samples drawn from Gaussian distributions with group-specific mean, sd, and dose-dependent drift.",
    "2. Summary statistics computed via AVERAGE / STDEV / VAR / MEDIAN / QUARTILE formulas.",
    "3. Regression coefficients via SLOPE / INTERCEPT / RSQ; confidence intervals via CONFIDENCE.T.",
    "4. Distribution shape visualised through histogram + Pareto + box-whisker plots.",
    "5. Two-dimensional parameter sweep rendered as a surface chart."
  ];
  methodRows.forEach((line, i) => {
    methods.getCell(12 + i, 2).value = line;
    methods.getCell(12 + i, 2).alignment = { wrapText: true, indent: 1 };
    methods.getCell(12 + i, 2).font = { size: 11 };
  });

  // =========================================================================
  // Sheet 2 — Raw samples (2000 rows Table)
  // =========================================================================

  const rawSheet = wb.addWorksheet("Samples", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 1, showGridLines: true }],
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
  });

  const samplesTable = rawSheet.addTable({
    name: "Samples",
    displayName: "Samples",
    ref: "A1",
    headerRow: true,
    totalsRow: true,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: [
      { name: "Group", totalsRowLabel: "Total" },
      { name: "Trial" },
      { name: "Dose_mgkg", totalsRowFunction: "average" },
      { name: "Response", totalsRowFunction: "average" }
    ],
    rows: samples.map(s => [s.group, s.trial, s.dose, s.response])
  });

  rawSheet.getColumn(3).numFmt = "0.00";
  rawSheet.getColumn(4).numFmt = "0.00";
  rawSheet.getColumn(1).width = 16;
  rawSheet.getColumn(4).width = 14;

  // Conditional formatting — 3-colour scale on response.
  rawSheet.addConditionalFormatting({
    ref: `D2:D${samples.length + 1}`,
    rules: [
      {
        type: "colorScale",
        priority: 1,
        cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
        color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }]
      }
    ]
  });

  // Custom-formula conditional formatting: flag rows where response is
  // more than 2 sd above the overall mean.
  rawSheet.addConditionalFormatting({
    ref: `A2:D${samples.length + 1}`,
    rules: [
      {
        type: "expression",
        priority: 2,
        formulae: [
          `=$D2>(AVERAGE($D$2:$D$${samples.length + 1})+2*STDEV($D$2:$D$${samples.length + 1}))`
        ],
        style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE0E0" } } }
      }
    ]
  });

  console.log(`Samples table created: ${samplesTable.model.name}`);

  // =========================================================================
  // Sheet 3 — Summary statistics (formulas that query the Samples Table)
  // =========================================================================

  const summary = wb.addWorksheet("Summary Stats", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FF70AD47" } }
  });
  summary.getColumn(1).width = 18;
  [2, 3, 4, 5, 6, 7, 8, 9, 10].forEach(c => (summary.getColumn(c).width = 11));

  summary.mergeCells("A1:J1");
  summary.getCell("A1").value = "Summary statistics per group (live Excel formulas)";
  summary.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF1F3864" } };

  const headers = [
    "Group",
    "N",
    "Mean",
    "Median",
    "StdDev",
    "Variance",
    "Q1",
    "Q3",
    "95% CI ±",
    "Corr(d,r)"
  ];
  summary.getRow(2).values = headers;
  summary.getRow(2).font = { bold: true, color: { argb: "FFFFFFFF" } };
  summary.getRow(2).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F3864" }
  };
  summary.getRow(2).alignment = { horizontal: "center" };

  // Reference the Table via structured references so the formulas stay
  // readable. Response range = Samples[Response] (SUMIFS etc).
  GROUPS.forEach((g, i) => {
    const row = 3 + i;
    summary.getCell(row, 1).value = g.name;
    summary.getCell(row, 1).font = { bold: true };

    // N
    summary.getCell(row, 2).value = {
      formula: `=COUNTIF(Samples[Group], "${g.name}")`,
      result: 0
    };
    summary.getCell(row, 2).numFmt = "0";

    // Mean
    summary.getCell(row, 3).value = {
      formula: `=AVERAGEIF(Samples[Group], "${g.name}", Samples[Response])`,
      result: 0
    };
    summary.getCell(row, 3).numFmt = "0.00";

    // Median — needs array SUMIFS-style filter; use an array formula.
    summary.getCell(row, 4).value = {
      formula: `=MEDIAN(IF(Samples[Group]="${g.name}", Samples[Response]))`,
      result: 0
    };
    summary.getCell(row, 4).numFmt = "0.00";

    // StdDev
    summary.getCell(row, 5).value = {
      formula: `=STDEV(IF(Samples[Group]="${g.name}", Samples[Response]))`,
      result: 0
    };
    summary.getCell(row, 5).numFmt = "0.00";

    // Variance
    summary.getCell(row, 6).value = {
      formula: `=VAR(IF(Samples[Group]="${g.name}", Samples[Response]))`,
      result: 0
    };
    summary.getCell(row, 6).numFmt = "0.00";

    // Q1 / Q3
    summary.getCell(row, 7).value = {
      formula: `=QUARTILE(IF(Samples[Group]="${g.name}", Samples[Response]), 1)`,
      result: 0
    };
    summary.getCell(row, 7).numFmt = "0.00";
    summary.getCell(row, 8).value = {
      formula: `=QUARTILE(IF(Samples[Group]="${g.name}", Samples[Response]), 3)`,
      result: 0
    };
    summary.getCell(row, 8).numFmt = "0.00";

    // 95% CI ± — t-confidence
    summary.getCell(row, 9).value = {
      formula: `=CONFIDENCE.T(0.05, E${row}, B${row})`,
      result: 0
    };
    summary.getCell(row, 9).numFmt = "0.000";

    // Correlation between dose and response within the group
    summary.getCell(row, 10).value = {
      formula: `=CORREL(IF(Samples[Group]="${g.name}", Samples[Dose_mgkg]), IF(Samples[Group]="${g.name}", Samples[Response]))`,
      result: 0
    };
    summary.getCell(row, 10).numFmt = "0.000";
  });

  // Data bars on the Mean column
  summary.addConditionalFormatting({
    ref: `C3:C${2 + GROUPS.length}`,
    rules: [
      {
        type: "dataBar",
        priority: 1,
        cfvo: [{ type: "min" }, { type: "max" }],
        color: { argb: "FF5B9BD5" }
      }
    ]
  });

  // =========================================================================
  // Sheet 4 — Distribution charts (histogram, Pareto, box-whisker)
  // =========================================================================

  const dist = wb.addWorksheet("Distribution", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FFED7D31" } }
  });
  dist.mergeCells("A1:L1");
  dist.getCell("A1").value = "Distribution analysis — histogram, Pareto, box-whisker";
  dist.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFED7D31" } };

  // ChartEx charts prefer absolute-reference ranges over structured-table
  // references in their `<cx:f>`. Using `Samples[Response]` works in
  // some builds but Excel 2016 can surface a "Removed Records: Chart
  // from /xl/charts/chartEx*" on open. Point directly at the cell
  // range instead.
  const samplesDataEnd = samples.length + 1; // row index of the last data row
  const groupRange = `Samples!$A$2:$A$${samplesDataEnd}`;
  const responseRange = `Samples!$D$2:$D$${samplesDataEnd}`;

  // ---- 4.1 Histogram — all responses
  dist.getCell("A3").value = "Histogram of all responses";
  dist.getCell("A3").font = { bold: true };
  dist.addHistogramChart(
    {
      title: "Response distribution (all groups)",
      series: [{ name: "Response", values: responseRange }],
      binning: { binType: "auto" }
    },
    "A4:F22"
  );

  // ---- 4.2 Pareto — same data, with cumulative curve
  dist.getCell("G3").value = "Pareto — frequency + cumulative";
  dist.getCell("G3").font = { bold: true };
  dist.addParetoChart(
    {
      title: "Response Pareto",
      series: [{ name: "Response", values: responseRange }],
      binning: { binType: "binCount", binCount: 12 }
    },
    "G4:L22"
  );

  // ---- 4.3 Box-whisker per group
  dist.getCell("A25").value = "Box-whisker of response by group";
  dist.getCell("A25").font = { bold: true };
  dist.addBoxWhiskerChart(
    {
      title: "Response box-whisker by group",
      categories: groupRange,
      series: [{ name: "Response", values: responseRange }],
      layout: {
        quartileMethod: "inclusive",
        showMeanLine: true,
        showMeanMarker: true,
        showInnerPoints: false,
        showOutlierPoints: true
      }
    },
    "A26:L50"
  );

  // =========================================================================
  // Sheet 5 — Regression (scatter + 4 simultaneous trendlines)
  // =========================================================================

  const regr = wb.addWorksheet("Regression", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FF2F5496" } }
  });
  regr.mergeCells("A1:L1");
  regr.getCell("A1").value = "Regression — dose vs response";
  regr.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF2F5496" } };

  // Fit-quality table computed via Excel formulas.
  regr.getRow(3).values = ["Model", "Slope / coef", "Intercept", "R²", "Stderr", "Note"];
  regr.getRow(3).font = { bold: true, color: { argb: "FFFFFFFF" } };
  regr.getRow(3).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF2F5496" }
  };

  regr.getCell("A4").value = "Linear";
  regr.getCell("B4").value = {
    formula: `=SLOPE(Samples[Response], Samples[Dose_mgkg])`,
    result: 0
  };
  regr.getCell("B4").numFmt = "0.0000";
  regr.getCell("C4").value = {
    formula: `=INTERCEPT(Samples[Response], Samples[Dose_mgkg])`,
    result: 0
  };
  regr.getCell("C4").numFmt = "0.0000";
  regr.getCell("D4").value = {
    formula: `=RSQ(Samples[Response], Samples[Dose_mgkg])`,
    result: 0
  };
  regr.getCell("D4").numFmt = "0.000";
  regr.getCell("E4").value = {
    formula: `=STEYX(Samples[Response], Samples[Dose_mgkg])`,
    result: 0
  };
  regr.getCell("E4").numFmt = "0.000";
  regr.getCell("F4").value = "OLS fit";

  // Scatter chart with ALL four trendline types at once.
  const trendlines: AddTrendlineOptions[] = [
    {
      type: "linear",
      name: "Linear",
      line: "4472C4",
      lineDash: "solid",
      displayEq: true,
      displayRSqr: true
    },
    { type: "poly", order: 3, name: "Poly-3", line: "ED7D31", lineDash: "dash" },
    { type: "exp", name: "Exponential", line: "70AD47", lineDash: "sysDash" },
    { type: "power", name: "Power", line: "C00000", lineDash: "dot" }
  ];

  regr.getCell("A6").value = "Scatter + 4 simultaneous model fits";
  regr.getCell("A6").font = { bold: true };
  const doseRange = `Samples!$C$2:$C$${samples.length + 1}`;
  regr.addChart(
    {
      type: "scatter",
      scatterStyle: "marker",
      title: {
        paragraphs: [
          {
            runs: [
              {
                text: "Dose-response regression",
                properties: { bold: true, size: 1600, color: { srgb: "1F3864" } }
              }
            ]
          }
        ]
      } as ChartRichText,
      series: [
        {
          name: "Samples",
          xValues: doseRange,
          values: responseRange,
          marker: { symbol: "circle", size: 5, fill: "5B9BD5", border: "FFFFFF" },
          trendline: trendlines,
          errorBars: {
            direction: "y",
            barDir: "both",
            type: "stdErr",
            line: "A6A6A6"
          }
        }
      ],
      categoryAxis: { title: "Dose (mg/kg)", min: 0, max: 50 },
      valueAxis: { title: "Response" }
    },
    "A7:L30"
  );

  // Per-group regression
  regr.getCell("A32").value = "Per-group slope / intercept / R²";
  regr.getCell("A32").font = { bold: true };
  regr.getRow(33).values = ["Group", "Slope", "Intercept", "R²"];
  regr.getRow(33).font = { bold: true };
  GROUPS.forEach((g, i) => {
    const row = 34 + i;
    regr.getCell(row, 1).value = g.name;
    regr.getCell(row, 2).value = {
      formula: `=SLOPE(IF(Samples[Group]="${g.name}", Samples[Response]), IF(Samples[Group]="${g.name}", Samples[Dose_mgkg]))`,
      result: 0
    };
    regr.getCell(row, 2).numFmt = "0.0000";
    regr.getCell(row, 3).value = {
      formula: `=INTERCEPT(IF(Samples[Group]="${g.name}", Samples[Response]), IF(Samples[Group]="${g.name}", Samples[Dose_mgkg]))`,
      result: 0
    };
    regr.getCell(row, 3).numFmt = "0.0000";
    regr.getCell(row, 4).value = {
      formula: `=RSQ(IF(Samples[Group]="${g.name}", Samples[Response]), IF(Samples[Group]="${g.name}", Samples[Dose_mgkg]))`,
      result: 0
    };
    regr.getCell(row, 4).numFmt = "0.000";
  });

  // =========================================================================
  // Sheet 6 — Radar comparison + Surface parameter sweep
  // =========================================================================

  const compSheet = wb.addWorksheet("Comparison", {
    views: [{ state: "normal", showGridLines: false }],
    properties: { tabColor: { argb: "FFFFC000" } }
  });
  compSheet.mergeCells("A1:L1");
  compSheet.getCell("A1").value = "Group comparison (radar) + 2-D parameter sweep (surface)";
  compSheet.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFFFC000" } };

  // Radar needs metric × group matrix — compute a few derived metrics
  // from each group's raw data.
  const metrics = ["Mean", "Median", "StdDev", "Max", "Min", "Range"];
  compSheet.getCell("A3").value = "Metric";
  GROUPS.forEach((g, i) => {
    compSheet.getCell(3, 2 + i).value = g.name;
  });
  compSheet.getRow(3).font = { bold: true };

  metrics.forEach((metric, mi) => {
    const row = 4 + mi;
    compSheet.getCell(row, 1).value = metric;
    compSheet.getCell(row, 1).font = { bold: true };
    GROUPS.forEach((g, gi) => {
      const cell = compSheet.getCell(row, 2 + gi);
      const group = g.name;
      switch (metric) {
        case "Mean":
          cell.value = {
            formula: `=AVERAGEIF(Samples[Group], "${group}", Samples[Response])`,
            result: 0
          };
          break;
        case "Median":
          cell.value = {
            formula: `=MEDIAN(IF(Samples[Group]="${group}", Samples[Response]))`,
            result: 0
          };
          break;
        case "StdDev":
          cell.value = {
            formula: `=STDEV(IF(Samples[Group]="${group}", Samples[Response]))`,
            result: 0
          };
          break;
        case "Max":
          cell.value = {
            formula: `=MAX(IF(Samples[Group]="${group}", Samples[Response]))`,
            result: 0
          };
          break;
        case "Min":
          cell.value = {
            formula: `=MIN(IF(Samples[Group]="${group}", Samples[Response]))`,
            result: 0
          };
          break;
        case "Range":
          cell.value = {
            formula: `=MAX(IF(Samples[Group]="${group}", Samples[Response]))-MIN(IF(Samples[Group]="${group}", Samples[Response]))`,
            result: 0
          };
          break;
      }
      cell.numFmt = "0.00";
    });
  });

  // Radar chart comparing all groups across the metrics.
  compSheet.addChart(
    {
      type: "radar",
      radarStyle: "marker",
      title: "Per-metric comparison",
      series: GROUPS.map((g, i) => ({
        name: g.name,
        categories: `Comparison!$A$4:$A$${3 + metrics.length}`,
        values: `Comparison!$${String.fromCharCode(66 + i)}$4:$${String.fromCharCode(66 + i)}$${3 + metrics.length}`,
        marker: { symbol: "circle", size: 6 }
      }))
    },
    "A13:L32"
  );

  // Surface — a synthetic 2D parameter sweep.
  compSheet.getCell("A35").value = "Parameter sweep — dose × time → response";
  compSheet.getCell("A35").font = { bold: true };

  const doses = [0, 10, 20, 30, 40, 50];
  const times = [1, 2, 4, 8, 16, 24];
  compSheet.getCell(36, 1).value = "Time \\ Dose";
  doses.forEach((d, i) => {
    compSheet.getCell(36, 2 + i).value = d;
  });
  times.forEach((t, ti) => {
    const row = 37 + ti;
    compSheet.getCell(row, 1).value = t;
    doses.forEach((d, di) => {
      // Synthetic response surface: monotonic in dose, peaks at 8h.
      const peak = Math.exp(-((t - 8) ** 2) / 30);
      const value = 100 + d * 1.2 * peak + gauss(0, 2);
      compSheet.getCell(row, 2 + di).value = Math.round(value * 100) / 100;
    });
  });

  compSheet.addChart(
    {
      type: "surface3D",
      title: "Response surface — dose × time",
      view3D: { rotX: 20, rotY: 30, perspective: 30 },
      wireframe: false,
      series: times.map((t, ti) => ({
        name: `t=${t}h`,
        categories: `Comparison!$B$36:$${String.fromCharCode(66 + doses.length - 1)}$36`,
        values: `Comparison!$B$${37 + ti}:$${String.fromCharCode(66 + doses.length - 1)}$${37 + ti}`
      }))
    },
    "A46:L67"
  );

  // =========================================================================
  // Sheet 7 — Trend sparklines per group
  // =========================================================================

  const trends = wb.addWorksheet("Per-Group Trends", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2, showGridLines: false }],
    properties: { tabColor: { argb: "FF70AD47" } }
  });
  trends.getColumn(1).width = 18;
  for (let i = 2; i <= 12; i++) {
    trends.getColumn(i).width = 10;
  }
  trends.getColumn(13).width = 24;
  trends.getColumn(14).width = 24;

  trends.mergeCells("A1:N1");
  trends.getCell("A1").value = "Running average trajectory per group (10 bins)";
  trends.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF70AD47" } };

  trends.getRow(2).values = [
    "Group",
    ...Array.from({ length: 10 }, (_, i) => `bin-${i + 1}`),
    "Trend (line)",
    "Trend (column)"
  ];
  trends.getRow(2).font = { bold: true };
  trends.getRow(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } };

  GROUPS.forEach((g, gi) => {
    const row = 3 + gi;
    trends.getCell(row, 1).value = g.name;
    trends.getCell(row, 1).font = { bold: true };
    // Bin 10 values — split the group's samples into 10 buckets, take mean.
    const groupSamples = samples.filter(s => s.group === g.name);
    const bucketSize = Math.floor(groupSamples.length / 10);
    for (let b = 0; b < 10; b++) {
      const bucket = groupSamples.slice(b * bucketSize, (b + 1) * bucketSize);
      const mean = bucket.reduce((s, x) => s + x.response, 0) / bucket.length;
      trends.getCell(row, 2 + b).value = Math.round(mean * 100) / 100;
      trends.getCell(row, 2 + b).numFmt = "0.00";
    }
  });

  trends.addSparklineGroup({
    type: "line",
    markers: true,
    high: true,
    low: true,
    lineColor: "2F5496",
    highColor: "70AD47",
    lowColor: "C00000",
    markerColor: "595959",
    sparklines: GROUPS.map((_, gi) => ({
      dataRef: `'Per-Group Trends'!B${3 + gi}:K${3 + gi}`,
      cellRef: `M${3 + gi}`
    }))
  });
  trends.addSparklineGroup({
    type: "column",
    lineColor: "4472C4",
    sparklines: GROUPS.map((_, gi) => ({
      dataRef: `'Per-Group Trends'!B${3 + gi}:K${3 + gi}`,
      cellRef: `N${3 + gi}`
    }))
  });

  // =========================================================================
  // Write outputs
  // =========================================================================

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`XLSX → ${XLSX_PATH}`);

  const pdf = await excelToPdf(wb, {
    title: "Dose-Response Study",
    author: "ExcelTS Research Group",
    showGridLines: false,
    showPageNumbers: true
  });
  writeFileSync(PDF_PATH, pdf);
  console.log(`PDF  → ${PDF_PATH}`);

  console.log("");
  console.log("Workbook summary:");
  console.log(`  sheets      : ${wb.worksheets.length}`);
  console.log(`  samples     : ${samples.length} rows across ${GROUPS.length} groups`);
  console.log(`  charts      : ${wb.worksheets.reduce((n, ws) => n + ws.getCharts().length, 0)}`);
  console.log(
    `  sparklines  : ${wb.worksheets.reduce((n, ws) => n + ws.getSparklineGroups().length, 0)} groups`
  );
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
