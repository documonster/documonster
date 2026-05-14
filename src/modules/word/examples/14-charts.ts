/**
 * Word Example 14 — Charts
 *
 * Covers the built-in DrawingML chart types via the `chart()` builder:
 *   - Column chart (clustered + stacked)
 *   - Line chart with markers + trendline
 *   - Pie / Doughnut
 *   - Area chart
 *   - Scatter + smooth scatter
 *   - Bar chart with data labels
 *   - Combo chart (column + line on secondary axis)
 *   - Custom titles, axis labels, legend position
 *   - Edge case: empty series, single category, very many categories,
 *     negative values, zero/missing values.
 *
 * Output: tmp/word-examples/14-charts.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, chart, cmToEmu, toBuffer } from "../index";
import type { ChartSeries } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word — Charts", 1);

const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];

// ---------------------------------------------------------------------------
// 1. Column chart (clustered)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Clustered column chart", 2);
Document.addContent(
  doc,
  chart({
    type: "column",
    title: "Quarterly revenue by region",
    legend: "b",
    width: cmToEmu(15),
    height: cmToEmu(8),
    categoryAxis: { title: "Quarter" },
    valueAxis: { title: "Revenue ($M)", min: 0 },
    series: [
      {
        name: "North",
        categories: QUARTERS,
        values: [1.2, 1.5, 1.8, 2.1],
        color: "4472C4",
        showDataLabels: true
      },
      {
        name: "South",
        categories: QUARTERS,
        values: [0.9, 1.1, 1.3, 1.5],
        color: "ED7D31"
      },
      {
        name: "East",
        categories: QUARTERS,
        values: [0.7, 0.8, 0.95, 1.1],
        color: "70AD47"
      }
    ]
  })
);

// ---------------------------------------------------------------------------
// 2. Stacked column
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Stacked column", 2);
Document.addContent(
  doc,
  chart({
    type: "columnStacked",
    title: "Stacked revenue",
    legend: "r",
    width: cmToEmu(15),
    height: cmToEmu(8),
    series: [
      { name: "North", categories: QUARTERS, values: [1.2, 1.5, 1.8, 2.1], color: "4472C4" },
      { name: "South", categories: QUARTERS, values: [0.9, 1.1, 1.3, 1.5], color: "ED7D31" },
      { name: "East", categories: QUARTERS, values: [0.7, 0.8, 0.95, 1.1], color: "70AD47" }
    ]
  })
);

// ---------------------------------------------------------------------------
// 3. Line chart with markers + trendline
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Line chart with trendline", 2);
Document.addContent(
  doc,
  chart({
    type: "lineMarked",
    title: "Stock price (with linear trendline)",
    legend: "b",
    width: cmToEmu(15),
    height: cmToEmu(8),
    series: [
      {
        name: "AAPL",
        categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        values: [150, 155, 152, 160, 168, 172],
        color: "1F4E79",
        trendline: { type: "linear", displayEquation: true, displayRSquared: true }
      }
    ]
  })
);

// ---------------------------------------------------------------------------
// 4. Pie & doughnut
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Pie & doughnut", 2);
const pieSeries: ChartSeries = {
  name: "Market share",
  categories: ["Chrome", "Safari", "Firefox", "Edge", "Other"],
  values: [62, 19, 8, 5, 6],
  pointColors: ["4472C4", "ED7D31", "A5A5A5", "FFC000", "70AD47"],
  showDataLabels: true
};
Document.addContent(
  doc,
  chart({ type: "pie", title: "Browser share", legend: "r", series: [pieSeries] })
);
Document.addContent(
  doc,
  chart({ type: "doughnut", title: "Browser share (doughnut)", legend: "r", series: [pieSeries] })
);

// ---------------------------------------------------------------------------
// 5. Area
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Area chart", 2);
Document.addContent(
  doc,
  chart({
    type: "area",
    title: "Cumulative downloads",
    legend: "b",
    series: [
      {
        name: "v1.0",
        categories: ["Jan", "Feb", "Mar", "Apr", "May"],
        values: [200, 450, 720, 1010, 1380],
        color: "70AD47"
      }
    ]
  })
);

// ---------------------------------------------------------------------------
// 6. Scatter
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Scatter plots", 2);
Document.addContent(
  doc,
  chart({
    type: "scatter",
    title: "Price vs sqft",
    legend: "none",
    categoryAxis: { title: "Sqft" },
    valueAxis: { title: "Price ($K)" },
    series: [
      {
        name: "Listings",
        categories: ["1500", "1800", "2200", "2500", "2900", "3300"],
        values: [350, 410, 490, 540, 620, 690],
        color: "C00000"
      }
    ]
  })
);
Document.addContent(
  doc,
  chart({
    type: "scatterSmooth",
    title: "Smooth-line scatter",
    legend: "none",
    series: [
      {
        name: "trend",
        categories: ["1", "2", "3", "4", "5", "6"],
        values: [1, 4, 2.5, 3.8, 5.2, 4.6],
        color: "4472C4"
      }
    ]
  })
);

// ---------------------------------------------------------------------------
// 7. Bar with data labels
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Bar chart", 2);
Document.addContent(
  doc,
  chart({
    type: "bar",
    title: "Defects by component",
    legend: "none",
    series: [
      {
        name: "Defects",
        categories: ["Auth", "DB", "UI", "API", "Network"],
        values: [12, 5, 18, 9, 3],
        color: "ED7D31",
        showDataLabels: true
      }
    ]
  })
);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Single category
Document.addContent(
  doc,
  chart({
    type: "column",
    title: "Single category",
    series: [{ name: "x", categories: ["only"], values: [42], color: "4472C4" }]
  })
);

// Negative values
Document.addContent(
  doc,
  chart({
    type: "column",
    title: "P&L (with negatives)",
    series: [{ name: "P/L", categories: QUARTERS, values: [120, -45, 30, -10], color: "C00000" }]
  })
);

// Many categories (12 months)
Document.addContent(
  doc,
  chart({
    type: "line",
    title: "Monthly active users (12 categories)",
    series: [
      {
        name: "MAU",
        categories: [
          "Jan",
          "Feb",
          "Mar",
          "Apr",
          "May",
          "Jun",
          "Jul",
          "Aug",
          "Sep",
          "Oct",
          "Nov",
          "Dec"
        ],
        values: [120, 132, 145, 160, 178, 190, 210, 225, 240, 260, 285, 312],
        color: "70AD47"
      }
    ]
  })
);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "14-charts.docx"), buf);
console.log(`  → 14-charts.docx (${buf.length} bytes)`);
