/**
 * Sales Dashboard — a regional BI dashboard driven by pivot charts.
 *
 * Focuses on the BI side of the stack: pivot tables + pivot charts +
 * slicer-style filter metadata + a region map + a customer cohort retention
 * matrix. Designed so that opening the workbook in Excel refreshes the
 * pivot cache against the source table and the pivot charts auto-update.
 *
 * Features covered:
 *   - 10000-row transactions Table across 3 years, 8 regions, 20 cities,
 *     12 product lines, 4 channels, 6 customer segments
 *   - Pivot tables with 2-level rows, multi-column fields, page filters
 *   - Pivot charts (bar, line, pie) wired to pivot tables with field
 *     buttons + drop-zone metadata + refreshOnOpen
 *   - Pivot chartsheet for a full-page regional view
 *   - Classic region map (ChartEx) — country revenue heat map
 *   - Funnel chart for conversion pipeline
 *   - Treemap for product hierarchy
 *   - Combined combo chart (revenue bars + growth line)
 *   - Cohort retention matrix with 2-colour-scale conditional formatting
 *     (classic BI "retention triangle")
 *   - Rich sparkline group per region (line + column + win-loss)
 *   - Customer-segment KPI cards with icon-set ratings
 *   - Cross-sheet hyperlinks inside a navigation ribbon
 *   - Print-ready page setup + scaled fit + repeat print titles
 *
 * Output:
 *   tmp/sales-dashboard.xlsx
 *   tmp/sales-dashboard.pdf
 *
 * Usage:
 *   npx tsx src/modules/excel/examples/sales-dashboard.ts
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import { type ChartRichText } from "@excel/chart/index";
import { excelToPdf } from "@pdf/excel-bridge";

import { Workbook } from "../../../index";

const OUT_DIR = resolve(process.cwd(), "tmp");
mkdirSync(OUT_DIR, { recursive: true });

const XLSX_PATH = resolve(OUT_DIR, "sales-dashboard.xlsx");
const PDF_PATH = resolve(OUT_DIR, "sales-dashboard.pdf");

// Reproducible PRNG.
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
const rng = mulberry32(0xd15ea5e);

const REGIONS = [
  "Americas-North",
  "Americas-South",
  "EMEA-North",
  "EMEA-South",
  "APAC-East",
  "APAC-SouthEast",
  "APAC-South",
  "Middle-East"
];
const COUNTRIES = [
  "United States",
  "Canada",
  "Brazil",
  "Argentina",
  "United Kingdom",
  "Germany",
  "France",
  "Italy",
  "Spain",
  "Japan",
  "China",
  "South Korea",
  "Singapore",
  "Thailand",
  "India",
  "Australia",
  "Saudi Arabia",
  "United Arab Emirates",
  "Mexico",
  "Netherlands"
];
const PRODUCT_LINES = [
  { line: "Platform", family: "Software" },
  { line: "Analytics", family: "Software" },
  { line: "Security", family: "Software" },
  { line: "Collaboration", family: "Software" },
  { line: "Hardware-X1", family: "Hardware" },
  { line: "Hardware-X2", family: "Hardware" },
  { line: "Hardware-Pro", family: "Hardware" },
  { line: "Consulting-Basic", family: "Services" },
  { line: "Consulting-Pro", family: "Services" },
  { line: "Training", family: "Services" },
  { line: "Support-Std", family: "Services" },
  { line: "Support-Enterprise", family: "Services" }
];
const CHANNELS = ["Direct", "Online", "Partner", "Retail"];
const SEGMENTS = ["SMB", "Mid-Market", "Enterprise", "Strategic", "Government", "Non-Profit"];
const FUNNEL_STAGES = ["Leads", "MQL", "SQL", "Opportunity", "Proposal", "Closed Won"];

interface Txn {
  Date: Date;
  Year: number;
  Quarter: string;
  Month: string;
  Region: string;
  Country: string;
  ProductLine: string;
  ProductFamily: string;
  Channel: string;
  Segment: string;
  Units: number;
  Revenue: number;
  Discount: number;
  Cost: number;
  Profit: number;
  Margin: number;
}

function generateTransactions(): Txn[] {
  const rows: Txn[] = [];
  for (const year of [2023, 2024, 2025]) {
    for (let month = 0; month < 12; month++) {
      const quarter = `Q${Math.floor(month / 3) + 1}`;
      for (let i = 0; i < 280; i++) {
        const day = 1 + Math.floor(rng() * 27);
        const region = REGIONS[Math.floor(rng() * REGIONS.length)];
        const country = COUNTRIES[Math.floor(rng() * COUNTRIES.length)];
        const product = PRODUCT_LINES[Math.floor(rng() * PRODUCT_LINES.length)];
        const channel = CHANNELS[Math.floor(rng() * CHANNELS.length)];
        const segment = SEGMENTS[Math.floor(rng() * SEGMENTS.length)];
        const units = Math.floor(1 + rng() * 40);
        const unitPrice =
          product.family === "Software"
            ? 800 + rng() * 4000
            : product.family === "Hardware"
              ? 400 + rng() * 1500
              : 200 + rng() * 1500;
        const gross = Math.round(units * unitPrice * 100) / 100;
        const discount = Math.round(gross * rng() * 0.2 * 100) / 100;
        const revenue = Math.round((gross - discount) * 100) / 100;
        const cost = Math.round(revenue * (0.35 + rng() * 0.3) * 100) / 100;
        const profit = Math.round((revenue - cost) * 100) / 100;
        const margin = Math.round((profit / revenue) * 10000) / 10000;
        rows.push({
          Date: new Date(Date.UTC(year, month, day)),
          Year: year,
          Quarter: quarter,
          Month: [
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
          ][month],
          Region: region,
          Country: country,
          ProductLine: product.line,
          ProductFamily: product.family,
          Channel: channel,
          Segment: segment,
          Units: units,
          Revenue: revenue,
          Discount: discount,
          Cost: cost,
          Profit: profit,
          Margin: margin
        });
      }
    }
  }
  return rows;
}

async function main(): Promise<void> {
  const wb = new Workbook();
  wb.title = "Sales Dashboard";
  wb.subject = "Regional BI dashboard";
  wb.creator = "ExcelTS sales-dashboard example";
  wb.keywords = "sales, pivot, regions, channel, treemap, regionMap";

  const txns = generateTransactions();
  console.log(`Generated ${txns.length} transactions`);

  // =========================================================================
  // Sheet 1 — Transactions (10 000-row Table)
  // =========================================================================

  const txSheet = wb.addWorksheet("Transactions", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 1, showGridLines: true }],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      printTitlesRow: "1:1"
    },
    headerFooter: {
      oddHeader: '&L&"Calibri,Bold"Transactions&R&D',
      oddFooter: "&LExcelTS&CPage &P / &N&R&F"
    }
  });

  const txnTable = txSheet.addTable({
    name: "Transactions",
    displayName: "Transactions",
    ref: "A1",
    headerRow: true,
    totalsRow: true,
    style: { theme: "TableStyleMedium6", showRowStripes: true },
    columns: [
      { name: "Date", totalsRowLabel: "Totals" },
      { name: "Year" },
      { name: "Quarter" },
      { name: "Month" },
      { name: "Region" },
      { name: "Country" },
      { name: "ProductLine" },
      { name: "ProductFamily" },
      { name: "Channel" },
      { name: "Segment" },
      { name: "Units", totalsRowFunction: "sum" },
      { name: "Revenue", totalsRowFunction: "sum" },
      { name: "Discount", totalsRowFunction: "sum" },
      { name: "Cost", totalsRowFunction: "sum" },
      { name: "Profit", totalsRowFunction: "sum" },
      { name: "Margin", totalsRowFunction: "average" }
    ],
    rows: txns.map(t => [
      t.Date,
      t.Year,
      t.Quarter,
      t.Month,
      t.Region,
      t.Country,
      t.ProductLine,
      t.ProductFamily,
      t.Channel,
      t.Segment,
      t.Units,
      t.Revenue,
      t.Discount,
      t.Cost,
      t.Profit,
      t.Margin
    ])
  });

  txSheet.getColumn(1).numFmt = "yyyy-mm-dd";
  [12, 13, 14, 15].forEach(c => (txSheet.getColumn(c).numFmt = "$#,##0.00"));
  txSheet.getColumn(16).numFmt = "0.0%";
  [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16].forEach(c => {
    txSheet.getColumn(c).width = c === 1 ? 12 : c === 5 || c === 6 ? 18 : 14;
  });

  // Data bar on revenue
  txSheet.addConditionalFormatting({
    ref: `L2:L${txns.length + 1}`,
    rules: [
      {
        type: "dataBar",
        priority: 1,
        cfvo: [{ type: "min" }, { type: "max" }],
        color: { argb: "FF70AD47" }
      }
    ]
  });

  console.log(`Transactions table: ${txnTable.model.name}`);

  // =========================================================================
  // Sheet 2 — Pivot Core (pivot tables + pivot charts)
  // =========================================================================

  const pivotSheet = wb.addWorksheet("Pivot Core", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 3, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1 },
    properties: { tabColor: { argb: "FF2F5496" } }
  });

  pivotSheet.mergeCells("A1:P1");
  pivotSheet.getCell("A1").value =
    "Pivot core — revenue by region / product family / channel / year";
  pivotSheet.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF2F5496" } };

  // Pivot 1: Region × Year revenue
  //
  // `ref` anchors the pivot body so the three pivots on this sheet do not
  // overlap when Excel refreshes the cache. Without distinct anchors every
  // pivot defaults to A3 and Excel reports "there's already a PivotTable
  // there" on open.
  const regionYearPivot = pivotSheet.addPivotTable({
    sourceTable: txnTable,
    rows: ["Region"],
    columns: ["Year"],
    values: ["Revenue"],
    metric: "sum",
    ref: "A3"
  });

  // Pivot chart on that pivot — bar + data labels
  pivotSheet.addPivotChart(
    regionYearPivot,
    {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: "Revenue by region × year",
      series: [
        {
          name: "Revenue",
          categories: "Transactions[Region]",
          values: "Transactions[Revenue]",
          dataLabels: { showVal: true, numFmt: "$#,##0", position: "outEnd" }
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
    "H3:P23"
  );

  // Pivot 2: Product family × Channel — anchored far enough below Pivot 1
  // that its expanded size (~20 rows × 6 cols) does not collide.
  const prodChannelPivot = pivotSheet.addPivotTable({
    sourceTable: txnTable,
    rows: ["ProductFamily", "ProductLine"],
    columns: ["Channel"],
    values: ["Revenue"],
    metric: "sum",
    ref: "A20"
  });

  pivotSheet.addPivotChart(
    prodChannelPivot,
    {
      type: "pie",
      varyColors: true,
      title: "Revenue share by product family",
      series: [
        {
          name: "Revenue",
          categories: "Transactions[ProductFamily]",
          values: "Transactions[Revenue]",
          dataLabels: {
            showPercent: true,
            showCatName: true,
            position: "outEnd",
            separator: " • "
          }
        }
      ],
      pivotChartOptions: { refreshOnOpen: true, showExpandCollapseFieldButtons: true }
    },
    "H25:P45"
  );

  // Pivot 3: Segment × Region with page filter on Year. Anchored below Pivot 2;
  // the `Year` page filter lives at row 45, a blank separator follows, and
  // the pivot body begins at row 47.
  const segmentRegionPivot = pivotSheet.addPivotTable({
    sourceTable: txnTable,
    rows: ["Segment"],
    columns: ["Region"],
    values: ["Revenue", "Profit"],
    pages: ["Year"],
    metric: "sum",
    ref: "A45"
  });

  pivotSheet.addPivotChart(
    segmentRegionPivot,
    {
      type: "bar",
      barDir: "bar",
      grouping: "stacked",
      title: "Revenue by segment × region (stacked)",
      series: [
        {
          name: "Revenue",
          categories: "Transactions[Segment]",
          values: "Transactions[Revenue]"
        },
        {
          name: "Profit",
          categories: "Transactions[Segment]",
          values: "Transactions[Profit]"
        }
      ],
      pivotChartOptions: {
        dropZonesVisible: true,
        dropZoneFilter: true,
        dropZoneCategories: true,
        dropZoneSeries: true,
        dropZoneData: true,
        refreshOnOpen: true
      }
    },
    "H47:P67"
  );

  // Combo pivot chart — revenue bars + YoY growth line (secondary axis)
  pivotSheet.addPivotComboChart(
    regionYearPivot,
    {
      title: "Revenue (bars) + growth (line, secondary)",
      groups: [
        {
          type: "bar",
          barDir: "col",
          series: [
            {
              name: "Revenue",
              categories: "Transactions[Region]",
              values: "Transactions[Revenue]",
              fill: "4472C4"
            }
          ]
        },
        {
          type: "line",
          useSecondaryAxis: true,
          series: [
            {
              name: "Profit",
              categories: "Transactions[Region]",
              values: "Transactions[Profit]",
              line: "ED7D31",
              lineWidth: 2.5,
              marker: { symbol: "diamond", size: 8, fill: "ED7D31", border: "FFFFFF" }
            }
          ]
        }
      ]
    },
    "H69:P89"
  );

  // =========================================================================
  // Sheet 3 — Funnel + Treemap + Sunburst
  // =========================================================================

  const funnelSheet = wb.addWorksheet("Pipeline", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FFED7D31" } }
  });

  funnelSheet.mergeCells("A1:L1");
  funnelSheet.getCell("A1").value = "Sales pipeline — funnel, product hierarchy, conversion";
  funnelSheet.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFED7D31" } };

  // Synthetic funnel — realistic waterfall drop-offs
  const funnelValues = [10000, 5400, 2900, 1800, 1200, 900];
  const funnelHidden = wb.addWorksheet("Pipeline-Data", { state: "hidden" });
  funnelHidden.getRow(1).values = ["Stage", "Count"];
  FUNNEL_STAGES.forEach((stage, i) => {
    funnelHidden.addRow([stage, funnelValues[i]]);
  });

  funnelSheet.getCell("A3").value = "Conversion funnel — FY25";
  funnelSheet.getCell("A3").font = { bold: true };
  funnelSheet.addFunnelChart(
    {
      title: "FY25 sales conversion funnel",
      categories: `'Pipeline-Data'!$A$2:$A$${1 + FUNNEL_STAGES.length}`,
      series: [
        {
          name: "Count",
          values: `'Pipeline-Data'!$B$2:$B$${1 + FUNNEL_STAGES.length}`
        }
      ]
    },
    "A4:F26"
  );

  // Treemap — product line nested under family
  const treemapHidden = wb.addWorksheet("Treemap-Data", { state: "hidden" });
  treemapHidden.getRow(1).values = ["Family", "Line", "Revenue"];
  const treemapAgg = new Map<string, number>();
  for (const t of txns) {
    const k = `${t.ProductFamily}|${t.ProductLine}`;
    treemapAgg.set(k, (treemapAgg.get(k) ?? 0) + t.Revenue);
  }
  const treemapEntries = Array.from(treemapAgg.entries()).map(([k, v]) => {
    const [family, line] = k.split("|");
    return [family, line, Math.round(v * 100) / 100] as const;
  });
  treemapEntries.forEach(row => treemapHidden.addRow([row[0], row[1], row[2]]));

  funnelSheet.getCell("G3").value = "Product hierarchy (treemap)";
  funnelSheet.getCell("G3").font = { bold: true };
  funnelSheet.addTreemapChart(
    {
      title: "Revenue by product family → line",
      categories: `'Treemap-Data'!$B$2:$B$${1 + treemapEntries.length}`,
      series: [
        {
          name: "Revenue",
          values: `'Treemap-Data'!$C$2:$C$${1 + treemapEntries.length}`,
          hierarchy: [`'Treemap-Data'!$A$2:$A$${1 + treemapEntries.length}`]
        }
      ],
      layout: { parentLabelLayout: "banner" }
    },
    "G4:L26"
  );

  // Sunburst — region → country → revenue
  const sunburstHidden = wb.addWorksheet("Sunburst-Data", { state: "hidden" });
  sunburstHidden.getRow(1).values = ["Region", "Country", "Revenue"];
  const sunburstAgg = new Map<string, number>();
  for (const t of txns) {
    const k = `${t.Region}|${t.Country}`;
    sunburstAgg.set(k, (sunburstAgg.get(k) ?? 0) + t.Revenue);
  }
  const sunburstEntries = Array.from(sunburstAgg.entries()).map(([k, v]) => {
    const [region, country] = k.split("|");
    return [region, country, Math.round(v * 100) / 100] as const;
  });
  sunburstEntries.forEach(row => sunburstHidden.addRow([row[0], row[1], row[2]]));

  funnelSheet.getCell("A28").value = "Region → country sunburst";
  funnelSheet.getCell("A28").font = { bold: true };
  funnelSheet.addSunburstChart(
    {
      title: "Revenue by region → country",
      categories: `'Sunburst-Data'!$B$2:$B$${1 + sunburstEntries.length}`,
      series: [
        {
          name: "Revenue",
          values: `'Sunburst-Data'!$C$2:$C$${1 + sunburstEntries.length}`,
          hierarchy: [`'Sunburst-Data'!$A$2:$A$${1 + sunburstEntries.length}`]
        }
      ]
    },
    "A29:F52"
  );

  // Region map — revenue per country
  const mapHidden = wb.addWorksheet("Map-Data", { state: "hidden" });
  mapHidden.getRow(1).values = ["Country", "Revenue"];
  const byCountry = new Map<string, number>();
  for (const t of txns) {
    byCountry.set(t.Country, (byCountry.get(t.Country) ?? 0) + t.Revenue);
  }
  const mapEntries = Array.from(byCountry.entries());
  mapEntries.forEach(([country, rev]) => mapHidden.addRow([country, Math.round(rev * 100) / 100]));

  funnelSheet.getCell("G28").value = "Revenue by country (region map)";
  funnelSheet.getCell("G28").font = { bold: true };
  funnelSheet.addRegionMapChart(
    {
      title: "Revenue by country",
      categories: `'Map-Data'!$A$2:$A$${1 + mapEntries.length}`,
      series: [
        {
          name: "Revenue",
          values: `'Map-Data'!$B$2:$B$${1 + mapEntries.length}`
        }
      ],
      layout: { projection: "robinson", regionLabels: "bestFit", geoMappingLevel: "country" }
    },
    "G29:L52"
  );

  // =========================================================================
  // Sheet 4 — Cohort retention matrix
  // =========================================================================

  const cohort = wb.addWorksheet("Cohort Retention", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FF70AD47" } }
  });
  cohort.getColumn(1).width = 18;
  for (let c = 2; c <= 13; c++) {
    cohort.getColumn(c).width = 8;
  }

  cohort.mergeCells("A1:M1");
  cohort.getCell("A1").value = "Customer cohort retention (%) — synthetic";
  cohort.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF70AD47" } };

  cohort.getRow(2).values = ["Cohort", ...Array.from({ length: 12 }, (_, i) => `M+${i}`)];
  cohort.getRow(2).font = { bold: true };
  cohort.getRow(2).alignment = { horizontal: "center" };
  cohort.getRow(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } };

  // Generate a classic retention "triangle" — each cohort decays
  // exponentially + some jitter.
  const cohorts = ["2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4", "2025-Q1", "2025-Q2", "2025-Q3"];
  cohorts.forEach((c, ci) => {
    const row = 3 + ci;
    cohort.getCell(row, 1).value = c;
    cohort.getCell(row, 1).font = { bold: true };
    for (let m = 0; m <= 11; m++) {
      // Cells beyond what has "happened yet" stay blank.
      if (m + ci * 3 > 20) {
        continue;
      }
      // Base retention: 100%, 85%, 72%, 65%, 58%, 54%, 51%, 48%, 46%, 45%, 44%, 43%
      const decay = [1, 0.85, 0.72, 0.65, 0.58, 0.54, 0.51, 0.48, 0.46, 0.45, 0.44, 0.43];
      const value = decay[m] * (0.95 + rng() * 0.1);
      cohort.getCell(row, 2 + m).value = Math.round(value * 1000) / 1000;
      cohort.getCell(row, 2 + m).numFmt = "0.0%";
      cohort.getCell(row, 2 + m).alignment = { horizontal: "center" };
    }
  });

  // Classic BI heat map — color scale across the whole matrix.
  cohort.addConditionalFormatting({
    ref: `B3:M${2 + cohorts.length}`,
    rules: [
      {
        type: "colorScale",
        priority: 1,
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: 0.5 },
          { type: "num", value: 1 }
        ],
        color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }]
      }
    ]
  });

  // =========================================================================
  // Sheet 5 — Regional sparklines + KPIs
  // =========================================================================

  const regional = wb.addWorksheet("Regional KPIs", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FFFFC000" } }
  });
  regional.getColumn(1).width = 22;
  for (let c = 2; c <= 13; c++) {
    regional.getColumn(c).width = 10;
  }
  [14, 15, 16].forEach(c => (regional.getColumn(c).width = 22));

  regional.mergeCells("A1:P1");
  regional.getCell("A1").value = "Regional KPIs — revenue trajectory + sparklines";
  regional.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFFFC000" } };

  const monthHeaders = [
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
  ];
  regional.getRow(2).values = [
    "Region",
    ...monthHeaders,
    "Revenue trend",
    "Units trend",
    "MoM win/loss"
  ];
  regional.getRow(2).font = { bold: true };
  regional.getRow(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };

  // Aggregate 2025 revenue per region × month
  REGIONS.forEach((region, ri) => {
    const row = 3 + ri;
    regional.getCell(row, 1).value = region;
    regional.getCell(row, 1).font = { bold: true };
    for (let m = 0; m < 12; m++) {
      const v = txns
        .filter(t => t.Region === region && t.Year === 2025 && t.Month === monthHeaders[m])
        .reduce((s, t) => s + t.Revenue, 0);
      regional.getCell(row, 2 + m).value = v;
      regional.getCell(row, 2 + m).numFmt = "$#,##0";
    }
  });

  // Revenue line sparkline
  regional.addSparklineGroup({
    type: "line",
    markers: true,
    high: true,
    low: true,
    first: true,
    last: true,
    lineColor: "4472C4",
    highColor: "70AD47",
    lowColor: "C00000",
    firstColor: "5B9BD5",
    lastColor: "FFC000",
    markerColor: "595959",
    sparklines: REGIONS.map((_, ri) => ({
      dataRef: `'Regional KPIs'!B${3 + ri}:M${3 + ri}`,
      cellRef: `N${3 + ri}`
    }))
  });

  // Units column sparkline (reuse revenue proxy)
  regional.addSparklineGroup({
    type: "column",
    negative: true,
    lineColor: "ED7D31",
    sparklines: REGIONS.map((_, ri) => ({
      dataRef: `'Regional KPIs'!B${3 + ri}:M${3 + ri}`,
      cellRef: `O${3 + ri}`
    }))
  });

  // Win/loss sparkline — month-over-month sign changes
  const deltaSheet = wb.addWorksheet("Regional-Deltas", { state: "hidden" });
  REGIONS.forEach((region, ri) => {
    const monthly: number[] = [];
    for (let m = 0; m < 12; m++) {
      monthly.push(
        txns
          .filter(t => t.Region === region && t.Year === 2025 && t.Month === monthHeaders[m])
          .reduce((s, t) => s + t.Revenue, 0)
      );
    }
    const signs = monthly.slice(1).map((v, i) => Math.sign(v - monthly[i]));
    signs.forEach((sign, i) => {
      deltaSheet.getCell(ri + 1, i + 1).value = sign;
    });
  });
  regional.addSparklineGroup({
    type: "stacked",
    negative: true,
    lineColor: "70AD47",
    negativeColor: "C00000",
    axisColor: "A6A6A6",
    displayXAxis: true,
    sparklines: REGIONS.map((_, ri) => ({
      dataRef: `'Regional-Deltas'!A${1 + ri}:K${1 + ri}`,
      cellRef: `P${3 + ri}`
    }))
  });

  // KPI cards row — show each segment's revenue + rating
  const kpiRow = 3 + REGIONS.length + 3;
  regional.mergeCells(kpiRow - 1, 1, kpiRow - 1, 12);
  regional.getCell(kpiRow - 1, 1).value = "Segment KPI cards (with icon ratings)";
  regional.getCell(kpiRow - 1, 1).font = { size: 12, bold: true, color: { argb: "FF2F5496" } };

  regional.getRow(kpiRow).values = ["Segment", "Revenue", "YoY vs FY24", "Rating"];
  regional.getRow(kpiRow).font = { bold: true };
  regional.getRow(kpiRow).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9E1F2" }
  };

  const ratingValues: number[] = [];
  SEGMENTS.forEach((segment, si) => {
    const row = kpiRow + 1 + si;
    const fy25 = txns
      .filter(t => t.Segment === segment && t.Year === 2025)
      .reduce((s, t) => s + t.Revenue, 0);
    const fy24 = txns
      .filter(t => t.Segment === segment && t.Year === 2024)
      .reduce((s, t) => s + t.Revenue, 0);
    const yoy = fy24 ? (fy25 - fy24) / fy24 : 0;
    regional.getCell(row, 1).value = segment;
    regional.getCell(row, 2).value = fy25;
    regional.getCell(row, 2).numFmt = "$#,##0";
    regional.getCell(row, 3).value = yoy;
    regional.getCell(row, 3).numFmt = "+0.0%;-0.0%";
    // Rating = normalised YoY
    regional.getCell(row, 4).value = yoy;
    regional.getCell(row, 4).numFmt = "";
    ratingValues.push(yoy);
  });
  regional.addConditionalFormatting({
    ref: `D${kpiRow + 1}:D${kpiRow + SEGMENTS.length}`,
    rules: [
      {
        type: "iconSet",
        priority: 1,
        iconSet: "5Rating",
        showValue: false,
        cfvo: [
          { type: "num", value: -0.1 },
          { type: "num", value: 0 },
          { type: "num", value: 0.05 },
          { type: "num", value: 0.1 },
          { type: "num", value: 0.2 }
        ]
      }
    ]
  });

  // =========================================================================
  // Sheet 6 — Navigation hub
  // =========================================================================

  const nav = wb.addWorksheet("Start", {
    views: [{ state: "normal", showGridLines: false, showRowColHeaders: false }]
  });
  nav.getColumn(1).width = 4;
  nav.getColumn(2).width = 36;
  nav.getColumn(3).width = 64;

  nav.mergeCells("B2:C2");
  nav.getCell("B2").value = {
    richText: [
      { text: "Sales ", font: { size: 28, bold: true, color: { argb: "FF2F5496" } } },
      { text: "BI Dashboard", font: { size: 28, bold: true, color: { argb: "FFED7D31" } } }
    ]
  };

  nav.mergeCells("B3:C3");
  nav.getCell("B3").value =
    `${txns.length.toLocaleString()} transactions · ${REGIONS.length} regions · ${COUNTRIES.length} countries · ${PRODUCT_LINES.length} product lines`;
  nav.getCell("B3").font = { italic: true, color: { argb: "FF7F7F7F" } };

  const links = [
    {
      sheet: "Transactions",
      label: "🧾 Raw transactions",
      desc: "10 000-row Table with autoFilter + data bars"
    },
    {
      sheet: "Pivot Core",
      label: "🔀 Pivot core",
      desc: "3 pivots + 4 pivot charts with drop-zone metadata"
    },
    {
      sheet: "Pipeline",
      label: "🚀 Pipeline & hierarchy",
      desc: "Funnel, treemap, sunburst, region map"
    },
    {
      sheet: "Cohort Retention",
      label: "♻️ Cohort retention",
      desc: "Classic retention triangle with heat-map colour scale"
    },
    {
      sheet: "Regional KPIs",
      label: "📈 Regional KPIs",
      desc: "Per-region sparklines + segment KPI cards"
    },
    { sheet: "Executive Chart", label: "🎯 Executive chart", desc: "Full-page pivot chartsheet" }
  ];
  links.forEach((link, i) => {
    const row = 5 + i;
    const cell = nav.getCell(row, 2);
    cell.value = { text: link.label, hyperlink: `#'${link.sheet}'!A1` };
    cell.font = { size: 13, color: { argb: "FF0563C1" }, underline: true };
    nav.getCell(row, 3).value = link.desc;
    nav.getCell(row, 3).font = { color: { argb: "FF595959" } };
  });

  wb.views = [
    {
      x: 0,
      y: 0,
      width: 30000,
      height: 18000,
      firstSheet: 0,
      activeTab: wb.worksheets.findIndex(ws => ws.name === "Start"),
      visibility: "visible"
    }
  ];

  // =========================================================================
  // Pivot chartsheet — full-page executive chart
  // =========================================================================

  const executiveTitle: ChartRichText = {
    paragraphs: [
      {
        runs: [
          {
            text: "FY25 ",
            properties: { bold: true, size: 2800, color: { srgb: "2F5496" } }
          },
          {
            text: "Revenue by region",
            properties: { size: 2400, color: { srgb: "1F3864" } }
          }
        ]
      }
    ]
  };

  wb.addPivotChartsheet("Executive Chart", regionYearPivot, {
    zoomToFit: true,
    pageMargins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    pageSetup: { orientation: "landscape", paperSize: 9, horizontalDpi: 300, verticalDpi: 300 },
    chart: {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: executiveTitle,
      series: [
        {
          name: "Revenue",
          categories: "Transactions[Region]",
          values: "Transactions[Revenue]",
          fill: "2F5496",
          dataLabels: { showVal: true, numFmt: "$#,##0", position: "outEnd" }
        }
      ],
      legendPosition: "b",
      dataTable: { showHorzBorder: true, showVertBorder: true, showKeys: true },
      pivotChartOptions: {
        dropZonesVisible: true,
        dropZoneFilter: true,
        dropZoneCategories: true,
        dropZoneData: true,
        dropZoneSeries: true,
        refreshOnOpen: true,
        showExpandCollapseFieldButtons: true
      }
    }
  });

  // =========================================================================
  // Write outputs
  // =========================================================================

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`XLSX → ${XLSX_PATH}`);

  const pdf = await excelToPdf(wb, {
    title: "Sales BI Dashboard — FY23-25",
    author: "ExcelTS",
    showGridLines: false,
    showPageNumbers: true
  });
  writeFileSync(PDF_PATH, pdf);
  console.log(`PDF  → ${PDF_PATH}`);

  console.log("");
  console.log("Workbook summary:");
  console.log(`  sheets           : ${wb.worksheets.length}`);
  console.log(`  chartsheets      : ${wb.chartsheets.length}`);
  console.log(`  transactions     : ${txns.length}`);
  console.log(`  pivot tables     : ${pivotSheet.pivotTables?.length ?? 3} (Pivot Core sheet)`);
  console.log(
    `  charts           : ${wb.worksheets.reduce((n, ws) => n + ws.getCharts().length, 0)}`
  );
  console.log(
    `  sparkline groups : ${wb.worksheets.reduce((n, ws) => n + ws.getSparklineGroups().length, 0)}`
  );
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
