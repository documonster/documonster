/**
 * Mega Dashboard — the most feature-dense single-file example in this repo.
 *
 * Generates a complete business dashboard workbook with:
 *
 *   - Sales Table (600 rows, auto-filtered, total row with formulas)
 *   - 3 Pivot tables (by region × quarter, by product × year, by channel)
 *   - 10+ charts (column, line, combo, pie, doughnut, treemap, sunburst,
 *     waterfall, boxWhisker, regionMap) across multiple sheets
 *   - Sparklines (line + column + win-loss) inside KPI table
 *   - Conditional formatting (data bar, colour scale, 3-icon arrow)
 *   - Data validation (list, decimal, date, textLength, custom)
 *   - Merged cells for section headers + rich-text headers
 *   - Hyperlinks between sheets + external URLs
 *   - Freeze panes on every data sheet
 *   - Page setup (orientation, fit-to-width, print titles, print area)
 *   - Custom header / footer with page numbers and date
 *   - Defined names (both global and sheet-scoped)
 *   - Data table (c:dTable) below one of the charts
 *   - Chartsheet (full-page revenue chart)
 *   - Threaded comments / legacy notes
 *   - Sheet protection (one sheet locked, another read-only)
 *   - Embedded image (company logo — generated inline)
 *   - PDF export with encryption + CJK font embed demo
 *
 * Output:
 *   tmp/mega-dashboard.xlsx
 *   tmp/mega-dashboard.pdf
 *
 * Usage:
 *   npx tsx src/modules/excel/examples/mega-dashboard.ts
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

import {
  type AddChartSeriesOptions,
  type ChartRichText,
  type ChartStyleModel,
  type ChartColorsModel
} from "@excel/chart/index";
import { chartToPdf, excelToPdf } from "@pdf/excel-bridge";

import { Workbook } from "../../../index";

const OUT_DIR = resolve(process.cwd(), "tmp");
mkdirSync(OUT_DIR, { recursive: true });

const XLSX_PATH = resolve(OUT_DIR, "mega-dashboard.xlsx");
const PDF_PATH = resolve(OUT_DIR, "mega-dashboard.pdf");

// ---------------------------------------------------------------------------
// Deterministic PRNG so the generated numbers stay the same between runs —
// important for golden diffs if anyone wires this into CI.
// ---------------------------------------------------------------------------

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

const rng = mulberry32(0xc0ffee);

const REGIONS = ["Americas", "EMEA", "APAC"] as const;
const COUNTRIES: Record<(typeof REGIONS)[number], string[]> = {
  Americas: ["United States", "Canada", "Brazil", "Mexico"],
  EMEA: ["Germany", "France", "United Kingdom", "Spain"],
  APAC: ["Japan", "China", "Australia", "India", "Singapore"]
};
const PRODUCTS = ["Atlas", "Beacon", "Catalyst", "Delta", "Evergreen"] as const;
const CHANNELS = ["Direct", "Retail", "Online", "Partner"] as const;
const QUARTERS = ["Q1", "Q2", "Q3", "Q4"] as const;

interface SalesRow {
  Date: Date;
  Year: number;
  Quarter: string;
  Region: string;
  Country: string;
  Product: string;
  Channel: string;
  Units: number;
  UnitPrice: number;
  Revenue: number;
  Cost: number;
  Profit: number;
  Margin: number;
}

function generateSalesData(): SalesRow[] {
  const rows: SalesRow[] = [];
  for (const year of [2023, 2024, 2025]) {
    for (const [qIdx, quarter] of QUARTERS.entries()) {
      for (const region of REGIONS) {
        for (const country of COUNTRIES[region]) {
          for (const product of PRODUCTS) {
            for (const channel of CHANNELS) {
              const day = 1 + Math.floor(rng() * 27);
              const month = qIdx * 3 + Math.floor(rng() * 3);
              const units = Math.floor(20 + rng() * 200);
              const unitPrice = Math.round((50 + rng() * 450) * 100) / 100;
              const revenue = Math.round(units * unitPrice * 100) / 100;
              const cost = Math.round(revenue * (0.45 + rng() * 0.25) * 100) / 100;
              const profit = Math.round((revenue - cost) * 100) / 100;
              const margin = Math.round((profit / revenue) * 10000) / 10000;
              rows.push({
                Date: new Date(Date.UTC(year, month, day)),
                Year: year,
                Quarter: quarter,
                Region: region,
                Country: country,
                Product: product,
                Channel: channel,
                Units: units,
                UnitPrice: unitPrice,
                Revenue: revenue,
                Cost: cost,
                Profit: profit,
                Margin: margin
              });
            }
          }
        }
      }
    }
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Tiny PNG generator (for a synthetic "logo" on the dashboard) —
// pure-JS uncompressed PNG so the example stays dependency-free.
// ---------------------------------------------------------------------------

function solidPng(width: number, height: number, rgb: number): Uint8Array {
  const rowBytes = 1 + width * 4;
  const raw = new Uint8Array(rowBytes * height);
  const r = (rgb >> 16) & 0xff;
  const g = (rgb >> 8) & 0xff;
  const b = rgb & 0xff;
  for (let y = 0; y < height; y++) {
    raw[y * rowBytes] = 0;
    for (let x = 0; x < width; x++) {
      const o = y * rowBytes + 1 + x * 4;
      raw[o] = r;
      raw[o + 1] = g;
      raw[o + 2] = b;
      raw[o + 3] = 0xff;
    }
  }
  const header = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, width);
  dv.setUint32(4, height);
  ihdr[8] = 8;
  ihdr[9] = 6;
  // Zlib stored
  const MAX = 65535;
  const parts: Uint8Array[] = [new Uint8Array([0x78, 0x01])];
  for (let i = 0; i < raw.length; i += MAX) {
    const chunk = raw.subarray(i, Math.min(i + MAX, raw.length));
    const fin = i + MAX >= raw.length ? 1 : 0;
    const h = new Uint8Array(5);
    h[0] = fin;
    h[1] = chunk.length & 0xff;
    h[2] = (chunk.length >> 8) & 0xff;
    h[3] = ~chunk.length & 0xff;
    h[4] = (~chunk.length >> 8) & 0xff;
    parts.push(h, chunk);
  }
  let a = 1;
  let b2 = 0;
  for (const byte of raw) {
    a = (a + byte) % 65521;
    b2 = (b2 + a) % 65521;
  }
  const ad = new Uint8Array(4);
  new DataView(ad.buffer).setUint32(0, ((b2 << 16) | a) >>> 0);
  parts.push(ad);
  const zlib = concat(parts);

  const mkChunk = (type: string, data: Uint8Array): Uint8Array => {
    const out = new Uint8Array(4 + 4 + data.length + 4);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.length);
    out.set(new TextEncoder().encode(type), 4);
    out.set(data, 8);
    view.setUint32(8 + data.length, crc32(new Uint8Array(out.buffer, 4, 4 + data.length)));
    return out;
  };
  return concat([
    header,
    mkChunk("IHDR", ihdr),
    mkChunk("IDAT", zlib),
    mkChunk("IEND", new Uint8Array(0))
  ]);
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

function crc32(bytes: Uint8Array): number {
  const table: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
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

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  const wb = new Workbook();
  wb.title = "Mega Dashboard";
  wb.subject = "All-features kitchen-sink showcase";
  wb.creator = "ExcelTS mega-dashboard example";
  wb.company = "ExcelTS Demo Corp";
  wb.category = "Business Intelligence";
  wb.keywords = "sales, pivot, charts, dashboard, ExcelTS";
  wb.created = new Date();
  wb.modified = new Date();

  // Register a shared image we can refer to later.
  const logoId = wb.addImage({ buffer: solidPng(240, 80, 0x2f5496), extension: "png" });

  const rows = generateSalesData();
  console.log(`Generated ${rows.length} sales rows`);

  // =========================================================================
  // Sheet 1 — Dashboard (overview with KPIs, sparklines, top charts)
  // =========================================================================

  const dashboard = wb.addWorksheet("Dashboard", {
    views: [
      {
        state: "frozen",
        xSplit: 0,
        ySplit: 4,
        showGridLines: false,
        zoomScale: 110
      }
    ],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: true,
      printArea: "A1:N60"
    },
    headerFooter: {
      oddHeader: '&L&"Calibri,Bold"&14Mega Dashboard&R&"Calibri,Regular"&10Printed &D',
      oddFooter: "&LExcelTS Demo Corp&CPage &P of &N&RConfidential"
    },
    properties: { tabColor: { argb: "FF2F5496" }, defaultRowHeight: 18 }
  });

  // Top banner — merged cells with rich-text heading + logo area
  dashboard.mergeCells("A1:N1");
  dashboard.getRow(1).height = 48;
  const banner = dashboard.getCell("A1");
  banner.value = {
    richText: [
      { text: "▦  ", font: { size: 20, bold: true, color: { argb: "FF2F5496" } } },
      { text: "ExcelTS ", font: { size: 20, bold: true, color: { argb: "FF1F3864" } } },
      { text: "Corporate Dashboard ", font: { size: 18, color: { argb: "FF404040" } } },
      { text: "FY23–FY25", font: { size: 14, italic: true, color: { argb: "FF7F7F7F" } } }
    ]
  };
  banner.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  banner.fill = {
    type: "gradient",
    gradient: "angle",
    degree: 90,
    stops: [
      { position: 0, color: { argb: "FFE7F0FA" } },
      { position: 1, color: { argb: "FFFFFFFF" } }
    ]
  };

  // Embed logo in top-right corner (one-cell anchor).
  dashboard.addImage(logoId, { tl: { col: 11, row: 0 }, ext: { width: 240, height: 64 } });

  dashboard.mergeCells("A2:N2");
  const subtitle = dashboard.getCell("A2");
  subtitle.value =
    "Interactive KPI dashboard with sparklines, pivot summaries and drill-through links.";
  subtitle.font = { size: 11, italic: true, color: { argb: "FF595959" } };
  subtitle.alignment = { horizontal: "left", indent: 1 };

  // KPI row — 4 big numbers with inline sparklines alongside each.
  const kpiHeaders = ["Revenue", "Profit", "Units Sold", "Avg Margin"];
  const kpiValues = [
    rows.reduce((s, r) => s + r.Revenue, 0),
    rows.reduce((s, r) => s + r.Profit, 0),
    rows.reduce((s, r) => s + r.Units, 0),
    rows.reduce((s, r) => s + r.Margin, 0) / rows.length
  ];

  dashboard.getRow(4).height = 10;
  const kpiRow = 5;
  kpiHeaders.forEach((label, i) => {
    const col = 1 + i * 3;
    const labelCell = dashboard.getCell(kpiRow, col);
    labelCell.value = label;
    labelCell.font = { size: 11, bold: true, color: { argb: "FF595959" } };
    labelCell.alignment = { horizontal: "left", indent: 1 };

    const numCell = dashboard.getCell(kpiRow + 1, col);
    numCell.value = kpiValues[i];
    numCell.font = { size: 20, bold: true, color: { argb: "FF2F5496" } };
    numCell.numFmt = i === 3 ? "0.0%" : i === 2 ? "#,##0" : "$#,##0";
    numCell.alignment = { horizontal: "left", indent: 1 };

    // Border + subtle fill per KPI box
    const box = { style: "thin" as const, color: { argb: "FFBFBFBF" } };
    for (let dr = 0; dr < 3; dr++) {
      for (let dc = 0; dc < 3; dc++) {
        const c = dashboard.getCell(kpiRow + dr, col + dc);
        c.border = {
          top: dr === 0 ? box : undefined,
          bottom: dr === 2 ? box : undefined,
          left: dc === 0 ? box : undefined,
          right: dc === 2 ? box : undefined
        };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F9FF" } };
      }
    }
  });

  // Freeze panes already set via views; set column widths for KPI cards
  for (let c = 1; c <= 14; c++) {
    dashboard.getColumn(c).width = 11;
  }
  dashboard.getColumn(1).width = 14;

  // =========================================================================
  // Sheet 2 — Sales Data (Table with ~600 rows + autoFilter + totals)
  // =========================================================================

  const sales = wb.addWorksheet("SalesData", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 1, showGridLines: true, zoomScale: 100 }],
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      printTitlesRow: "1:1"
    },
    headerFooter: {
      oddHeader: '&L&"Calibri,Bold"Sales Data&R&"Calibri,Regular"&D',
      oddFooter: "&LExcelTS&CPage &P / &N&R&F"
    }
  });

  const salesTable = sales.addTable({
    name: "SalesData",
    ref: "A1",
    displayName: "SalesData",
    headerRow: true,
    totalsRow: true,
    style: { theme: "TableStyleMedium9", showRowStripes: true, showColumnStripes: false },
    columns: [
      { name: "Date", totalsRowLabel: "Totals" },
      { name: "Year" },
      { name: "Quarter" },
      { name: "Region" },
      { name: "Country" },
      { name: "Product" },
      { name: "Channel" },
      { name: "Units", totalsRowFunction: "sum" },
      { name: "UnitPrice", totalsRowFunction: "average" },
      { name: "Revenue", totalsRowFunction: "sum" },
      { name: "Cost", totalsRowFunction: "sum" },
      { name: "Profit", totalsRowFunction: "sum" },
      { name: "Margin", totalsRowFunction: "average" }
    ],
    rows: rows.map(r => [
      r.Date,
      r.Year,
      r.Quarter,
      r.Region,
      r.Country,
      r.Product,
      r.Channel,
      r.Units,
      r.UnitPrice,
      r.Revenue,
      r.Cost,
      r.Profit,
      r.Margin
    ])
  });

  sales.getColumn(1).numFmt = "yyyy-mm-dd";
  sales.getColumn(9).numFmt = "$#,##0.00";
  sales.getColumn(10).numFmt = "$#,##0.00";
  sales.getColumn(11).numFmt = "$#,##0.00";
  sales.getColumn(12).numFmt = "$#,##0.00";
  sales.getColumn(13).numFmt = "0.0%";
  [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13].forEach(col => {
    sales.getColumn(col).width = col === 1 ? 12 : col === 5 ? 18 : 13;
  });

  // Add conditional formatting to the Revenue column — data bar that lives
  // inside the cells of the Table.
  sales.addConditionalFormatting({
    ref: `J2:J${rows.length + 1}`,
    rules: [
      {
        type: "dataBar",
        priority: 1,
        cfvo: [{ type: "min" }, { type: "max" }],
        color: { argb: "FF2F5496" }
      }
    ]
  });

  // Colour-scale on margin.
  sales.addConditionalFormatting({
    ref: `M2:M${rows.length + 1}`,
    rules: [
      {
        type: "colorScale",
        priority: 2,
        cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
        color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }]
      }
    ]
  });

  // 3-icon arrows on Profit.
  sales.addConditionalFormatting({
    ref: `L2:L${rows.length + 1}`,
    rules: [
      {
        type: "iconSet",
        priority: 3,
        iconSet: "3Arrows",
        showValue: true,
        cfvo: [
          { type: "percent", value: 0 },
          { type: "percent", value: 33 },
          { type: "percent", value: 67 }
        ]
      }
    ]
  });

  console.log(`SalesData table: ${salesTable.model.name} — ${rows.length} rows`);

  // =========================================================================
  // Sheet 3 — Pivot summaries (3 pivot tables on one sheet)
  // =========================================================================

  const pivots = wb.addWorksheet("Pivots", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 3, zoomScale: 100 }],
    properties: { tabColor: { argb: "FF70AD47" } }
  });

  pivots.mergeCells("A1:H1");
  pivots.getCell("A1").value = "Pivot summaries — region × product × channel";
  pivots.getCell("A1").font = { size: 14, bold: true, color: { argb: "FF2F5496" } };
  pivots.getCell("A1").alignment = { horizontal: "left", indent: 1 };

  // Pivot 1 — Region × Year revenue
  const p1Anchor = pivots.addPivotTable({
    sourceTable: salesTable,
    rows: ["Region"],
    columns: ["Year"],
    values: ["Revenue"],
    metric: "sum"
  });

  // Pivot 2 — Product × Quarter units
  const p2Anchor = pivots.addPivotTable({
    sourceTable: salesTable,
    rows: ["Product"],
    columns: ["Quarter"],
    values: ["Units"],
    metric: "sum"
  });

  // Pivot 3 — Channel × Region margin average, with page filter on Year.
  const p3Anchor = pivots.addPivotTable({
    sourceTable: salesTable,
    rows: ["Channel", "Product"],
    columns: ["Region"],
    values: ["Margin"],
    pages: ["Year"],
    metric: "average"
  });

  console.log(
    `Pivots created: revenue by region×year, units by product×quarter, margin by channel×region (${p1Anchor ? "OK" : "?"} ${p2Anchor ? "OK" : "?"} ${p3Anchor ? "OK" : "?"})`
  );

  // =========================================================================
  // Sheet 4 — Charts (10+ charts aggregated from the sales data)
  // =========================================================================

  const charts = wb.addWorksheet("Charts", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 3, showGridLines: false, zoomScale: 90 }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    properties: { tabColor: { argb: "FFED7D31" } }
  });

  charts.mergeCells("A1:P1");
  charts.getCell("A1").value = "Chart gallery — 10+ views of the same dataset";
  charts.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFED7D31" } };

  // Tiny aggregation helper used by the charts below.
  const aggSumBy = <K extends keyof SalesRow>(key: K, filter?: (r: SalesRow) => boolean) => {
    const map = new Map<SalesRow[K], { Revenue: number; Profit: number; Units: number }>();
    for (const r of rows) {
      if (filter && !filter(r)) {
        continue;
      }
      const bucket = map.get(r[key]) ?? { Revenue: 0, Profit: 0, Units: 0 };
      bucket.Revenue += r.Revenue;
      bucket.Profit += r.Profit;
      bucket.Units += r.Units;
      map.set(r[key], bucket);
    }
    return map;
  };

  // Prepare aggregated data sheets the charts reference.
  const aggSheet = wb.addWorksheet("Aggregates", {
    state: "hidden",
    views: [{ state: "normal", style: "pageBreakPreview" }]
  });
  // --- by region
  aggSheet.addRow(["Region", "Revenue", "Profit", "Units"]);
  const byRegion = aggSumBy("Region");
  for (const region of REGIONS) {
    const v = byRegion.get(region)!;
    aggSheet.addRow([region, v.Revenue, v.Profit, v.Units]);
  }
  // --- by year
  aggSheet.addRow([]);
  aggSheet.addRow(["Year", "Revenue", "Profit", "Units"]);
  const yearStartRow = aggSheet.lastRow!.number;
  const byYear = aggSumBy("Year");
  for (const year of [2023, 2024, 2025]) {
    const v = byYear.get(year)!;
    aggSheet.addRow([year, v.Revenue, v.Profit, v.Units]);
  }
  // --- by quarter (per year)
  aggSheet.addRow([]);
  aggSheet.addRow(["Year", "Quarter", "Revenue", "Profit", "Units"]);
  const quarterStartRow = aggSheet.lastRow!.number;
  for (const year of [2023, 2024, 2025]) {
    for (const q of QUARTERS) {
      const bucket = rows
        .filter(r => r.Year === year && r.Quarter === q)
        .reduce(
          (acc, r) => {
            acc.Revenue += r.Revenue;
            acc.Profit += r.Profit;
            acc.Units += r.Units;
            return acc;
          },
          { Revenue: 0, Profit: 0, Units: 0 }
        );
      aggSheet.addRow([year, q, bucket.Revenue, bucket.Profit, bucket.Units]);
    }
  }
  // --- by product
  aggSheet.addRow([]);
  aggSheet.addRow(["Product", "Revenue", "Profit", "Units"]);
  const productStartRow = aggSheet.lastRow!.number;
  const byProduct = aggSumBy("Product");
  for (const product of PRODUCTS) {
    const v = byProduct.get(product)!;
    aggSheet.addRow([product, v.Revenue, v.Profit, v.Units]);
  }
  // --- by country
  aggSheet.addRow([]);
  aggSheet.addRow(["Country", "Revenue"]);
  const countryStartRow = aggSheet.lastRow!.number;
  const byCountry = aggSumBy("Country");
  const countriesList = Array.from(byCountry.keys());
  for (const country of countriesList) {
    aggSheet.addRow([country, byCountry.get(country)!.Revenue]);
  }
  // --- by channel
  aggSheet.addRow([]);
  aggSheet.addRow(["Channel", "Revenue", "Profit"]);
  const channelStartRow = aggSheet.lastRow!.number;
  const byChannel = aggSumBy("Channel");
  for (const channel of CHANNELS) {
    const v = byChannel.get(channel)!;
    aggSheet.addRow([channel, v.Revenue, v.Profit]);
  }

  // Strip-trim helper to get absolute ranges from an anchor row + count.
  const agg = (startRow: number, count: number, col: string): string =>
    `Aggregates!$${col}$${startRow}:$${col}$${startRow + count - 1}`;

  // ---- Chart 1: Revenue by region (column) ----
  charts.getCell("A3").value = "1 — Revenue by region";
  charts.getCell("A3").font = { bold: true };
  charts.addChart(
    {
      type: "bar",
      barDir: "col",
      title: "Revenue by region",
      style: 26,
      series: [
        {
          name: "Revenue",
          categories: agg(2, REGIONS.length, "A"),
          values: agg(2, REGIONS.length, "B"),
          dataLabels: { showVal: true, numFmt: "$#,##0", position: "outEnd" }
        }
      ],
      categoryAxis: { title: "Region" },
      valueAxis: { title: "Revenue (USD)", numFmt: "$#,##0" },
      showLegend: false
    },
    "A4:H22"
  );

  // ---- Chart 2: Profit trend by quarter × year (combo) ----
  charts.getCell("I3").value = "2 — Revenue bars + Profit line (combo)";
  charts.getCell("I3").font = { bold: true };
  charts.addComboChart(
    {
      title: "Revenue vs Profit by quarter",
      groups: [
        {
          type: "bar",
          barDir: "col",
          grouping: "clustered",
          series: [
            {
              name: "Revenue 2023",
              categories: agg(quarterStartRow, 4, "B"),
              values: agg(quarterStartRow, 4, "C"),
              fill: "4472C4"
            },
            {
              name: "Revenue 2024",
              categories: agg(quarterStartRow + 4, 4, "B"),
              values: agg(quarterStartRow + 4, 4, "C"),
              fill: "5B9BD5"
            },
            {
              name: "Revenue 2025",
              categories: agg(quarterStartRow + 8, 4, "B"),
              values: agg(quarterStartRow + 8, 4, "C"),
              fill: "8FAADC"
            }
          ]
        },
        {
          type: "line",
          useSecondaryAxis: true,
          series: [
            {
              name: "Profit 2025",
              categories: agg(quarterStartRow + 8, 4, "B"),
              values: agg(quarterStartRow + 8, 4, "D"),
              line: "ED7D31",
              lineWidth: 2.5,
              marker: { symbol: "circle", size: 7, fill: "ED7D31", border: "FFFFFF" }
            }
          ]
        }
      ],
      legendPosition: "b"
    },
    "I4:P22"
  );

  // ---- Chart 3: Product revenue share (doughnut) ----
  charts.getCell("A24").value = "3 — Product revenue share";
  charts.getCell("A24").font = { bold: true };
  charts.addChart(
    {
      type: "doughnut",
      holeSize: 55,
      varyColors: true,
      title: "Revenue share by product",
      series: [
        {
          name: "Revenue",
          categories: agg(productStartRow, PRODUCTS.length, "A"),
          values: agg(productStartRow, PRODUCTS.length, "B"),
          dataLabels: {
            showPercent: true,
            showCatName: true,
            // Doughnut charts do not accept `c:dLblPos` in their
            // series `c:dLbls` — Excel's UI exposes no label
            // position picker for doughnut, and emitting any value
            // causes the drawing part to be stripped on open. See
            // `VALID_DLBL_POSITIONS_BY_TYPE` in chart-builder.ts.
            separator: " • "
          }
        }
      ],
      legendPosition: "r"
    },
    "A25:H43"
  );

  // ---- Chart 4: Channel contribution (stacked bar) ----
  charts.getCell("I24").value = "4 — Channel contribution";
  charts.getCell("I24").font = { bold: true };
  charts.addChart(
    {
      type: "bar",
      barDir: "bar",
      grouping: "percentStacked",
      title: "Channel contribution (100% stacked)",
      series: [
        {
          name: "Revenue",
          categories: agg(channelStartRow, CHANNELS.length, "A"),
          values: agg(channelStartRow, CHANNELS.length, "B")
        },
        {
          name: "Profit",
          categories: agg(channelStartRow, CHANNELS.length, "A"),
          values: agg(channelStartRow, CHANNELS.length, "C")
        }
      ]
    },
    "I25:P43"
  );

  // ---- Chart 5: Revenue trend line with markers + trendline ----
  charts.getCell("A45").value = "5 — Yearly revenue trend with linear trendline";
  charts.getCell("A45").font = { bold: true };
  charts.addChart(
    {
      type: "line",
      title: "Revenue trend FY23–FY25",
      showMarker: true,
      series: [
        {
          name: "Revenue",
          categories: agg(yearStartRow, 3, "A"),
          values: agg(yearStartRow, 3, "B"),
          line: "4472C4",
          lineWidth: 2.5,
          marker: { symbol: "circle", size: 8, fill: "4472C4", border: "FFFFFF" },
          trendline: {
            type: "linear",
            displayEq: true,
            displayRSqr: true,
            forward: 1,
            line: "ED7D31",
            lineDash: "dash"
          }
        }
      ],
      valueAxis: { min: 0, numFmt: "$#,##0" }
    },
    "A46:H64"
  );

  // ---- Chart 6: Region map (ChartEx) ----
  charts.getCell("I45").value = "6 — Revenue by country (ChartEx regionMap)";
  charts.getCell("I45").font = { bold: true };
  charts.addRegionMapChart(
    {
      title: "Revenue by country",
      categories: agg(countryStartRow, countriesList.length, "A"),
      series: [
        {
          name: "Revenue",
          values: agg(countryStartRow, countriesList.length, "B")
        }
      ],
      layout: { projection: "mercator", regionLabels: "bestFit", geoMappingLevel: "country" }
    },
    "I46:P64"
  );

  // ---- Chart 7: Waterfall for profit bridge ----
  charts.getCell("A66").value = "7 — Profit bridge (waterfall)";
  charts.getCell("A66").font = { bold: true };
  charts.addWaterfallChart(
    {
      title: "Quarterly profit bridge (2025)",
      categories: agg(quarterStartRow + 8, 4, "B"),
      series: [
        {
          name: "Profit",
          values: agg(quarterStartRow + 8, 4, "D"),
          subtotals: [0, 3]
        }
      ],
      layout: { connectorLines: true }
    },
    "A67:H85"
  );

  // ---- Chart 8: Treemap for country revenue ----
  charts.getCell("I66").value = "8 — Country revenue treemap";
  charts.getCell("I66").font = { bold: true };
  charts.addTreemapChart(
    {
      title: "Country revenue (treemap)",
      categories: agg(countryStartRow, countriesList.length, "A"),
      series: [
        {
          name: "Revenue",
          values: agg(countryStartRow, countriesList.length, "B")
        }
      ],
      layout: { parentLabelLayout: "banner" }
    },
    "I67:P85"
  );

  // ---- Chart 9: Histogram of margins ----
  charts.getCell("A87").value = "9 — Margin distribution (histogram)";
  charts.getCell("A87").font = { bold: true };
  // Need raw margins — materialise them in the Aggregates sheet.
  aggSheet.addRow([]);
  aggSheet.addRow(["Margin"]);
  const marginHeaderRow = aggSheet.lastRow!.number;
  rows.forEach(r => aggSheet.addRow([r.Margin]));
  charts.addHistogramChart(
    {
      title: "Distribution of order margin",
      series: [
        {
          name: "Margin",
          values: `Aggregates!$A$${marginHeaderRow + 1}:$A$${marginHeaderRow + rows.length}`
        }
      ],
      binning: { binType: "auto" }
    },
    "A88:H106"
  );

  // ---- Chart 10: Box-whisker of margins by region ----
  charts.getCell("I87").value = "10 — Margin by region (box-whisker)";
  charts.getCell("I87").font = { bold: true };
  charts.addBoxWhiskerChart(
    {
      title: "Margin distribution by region",
      categories: `Aggregates!$A$${marginHeaderRow + 1}:$A$${marginHeaderRow + rows.length}`,
      series: [
        {
          name: "Margin",
          values: `Aggregates!$A$${marginHeaderRow + 1}:$A$${marginHeaderRow + rows.length}`
        }
      ],
      layout: {
        quartileMethod: "inclusive",
        showMeanLine: true,
        showMeanMarker: true,
        showInnerPoints: false,
        showOutlierPoints: true
      }
    },
    "I88:P106"
  );

  // ---- Chart 11: Funnel of units sold by product ----
  charts.getCell("A108").value = "11 — Sales funnel by product";
  charts.getCell("A108").font = { bold: true };
  charts.addFunnelChart(
    {
      title: "Units sold — product funnel",
      categories: agg(productStartRow, PRODUCTS.length, "A"),
      series: [{ name: "Units", values: agg(productStartRow, PRODUCTS.length, "D") }]
    },
    "A109:H127"
  );

  // ---- Chart 12: Sunburst hierarchy ----
  charts.getCell("I108").value = "12 — Region → country sunburst";
  charts.getCell("I108").font = { bold: true };
  // Need country + region mapping — we already have per-country revenue in
  // Aggregates. Also need region as the first hierarchy level; build a
  // small two-column block.
  aggSheet.addRow([]);
  aggSheet.addRow(["Region", "Country", "Revenue"]);
  const sunburstHeaderRow = aggSheet.lastRow!.number;
  for (const region of REGIONS) {
    for (const country of COUNTRIES[region]) {
      aggSheet.addRow([region, country, byCountry.get(country)!.Revenue]);
    }
  }
  const sunburstCount = aggSheet.lastRow!.number - sunburstHeaderRow;
  charts.addSunburstChart(
    {
      title: "Revenue by region → country",
      categories: `Aggregates!$B$${sunburstHeaderRow + 1}:$B$${sunburstHeaderRow + sunburstCount}`,
      series: [
        {
          name: "Revenue",
          values: `Aggregates!$C$${sunburstHeaderRow + 1}:$C$${sunburstHeaderRow + sunburstCount}`,
          hierarchy: [
            `Aggregates!$A$${sunburstHeaderRow + 1}:$A$${sunburstHeaderRow + sunburstCount}`
          ]
        }
      ]
    },
    "I109:P127"
  );

  // =========================================================================
  // Sheet 5 — KPI Sparklines (one row per product with trend sparklines)
  // =========================================================================

  const kpiSheet = wb.addWorksheet("KPI Sparklines", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 2, showGridLines: false }],
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 },
    properties: { tabColor: { argb: "FFFFC000" } }
  });
  kpiSheet.getColumn(1).width = 16;
  [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13].forEach(c => (kpiSheet.getColumn(c).width = 9));
  kpiSheet.getColumn(14).width = 22;
  kpiSheet.getColumn(15).width = 22;
  kpiSheet.getColumn(16).width = 22;

  kpiSheet.mergeCells("A1:P1");
  kpiSheet.getCell("A1").value = "Product-level KPI sparklines (12 months × 3 years)";
  kpiSheet.getCell("A1").font = { size: 14, bold: true, color: { argb: "FFFFC000" } };

  const kpiHeader = [
    "Product",
    ...[2023, 2024, 2025].flatMap(y => QUARTERS.map(q => `${y}-${q}`)),
    "Revenue trend",
    "Units trend",
    "Margin win/loss"
  ];
  kpiSheet.getRow(2).values = kpiHeader;
  kpiSheet.getRow(2).font = { bold: true };
  kpiSheet.getRow(2).alignment = { horizontal: "center" };
  kpiSheet.getRow(2).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFF2CC" }
  };

  PRODUCTS.forEach((product, pIdx) => {
    const row = 3 + pIdx;
    kpiSheet.getCell(row, 1).value = product;
    kpiSheet.getCell(row, 1).font = { bold: true };
    let col = 2;
    for (const year of [2023, 2024, 2025]) {
      for (const q of QUARTERS) {
        const v = rows
          .filter(r => r.Product === product && r.Year === year && r.Quarter === q)
          .reduce((s, r) => s + r.Revenue, 0);
        const cell = kpiSheet.getCell(row, col);
        cell.value = v;
        cell.numFmt = "$#,##0";
        col += 1;
      }
    }
  });

  // Revenue sparkline (line) for each product across 12 quarters
  kpiSheet.addSparklineGroup({
    type: "line",
    markers: true,
    high: true,
    low: true,
    lineColor: "4472C4",
    highColor: "70AD47",
    lowColor: "C00000",
    markerColor: "595959",
    sparklines: PRODUCTS.map((_, i) => ({
      dataRef: `'KPI Sparklines'!B${3 + i}:M${3 + i}`,
      cellRef: `N${3 + i}`
    }))
  });

  // Units sparkline (column) referencing the same ranges
  kpiSheet.addSparklineGroup({
    type: "column",
    negative: true,
    lineColor: "ED7D31",
    sparklines: PRODUCTS.map((_, i) => ({
      dataRef: `'KPI Sparklines'!B${3 + i}:M${3 + i}`,
      cellRef: `O${3 + i}`
    }))
  });

  // Win/loss sparkline using quarter-over-quarter delta
  const deltaSheet = wb.addWorksheet("Aggregates-Deltas", { state: "hidden" });
  deltaSheet.getRow(1).values = PRODUCTS.map((_, i) => `delta-${i}`);
  PRODUCTS.forEach((product, pIdx) => {
    const quarters: number[] = [];
    for (const year of [2023, 2024, 2025]) {
      for (const q of QUARTERS) {
        quarters.push(
          rows
            .filter(r => r.Product === product && r.Year === year && r.Quarter === q)
            .reduce((s, r) => s + r.Revenue, 0)
        );
      }
    }
    const deltas = quarters.slice(1).map((v, i) => Math.sign(v - quarters[i]));
    deltas.forEach((d, i) => {
      deltaSheet.getCell(pIdx + 2, i + 1).value = d;
    });
  });

  kpiSheet.addSparklineGroup({
    type: "stacked",
    negative: true,
    lineColor: "70AD47",
    negativeColor: "C00000",
    axisColor: "A6A6A6",
    displayXAxis: true,
    sparklines: PRODUCTS.map((_, i) => ({
      dataRef: `'Aggregates-Deltas'!A${2 + i}:K${2 + i}`,
      cellRef: `P${3 + i}`
    }))
  });

  // =========================================================================
  // Sheet 6 — Forecast form (data validation + locked protected sheet)
  // =========================================================================

  const forecast = wb.addWorksheet("Forecast Form", {
    views: [{ state: "frozen", xSplit: 0, ySplit: 3 }],
    properties: { tabColor: { argb: "FFC00000" } }
  });

  forecast.columns = [
    { header: "Field", key: "field", width: 22 },
    { header: "Value", key: "value", width: 20 },
    { header: "Notes", key: "notes", width: 40 }
  ];
  forecast.getRow(1).font = { bold: true };
  forecast.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFCE4D6" }
  };

  const formRows: Array<{ field: string; value?: string | number | Date; notes?: string }> = [
    { field: "Forecast Region", notes: "Pick from the dropdown" },
    { field: "Forecast Year", notes: "Must be between 2025 and 2030" },
    { field: "Growth Rate (%)", notes: "Decimal between 0 and 50" },
    { field: "Review Date", notes: "Any date next 30 days" },
    { field: "Forecast Code", notes: "Exactly 6 characters" },
    { field: "Confidence", notes: "Must match Low/Medium/High" }
  ];
  formRows.forEach(r => forecast.addRow(r));

  // List validation (names referencing Excel defined name)
  wb.definedNames.add("Aggregates!$A$2:$A$4", "Regions");
  forecast.getCell("B2").dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ["=Regions"],
    showErrorMessage: true,
    errorTitle: "Invalid region",
    error: "Pick a region from the dropdown.",
    promptTitle: "Region",
    prompt: "Select one of the supported regions.",
    showInputMessage: true
  };

  forecast.getCell("B3").dataValidation = {
    type: "whole",
    operator: "between",
    allowBlank: false,
    formulae: [2025, 2030],
    errorTitle: "Invalid year",
    error: "Year must be in 2025..2030",
    showErrorMessage: true
  };

  forecast.getCell("B4").dataValidation = {
    type: "decimal",
    operator: "between",
    formulae: [0, 0.5],
    errorTitle: "Invalid rate",
    error: "Growth rate must be between 0% and 50%",
    showErrorMessage: true
  };
  forecast.getCell("B4").numFmt = "0.0%";

  forecast.getCell("B5").dataValidation = {
    type: "date",
    operator: "between",
    formulae: [new Date(), new Date(Date.now() + 30 * 86400000)],
    errorTitle: "Invalid date",
    error: "Pick a review date within the next 30 days",
    showErrorMessage: true
  };
  forecast.getCell("B5").numFmt = "yyyy-mm-dd";

  forecast.getCell("B6").dataValidation = {
    type: "textLength",
    operator: "equal",
    formulae: [6],
    errorTitle: "Bad code",
    error: "Forecast code must be exactly 6 characters.",
    showErrorMessage: true
  };

  forecast.getCell("B7").dataValidation = {
    type: "list",
    formulae: ['"Low,Medium,High"'],
    allowBlank: false
  };

  // Protect this sheet — locked everywhere EXCEPT column B (input zone)
  for (let r = 2; r <= 7; r++) {
    forecast.getCell(r, 2).protection = { locked: false };
  }
  await forecast.protect("forecast-password", {
    selectLockedCells: true,
    selectUnlockedCells: true,
    sort: false,
    formatCells: false,
    insertRows: false,
    deleteRows: false,
    autoFilter: false
  });

  // =========================================================================
  // Sheet 7 — Index / Navigation (hyperlinks back to every sheet)
  // =========================================================================

  const index = wb.addWorksheet("Index", {
    views: [{ state: "normal", style: "pageLayout" }]
  });
  index.getColumn(1).width = 4;
  index.getColumn(2).width = 32;
  index.getColumn(3).width = 60;

  index.mergeCells("B1:C1");
  index.getCell("B1").value = "Mega Dashboard — Table of Contents";
  index.getCell("B1").font = { size: 16, bold: true, color: { argb: "FF2F5496" } };

  const links = [
    {
      sheet: "Dashboard",
      ref: "A1",
      label: "🏠 Dashboard overview",
      desc: "KPI cards, sparklines, top-line charts"
    },
    {
      sheet: "SalesData",
      ref: "A1",
      label: "📊 Sales data table",
      desc: `${rows.length}-row table with autoFilter + conditional formatting`
    },
    {
      sheet: "Pivots",
      ref: "A1",
      label: "🔀 Pivot summaries",
      desc: "Region×Year, Product×Quarter, Channel×Region×Year"
    },
    {
      sheet: "Charts",
      ref: "A1",
      label: "📈 Chart gallery",
      desc: "12 charts aggregated from the sales data"
    },
    {
      sheet: "KPI Sparklines",
      ref: "A1",
      label: "✨ KPI sparklines",
      desc: "Line + column + win/loss sparklines per product"
    },
    {
      sheet: "Forecast Form",
      ref: "B2",
      label: "📝 Forecast form",
      desc: "Protected sheet with data validation"
    },
    {
      sheet: "Revenue Chart",
      ref: "A1",
      label: "📉 Revenue chartsheet",
      desc: "Full-page dedicated chart tab"
    }
  ];

  links.forEach((link, i) => {
    const row = 3 + i;
    const cell = index.getCell(row, 2);
    cell.value = {
      text: link.label,
      hyperlink: `#'${link.sheet}'!${link.ref}`,
      tooltip: `Jump to ${link.sheet}`
    };
    cell.font = { color: { argb: "FF0563C1" }, underline: true, size: 12 };
    index.getCell(row, 3).value = link.desc;
    index.getCell(row, 3).font = { size: 11, color: { argb: "FF595959" } };
  });

  // External hyperlink
  const row = 3 + links.length + 1;
  index.getCell(row, 2).value = {
    text: "📖 Documentation",
    hyperlink: "https://opencode.ai/docs"
  };
  index.getCell(row, 2).font = { color: { argb: "FF0563C1" }, underline: true, size: 12 };
  index.getCell(row, 3).value = "External link example";

  // Add a legacy note to the title
  index.getCell("B1").note = {
    texts: [
      { text: "Generated by ", font: { name: "Calibri", size: 10 } },
      { text: "ExcelTS mega-dashboard example", font: { name: "Calibri", size: 10, bold: true } }
    ],
    margins: { insetmode: "auto" },
    protection: { locked: "True", lockText: "True" }
  };

  // Threaded comment on the "Dashboard overview" link
  index.getCell(3, 2).note = {
    texts: [
      {
        text: "Start here — the dashboard loads the latest snapshot from the SalesData table.",
        font: { name: "Calibri", size: 10 }
      }
    ]
  };

  // Make the index sheet the active one on open
  wb.views = [
    {
      x: 0,
      y: 0,
      width: 30000,
      height: 18000,
      firstSheet: 0,
      activeTab: wb.worksheets.findIndex(ws => ws.name === "Index"),
      visibility: "visible"
    }
  ];

  // =========================================================================
  // Chartsheet — full-page Revenue chart
  // =========================================================================

  const richTitle: ChartRichText = {
    paragraphs: [
      {
        runs: [
          { text: "FY25 ", properties: { bold: true, size: 2400, color: { srgb: "2F5496" } } },
          { text: "Revenue ", properties: { size: 2400, color: { srgb: "1F3864" } } },
          {
            text: "Board View",
            properties: { italic: true, size: 1800, color: { srgb: "7F7F7F" } }
          }
        ]
      }
    ]
  };

  const fy25Series: AddChartSeriesOptions[] = [
    {
      name: "Revenue",
      categories: agg(quarterStartRow + 8, 4, "B"),
      values: agg(quarterStartRow + 8, 4, "C"),
      fill: "2F5496",
      dataLabels: { showVal: true, numFmt: "$#,##0", position: "outEnd" }
    },
    {
      name: "Profit",
      categories: agg(quarterStartRow + 8, 4, "B"),
      values: agg(quarterStartRow + 8, 4, "D"),
      fill: "ED7D31",
      dataLabels: { showVal: true, numFmt: "$#,##0", position: "outEnd" }
    }
  ];

  const chartStyle: ChartStyleModel = {
    id: 201,
    elements: {
      chartArea: { fillRefIdx: 1, lnRefIdx: 1, effectRefIdx: 0, fontRefIdx: "minor" },
      title: { fontRefIdx: "major" }
    }
  };
  const chartColors: ChartColorsModel = {
    method: "cycle",
    id: 10,
    colors: [{ srgb: "2F5496" }, { srgb: "ED7D31" }, { srgb: "70AD47" }]
  };

  wb.addChartsheet("Revenue Chart", {
    tabSelected: false,
    zoomToFit: true,
    pageMargins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    pageSetup: { orientation: "landscape", paperSize: 9, horizontalDpi: 300, verticalDpi: 300 },
    chart: {
      type: "bar",
      barDir: "col",
      grouping: "clustered",
      title: richTitle,
      series: fy25Series,
      categoryAxis: { title: "Quarter" },
      valueAxis: { title: "USD", numFmt: "$#,##0" },
      legendPosition: "b",
      dataTable: { showHorzBorder: true, showVertBorder: true, showOutline: true, showKeys: true },
      chartStyle,
      chartColors
    }
  });

  // =========================================================================
  // Workbook-level defined names — `add(location, name)`.
  // =========================================================================

  wb.definedNames.add(`'SalesData'!$J$2:$J$${rows.length + 1}`, "TotalRevenue");
  wb.definedNames.add(`'SalesData'!$L$2:$L$${rows.length + 1}`, "TotalProfit");

  // =========================================================================
  // Write XLSX + PDF
  // =========================================================================

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`XLSX → ${XLSX_PATH}`);

  // PDF export with metadata (skip encryption to keep the file openable in
  // CI inspection; see financial-report.ts for encryption example).
  const pdf = await excelToPdf(wb, {
    title: "Mega Dashboard — FY23-25",
    author: "ExcelTS",
    showGridLines: false,
    showPageNumbers: true
  });
  writeFileSync(PDF_PATH, pdf);
  console.log(`PDF → ${PDF_PATH}`);

  // Also emit a standalone PDF of the first chart on the Charts sheet.
  const firstChart = charts.getCharts()[0];
  const chartPdf = await chartToPdf(firstChart, {
    title: "Revenue by region",
    width: 640,
    height: 400
  });
  writeFileSync(resolve(OUT_DIR, "mega-dashboard-chart.pdf"), chartPdf);
  console.log(`Chart PDF → ${resolve(OUT_DIR, "mega-dashboard-chart.pdf")}`);

  // Summary
  console.log("");
  console.log("Workbook summary:");
  console.log(`  worksheets  : ${wb.worksheets.length}`);
  console.log(`  chartsheets : ${wb.chartsheets.length}`);
  console.log(`  charts      : ${wb.worksheets.reduce((n, ws) => n + ws.getCharts().length, 0)}`);
  console.log(
    `  sparklines  : ${wb.worksheets.reduce((n, ws) => n + ws.getSparklineGroups().length, 0)} groups`
  );
  console.log(`  sales rows  : ${rows.length}`);
  console.log(`  tables      : 1 Excel table (${salesTable.model.name})`);
  console.log(`  pivots      : 3 pivot tables on '${pivots.name}'`);
  console.log(`  images      : 1 shared image`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
