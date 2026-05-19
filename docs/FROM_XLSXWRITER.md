# Migrating from XlsxWriter (Python)

[XlsxWriter](https://xlsxwriter.readthedocs.io/chart.html) is a Python, write-only library that is the de facto reference for "how a chart API should feel." ExcelTS is the closest TypeScript/JavaScript equivalent — its option shapes, preset names, and structural grouping were modelled after XlsxWriter's, with additions for reading, editing, rendering, and the full ChartEx surface that XlsxWriter does not expose.

Companion guides: [`FROM_OPENPYXL.md`](./FROM_OPENPYXL.md) · [`FROM_EXCELIZE.md`](./FROM_EXCELIZE.md) · [`FROM_POI.md`](./FROM_POI.md) · [`FROM_EXCELJS.md`](./FROM_EXCELJS.md)

## Cheat sheet — creation

| Task                        | XlsxWriter (Python)                         | ExcelTS (TypeScript)                                                                                                                          |
| --------------------------- | ------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| Create a workbook           | `Workbook("f.xlsx")`                        | `new Workbook()`                                                                                                                              |
| Add a worksheet             | `wb.add_worksheet()`                        | `wb.addWorksheet("S")`                                                                                                                        |
| Write cell values           | `ws.write(row, col, val)`                   | `ws.getCell(addr).value = val` or `ws.addRows([[…]])`                                                                                         |
| Create a chart object       | `wb.add_chart({"type": "bar"})`             | `ws.addChart({ type: "bar", series: [...] }, range)`                                                                                          |
| Add a series                | `chart.add_series({…})`                     | `chart.addSeries(series)` or in `addChart` options                                                                                            |
| Set title                   | `chart.set_title({"name": "T"})`            | `chart.title = "T"` or in `addChart` options                                                                                                  |
| Set x/y axis                | `chart.set_x_axis({…})` / `set_y_axis({…})` | Via getter: `chart.categoryAxis.title = "…"`, `chart.valueAxis.majorUnit = 10` — or pass `valueAxis: {…}` / `categoryAxis: {…}` to `addChart` |
| Set a built-in style (1–48) | `chart.set_style(37)`                       | `chart.setStyle(37)` (alias `chart.setBuiltInStyle(37)`)                                                                                      |
| Insert chart into sheet     | `ws.insert_chart("D2", chart)`              | Second argument of `addChart` (a range like `"D2:J10"`)                                                                                       |
| Combine two charts          | `primary.combine(secondary)`                | `ws.addComboChart({ groups: [{ type: "bar", … }, { type: "line", … }] }, range)`                                                              |
| Render preview              | ✗ (none)                                    | `chart.toSVG()` / `chart.toPNG()` / `chartToPdf(chart)` (zero-dependency preview)                                                             |
| Modern chart types          | ✗ (none)                                    | `ws.addChartEx({ type: "sunburst" / "waterfall" / … })`                                                                                       |

## Example 1 — End-to-end column chart

**Python (xlsxwriter):**

```python
import xlsxwriter
wb = xlsxwriter.Workbook("report.xlsx")
ws = wb.add_worksheet("Sales")
ws.write_row("A1", ["Q1", "Q2", "Q3", "Q4"])
ws.write_row("A2", [100, 180, 240, 310])

chart = wb.add_chart({"type": "column"})
chart.add_series({
    "name":       "Revenue",
    "categories": "=Sales!$A$1:$D$1",
    "values":     "=Sales!$A$2:$D$2",
    "data_labels": {"value": True},
    "trendline":  {"type": "linear"},
})
chart.set_title({"name": "Quarterly revenue"})
chart.set_x_axis({"name": "Quarter"})
chart.set_y_axis({"name": "USD", "major_unit": 50})
chart.set_style(37)

ws.insert_chart("F1", chart, {"x_scale": 1.2, "y_scale": 1.2})
wb.close()
```

**TypeScript (ExcelTS):**

```typescript
import { Workbook } from "@cjnoname/excelts";

const wb = new Workbook();
const ws = wb.addWorksheet("Sales");
ws.getRow(1).values = ["Q1", "Q2", "Q3", "Q4"];
ws.getRow(2).values = [100, 180, 240, 310];

const chartNum = ws.addChart(
  {
    type: "bar",
    barDir: "col", // "column" in xlsxwriter == barDir "col" in OOXML
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$1:$D$1",
        values: "Sales!$A$2:$D$2",
        dataLabels: { showVal: true },
        trendline: { type: "linear" }
      }
    ],
    title: "Quarterly revenue",
    categoryAxis: { title: "Quarter" },
    valueAxis: { title: "USD", majorUnit: 50 }
  },
  "F1:M14"
);
const chart = ws.getCharts().find(c => c.chartNumber === chartNum)!;
chart.setStyle(37);

await wb.xlsx.writeFile("report.xlsx"); // Node
// or: const bytes = await wb.xlsx.writeBuffer();  // browser / buffer
```

**Key translations**

- `"column"` → `{ type: "bar", barDir: "col" }` (OOXML nomenclature).
- `set_title({"name": "T"})` → `title: "T"` option, or `chart.title = "T"` later.
- `set_x_axis({"name": …})` → `categoryAxis: { title: … }` option.
- `set_y_axis({"major_unit": 50})` → `valueAxis: { majorUnit: 50 }` option.
- `data_labels: {"value": True}` → `dataLabels: { showVal: true }`.
- `trendline: {"type": "linear"}` → `trendline: { type: "linear" }` (identical).
- `insert_chart("F1", …, {"x_scale": 1.2, "y_scale": 1.2})` → range argument `"F1:M14"` sizes the chart directly (ExcelTS uses the anchor range; if you want scale-based sizing pick a larger range).
- `chart.set_style(37)` → `chart.setStyle(37)` or `setBuiltInStyle(37)`.

## Example 2 — Combined chart (Pareto-style)

**Python:**

```python
col = wb.add_chart({"type": "column"})
col.add_series({"name": "Count", "categories": "=S!$A$2:$A$6", "values": "=S!$B$2:$B$6"})

line = wb.add_chart({"type": "line"})
line.add_series({
    "name":    "Cumulative %",
    "categories": "=S!$A$2:$A$6",
    "values":     "=S!$C$2:$C$6",
    "y2_axis":    True,
})
col.combine(line)
col.set_title({"name": "Pareto"})
col.set_y2_axis({"name": "%", "max": 100})
ws.insert_chart("E2", col)
```

**TypeScript:**

```typescript
ws.addComboChart(
  {
    groups: [
      {
        type: "bar",
        barDir: "col",
        series: [{ name: "Count", categories: "S!$A$2:$A$6", values: "S!$B$2:$B$6" }]
      },
      {
        type: "line",
        series: [
          {
            name: "Cumulative %",
            categories: "S!$A$2:$A$6",
            values: "S!$C$2:$C$6",
            useSecondaryAxis: true
          }
        ]
      }
    ],
    title: "Pareto",
    secondaryValueAxis: { title: "%", max: 100 }
  },
  "E2:M18"
);
```

XlsxWriter documents that pie/scatter cannot be the primary chart in a combo. ExcelTS enforces the same via the `ChartOptionsError`: you'll get a typed error at build time rather than a mysterious write failure.

## Example 3 — Chart from a worksheet Table

XlsxWriter's `add_table` + `add_chart` requires you to manage row ranges manually. ExcelTS ships a helper:

```typescript
ws.addTable({
  name: "Sales",
  ref: "A1",
  columns: [{ name: "Month" }, { name: "Revenue" }, { name: "Profit" }],
  rows: [
    ["Jan", 100, 20],
    ["Feb", 180, 40],
    ["Mar", 240, 55]
  ]
});

// Structured references — chart tracks table expansion
ws.addChartFromTable(
  "Sales",
  { type: "bar", barDir: "col", categoryColumn: "Month", valueColumns: ["Revenue", "Profit"] },
  "E1:L15"
);

// Absolute refs — useful if the chart should freeze to the initial range
ws.addChartFromTable(
  "Sales",
  {
    type: "line",
    categoryColumn: "Month",
    valueColumns: ["Revenue"],
    structuredReferences: false
  },
  "E17:L31"
);
```

Same helper exists for ChartEx: `ws.addChartExFromTable("Sales", { type: "funnel" }, range)`.

## Example 4 — Chart from an in-memory dataset

XlsxWriter has no object-to-chart helper; you have to write cells first and remember the ranges. ExcelTS:

```typescript
ws.addChartFromRows(
  [
    { month: "Jan", revenue: 100, profit: 20 },
    { month: "Feb", revenue: 180, profit: 40 },
    { month: "Mar", revenue: 240, profit: 55 }
  ],
  {
    type: "bar",
    barDir: "col",
    x: "month",
    y: ["revenue", "profit"],
    startCell: "A1",
    title: "Q1 performance"
  },
  "E1:L15"
);
// The rows are written to A1:C4 (with headers) and the chart references
// those cells. A1 "month" cell holds the header; ranges use absolute refs.
```

ChartEx variant: `ws.addChartExFromRows(rows, { type: "treemap", x, y, startCell }, range)`.

## Example 5 — Load, patch a single series, save

XlsxWriter is write-only. ExcelTS round-trips:

```typescript
import { readFileSync } from "node:fs";

const wb = new Workbook();
await wb.xlsx.load(readFileSync("template.xlsx"));
const chart = wb.getWorksheet("Dashboard")!.getCharts()[0];

chart.mutate(
  model => {
    // Change the first series' values reference
    const group = model.chart.plotArea.chartTypes[0];
    group.series[0].val = {
      formula: "Dashboard!$B$2:$B$25",
      cache: undefined // force Excel to re-compute on open
    };
  },
  { preferRawPatch: true } // keep every byte XlsxWriter/Excel wrote
);

await wb.xlsx.writeFile("out.xlsx");
```

`preferRawPatch` lets ExcelTS apply a surgical byte patch for narrow edits (title, single series ref, grouping flags) so vendor extensions like `c15:layoutFlag` stay intact. If the edit can't patch cleanly the writer falls back to a structural rebuild; `templateMode: "strict"` turns that fallback into a hard error.

## Example 6 — Rendering preview (XlsxWriter has none)

```typescript
import { writeFileSync } from "node:fs";

const svg = chart.toSVG({ width: 640, height: 360 });
writeFileSync("chart.svg", svg);

const png = await chart.toPNG({ width: 640, height: 360, scale: 2 });
writeFileSync("chart.png", png);

// Standalone PDF — classic chart uses vector path, ChartEx uses vector too
import { chartToPdf } from "@cjnoname/excelts/pdf";
const pdf = await chartToPdf(chart, { title: "Quarterly revenue" });
writeFileSync("chart.pdf", pdf);
```

**Scope reminder**: this is a zero-dependency deterministic preview, not an Excel-pixel-perfect compositor. For production-grade output, use `chart.toSVG()` for dashboards and route the `.xlsx` through headless LibreOffice for Excel-identical PDFs — ExcelTS' byte-preserving round-trip makes that handoff safe. See README → "Rendering scope" for the full boundary list.

## Behavioural differences worth knowing

- **Chart sizing.** XlsxWriter uses `x_scale` / `y_scale` on `insert_chart`; ExcelTS sizes via the anchor range passed to `addChart` (`"D1:J10"` occupies cells D1 through J10). Use a larger range for a bigger chart.
- **Data reference format.** XlsxWriter accepts `[sheet, first_row, first_col, last_row, last_col]` arrays as a convenience; ExcelTS requires the Excel-style `"Sheet!$A$1:$A$10"` string. Fully-qualified (with sheet name) is safest.
- **No implicit cache refresh at save.** If you mutate a chart's value references but the target cells already hold the cache, ExcelTS keeps the old cache by default. Call `fillChartCaches(chart.chartModel, worksheet)` to regenerate explicitly.
- **No `chart.set_table()` / `chart.set_legend({"none": True})` shortcuts.** Use `dataTable: true` on `addChart` and `{ legendPos: "none" }` (ExcelTS uses OOXML nomenclature).
- **Pattern fills are structured, not enums.** XlsxWriter's 48 preset patterns become `pattern: { preset: "percent50", foreground: "FF0000", background: "FFFFFF" }` — same preset list, explicit shape.

## Features XlsxWriter has that ExcelTS maps differently

- **Sparklines** — XlsxWriter's `ws.add_sparkline(...)` becomes `ws.addSparkline({ ... })` (not a chart); covered in the main ExcelTS README.
- **`add_data_table`** — XlsxWriter exposes data tables only as an on/off toggle. ExcelTS supports the full `c:dTable` structure: `dataTable: { showHorzBorder, showVertBorder, showOutline, showKeys, spPr, txPr }`.

## Features ExcelTS has that XlsxWriter does not

- **Read / edit / mutate**: structured access to loaded charts, not write-only.
- **ChartEx**: sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap.
- **Pivot chart metadata**: `ws.addPivotChart(pivot, options, range)`.
- **Chartsheet with full `pageSetup` / `printOptions` / `state`**.
- **Preview renderer**: SVG/PNG/PDF from the model directly.
- **User-shape overlays**: `Chart.userShapesXml` / `setUserShapesXml`.
- **Vendor-extension preservation**: `templateMode: "strict"` + `Chart.unknownElements`.
- **Browser support**: ExcelTS is platform-agnostic; XlsxWriter is Python-only.

If you hit a gap in the mapping, open an issue with the XlsxWriter snippet and we'll add the translation.
