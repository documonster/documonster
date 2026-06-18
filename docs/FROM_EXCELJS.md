# Migrating from ExcelJS

Documonster started life as a fork of [ExcelJS](https://github.com/exceljs/exceljs) and keeps the same workbook/worksheet/cell object graph at the top. This guide covers the **chart-related** API differences, which is where Documonster diverges most sharply: ExcelJS has **no native chart creation API**, so most migration work for chart code is translating ExcelJS "export the Excel unchanged" usage into Documonster structured chart authoring.

The rest of the API (workbook metadata, worksheets, rows, cells, styles, merges, data validations, images, conditional formatting, tables, defined names, print setup) is **source-compatible in the common cases** — see `MIGRATION.md` for the full cross-version delta within Documonster itself.

## What ExcelJS does not do, and Documonster does

| Capability                            | ExcelJS                                                   | Documonster                                                                                                                     |
| ------------------------------------- | --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| Create chart programmatically         | ✗ no public API                                           | `worksheet.addChart(options, range)`                                                                                            |
| Parse chart XML into a model          | partial (keeps as opaque blob)                            | full `ChartModel` / `ChartExModel` with round-trip                                                                              |
| Edit loaded chart title/legend/series | ✗                                                         | `chart.title =`, `chart.legend =`, `chart.addSeries()`, `chart.mutate()`                                                        |
| Render chart preview (SVG/PNG/PDF)    | ✗                                                         | `chart.toSVG()`, `chart.toPNG()`, `chartToPdf(chart)`                                                                           |
| Modern ChartEx types                  | ✗ (strips on save)                                        | `worksheet.addChartEx(...)` for sunburst / treemap / waterfall / funnel / histogram / pareto / boxWhisker / regionMap           |
| Preserve c15/c16 vendor extensions    | ✓ byte-preserved on clean load; silently lost on any edit | ✓ byte-preserved clean; raw-patched for safe edits; `templateMode: "strict"` fails loudly when a rebuild would drop unknown XML |
| Chart inside chartsheet               | ✗                                                         | `workbook.addChartsheet(name, { chart: { … } })`                                                                                |
| Pivot chart metadata                  | ✗                                                         | `worksheet.addPivotChart(...)` with drop-zones, field buttons, `c14:pivotOptions`, `c16:pivotOptions16`                         |
| External (linked) images              | ✗ embeds bytes only                                       | `workbook.addImage({ link })` references a URL/file path (`TargetMode="External"`, `<a:blip r:link>`) — no bytes stored         |

## Translation cookbook

### Create a bar chart (ExcelJS had no direct equivalent)

Documonster:

```typescript
import { Workbook } from "documonster";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.addRows([
  ["Q1", 100],
  ["Q2", 180],
  ["Q3", 240]
]);

ws.addChart(
  {
    type: "bar",
    barDir: "col",
    series: [{ name: "Revenue", categories: "Sheet1!$A$1:$A$3", values: "Sheet1!$B$1:$B$3" }],
    title: "Quarterly revenue"
  },
  "D1:J10"
);

await wb.xlsx.writeBuffer();
```

ExcelJS workflow for the same result — either hand-edit the exported `chartN.xml` after `writeBuffer`, or use a template file with the chart already in place and patch cell values. Both are error-prone and break on edits. Documonster replaces them with structured authoring.

### Load a template, tweak one series, write it back

ExcelJS had to round-trip the bytes and hope nothing changed. Documonster:

```typescript
const wb = new Workbook();
await wb.xlsx.load(fs.readFileSync("template.xlsx"));
const chart = wb.getWorksheet("Dashboard")!.getCharts()[0];

chart.mutate(
  model => {
    model.chart.plotArea.chartTypes[0].series[0].val = {
      formula: "Dashboard!$B$2:$B$25",
      cache: undefined
    };
  },
  { preferRawPatch: true } // surgical byte patch when safe; rebuild otherwise
);

await wb.xlsx.writeBuffer();
```

`preferRawPatch` keeps every byte Excel wrote that Documonster doesn't have a structured setter for, including vendor `c15:` / `c16:` extensions. See `src/modules/excel/README.md` → "Strict template mode" for the opt-in strict error path.

### Read cached cell values from a loaded chart

ExcelJS exposes nothing here. Documonster:

```typescript
const chart = wb.getWorksheet("S")!.getCharts()[0];
for (const group of chart.chartTypes) {
  for (const series of group.series) {
    console.log(series.tx, series.val?.cache?.points);
  }
}
```

### Render a chart preview server-side

ExcelJS users historically piped `.xlsx` through LibreOffice headless or `pdfkit`. Documonster ships a built-in preview:

```typescript
import { writeFileSync } from "node:fs";

const svg = chart.toSVG({ width: 640, height: 360 });
writeFileSync("chart.svg", svg);

const png = await chart.toPNG({ width: 640, height: 360, scale: 2 });
writeFileSync("chart.png", png);

// Standalone PDF — classic chart uses vector path, ChartEx uses raster.
import { chartToPdf } from "documonster/pdf";
const pdf = await chartToPdf(chart, { title: "Quarterly revenue" });
writeFileSync("chart.pdf", pdf);
```

The preview is **intentionally deterministic**, not Excel-pixel-perfect. See README → "Rendering" for the scope boundaries.

## Behavioural differences worth knowing

- **Byte preservation when clean.** Both libraries preserve the loaded bytes when no user code touches a chart. Documonster extends this to include a _raw-patch_ path for narrow edits (title, legend, single series value reference) — ExcelJS re-serialised the whole file.
- **Chart references are live.** `worksheet.getCharts()` returns `Chart` instances wired to the workbook; mutating them is how you edit. ExcelJS returned inert bytes.
- **PivotChart is metadata-only.** See README → "Pivot chart note". The user-visible rendering is still Excel's job.
- **No `chart.set_x_axis` / `set_y_axis` setter.** Use `chart.categoryAxis` / `chart.valueAxis` (getters) with direct property assignment, or pass `valueAxis: { … }` to `addChart` at creation time. See `FROM_XLSXWRITER.md` for a full mapping.

## Features ExcelJS has that Documonster intentionally does not

- **No auto-generated charts from `.set({ chart: true })` shortcut.** ExcelJS never had this; some forks do. Documonster requires an explicit `addChart({ type, series, … })` call to keep the schema check straightforward.

Nothing in the non-chart ExcelJS surface has been removed — if a call site works in ExcelJS, it generally works in Documonster unless noted in `MIGRATION.md`.

## Chartsheet migration

ExcelJS treats chartsheets as unsupported — it loads the workbook but the chartsheet's chart body is inaccessible and lost on write. Documonster exposes them as first-class objects:

```typescript
// Create a chartsheet (a sheet with exactly one chart, no grid)
const chartsheet = workbook.addChartsheet("Revenue Dashboard", {
  chart: {
    type: "bar",
    barDir: "col",
    series: [{ name: "Revenue", categories: "Data!$A$1:$A$12", values: "Data!$B$1:$B$12" }],
    title: "Monthly revenue"
  },
  tabSelected: true,
  pageSetup: { orientation: "landscape", paperSize: 9, scale: 100 }
});

// Access the chart for further mutation
chartsheet.chart.title = "Quarterly revenue";

// ChartEx chartsheet (sunburst on its own sheet)
workbook.addChartsheet("Hierarchy", {
  chart: {
    type: "sunburst",
    series: [{ values: "Data!$B$1:$B$20", hierarchy: ["Data!$A$1:$A$20"] }]
  }
});

// Clone / rename / hide a chartsheet the same way as a worksheet
chartsheet.rename("Dashboard");
chartsheet.state = "veryHidden";
```

Chartsheets round-trip losslessly through `wb.xlsx.load` → `writeBuffer` including their `pageSetup`, `printOptions`, `tabSelected`, and visibility state. Pivot chartsheets are supported via `workbook.addPivotChartsheet(name, { pivotTable, chart })`.

## Pivot chart migration

ExcelJS has no pivot-chart surface — loaded pivot charts survive as opaque bytes only. Documonster gives you structured pivot-chart metadata:

```typescript
// Create the pivot table first
const pivot = ws.addPivotTable({
  name: "RegionSales",
  ref: "F3",
  source: "Data!A1:D1000",
  rows: [{ name: "Region" }],
  columns: [{ name: "Quarter" }],
  values: [{ name: "Revenue", subtotal: "sum" }]
});

// Then attach a pivot chart — the metadata is what Excel uses to keep
// the two in sync (pivotSource + pivotFmt + field buttons)
ws.addPivotChart(
  pivot,
  {
    type: "bar",
    series: [{ name: "Revenue", categories: "", values: "" }],
    title: "Revenue by region"
  },
  "D1:L20"
);
```

Documonster supports `c14:pivotOptions` (field buttons, drop-zone hints, `refreshOnOpen`) and `c16:pivotOptions16` (Office 2019+ expand/collapse buttons). The **rendering** still lives with Excel — Documonster preserves the metadata, hosts read it to draw the real chart.

## ChartEx (modern chart types) migration

This is where Documonster diverges most: ExcelJS does not support any Excel-2016+ chart (sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap). These are stored in a separate OOXML namespace (`cx:`) that ExcelJS silently strips.

```typescript
// Waterfall with subtotals
ws.addChartEx(
  {
    type: "waterfall",
    categories: "Data!$A$1:$A$6",
    series: [{ values: "Data!$B$1:$B$6" }],
    layout: {
      subtotals: [{ idx: 3 }, { idx: 6 }],
      connectorLines: true
    },
    title: "Cash flow waterfall"
  },
  "D1:J10"
);

// Sunburst with drill-down hierarchy
ws.addChartEx(
  {
    type: "sunburst",
    series: [
      {
        values: "Data!$D$1:$D$30",
        hierarchy: ["Data!$A$1:$A$30", "Data!$B$1:$B$30", "Data!$C$1:$C$30"]
      }
    ]
  },
  "D1:J10"
);

// Box & whisker with custom quartile method
ws.addChartEx(
  {
    type: "boxWhisker",
    categories: "Data!$A$1:$A$5",
    series: [{ values: "Data!$B$1:$B$5" }],
    layout: {
      quartileMethod: "inclusive",
      showMeanLine: true,
      showMeanMarker: true,
      showInnerPoints: false,
      showOutlierPoints: true
    }
  },
  "D1:J12"
);

// Region map with Albers projection (requires TopoJSON at render time)
ws.addChartEx(
  {
    type: "regionMap",
    categories: "Data!$A$1:$A$50", // country names
    series: [{ values: "Data!$B$1:$B$50" }],
    layout: { projection: "albers", regionLabels: "showAll" }
  },
  "D1:M20"
);
```

Need to stage data from a JavaScript array first? Use the table/rows helpers, which mirror the classic `chartOptionsFromTable` / `chartOptionsFromRows` utilities:

```typescript
// From an existing Table
ws.addChartExFromTable("SalesTable", { type: "funnel", categoryColumn: "Stage" }, "D1:J10");

// From a plain array (rows are written to the sheet first)
ws.addChartExFromRows(
  [
    { stage: "Visits", count: 1000 },
    { stage: "Sign-ups", count: 400 },
    { stage: "Purchases", count: 80 }
  ],
  { type: "funnel", x: "stage", y: "count", startCell: "A1" },
  "D1:J10"
);
```

ChartEx charts round-trip with full preservation of `layoutPr` fields. See `docs/COMPATIBILITY.md` → "ChartEx types" for the full feature matrix.

## Data table migration

ExcelJS has no `c:dTable` API. Documonster exposes it on both the builder and the renderer:

```typescript
ws.addChart(
  {
    type: "bar",
    series: [{ name: "Revenue", categories: "Data!$A$1:$A$4", values: "Data!$B$1:$B$4" }],
    dataTable: {
      showHorzBorder: true,
      showVertBorder: true,
      showOutline: true,
      showKeys: true // draws legend swatches next to series names
    }
  },
  "D1:J12"
);
```

The preview renderer (SVG/PNG/PDF) draws the data table as a grid below the plot area with the series-name column on the left, one row per series, and a header row of category names. Excel-side, the same metadata renders the data table natively. X-axis category labels are suppressed while the data table is visible (matching Excel's behaviour).

## User-shape overlay migration

ExcelJS preserves the `c:userShapes` reference on clean loads but cannot edit it and silently drops the backing drawing part on most rebuilds. Documonster persists the bytes:

```typescript
const chart = ws.getCharts()[0];

// Read the overlay drawing (undefined if the chart has no user shapes)
const existingXml = chart.userShapesXml;

// Replace — pass a DrawingML document whose root is c:userShapes
chart.setUserShapesXml(`<?xml version="1.0" encoding="UTF-8"?>
<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:relSizeAnchor>
    <c:from><c:x>0.1</c:x><c:y>0.05</c:y></c:from>
    <c:to><c:x>0.4</c:x><c:y>0.15</c:y></c:to>
    <c:sp>
      <c:nvSpPr><c:cNvPr id="1" name="Callout"/><c:cNvSpPr/></c:nvSpPr>
      <c:spPr/>
      <c:txBody>
        <a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>Peak!</a:t></a:r></a:p>
      </c:txBody>
    </c:sp>
  </c:relSizeAnchor>
</c:userShapes>`);

// Remove
chart.removeUserShapes();
```

The library deliberately does not reimplement the DrawingML shape model — the bytes stay opaque and the rels/content-types/part path are managed automatically. This keeps "byte-preserving" the headline guarantee for exotic overlay content.

## Unknown-element / strict template migration

ExcelJS' byte-preservation breaks as soon as you touch a chart — the re-serialisation discards any vendor-extension XML (`c15:` from Office 2013, `cx14:` from Office 2016) that didn't map to its structured model. Documonster gives you two escape hatches:

```typescript
// Diagnose before writing — returns c15:/cx14: vendor tags the parser observed
const chart = ws.getCharts()[0];
const unknown = chart.unknownElements;
if (unknown) {
  console.warn(
    "Chart carries vendor extensions that will be dropped by a structural rebuild:",
    unknown.map(e => e.path)
  );
}

// Raw-XML patch path for narrow edits (title, single series val ref, grouping flags)
chart.mutate(
  model => {
    model.chart.title = { text: { paragraphs: [{ runs: [{ text: "New title" }] }] } };
  },
  { preferRawPatch: true } // tries surgical byte patching first
);

// Strict mode — refuse any write that would force a rebuild that drops unknown XML
await wb.xlsx.writeBuffer({ templateMode: "strict" });
// or: await wb.xlsx.writeBuffer({ strictTemplateMode: true });
```

The strict writer's error message enumerates the specific unknown paths ("The loaded part contains unstructured XML at: cx:chartSpace/c15:vendorFoo, cx:chart/c16:customLayout"), so authors can decide between relaxing the mode or reshaping the mutation to land on a raw-patch-friendly path.

## Migration checklist

When porting an ExcelJS codebase to Documonster, run through this list:

- [ ] Any `chart.` property access? ExcelJS only exposes `chart.name` — Documonster adds `title`, `legend`, `chartTypes[]`, `categoryAxis`, `valueAxis`, `chartModel`, `chartExModel`, `unknownElements`.
- [ ] Template-driven charts (load → write unchanged)? Works identically with byte-level round-trip.
- [ ] Template + one edit? Use `chart.mutate(fn, { preferRawPatch: true })` to keep unknown XML intact.
- [ ] Post-write validation? Add `{ templateMode: "strict" }` to catch silent extension loss in CI.
- [ ] Chart rendering outside Excel? Swap `libreoffice-convert` / `pdfkit` for `chart.toSVG()` / `chart.toPNG()` / `chartToPdf(chart)` — all zero-dependency (see `README.md` → "Rendering scope" for the preview-vs-production boundary).
- [ ] Modern chart types (sunburst, treemap, waterfall, funnel, histogram, pareto, boxWhisker, regionMap)? Use `addChartEx` — they were unreachable in ExcelJS.
- [ ] Chartsheet contents? Use `addChartsheet(name, { chart })` — ExcelJS dropped the chart body.
- [ ] Pivot chart metadata? Use `addPivotChart(pivotTable, options, range)` — ExcelJS lost this entirely.

If you hit a gap, open an issue with the ExcelJS snippet and we'll add the mapping.
