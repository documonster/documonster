# Migrating from Apache POI (Java)

[Apache POI](https://poi.apache.org/components/spreadsheet/) is Java's de facto Excel library. Its XSSF (`.xlsx`) chart API is class-based like openpyxl but with a different abstraction: `XSSFChart`, `ChartDataFactory`, `XDDFChartData`, `XDDFLineChartData` etc. It is the open-source tool with the deepest chart-format coverage before ExcelTS — including partial ChartEx read support — so migration is mostly a surface translation rather than a capability expansion.

Companion guides: [`FROM_XLSXWRITER.md`](./FROM_XLSXWRITER.md) · [`FROM_OPENPYXL.md`](./FROM_OPENPYXL.md) · [`FROM_EXCELIZE.md`](./FROM_EXCELIZE.md) · [`FROM_EXCELJS.md`](./FROM_EXCELJS.md)

## Why migrate

- **Runtime**: switching a Node/TypeScript backend off a JVM dependency (POI + a Java runtime) can cut container size by 100 – 300 MB.
- **Browser support**: POI requires JVM; ExcelTS runs in Node _and_ modern browsers.
- **ChartEx**: POI's `cx:` support is raw-XML read-only; ExcelTS creates, reads, edits, and previews all 8 ChartEx types.
- **Preview renderer**: POI requires round-tripping through a separate image library (JasperReports, Apache Batik) to get a chart image; ExcelTS ships SVG/PNG/PDF out of the box.

## Philosophy map

| Aspect             | POI (Java)                                                                   | ExcelTS                                                                                               |
| ------------------ | ---------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Chart object       | `XSSFChart`                                                                  | `Chart` (returned from `ws.getCharts()`)                                                              |
| Chart data factory | `XDDFChartData chartData = chart.createData(...)`                            | Options object passed to `addChart({ type, series, ... })`                                            |
| Data source        | `XDDFDataSourcesFactory.fromNumericCellRange(ws, new CellRangeAddress(...))` | Excel-style A1 string: `"Sheet1!$B$2:$B$10"`                                                          |
| Series             | `chartData.addSeries(xSource, ySource)`                                      | `series: [{ categories, values, name }]` in options                                                   |
| Line type          | `XDDFChartData.Series s = ...` then `s.setSmooth(true)`                      | `series: [{ smooth: true }]`                                                                          |
| 3D rotation        | `XSSFChart.getCTChart().getView3D()` (raw XML)                               | `view3D: { rotX, rotY, depthPercent, rAngAx }` on options                                             |
| Edit loaded chart  | Reach into `chart.getCTChart()` and mutate raw XML                           | `chart.mutate(model => { ... }, { preferRawPatch: true })` — structured, with byte-preserving patches |
| Render chart image | External library (Batik, JasperReports)                                      | `chart.toSVG()` / `chart.toPNG()` / `chartToPdf(chart)` (zero-dependency)                             |

## Example 1 — Basic bar chart

**Java:**

```java
XSSFWorkbook wb = new XSSFWorkbook();
XSSFSheet sheet = wb.createSheet("Sales");

// Data
String[] months = {"Jan", "Feb", "Mar"};
Double[] revenue = {100.0, 180.0, 240.0};
Row header = sheet.createRow(0);
header.createCell(0).setCellValue("Month");
header.createCell(1).setCellValue("Revenue");
for (int i = 0; i < months.length; i++) {
    Row r = sheet.createRow(i + 1);
    r.createCell(0).setCellValue(months[i]);
    r.createCell(1).setCellValue(revenue[i]);
}

// Chart
XSSFDrawing drawing = sheet.createDrawingPatriarch();
XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 12, 15);
XSSFChart chart = drawing.createChart(anchor);
chart.setTitleText("Quarterly revenue");

XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);

XDDFDataSource<String> monthsSource =
    XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 3, 0, 0));
XDDFNumericalDataSource<Double> revenueSource =
    XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 3, 1, 1));

XDDFBarChartData data = (XDDFBarChartData) chart.createData(
    ChartTypes.BAR, bottomAxis, leftAxis);
data.setBarDirection(BarDirection.COL);
XDDFBarChartData.Series s = (XDDFBarChartData.Series) data.addSeries(monthsSource, revenueSource);
s.setTitle("Revenue", null);
chart.plot(data);

wb.write(new FileOutputStream("out.xlsx"));
wb.close();
```

**TypeScript:**

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const wb = new Workbook();
const ws = wb.addWorksheet("Sales");
ws.addRow(["Month", "Revenue"]);
ws.addRow(["Jan", 100]);
ws.addRow(["Feb", 180]);
ws.addRow(["Mar", 240]);

ws.addChart(
  {
    type: "bar",
    barDir: "col",
    series: [
      {
        name: "Revenue",
        categories: "Sales!$A$2:$A$4",
        values: "Sales!$B$2:$B$4"
      }
    ],
    title: "Quarterly revenue"
  },
  "E1:M16"
);

await wb.xlsx.writeFile("out.xlsx");
```

**Key translations**

- `XSSFDrawing drawing = sheet.createDrawingPatriarch()` + `XSSFChart chart = drawing.createChart(anchor)` → all handled by `addChart(options, range)`.
- `XDDFDataSourcesFactory.fromNumericCellRange(...)` → A1 string `"Sheet1!$B$2:$B$4"`.
- `data.setBarDirection(BarDirection.COL)` → `barDir: "col"`.
- `AxisPosition.BOTTOM / LEFT` axes are created implicitly; overrides go in `categoryAxis: { ... }` / `valueAxis: { ... }`.
- `s.setTitle("Revenue", null)` → `series: [{ name: "Revenue" }]`. Use `name: { formula: "Sheet1!$B$1" }` for cell-sourced titles (POI's `setTitle(CellReference, null)` equivalent).

## Example 2 — Line + secondary axis

**Java:**

```java
XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(
    ChartTypes.LINE, bottomAxis, rightAxis);
lineData.setVaryColors(true);

XDDFLineChartData.Series primary = (XDDFLineChartData.Series) lineData.addSeries(
    xSource, ySource1);
primary.setTitle("Revenue", null);

XDDFLineChartData.Series secondary = (XDDFLineChartData.Series) lineData.addSeries(
    xSource, ySource2);
secondary.setTitle("Margin %", null);
secondary.setSmooth(true);
// Binding to secondary axis requires reaching into the underlying CT types:
CTPlotArea plotArea = chart.getCTChart().getPlotArea();
CTValAx secondaryAxis = plotArea.addNewValAx();
// ... multiple more lines configuring secondary axis ...

chart.plot(lineData);
```

**TypeScript:**

```typescript
ws.addChart(
  {
    type: "line",
    series: [
      { name: "Revenue", categories: "S!$A$2:$A$13", values: "S!$B$2:$B$13" },
      {
        name: "Margin %",
        categories: "S!$A$2:$A$13",
        values: "S!$C$2:$C$13",
        smooth: true,
        useSecondaryAxis: true
      }
    ],
    title: "Revenue vs margin",
    secondaryValueAxis: { title: "%", max: 100 }
  },
  "E1:M15"
);
```

The secondary-axis wiring that POI requires reaching into `CTValAx` / `CTPlotArea` is hidden behind the single `useSecondaryAxis: true` flag.

## Example 3 — Combo (column + line)

**Java:**

```java
XDDFBarChartData bar = (XDDFBarChartData) chart.createData(
    ChartTypes.BAR, bottomAxis, leftAxis);
bar.setBarDirection(BarDirection.COL);
bar.addSeries(xSource, countSource).setTitle("Count", null);

XDDFLineChartData line = (XDDFLineChartData) chart.createData(
    ChartTypes.LINE, bottomAxis, rightAxis);
XDDFLineChartData.Series cumul = (XDDFLineChartData.Series) line.addSeries(xSource, cumSource);
cumul.setTitle("Cumulative %", null);

chart.plot(bar);
chart.plot(line);
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
  "E1:M18"
);
```

`addComboChart` keeps the two chart groups in one structured call — POI needs two `chart.plot()` invocations plus an explicit secondary axis setup. Both approaches produce the same OOXML, ExcelTS just needs fewer steps.

## Example 4 — Load existing chart and edit

**Java:**

```java
XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("existing.xlsx"));
for (Sheet s : wb) {
    XSSFSheet sheet = (XSSFSheet) s;
    XSSFDrawing drawing = sheet.getDrawingPatriarch();
    if (drawing == null) continue;
    for (XSSFChart chart : drawing.getCharts()) {
        // Title access — works via raw XML
        CTTitle title = chart.getCTChart().getTitle();
        // ... modifying the value series reference requires CT-level code
        CTBarChart barCT = chart.getCTChart().getPlotArea().getBarChartList().get(0);
        CTBarSer ser = barCT.getSerList().get(0);
        ser.getVal().getNumRef().setF("Data!$B$2:$B$50");
        ser.getVal().getNumRef().unsetNumCache(); // force Excel to recompute
    }
}
wb.write(new FileOutputStream("out.xlsx"));
```

**TypeScript:**

```typescript
await wb.xlsx.load(readFileSync("existing.xlsx"));
for (const ws of wb.worksheets) {
  for (const chart of ws.getCharts()) {
    chart.mutate(
      model => {
        const group = model.chart.plotArea.chartTypes[0];
        group.series[0].val = { formula: "Data!$B$2:$B$50", cache: undefined };
      },
      { preferRawPatch: true }
    );
  }
}
await wb.xlsx.writeBuffer();
```

`preferRawPatch: true` lets ExcelTS apply a surgical byte patch (no XML rebuild) for narrow edits — preserving every `c15:` / `c16:` vendor extension byte-for-byte. POI's CT-level code rebuilds the subtree on every edit, which can silently drop unknown XML.

## Example 5 — Chart style (1–48)

**Java:**

```java
XDDFChartStyle styleN = new XDDFChartStyle();
chart.getCTChart().addNewStyle().setVal((short) 37);
```

**TypeScript:**

```typescript
const chart = ws.getCharts().find(c => c.chartNumber === chartNum)!;
chart.setStyle(37); // alias: chart.setBuiltInStyle(37)
```

## Example 6 — ChartEx (POI partial read-only; ExcelTS full authoring)

POI can read a chartEx `cx:chart` part as raw XML but has no structured builder for it. ExcelTS:

```typescript
// Sunburst — POI has no structural equivalent
ws.addChartEx(
  {
    type: "sunburst",
    series: [
      {
        values: "Data!$D$1:$D$30",
        hierarchy: ["Data!$A$1:$A$30", "Data!$B$1:$B$30", "Data!$C$1:$C$30"]
      }
    ],
    title: "Org headcount"
  },
  "D1:J12"
);

// Box-and-whisker — POI can read cx:boxWhisker but not create it
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

// Region map — POI has no support
ws.addChartEx(
  {
    type: "regionMap",
    categories: "Data!$A$1:$A$50",
    series: [{ values: "Data!$B$1:$B$50" }],
    layout: { projection: "albers", regionLabels: "showAll" }
  },
  "D1:N25"
);
```

ChartEx helper APIs: `chartExOptionsFromTable` / `chartExOptionsFromRows` for the common "from-existing-data" cases — see `docs/FROM_OPENPYXL.md` Example 6 for equivalent patterns applied to `addChartEx`.

## Example 7 — Rendering preview (POI has none)

```typescript
import { writeFileSync } from "node:fs";

const svg = chart.toSVG({ width: 640, height: 360 });
writeFileSync("chart.svg", svg);

const png = await chart.toPNG({ width: 640, height: 360, scale: 2 });
writeFileSync("chart.png", png);

import { chartToPdf } from "@cj-tech-master/excelts/pdf";
const pdf = await chartToPdf(chart, { title: "Quarterly revenue" });
writeFileSync("chart.pdf", pdf);
```

**Scope reminder**: zero-dependency deterministic preview, not Excel-pixel-perfect. For publication-grade output route the `.xlsx` through `soffice --convert-to pdf` — ExcelTS' byte-preserving round-trip makes the handoff safe. See README → "Rendering scope" for the full boundary list.

## Behavioural differences worth knowing

- **Chart sizing.** POI's `XSSFClientAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2)` uses cell-based coordinates with EMU offsets. ExcelTS' range (`"E1:M18"`) is cell-based only — set column widths / row heights to tune the pixel size, or use the cell-offset form via `chart.rangeFromOffset({ ... })`.
- **No explicit axis objects.** POI requires creating `XDDFCategoryAxis` / `XDDFValueAxis` up front and passing them to `createData`. ExcelTS auto-creates the standard axis pair; pass `categoryAxis: { ... }` / `valueAxis: { ... }` only when you need non-default settings.
- **Multiple `chart.plot(data)`** becomes one `addComboChart({ groups: [...] })` call.
- **Raw CT access.** If you really need to drop to the CT level (new element not yet covered), reach for `chart.chartModel` (structured) and then call `chart.mutate(fn, { preferRawPatch })` — the model's type is declared in `src/modules/excel/chart/types.ts`.

## Features ExcelTS has that POI does not

- **ChartEx structural authoring** (POI is read-only).
- **Pivot chart metadata** with structured field buttons + `c14:pivotOptions` + `c16:pivotOptions16`.
- **Preview renderer** (SVG/PNG/PDF, zero-dependency).
- **Vendor-extension preservation** via `templateMode: "strict"` + `Chart.unknownElements`.
- **User-shape overlays** (`c:userShapes`) with byte-preserving programmatic replacement.
- **Browser support**.

## Features POI has that ExcelTS maps differently

- **`XDDFLineProperties` / `XDDFShapeProperties`** — ExcelTS exposes equivalent fields on `series[i].spPr` (fill/line/effects) and `dataLabels.spPr`. The schema matches OOXML; the field names are camelCase rather than Java's PascalCase.
- **`AxesManager`** — ExcelTS' axis model is inline on `addChart` options (`categoryAxis`, `valueAxis`, `secondaryValueAxis`, `secondaryCategoryAxis`) plus post-creation mutation via `chart.categoryAxis` / `chart.valueAxis` getters.
- **`ChartShapeProperties.setFillProperties(...)`** — pass `spPr: { solidFill: { srgbClr: "4472C4" } }` directly on the chart options or per-series.
- **`XSSFChart.getEmbedded()`** — POI's escape hatch to the raw CT. ExcelTS' equivalent is `chart.chartModel` (structured) + `chart.mutate(...)` (patching) + `chart.userShapesXml` (overlay bytes).

If you hit a gap in the mapping, open an issue with the POI snippet and we'll add the translation.
