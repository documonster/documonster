# Migrating from openpyxl (Python)

[openpyxl](https://openpyxl.readthedocs.io/en/stable/charts/introduction.html) is Python's class-based Excel library — `BarChart()`, `LineChart()`, `PieChart()` etc. are instantiated and composed with `Reference` objects pointing at cell ranges. Documonster adopts a flatter, options-object approach that maps cleanly from openpyxl's class graph.

Companion guides: [`FROM_XLSXWRITER.md`](./FROM_XLSXWRITER.md) · [`FROM_EXCELIZE.md`](./FROM_EXCELIZE.md) · [`FROM_POI.md`](./FROM_POI.md) · [`FROM_EXCELJS.md`](./FROM_EXCELJS.md)

## Philosophy

openpyxl models Excel's object graph directly — one class per OOXML element. Documonster keeps the underlying model (`ChartModel`, `ChartExModel`) but the **authoring surface** is an options object, similar to xlsxwriter's, so the common case (create a chart, add series, set title) is one `addChart({ … })` call instead of six `setAttr` calls on a chart instance.

| Aspect                  | openpyxl                                                                  | Documonster                                                                       |
| ----------------------- | ------------------------------------------------------------------------- | --------------------------------------------------------------------------------- |
| Chart instantiation     | `BarChart()`, `LineChart()`, …                                            | `ws.addChart({ type: "bar" / "line" / … })`                                       |
| Data sources            | `Reference(ws, min_col, min_row, max_col, max_row)` then `chart.add_data` | Excel-style A1 string: `"Sheet1!$B$2:$B$10"`                                      |
| Modern chart types      | ✗ (some raw-XML reading only)                                             | `ws.addChartEx({ type: "sunburst" / "waterfall" / … })`                           |
| Chart styles 1–48       | `chart.style = 37`                                                        | `chart.setStyle(37)`                                                              |
| Chart size              | `chart.width = 15; chart.height = 10` (cm)                                | Range argument: `"D1:J10"` (anchor-based)                                         |
| Rendering outside Excel | ✗ (none)                                                                  | `chart.toSVG()` / `chart.toPNG()` / `chartToPdf(chart)` (zero-dependency preview) |

## Example 1 — Simple bar chart

**Python:**

```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

wb = Workbook()
ws = wb.active
ws.title = "Sales"
for row in (
    ("Month", "Revenue", "Profit"),
    ("Jan", 100, 20),
    ("Feb", 180, 40),
    ("Mar", 240, 55),
):
    ws.append(row)

chart = BarChart()
chart.type = "col"             # "col" for column, "bar" for horizontal
chart.title = "Quarterly revenue"
chart.style = 10
chart.y_axis.title = "Amount (USD)"
chart.x_axis.title = "Month"

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=4)
cats = Reference(ws, min_col=1, min_row=2, max_col=1, max_row=4)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "E1")
wb.save("out.xlsx")
```

**TypeScript:**

```typescript
import { Workbook } from "documonster";

const wb = new Workbook();
const ws = wb.addWorksheet("Sales");
ws.addRow(["Month", "Revenue", "Profit"]);
ws.addRow(["Jan", 100, 20]);
ws.addRow(["Feb", 180, 40]);
ws.addRow(["Mar", 240, 55]);

ws.addChart(
  {
    type: "bar",
    barDir: "col",
    series: [
      { name: "Revenue", categories: "Sales!$A$2:$A$4", values: "Sales!$B$2:$B$4" },
      { name: "Profit", categories: "Sales!$A$2:$A$4", values: "Sales!$C$2:$C$4" }
    ],
    title: "Quarterly revenue",
    categoryAxis: { title: "Month" },
    valueAxis: { title: "Amount (USD)" }
  },
  "E1:L15"
);

const chart = ws.getCharts()[0];
chart.setStyle(10);

await wb.xlsx.writeFile("out.xlsx");
```

**Key translations**

- `chart.type = "col"` → `type: "bar", barDir: "col"` (OOXML nomenclature: the type is always `"bar"` for both orientations).
- `Reference(ws, min_col, min_row, max_col, max_row)` → `"Sheet!$A$1:$B$5"` A1-style string.
- `titles_from_data=True` → explicit `name: "Revenue"` / `name: { formula: "Sheet!$B$1" }` per series.
- `chart.y_axis.title = "..."` → `valueAxis: { title: "..." }`.
- `chart.x_axis.title = "..."` → `categoryAxis: { title: "..." }`.
- `ws.add_chart(chart, "E1")` → the range argument of `addChart` (e.g. `"E1:L15"`).

## Example 2 — Line chart with secondary axis

**Python:**

```python
from openpyxl.chart import LineChart, Reference

lc = LineChart()
lc.title = "Revenue vs margin"
data = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=13)
margin = Reference(ws, min_col=3, min_row=1, max_col=3, max_row=13)
lc.add_data(data, titles_from_data=True)
lc.add_data(margin, titles_from_data=True)

# Second series goes on the right axis
lc.y_axis.crosses = "autoZero"
lc.series[1].smooth = True
# ... second-axis config involves extra axis objects and is verbose

ws.add_chart(lc, "E1")
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
        useSecondaryAxis: true,
        smooth: true
      }
    ],
    title: "Revenue vs margin",
    secondaryValueAxis: { title: "%", max: 100 }
  },
  "E1:M15"
);
```

Documonster automatically inserts the secondary value axis when any series declares `useSecondaryAxis: true`, so there's no extra class graph to construct. The secondary-axis options (`secondaryValueAxis`, `secondaryCategoryAxis`) are siblings of `valueAxis` / `categoryAxis`.

## Example 3 — Pie chart with data labels

**Python:**

```python
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

pie = PieChart()
labels = Reference(ws, min_col=1, min_row=2, max_row=6)
data = Reference(ws, min_col=2, min_row=1, max_row=6)
pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.title = "Revenue by region"
pie.dataLabels = DataLabelList(showPercent=True, showCatName=True)
pie.style = 26
ws.add_chart(pie, "E1")
```

**TypeScript:**

```typescript
const chartNum = ws.addChart(
  {
    type: "pie",
    series: [
      {
        name: "Revenue",
        categories: "S!$A$2:$A$6",
        values: "S!$B$2:$B$6",
        dataLabels: {
          showPercent: true,
          showCatName: true,
          position: "outEnd" // openpyxl DataLabelList default is "bestFit"
        }
      }
    ],
    title: "Revenue by region"
  },
  "E1:L15"
);

const chart = ws.getCharts().find(c => c.chartNumber === chartNum)!;
chart.setStyle(26);
```

Data label flags map 1:1: `showCatName` / `showLegendKey` / `showPercent` / `showSerName` / `showVal` / `showBubbleSize`. The `position` values are the OOXML `DataLabelPosition` enum (`ctr`, `l`, `r`, `t`, `b`, `inBase`, `inEnd`, `outEnd`, `bestFit`, `inherit`, `inv`, `nextTo`).

## Example 4 — Scatter / bubble chart

**Python:**

```python
from openpyxl.chart import ScatterChart, Reference, Series

sc = ScatterChart()
sc.title = "Correlation"
sc.style = 13
sc.x_axis.title = "Predictor"
sc.y_axis.title = "Response"

xvalues = Reference(ws, min_col=1, min_row=2, max_row=100)
for i in range(2, 5):
    values = Reference(ws, min_col=i, min_row=2, max_row=100)
    series = Series(values, xvalues, title_from_data=False, title=f"Group {i - 1}")
    sc.series.append(series)

ws.add_chart(sc, "E1")
```

**TypeScript:**

```typescript
ws.addChart(
  {
    type: "scatter",
    scatterStyle: "lineMarker", // openpyxl's `style` 13 ≈ ScatterStyle.lineMarker
    series: [
      { name: "Group 1", xValues: "S!$A$2:$A$100", values: "S!$B$2:$B$100" },
      { name: "Group 2", xValues: "S!$A$2:$A$100", values: "S!$C$2:$C$100" },
      { name: "Group 3", xValues: "S!$A$2:$A$100", values: "S!$D$2:$D$100" }
    ],
    title: "Correlation",
    categoryAxis: { title: "Predictor" },
    valueAxis: { title: "Response" }
  },
  "E1:M20"
);
```

Documonster keeps xValues/values as separate fields rather than wrapping in a `Series` object. `scatterStyle` is the first-class axis: valid values are `lineMarker`, `line`, `marker`, `smooth`, `smoothMarker` — matching Excel's scatter sub-types.

## Example 5 — Loading a chart and editing it (openpyxl can't)

openpyxl can read _some_ loaded chart metadata (`ws._charts`) but cannot reliably mutate it without data loss. Documonster:

```typescript
await wb.xlsx.load(readFileSync("existing.xlsx"));
const chart = wb.getWorksheet("Data")!.getCharts()[0];

// Change the title
chart.title = "Revised title";

// Change a series values reference
chart.mutate(
  model => {
    const group = model.chart.plotArea.chartTypes[0];
    group.series[0].val = { formula: "Data!$B$2:$B$50", cache: undefined };
  },
  { preferRawPatch: true }
);

// Diagnose: what vendor extensions exist on this chart?
const unknown = chart.unknownElements;
if (unknown) {
  console.log(
    "Vendor extensions:",
    unknown.map(e => e.path)
  );
}

// Strict write — refuse to rebuild if that would drop unknown XML
await wb.xlsx.writeBuffer({ templateMode: "strict" });
```

If the loaded chart was authored by Excel 2013+ and carries `c15:datalabelsRange` or `c16:pivotOptions16` extensions, Documonster preserves them through the edit — openpyxl would silently lose them on re-serialisation.

## Example 6 — Modern chart types (openpyxl cannot create these)

```typescript
// Waterfall
ws.addChartEx(
  {
    type: "waterfall",
    categories: "Data!$A$1:$A$6",
    series: [{ values: "Data!$B$1:$B$6" }],
    layout: {
      subtotals: [{ idx: 3 }, { idx: 6 }],
      connectorLines: true
    }
  },
  "D1:J10"
);

// Sunburst
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
  "D1:J12"
);

// Histogram with manual bin count
ws.addChartEx(
  {
    type: "histogram",
    series: [{ values: "Data!$A$1:$A$1000" }],
    layout: { binning: { binType: "binCount", binCount: 20, intervalClosed: "r" } }
  },
  "D1:J12"
);

// Box & whisker with inclusive quartile method
ws.addChartEx(
  {
    type: "boxWhisker",
    categories: "Data!$A$1:$A$5",
    series: [{ values: "Data!$B$1:$B$5" }],
    layout: {
      quartileMethod: "inclusive",
      showMeanLine: true,
      showMeanMarker: true,
      showOutlierPoints: true
    }
  },
  "D1:J12"
);
```

See `docs/COMPATIBILITY.md` → "ChartEx types" for the complete type list with feature coverage.

## Behavioural differences

- **`Reference` vs A1 strings.** openpyxl's `Reference` lets you pass column/row integers; Documonster always uses Excel-style A1 notation. This is less terse but works directly with cells copy-pasted from the UI.
- **Default anchor.** openpyxl defaults to E15 + 15×7.5cm. Documonster makes the range explicit (`addChart(options, "E15:L30")`) — there's no implicit default.
- **No `chart.width` / `chart.height` cm properties.** Size is derived from the anchor range + worksheet column widths / row heights. To make a specific chart larger, pass a larger range.
- **Gauge / Sankey / radial-bar charts**: openpyxl supports these as _combined_ classic charts. Documonster builds the same via `addComboChart` with appropriate groups; see `src/modules/excel/__tests__/chart-builder.test.ts` for patterns.
- **Streaming writers.** openpyxl supports `WriteOnlyWorkbook`; Documonster has `WorkbookWriter` (see the main README → Streaming section). Chart support in both streaming modes is limited to adding pre-built models — dynamic editing during streaming is not supported in either library.

## Features Documonster has that openpyxl does not

- **Render preview**: SVG/PNG/PDF from the model directly (no Excel / LibreOffice roundtrip needed for previews).
- **ChartEx helper APIs**: `chartExOptionsFromTable` / `chartExOptionsFromRows`.
- **Pivot chart metadata**: `ws.addPivotChart(pivotTable, options, range)` — openpyxl reads only.
- **User-shape overlays**: byte-preserving programmatic API.
- **Vendor-extension preservation**: `templateMode: "strict"` + `Chart.unknownElements` catch silent loss.
- **Browser support**: Documonster runs in both Node and browsers; openpyxl is Python-only.

## Features openpyxl has that Documonster maps differently

- **`chart.legend.position`** → `legend: { legendPos: "r" | "l" | "t" | "b" | "tr" }` on the options object.
- **`chart.y_axis.crosses = "autoZero"`** → `valueAxis: { crosses: "autoZero" }`.
- **`chart.dataLabels.separator = "\n"`** → `dataLabels: { separator: "\n" }`.
- **`chart.add_data(..., from_rows=True)`** → transpose the data manually before passing to `addChart`. The A1 strings are direction-explicit, so "from rows" becomes the column arrangement you actually want.
- **`GaugeChart`** → construct via `addComboChart` with a doughnut + marker arrangement; see xlsxwriter's `Example: Gauge Chart` for the pattern (the OOXML is identical).

If you hit a gap in the mapping, open an issue with the openpyxl snippet and we'll add the translation.
