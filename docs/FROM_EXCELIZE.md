# Migrating from excelize (Go)

[excelize](https://xuri.me/excelize/en/) is Go's mainline Excel library. Its chart API is a single `AddChart` call that takes a JSON-ish options struct — closer in spirit to ExcelTS than openpyxl, but with a smaller chart-type coverage (no ChartEx, no pivot chart metadata, partial combo support).

Companion guides: [`FROM_XLSXWRITER.md`](./FROM_XLSXWRITER.md) · [`FROM_OPENPYXL.md`](./FROM_OPENPYXL.md) · [`FROM_POI.md`](./FROM_POI.md) · [`FROM_EXCELJS.md`](./FROM_EXCELJS.md)

## Why migrate

- **ChartEx**: excelize has no `cx:` support; ExcelTS ships all 8 modern chart types.
- **Edit loaded charts**: excelize can add and delete charts but cannot mutate an existing chart's structured fields; ExcelTS exposes `chart.mutate(fn, { preferRawPatch })`.
- **Preview rendering**: excelize has no built-in renderer. ExcelTS ships zero-dependency SVG/PNG/PDF previews.
- **Runtime**: if your pipeline already calls Node for other Excel work, consolidating on ExcelTS removes the Go dependency.

## Cheat sheet — creation

| Task               | excelize (Go)                     | ExcelTS (TypeScript)                                                      |
| ------------------ | --------------------------------- | ------------------------------------------------------------------------- |
| Create a workbook  | `f := excelize.NewFile()`         | `const wb = new Workbook()`                                               |
| Add a worksheet    | `f.NewSheet("Sheet2")`            | `wb.addWorksheet("Sheet2")`                                               |
| Write a cell       | `f.SetCellValue(sheet, "A1", 42)` | `ws.getCell("A1").value = 42`                                             |
| Create a chart     | `f.AddChart(sheet, cell, opts)`   | `ws.addChart(opts, range)`                                                |
| Delete a chart     | `f.DeleteChart(sheet, cell)`      | `chart.remove()`                                                          |
| Save               | `f.SaveAs("out.xlsx")`            | `await wb.xlsx.writeFile("out.xlsx")`                                     |
| Save to memory     | `f.Write(buffer)`                 | `await wb.xlsx.writeBuffer()`                                             |
| Modern chart types | ✗                                 | `ws.addChartEx({ type: "sunburst" / "waterfall" / … })`                   |
| Render preview     | ✗                                 | `chart.toSVG()` / `chart.toPNG()` / `chartToPdf(chart)` (zero-dependency) |

## Example 1 — Basic column chart

**Go:**

```go
package main

import (
    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    for i, row := range [][]interface{}{
        {"Month", "Revenue", "Profit"},
        {"Jan", 100, 20},
        {"Feb", 180, 40},
        {"Mar", 240, 55},
    } {
        cell, _ := excelize.CoordinatesToCellName(1, i+1)
        f.SetSheetRow("Sheet1", cell, &row)
    }
    if err := f.AddChart("Sheet1", "E1", &excelize.Chart{
        Type: excelize.Col,
        Series: []excelize.ChartSeries{
            {Name: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$4", Values: "Sheet1!$B$2:$B$4"},
            {Name: "Sheet1!$C$1", Categories: "Sheet1!$A$2:$A$4", Values: "Sheet1!$C$2:$C$4"},
        },
        Title:  []excelize.RichTextRun{{Text: "Quarterly revenue"}},
        Legend: excelize.ChartLegend{Position: "bottom"},
    }); err != nil {
        panic(err)
    }
    f.SaveAs("out.xlsx")
}
```

**TypeScript:**

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.addRow(["Month", "Revenue", "Profit"]);
ws.addRow(["Jan", 100, 20]);
ws.addRow(["Feb", 180, 40]);
ws.addRow(["Mar", 240, 55]);

ws.addChart(
  {
    type: "bar",
    barDir: "col", // excelize's "Col" → OOXML barDir "col"
    series: [
      {
        name: { formula: "Sheet1!$B$1" },
        categories: "Sheet1!$A$2:$A$4",
        values: "Sheet1!$B$2:$B$4"
      },
      {
        name: { formula: "Sheet1!$C$1" },
        categories: "Sheet1!$A$2:$A$4",
        values: "Sheet1!$C$2:$C$4"
      }
    ],
    title: "Quarterly revenue",
    legend: { legendPos: "b" }
  },
  "E1:L15"
);

await wb.xlsx.writeFile("out.xlsx");
```

**Key translations**

- `excelize.Col` → `type: "bar", barDir: "col"` (OOXML nomenclature).
- `Series.Name` as `"Sheet!$B$1"` → `name: { formula: "Sheet1!$B$1" }` (cell-sourced). For literal strings: `name: "Revenue"`.
- `Legend.Position: "bottom"` → `legend: { legendPos: "b" }`. Values: `"t"`, `"b"`, `"l"`, `"r"`, `"tr"` (top-right).
- `Title []RichTextRun` → plain `title: "..."` for simple text, `title: { paragraphs: [...] }` for rich text with formatting.
- Chart anchor in excelize is a single cell; ExcelTS uses a range to size the chart (`"E1:L15"`).

## Example 2 — Line chart with markers

**Go:**

```go
f.AddChart("Sheet1", "E1", &excelize.Chart{
    Type: excelize.Line,
    Series: []excelize.ChartSeries{{
        Name:       "Revenue",
        Categories: "Sheet1!$A$2:$A$13",
        Values:     "Sheet1!$B$2:$B$13",
        Marker: excelize.ChartMarker{
            Symbol: "diamond",
            Size:   8,
        },
    }},
    PlotArea: excelize.ChartPlotArea{
        ShowBubbleSize:  false,
        ShowCatName:     false,
        ShowLeaderLines: false,
    },
})
```

**TypeScript:**

```typescript
ws.addChart(
  {
    type: "line",
    series: [
      {
        name: "Revenue",
        categories: "Sheet1!$A$2:$A$13",
        values: "Sheet1!$B$2:$B$13",
        marker: {
          symbol: "diamond",
          size: 8
        }
      }
    ]
  },
  "E1:M15"
);
```

Marker symbols accepted: `circle`, `dash`, `diamond`, `dot`, `none`, `picture`, `plus`, `square`, `star`, `triangle`, `x`, `auto`. Same list as excelize, with `none` being the "no marker" option (excelize uses an empty string).

## Example 3 — Combo chart (column + line)

excelize supports combos but requires the less ergonomic `ChartType: excelize.Combo` with the secondary chart nested in `PlotArea`. ExcelTS is explicit:

```typescript
ws.addComboChart(
  {
    groups: [
      {
        type: "bar",
        barDir: "col",
        series: [
          {
            name: "Count",
            categories: "S!$A$2:$A$6",
            values: "S!$B$2:$B$6"
          }
        ]
      },
      {
        type: "line",
        useSecondaryAxis: true,
        series: [
          {
            name: "Cumulative %",
            categories: "S!$A$2:$A$6",
            values: "S!$C$2:$C$6"
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

## Example 4 — Pie / Doughnut with per-point colours

**Go:**

```go
f.AddChart("Sheet1", "E1", &excelize.Chart{
    Type: excelize.Doughnut,
    Series: []excelize.ChartSeries{{
        Name:       "Revenue",
        Categories: "Sheet1!$A$2:$A$6",
        Values:     "Sheet1!$B$2:$B$6",
    }},
    HoleSize: 40,
})
```

excelize has no per-point colour override — you have to edit the raw XML. ExcelTS:

```typescript
ws.addChart(
  {
    type: "doughnut",
    holeSize: 40,
    series: [
      {
        name: "Revenue",
        categories: "Sheet1!$A$2:$A$6",
        values: "Sheet1!$B$2:$B$6",
        dataPoints: [
          { idx: 0, spPr: { solidFill: { srgbClr: "4472C4" } } },
          { idx: 1, spPr: { solidFill: { srgbClr: "ED7D31" } } },
          { idx: 2, spPr: { solidFill: { srgbClr: "A5A5A5" } } },
          { idx: 3, spPr: { solidFill: { srgbClr: "FFC000" } } },
          { idx: 4, spPr: { solidFill: { srgbClr: "5B9BD5" } } }
        ]
      }
    ]
  },
  "E1:L15"
);
```

`dataPoints` is an array of per-point overrides by `idx` — the same semantic as xlsxwriter's `points` list.

## Example 5 — Read existing chart and render

excelize 2.8+ can read chart _info_ but not mutate structured fields. ExcelTS:

```typescript
import { readFileSync } from "node:fs";

const wb = new Workbook();
await wb.xlsx.load(readFileSync("existing.xlsx"));

for (const ws of wb.worksheets) {
  for (const chart of ws.getCharts()) {
    console.log(
      `Chart in ${ws.name}: type=${chart.chartTypes[0]?.type ?? "chartEx"}`,
      `title=${chart.title}`
    );

    // Render a preview of each chart for a report
    const svg = chart.toSVG({ width: 500, height: 300 });
    writeFileSync(`${ws.name}-${chart.chartNumber}.svg`, svg);

    // Mutate — change the title in place, keeping all other bytes intact
    chart.mutate(
      model => {
        model.chart.title = {
          text: { paragraphs: [{ runs: [{ text: "Revised" }] }] }
        };
      },
      { preferRawPatch: true }
    );
  }
}

await wb.xlsx.writeBuffer();
```

## Example 6 — Modern chart types (not supported in excelize)

```typescript
// Sunburst
ws.addChartEx(
  {
    type: "sunburst",
    series: [
      {
        values: "Data!$C$1:$C$30",
        hierarchy: ["Data!$A$1:$A$30", "Data!$B$1:$B$30"]
      }
    ]
  },
  "D1:J12"
);

// Waterfall
ws.addChartEx(
  {
    type: "waterfall",
    categories: "Data!$A$1:$A$6",
    series: [{ values: "Data!$B$1:$B$6" }],
    layout: { subtotals: [{ idx: 3 }, { idx: 6 }], connectorLines: true }
  },
  "D1:J12"
);

// Region map with Albers projection
ws.addChartEx(
  {
    type: "regionMap",
    categories: "Data!$A$1:$A$50",
    series: [{ values: "Data!$B$1:$B$50" }],
    layout: { projection: "albers", regionLabels: "showAll" }
  },
  "D1:M25"
);
```

## Behavioural differences

- **Chart sizing.** excelize uses `Dimension{Width: 480, Height: 260}` in pixels; ExcelTS uses a range like `"E1:M20"` that follows the worksheet's column widths / row heights. For fixed pixel sizing, look at `chart.rangeFromOffset({ x, y, width, height })` in the Excel module API.
- **Data reference format.** excelize accepts `"Sheet1!$B$1"` and `[sheet, col, row]` forms; ExcelTS uses A1-only. Always sheet-qualified is safest.
- **Series name source.** excelize treats `Name` as a formula when it starts with `=` and as a literal otherwise. ExcelTS is explicit: `name: "literal"` or `name: { formula: "Sheet1!$B$1" }`.
- **Plot area styling.** excelize's `PlotArea` carries `ShowXxx` boolean toggles. ExcelTS folds these into each series' `dataLabels: { showCatName, showVal, ... }` + the chart-level `legend`, which matches the OOXML structure.

## Features ExcelTS has that excelize does not

- **ChartEx** (all 8 modern chart types).
- **Pivot chart metadata** (`addPivotChart`, `c14:pivotOptions`, `c16:pivotOptions16`).
- **Edit loaded charts** with byte-preserving raw patches.
- **Chartsheet** (single-chart sheet) with full `pageSetup` / `printOptions` / `state`.
- **Preview renderer** (SVG/PNG/PDF, zero-dependency).
- **User-shape overlays** (byte-preserving programmatic API).
- **Vendor-extension preservation** (`templateMode: "strict"` + `Chart.unknownElements`).
- **Browser support**.

## Features excelize has that are marked elsewhere

- **`AddPictureFromBytes` / `AddShape`** — ExcelTS' `ws.addImage({ buffer, extension }, range)` and `ws.addShape(...)` are the equivalents, not chart-related.
- **Sparklines** — ExcelTS has `ws.addSparkline({ ... })`; covered in the main README.

If you hit a gap in the mapping, open an issue with the excelize snippet and we'll add the translation.
