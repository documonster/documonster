/**
 * Structural edits that keep the workbook consistent.
 *
 * Builds a sales sheet with a table, same-sheet and cross-sheet formulas, a filled (shared) formula, conditional
 * formatting, a data validation, a chart, a sparkline, a shape, a checkbox, a page break and defined names (one of
 * them sheet-scoped) — then inserts and deletes rows and columns, renames the sheet and removes another. Every reference follows the cells it named, as it would in Excel, so the file this writes
 * opens with the totals, highlights, drop-downs and chart all still pointing at the data.
 *
 * Run: pnpm example --filter structural-edits
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { writeBothFormats } from "@excel/examples/utils/write-both";
import {
  Cell,
  Chart,
  Column,
  ConditionalFormatting,
  DefinedNames,
  Form,
  Image,
  Row,
  Sparkline,
  Table,
  Workbook,
  Worksheet
} from "@excel/index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/excel-examples"
);
fs.mkdirSync(outDir, { recursive: true });
const filename = process.argv[2] ?? path.join(outDir, "structural-edits.xlsx");

const wb = Workbook.create();
const sales = Workbook.addWorksheet(wb, "Sales");
const scratch = Workbook.addWorksheet(wb, "Scratch");
const summary = Workbook.addWorksheet(wb, "Summary");

// --- the data: a title, a table at A3, and formulas around it -------------------------------------------------
Cell.setValue(sales, "A1", "Quarterly sales");
Table.add(sales, {
  name: "Orders",
  ref: "A3",
  totalsRow: true,
  style: { theme: "TableStyleMedium9", showRowStripes: true },
  columns: [
    { name: "Region", totalsRowLabel: "Total" },
    { name: "Q1", totalsRowFunction: "sum" },
    { name: "Q2", totalsRowFunction: "sum" }
  ],
  rows: [
    ["North", 120, 135],
    ["South", 90, 110],
    ["East", 150, 140],
    ["West", 80, 95]
  ]
});
// D4:D7 — growth per region, filled as one shared formula.
Cell.setValue(sales, "D3", "Growth");
Worksheet.fillFormula(sales, "D4:D7", "C4/B4-1");
Column.setNumFmt(sales, "D", "0.0%");
// F1 / F2 — a same-sheet sum, and a lookup of the cell above the totals row.
Cell.setValue(sales, "F1", { formula: "SUM(B4:C7)" });
Cell.setValue(sales, "F2", { formula: "MAX($B$4:$B$7)" });

// A sparkline per region, a highlight for strong quarters, and a drop-down of regions.
Sparkline.add(sales, {
  type: "column",
  sparklines: [4, 5, 6, 7].map(row => ({ dataRef: `Sales!B${row}:C${row}`, cellRef: `E${row}` }))
});
ConditionalFormatting.add(sales, {
  ref: "B4:C7",
  rules: [
    {
      type: "expression",
      formulae: ["B4>=$F$2"],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFC6EFCE" } } }
    }
  ]
});
Cell.setValue(sales, "H3", "Pick a region");
Cell.setValidation(sales, "H4", {
  type: "list",
  allowBlank: true,
  formulae: ["$A$4:$A$7"]
});

// Names: a reference name and a formula name.
const names = Workbook.getDefinedNames(wb);
DefinedNames.add(names, "Sales!$B$4:$C$7", "Quarters");
DefinedNames.addFormula(names, "BestQ1", "MAX(Sales!$B$4:$B$7)");

// A shape, a checkbox and a page break on the data sheet. The shape is placed by a range, so it moves with its cells
// but keeps its size; the checkbox is written "don't move or size with cells", so it stays put while the cell it is
// linked to moves; the break below the totals row moves with the row after it.
Image.addShape(sales, {
  type: "roundRect",
  range: "J3:L6",
  fillColor: "FFF2CC",
  lineColor: "BF9000",
  text: "Reviewed by finance"
});
Cell.setValue(sales, "H6", "Reviewed?");
Form.addCheckbox(sales, "H7", { link: "I7", text: "Reviewed" });
Row.addPageBreak(sales, 8);

// --- the summary sheet reads the data from across the workbook --------------------------------------------------
Row.addMany(summary, [
  ["Total Q1", { formula: "Sales!B8" }],
  ["Total Q2", { formula: "Sales!C8" }],
  ["All quarters", { formula: "SUM(Quarters)" }],
  ["Best Q1", { formula: "BestQ1" }],
  ["North Q2", { formula: "Sales!$C$4" }]
]);
Column.setWidth(summary, "A", 16);
Chart.place(
  summary,
  { tl: { col: 3, row: 1 }, br: { col: 11, row: 18 } },
  {
    type: "bar",
    barDir: "col",
    title: "Sales by region",
    series: [
      { name: "Q1", categories: "Sales!$A$4:$A$7", values: "Sales!$B$4:$B$7" },
      { name: "Q2", categories: "Sales!$A$4:$A$7", values: "Sales!$C$4:$C$7" }
    ]
  }
);

// Sheet-scoped names: one on Summary (tab 2) and one on Scratch (tab 1), which is removed below. Removing Scratch
// removes its name and renumbers Summary's to the tab Summary then occupies.
Cell.setValue(scratch, "A1", "temporary");
DefinedNames.setModel(names, [
  ...DefinedNames.getModel(names),
  { name: "LocalTotal", ranges: ["Summary!$B$1:$B$2"], localSheetId: 2 },
  { name: "ScratchCell", ranges: ["Scratch!$A$1"], localSheetId: 1 }
]);
Cell.setValue(summary, "A7", "Both quarters");
Cell.setValue(summary, "B7", { formula: "SUM(LocalTotal)" });

// --- the edits --------------------------------------------------------------------------------------------------
// 1. A subtitle row above the table: everything below moves down one.
Row.insert(sales, 2, ["Figures in thousands"]);
// 2. A new region inside the table (between East and West): the table, its totals, the sums, the formats, the
//    validation list, the names and the chart all grow over it.
Row.insert(sales, 8, ["Central", 100, 125]);
// 3. A column for a code, before the data: everything moves right one column.
Column.splice(sales, 1, 0, [null, null, null, "Code", "N", "S", "E", "C", "W"]);
// 4. Drop the South row: ranges over it shrink, and the shared growth formula is restated where it must be.
Row.splice(sales, 6, 1);
// 5. A clearer name for the sheet: every formula, name and series that names it follows.
Worksheet.setName(sales, "Sales 2026");
// 6. Drop the scratch sheet: its sheet-scoped name goes with it, and Summary's name follows Summary to tab 1.
Workbook.removeWorksheet(wb, "Scratch");

// --- what the references became ------------------------------------------------------------------------------
const orders = Table.get(sales, "Orders")!;
console.log("Table Orders:          ", Table.getModel(orders).tableRef);
console.log("Sales 2026!G1:         ", Cell.getFormula(sales, "G1"));
console.log("Summary!B1:            ", Cell.getFormula(summary, "B1"));
console.log("Summary!B5:            ", Cell.getFormula(summary, "B5"));
console.log(
  "Quarters:              ",
  DefinedNames.list(names).find(n => n.name === "Quarters")?.ranges
);
console.log(
  "BestQ1:                ",
  DefinedNames.list(names).find(n => n.name === "BestQ1")?.ranges
);
console.log("Highlight applies to:  ", ConditionalFormatting.list(sales)[0]?.ref);
console.log("Region list:           ", Cell.getValidation(sales, "I5"));
const shape = Worksheet.getModel(sales).shapes?.[0]?.anchorRange;
console.log(
  "Shape anchor (0-based):",
  shape && [shape.tl.nativeCol, shape.tl.nativeRow, shape.br?.nativeCol, shape.br?.nativeRow]
);
console.log("Checkbox link:         ", Form.list(sales)[0]?.model.link);
console.log(
  "Page breaks below rows:",
  sales.rowBreaks.map(b => b.id)
);
console.log(
  "Scoped names:          ",
  DefinedNames.list(names)
    .filter(n => n.localSheetId !== undefined)
    .map(n => `${n.name}@${n.localSheetId}`)
);
const series = Chart.getSeries(Chart.list(summary)[0], 0) as {
  val?: { numRef?: { formula?: string } };
};
console.log("Chart series Q1:       ", series.val?.numRef?.formula);

await writeBothFormats(wb, filename);
console.log("Wrote", filename);
