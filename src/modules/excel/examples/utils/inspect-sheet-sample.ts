/**
 * The sample workbook `inspect-sheet.ts` describes when run without arguments, and the checks on that
 * description — so running the example verifies its output rather than only running it.
 *
 * Kept apart from the example so the recipe itself stays short enough to copy.
 */
import assert from "node:assert/strict";

import type { RegionInfo, SheetInfo } from "@excel/examples/inspect-sheet";
import { Cell, Table, Workbook, Worksheet } from "@excel/index";

/**
 * One of each case the description has to tell apart. Written to disk and read back by the caller, because
 * shared formulas are what a file loads as, not what an in-memory build necessarily holds.
 */
export async function buildSample(file: string): Promise<void> {
  const wb = Workbook.create();

  // Two tables stacked in the same columns, one with a totals row, under a title.
  const orders = Workbook.addWorksheet(wb, "Orders");
  Cell.setValue(orders, "A1", "Order book — Q3");
  Table.add(orders, {
    name: "Orders",
    ref: "A3",
    headerRow: true,
    totalsRow: true,
    columns: [
      { name: "Item", totalsRowLabel: "Total" },
      { name: "Qty", totalsRowFunction: "sum", totalsRowResult: 8 },
      { name: "Price" },
      { name: "Amount", totalsRowFunction: "sum", totalsRowResult: 22.5 }
    ],
    rows: [
      ["Widget", 3, 2.5, null],
      ["Gadget", 1, 10, null],
      ["Gizmo", 4, 1.25, null]
    ]
  });
  Worksheet.fillFormula(orders, "D4:D6", "B4*C4", [7.5, 10, 5]);
  Table.add(orders, {
    name: "Stock",
    ref: "A10",
    headerRow: true,
    columns: [{ name: "Item" }, { name: "On hand" }],
    rows: [
      ["Widget", 40],
      ["Gizmo", 7]
    ]
  });

  // An auto-filter below an export note.
  const filtered = Workbook.addWorksheet(wb, "Filtered");
  Worksheet.addRows(filtered, [
    ["Exported 2026-10-03"],
    [],
    ["Region", "Sales"],
    ["North", 120],
    ["South", 95]
  ]);
  Worksheet.setAutoFilter(filtered, "A3:B5");

  // Frozen rows above the data, and a column whose first data row is blank.
  const frozen = Workbook.addWorksheet(wb, "Frozen");
  Worksheet.addRows(frozen, [
    ["Monthly report"],
    ["Month", "Visits", "Notes"],
    ["Jan", 10, null],
    ["Feb", 12, "launch"]
  ]);
  Worksheet.freeze(frozen, 0, 2);

  // A header row that is only the default, a two-column array formula and a running total.
  const plain = Workbook.addWorksheet(wb, "Plain");
  Worksheet.addRows(plain, [
    ["x", "x²", "x³", "running"],
    [1, null, null, null],
    [2, null, null, null],
    [3, null, null, null]
  ]);
  Worksheet.fillFormula(
    plain,
    "B2:C4",
    "A2:A4^{2,3}",
    [
      [1, 1],
      [4, 8],
      [9, 27]
    ],
    "array"
  );
  Cell.setValue(plain, "D2", { formula: "A2", result: 1 });
  Cell.setValue(plain, "D3", { formula: "D2+A3", result: 3 });
  Cell.setValue(plain, "D4", { formula: "D3+A4", result: 6 });

  // A table directly under the remainder's header row, an array formula anchored *in* that header row, a
  // table with no header row, and a gap that makes the sheet sparse.
  const edges = Workbook.addWorksheet(wb, "Edges");
  Cell.setValue(edges, "A1", "Notes");
  Table.add(edges, {
    name: "Inner",
    ref: "A2",
    headerRow: true,
    columns: [{ name: "Key" }, { name: "Val" }],
    rows: [["k", 1]]
  });
  Cell.setValue(edges, "A6", "footnote");
  Worksheet.fillFormula(edges, "F1:F3", "ROW()", [1, 2, 3], "array");
  Table.add(edges, {
    name: "Headless",
    ref: "H5",
    headerRow: false,
    columns: [{ name: "P" }, { name: "Q" }],
    rows: [[1, 2]]
  });
  Cell.setValue(edges, "A40", "end");

  // One column per kind of value, so the typed first-data-row reading is pinned for each. H2 is merged
  // into G2 and reads as its master.
  const kinds = Workbook.addWorksheet(wb, "Kinds");
  Worksheet.addRow(kinds, ["check", "err", "divErr", "rich", "link", "uncached", "date", "merged"]);
  Cell.setValue(kinds, "A2", { checkbox: false });
  Cell.setValue(kinds, "B2", { error: "#N/A" });
  Cell.setValue(kinds, "C2", { formula: "1/0", result: { error: "#DIV/0!" } });
  Cell.setValue(kinds, "D2", { richText: [{ text: "ri" }, { text: "ch" }] });
  Cell.setValue(kinds, "E2", { text: "docs", hyperlink: "https://example.com" });
  Cell.setValue(kinds, "F2", { formula: "NOW()" });
  Cell.setValue(kinds, "G2", new Date(Date.UTC(2026, 9, 3)));
  Worksheet.merge(kinds, "G2:H2");

  await Workbook.writeFile(wb, file);
}

/** Pin the description of {@link buildSample}'s workbook. */
export function verifySample(sheets: readonly SheetInfo[]): void {
  const sheet = (name: string) => sheets.find(s => s.sheet === name)!;
  const region = (name: string, kind: RegionInfo["kind"], id?: string) =>
    sheet(name).regions.find(g => g.kind === kind && (id === undefined || g.name === id))!;
  const column = (g: RegionInfo, letter: string) => g.columns.find(c => c.letter === letter)!;

  // Two tables stacked in the same columns, each with its own header; the totals row is not data.
  const orders = region("Orders", "table", "Orders");
  assert.deepEqual(
    [orders.headerRow, orders.firstDataRow, orders.lastDataRow, orders.totalsRow],
    [3, 4, 6, 7]
  );
  assert.equal(column(orders, "D").total, "22.5");
  assert.deepEqual(column(orders, "D").formulas, { own: 1, shared: 2, array: 0 });
  const stock = region("Orders", "table", "Stock");
  assert.equal(stock.headerRow, 10);
  assert.equal(column(stock, "A").firstDataValue, "Widget");
  // The title above the first table is not a region of its own.
  assert.equal(sheet("Orders").regions.length, 2);

  // A typed value beside its display.
  const filtered = region("Filtered", "autoFilter");
  assert.equal(filtered.headerSource, "autoFilter");
  assert.equal(column(filtered, "B").firstDataValue, 120);
  assert.equal(column(filtered, "B").firstDataText, "120");

  // A guessed header, and a blank first data row with a value further down.
  const frozen = region("Frozen", "sheet");
  assert.deepEqual([frozen.headerRow, frozen.headerSource], [2, "frozenRows"]);
  assert.deepEqual(
    [column(frozen, "C").firstDataValue, column(frozen, "C").sample],
    [null, "launch"]
  );

  // A two-column array formula: both columns point at the one anchor.
  const plain = region("Plain", "sheet");
  assert.equal(plain.headerSource, "default");
  assert.deepEqual(column(plain, "C").firstFormula, {
    address: "C2",
    kind: "array",
    text: "A2:A4^{2,3}",
    anchor: "B2"
  });
  assert.deepEqual(column(plain, "D").formulas, { own: 3, shared: 0, array: 0 });

  const edges = sheet("Edges");
  const rest = region("Edges", "sheet");
  // A2 is the Inner table's header, not the remainder's first data row.
  assert.deepEqual([column(rest, "A").header, column(rest, "A").firstDataValue], ["Notes", null]);
  assert.equal(column(rest, "A").sample, "footnote");
  // Anchored in the skipped header row, the array is still found for the cells it covers.
  assert.deepEqual(column(rest, "F").formulas, { own: 0, shared: 0, array: 2 });
  assert.equal(column(rest, "F").firstFormula?.anchor, "F1");
  // Columns come from the definition when the table has no header row.
  const headless = region("Edges", "table", "Headless");
  assert.deepEqual([headless.headerRow, headless.firstDataRow], [null, 5]);
  assert.deepEqual(
    headless.columns.map(c => [c.letter, c.header, c.firstDataValue]),
    [
      ["H", "P", 1],
      ["I", "Q", 2]
    ]
  );
  // The last row in use, and how many rows hold anything.
  assert.deepEqual([edges.rowCount, edges.actualRowCount], [40, 6]);

  // Each kind of value keeps its type.
  const kinds = region("Kinds", "sheet");
  const first = (letter: string) => column(kinds, letter).firstDataValue;
  assert.equal(first("A"), false);
  assert.deepEqual(first("B"), { error: "#N/A" });
  assert.deepEqual(first("C"), { error: "#DIV/0!" });
  assert.equal(column(kinds, "C").firstDataText, "#DIV/0!");
  assert.equal(first("D"), "rich");
  assert.equal(first("E"), "docs");
  assert.equal(first("F"), null);
  assert.deepEqual(first("G"), new Date(Date.UTC(2026, 9, 3)));
  assert.deepEqual(first("H"), new Date(Date.UTC(2026, 9, 3)));
}
