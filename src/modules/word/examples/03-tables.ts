/**
 * Word Example 03 — Tables (advanced)
 *
 * Covers:
 *   - Simple 2D-array tables (the convenience API)
 *   - Custom column widths, percent vs twips widths
 *   - Cell shading / striped tables
 *   - Per-cell borders, mixed border styles, "no borders"
 *   - Horizontal merge (gridSpan) & vertical merge (vMerge restart/continue)
 *   - Nested tables (a table inside a cell)
 *   - Cell with multiple paragraphs / cell with images would be in 07-images
 *   - Vertical alignment, text rotation
 *   - Repeating header row across pages
 *   - cantSplit (keep row together) and explicit row heights
 *   - Floating table (text wraps around it)
 *   - Edge case: empty table cell, single-cell table, super-wide table
 *
 * Output: tmp/word-examples/03-tables.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  text,
  bold,
  cell,
  row,
  table,
  simpleTable,
  border,
  gridBorders,
  cmToTwips,
  toBuffer
} from "../index";
import type { TableCell, TableRow } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word Module — Tables", 1);

// ---------------------------------------------------------------------------
// 1. Simple 2D-array convenience API (with header row)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Simple 2D-array table", 2);
Document.addTable(
  doc,
  [
    ["Product", "Price", "Qty"],
    ["Widget", "$10", "100"],
    ["Gadget", "$25", "50"]
  ],
  { headerRow: true, borders: true }
);

// ---------------------------------------------------------------------------
// 2. Custom column widths in twips, header shading, striped body
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Striped table with custom widths", 2);
const headerCell = (label: string): TableCell =>
  cell([textParagraph(label, { run: { bold: true, color: "FFFFFF" } })], {
    shading: { fill: "1F4E79", pattern: "clear" },
    verticalAlign: "center"
  });
const stripedRows: TableRow[] = [
  row(
    [headerCell("Region"), headerCell("Q1"), headerCell("Q2"), headerCell("Q3"), headerCell("Q4")],
    {
      tableHeader: true
    }
  )
];
const stripeFills = ["FFFFFF", "F2F2F2"];
for (const [i, r] of [
  ["North", "120", "140", "160", "190"],
  ["South", "90", "110", "130", "150"],
  ["East", "200", "210", "220", "240"],
  ["West", "75", "82", "95", "110"]
].entries()) {
  stripedRows.push(
    row(
      r.map(v =>
        cell(v, {
          shading: { fill: stripeFills[i % 2], pattern: "clear" },
          verticalAlign: "center"
        })
      )
    )
  );
}
Document.addTableElement(
  doc,
  table(
    stripedRows,
    {
      width: { value: 5000, type: "pct" },
      borders: gridBorders(4, "BFBFBF"),
      cellMargins: {
        top: { value: 80, type: "dxa" },
        bottom: { value: 80, type: "dxa" },
        left: { value: 100, type: "dxa" },
        right: { value: 100, type: "dxa" }
      }
    },
    [cmToTwips(4), cmToTwips(2.5), cmToTwips(2.5), cmToTwips(2.5), cmToTwips(2.5)]
  )
);

// ---------------------------------------------------------------------------
// 3. Mixed borders & "no borders" table
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Mixed borders / borderless", 2);
Document.addTableElement(
  doc,
  table(
    [
      row([
        cell("thick top", {
          borders: { top: border("single", 24, "C00000"), bottom: border("single", 4, "auto") }
        }),
        cell("dashed", {
          borders: { top: border("dashed", 8, "0070C0"), bottom: border("dashed", 8, "0070C0") }
        }),
        cell("double", {
          borders: { top: border("double", 6, "70AD47"), bottom: border("double", 6, "70AD47") }
        })
      ]),
      row([cell("none", { borders: { top: border("nil") } }), cell("none"), cell("none")])
    ],
    {
      width: { value: 5000, type: "pct" },
      borders: {
        top: border("nil"),
        bottom: border("nil"),
        left: border("nil"),
        right: border("nil")
      }
    }
  )
);

// ---------------------------------------------------------------------------
// 4. Horizontal merge (gridSpan)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Horizontal merge (gridSpan)", 2);
Document.addTableElement(
  doc,
  table(
    [
      row([cell("Header spanning all 3 columns", { gridSpan: 3 })]),
      row([cell("A1"), cell("B1"), cell("C1")]),
      row([cell("A2 + B2 merged", { gridSpan: 2 }), cell("C2")])
    ],
    { width: { value: 5000, type: "pct" }, borders: gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 5. Vertical merge (vMerge restart / continue)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Vertical merge (vMerge)", 2);
Document.addTableElement(
  doc,
  table(
    [
      row([
        cell("Group A", { verticalMerge: "restart", verticalAlign: "center" }),
        cell("Item A1"),
        cell("100")
      ]),
      row([cell("", { verticalMerge: "continue" }), cell("Item A2"), cell("120")]),
      row([cell("", { verticalMerge: "continue" }), cell("Item A3"), cell("140")]),
      row([
        cell("Group B", { verticalMerge: "restart", verticalAlign: "center" }),
        cell("Item B1"),
        cell("50")
      ]),
      row([cell("", { verticalMerge: "continue" }), cell("Item B2"), cell("60")])
    ],
    { width: { value: 5000, type: "pct" }, borders: gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 6. Nested table inside a cell (and a cell with multiple paragraphs)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Nested table & multi-paragraph cell", 2);
const nested = table(
  [row([cell("inner-1A"), cell("inner-1B")]), row([cell("inner-2A"), cell("inner-2B")])],
  { width: { value: 5000, type: "pct" }, borders: gridBorders(2, "808080") }
);
Document.addTableElement(
  doc,
  table(
    [
      row([
        cell("outer-1"),
        cell([
          textParagraph("Cell with multiple paragraphs:"),
          paragraph([bold("First "), text("then second.")]),
          textParagraph("…and a third one.")
        ])
      ]),
      row([cell("outer-2"), cell([textParagraph("Nested below ↓"), nested])])
    ],
    { width: { value: 5000, type: "pct" }, borders: gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 7. Vertical alignment & text direction
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Vertical alignment & text direction", 2);
Document.addTableElement(
  doc,
  table(
    [
      row(
        [
          cell("top", { verticalAlign: "top" }),
          cell("center", { verticalAlign: "center" }),
          cell("bottom", { verticalAlign: "bottom" }),
          cell("rotated\nleft→top", { textDirection: "btLr", verticalAlign: "center" }),
          cell("rotated\ntop→right", { textDirection: "tbRl", verticalAlign: "center" })
        ],
        { height: { value: cmToTwips(3), rule: "atLeast" } }
      )
    ],
    { width: { value: 5000, type: "pct" }, borders: gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 8. Header repeats / cantSplit / explicit row height
// ---------------------------------------------------------------------------
Document.addHeading(doc, "8. Header repeat + cantSplit", 2);
const repeatRows: TableRow[] = [row([headerCell("#"), headerCell("Value")], { tableHeader: true })];
for (let i = 1; i <= 30; i++) {
  repeatRows.push(
    row([cell(`${i}`), cell(`row ${i} content`)], {
      cantSplit: true,
      height: { value: cmToTwips(0.6), rule: "atLeast" }
    })
  );
}
Document.addTableElement(
  doc,
  table(repeatRows, { width: { value: 5000, type: "pct" }, borders: gridBorders() })
);

// ---------------------------------------------------------------------------
// 9. Floating table — text wraps around it
// ---------------------------------------------------------------------------
Document.addHeading(doc, "9. Floating table", 2);
Document.addParagraph(
  doc,
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit. The table on the right floats and the body text wraps. ".repeat(
    4
  )
);
Document.addTableElement(
  doc,
  table([row([cell("Float-1"), cell("Float-2")]), row([cell("Float-3"), cell("Float-4")])], {
    width: { value: cmToTwips(6), type: "dxa" },
    borders: gridBorders(),
    float: {
      horizontalAnchor: "page",
      verticalAnchor: "text",
      leftFromText: 180,
      rightFromText: 180,
      topFromText: 0,
      bottomFromText: 0,
      absoluteHorizontalPosition: cmToTwips(13),
      absoluteVerticalPosition: 0
    }
  })
);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Empty cell — must still contain at least one (empty) paragraph
Document.addTableElement(
  doc,
  table([row([cell("non-empty"), cell([textParagraph("")])])], {
    width: { value: 2000, type: "dxa" },
    borders: gridBorders()
  })
);

// Single-cell single-row table
Document.addTableElement(
  doc,
  table([row([cell("only cell")])], {
    width: { value: 5000, type: "pct" },
    borders: gridBorders()
  })
);

// Wide table (10 columns)
Document.addTable(
  doc,
  [
    Array.from({ length: 10 }, (_, i) => `H${i + 1}`),
    Array.from({ length: 10 }, (_, i) => `${i * 11}`),
    Array.from({ length: 10 }, (_, i) => `${(i + 1) * 7}`)
  ],
  { headerRow: true, borders: true }
);

// "Ragged" table — Word's strict OOXML schema requires every row to fill
// the full grid (sum of gridSpan + gridBefore + gridAfter must equal the
// tblGrid column count). To produce a *visually* ragged row we use
// gridBefore / gridAfter to leave grid units empty rather than dropping
// cells entirely.
Document.addTableElement(
  doc,
  table(
    [
      row([cell("a"), cell("b"), cell("c")]),
      // Row 2 visually has one cell, padded by leaving 2 trailing grid units empty.
      row([cell("only-one")], { gridAfter: 2, widthAfter: { value: 2666, type: "dxa" } }),
      // Row 3 visually has two cells with the third grid unit blank.
      row([cell("x"), cell("y")], { gridAfter: 1, widthAfter: { value: 1333, type: "dxa" } })
    ],
    { width: { value: 4000, type: "dxa" }, borders: gridBorders() }
  )
);

// simpleTable() — the standalone builder Document.addTable wraps internally.
// Useful when you need the Table value without immediately attaching it to
// a doc (e.g. to nest it).
const standaloneTable = simpleTable(
  [
    ["A", "B"],
    ["1", "2"]
  ],
  { headerRow: true, borders: true, columnWidths: [cmToTwips(3), cmToTwips(3)] }
);
Document.addTableElement(doc, standaloneTable);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "03-tables.docx"), buf);
console.log(`  → 03-tables.docx (${buf.length} bytes)`);
