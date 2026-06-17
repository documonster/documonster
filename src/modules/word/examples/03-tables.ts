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

import { Document, Build, Io, Units } from "../index";
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
  Build.cell([Build.textParagraph(label, { run: { bold: true, color: "FFFFFF" } })], {
    shading: { fill: "1F4E79", pattern: "clear" },
    verticalAlign: "center"
  });
const stripedRows: TableRow[] = [
  Build.row(
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
    Build.row(
      r.map(v =>
        Build.cell(v, {
          shading: { fill: stripeFills[i % 2], pattern: "clear" },
          verticalAlign: "center"
        })
      )
    )
  );
}
Document.addTableElement(
  doc,
  Build.table(
    stripedRows,
    {
      width: { value: 5000, type: "pct" },
      borders: Build.gridBorders(4, "BFBFBF"),
      cellMargins: {
        top: { value: 80, type: "dxa" },
        bottom: { value: 80, type: "dxa" },
        left: { value: 100, type: "dxa" },
        right: { value: 100, type: "dxa" }
      }
    },
    [
      Units.cmToTwips(4),
      Units.cmToTwips(2.5),
      Units.cmToTwips(2.5),
      Units.cmToTwips(2.5),
      Units.cmToTwips(2.5)
    ]
  )
);

// ---------------------------------------------------------------------------
// 3. Mixed borders & "no borders" table
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Mixed borders / borderless", 2);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row([
        Build.cell("thick top", {
          borders: {
            top: Build.border("single", 24, "C00000"),
            bottom: Build.border("single", 4, "auto")
          }
        }),
        Build.cell("dashed", {
          borders: {
            top: Build.border("dashed", 8, "0070C0"),
            bottom: Build.border("dashed", 8, "0070C0")
          }
        }),
        Build.cell("double", {
          borders: {
            top: Build.border("double", 6, "70AD47"),
            bottom: Build.border("double", 6, "70AD47")
          }
        })
      ]),
      Build.row([
        Build.cell("none", { borders: { top: Build.border("nil") } }),
        Build.cell("none"),
        Build.cell("none")
      ])
    ],
    {
      width: { value: 5000, type: "pct" },
      borders: {
        top: Build.border("nil"),
        bottom: Build.border("nil"),
        left: Build.border("nil"),
        right: Build.border("nil")
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
  Build.table(
    [
      Build.row([Build.cell("Header spanning all 3 columns", { gridSpan: 3 })]),
      Build.row([Build.cell("A1"), Build.cell("B1"), Build.cell("C1")]),
      Build.row([Build.cell("A2 + B2 merged", { gridSpan: 2 }), Build.cell("C2")])
    ],
    { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 5. Vertical merge (vMerge restart / continue)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Vertical merge (vMerge)", 2);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row([
        Build.cell("Group A", { verticalMerge: "restart", verticalAlign: "center" }),
        Build.cell("Item A1"),
        Build.cell("100")
      ]),
      Build.row([
        Build.cell("", { verticalMerge: "continue" }),
        Build.cell("Item A2"),
        Build.cell("120")
      ]),
      Build.row([
        Build.cell("", { verticalMerge: "continue" }),
        Build.cell("Item A3"),
        Build.cell("140")
      ]),
      Build.row([
        Build.cell("Group B", { verticalMerge: "restart", verticalAlign: "center" }),
        Build.cell("Item B1"),
        Build.cell("50")
      ]),
      Build.row([
        Build.cell("", { verticalMerge: "continue" }),
        Build.cell("Item B2"),
        Build.cell("60")
      ])
    ],
    { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 6. Nested table inside a cell (and a cell with multiple paragraphs)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Nested table & multi-paragraph cell", 2);
const nested = Build.table(
  [
    Build.row([Build.cell("inner-1A"), Build.cell("inner-1B")]),
    Build.row([Build.cell("inner-2A"), Build.cell("inner-2B")])
  ],
  { width: { value: 5000, type: "pct" }, borders: Build.gridBorders(2, "808080") }
);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row([
        Build.cell("outer-1"),
        Build.cell([
          Build.textParagraph("Cell with multiple paragraphs:"),
          Build.paragraph([Build.bold("First "), Build.text("then second.")]),
          Build.textParagraph("…and a third one.")
        ])
      ]),
      Build.row([
        Build.cell("outer-2"),
        Build.cell([Build.textParagraph("Nested below ↓"), nested])
      ])
    ],
    { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 7. Vertical alignment & text direction
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Vertical alignment & text direction", 2);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row(
        [
          Build.cell("top", { verticalAlign: "top" }),
          Build.cell("center", { verticalAlign: "center" }),
          Build.cell("bottom", { verticalAlign: "bottom" }),
          Build.cell("rotated\nleft→top", { textDirection: "btLr", verticalAlign: "center" }),
          Build.cell("rotated\ntop→right", { textDirection: "tbRl", verticalAlign: "center" })
        ],
        { height: { value: Units.cmToTwips(3), rule: "atLeast" } }
      )
    ],
    { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() }
  )
);

// ---------------------------------------------------------------------------
// 8. Header repeats / cantSplit / explicit row height
// ---------------------------------------------------------------------------
Document.addHeading(doc, "8. Header repeat + cantSplit", 2);
const repeatRows: TableRow[] = [
  Build.row([headerCell("#"), headerCell("Value")], { tableHeader: true })
];
for (let i = 1; i <= 30; i++) {
  repeatRows.push(
    Build.row([Build.cell(`${i}`), Build.cell(`row ${i} content`)], {
      cantSplit: true,
      height: { value: Units.cmToTwips(0.6), rule: "atLeast" }
    })
  );
}
Document.addTableElement(
  doc,
  Build.table(repeatRows, { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() })
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
  Build.table(
    [
      Build.row([Build.cell("Float-1"), Build.cell("Float-2")]),
      Build.row([Build.cell("Float-3"), Build.cell("Float-4")])
    ],
    {
      width: { value: Units.cmToTwips(6), type: "dxa" },
      borders: Build.gridBorders(),
      float: {
        horizontalAnchor: "page",
        verticalAnchor: "text",
        leftFromText: 180,
        rightFromText: 180,
        topFromText: 0,
        bottomFromText: 0,
        absoluteHorizontalPosition: Units.cmToTwips(13),
        absoluteVerticalPosition: 0
      }
    }
  )
);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Empty cell — must still contain at least one (empty) paragraph
Document.addTableElement(
  doc,
  Build.table([Build.row([Build.cell("non-empty"), Build.cell([Build.textParagraph("")])])], {
    width: { value: 2000, type: "dxa" },
    borders: Build.gridBorders()
  })
);

// Single-cell single-row table
Document.addTableElement(
  doc,
  Build.table([Build.row([Build.cell("only cell")])], {
    width: { value: 5000, type: "pct" },
    borders: Build.gridBorders()
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
  Build.table(
    [
      Build.row([Build.cell("a"), Build.cell("b"), Build.cell("c")]),
      // Row 2 visually has one cell, padded by leaving 2 trailing grid units empty.
      Build.row([Build.cell("only-one")], {
        gridAfter: 2,
        widthAfter: { value: 2666, type: "dxa" }
      }),
      // Row 3 visually has two cells with the third grid unit blank.
      Build.row([Build.cell("x"), Build.cell("y")], {
        gridAfter: 1,
        widthAfter: { value: 1333, type: "dxa" }
      })
    ],
    { width: { value: 4000, type: "dxa" }, borders: Build.gridBorders() }
  )
);

// simpleTable() — the standalone builder Document.addTable wraps internally.
// Useful when you need the Table value without immediately attaching it to
// a doc (e.g. to nest it).
const standaloneTable = Build.simpleTable(
  [
    ["A", "B"],
    ["1", "2"]
  ],
  { headerRow: true, borders: true, columnWidths: [Units.cmToTwips(3), Units.cmToTwips(3)] }
);
Document.addTableElement(doc, standaloneTable);

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "03-tables.docx"), buf);
console.log(`  → 03-tables.docx (${buf.length} bytes)`);
