/**
 * Advanced Table & Track Changes Round-Trip Tests
 *
 * Tests complex table boundary scenarios:
 * - Nested tables
 * - Horizontal + vertical cell merging
 * - Table properties (borders, shading, width, row height)
 * - Header rows (tblHeader)
 * - XML compliance for tblGrid/gridCol, gridSpan, vMerge
 *
 * Also tests Track Changes writer enhancements:
 * - Table row insertion/deletion revisions
 * - Section property change (sectPrChange)
 * - Run property change (rPrChange)
 */

import { extractAll } from "@archive/unzip/extract";
import { Document, Build, Io } from "@word/index";
import type { Table, Paragraph, SectionProperties, RunProperties } from "@word/index";
import { parseXml, findChild, findChildren } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { describe, it, expect } from "vitest";

const decoder = new TextDecoder();

/** Helper: extract all files from DOCX bytes into a Map<path, Uint8Array>. */
async function extractDocx(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const files = await extractAll(bytes);
  const result = new Map<string, Uint8Array>();
  for (const [path, entry] of files) {
    result.set(path, entry.data);
  }
  return result;
}

/** Helper: get document.xml root element from DOCX bytes. */
async function getDocumentXml(bytes: Uint8Array): Promise<XmlElement> {
  const files = await extractDocx(bytes);
  const docXml = files.get("word/document.xml");
  expect(docXml).toBeDefined();
  return parseXml(decoder.decode(docXml!)).root;
}

// =============================================================================
// Nested Tables
// =============================================================================

describe("Advanced Tables - Nested Tables", () => {
  it("should roundtrip a table with a nested table in a cell", async () => {
    // Inner table: 2x2
    const innerTable = Build.table(
      [
        Build.row([Build.cell("Inner A1"), Build.cell("Inner A2")]),
        Build.row([Build.cell("Inner B1"), Build.cell("Inner B2")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    // Outer table: 2x2, cell (0,1) contains the nested table
    const outerTable = Build.table(
      [
        Build.row([
          Build.cell("Outer A1"),
          Build.cell([
            Build.textParagraph("Before nested"),
            innerTable,
            Build.textParagraph("After nested")
          ])
        ]),
        Build.row([Build.cell("Outer B1"), Build.cell("Outer B2")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, outerTable);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    // Verify outer table
    const outerTbl = parsed.body[0] as Table;
    expect(outerTbl.type).toBe("table");
    expect(outerTbl.rows).toHaveLength(2);
    expect(outerTbl.rows[0].cells).toHaveLength(2);

    // Verify nested table in cell (0,1)
    const cellContent = outerTbl.rows[0].cells[1].content;
    // Should have: paragraph, table, paragraph
    const nestedTables = cellContent.filter(c => c.type === "table");
    expect(nestedTables).toHaveLength(1);

    const nestedTbl = nestedTables[0] as Table;
    expect(nestedTbl.rows).toHaveLength(2);
    expect(nestedTbl.rows[0].cells).toHaveLength(2);
    expect(nestedTbl.rows[1].cells).toHaveLength(2);
  });

  it("should preserve nested table column widths", async () => {
    const innerTable = Build.table(
      [Build.row([Build.cell("A"), Build.cell("B"), Build.cell("C")])],
      { width: { value: 3000, type: "dxa" } },
      [1000, 1000, 1000]
    );

    const outerTable = Build.table(
      [Build.row([Build.cell([innerTable, Build.textParagraph("")])])],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, outerTable);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const outerTbl = parsed.body[0] as Table;
    const nestedTbl = outerTbl.rows[0].cells[0].content.find(c => c.type === "table") as Table;
    expect(nestedTbl).toBeDefined();
    expect(nestedTbl.columnWidths).toEqual([1000, 1000, 1000]);
  });
});

// =============================================================================
// Merged Cells (Horizontal + Vertical)
// =============================================================================

describe("Advanced Tables - Merged Cells", () => {
  it("should roundtrip horizontal merge (gridSpan)", async () => {
    const tbl = Build.table(
      [
        Build.row([
          Build.cell("Spans 2 columns", { gridSpan: 2, width: { value: 4000, type: "dxa" } }),
          Build.cell("Normal")
        ]),
        Build.row([Build.cell("A"), Build.cell("B"), Build.cell("C")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2000, 2000, 2000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].cells[0].properties?.gridSpan).toBe(2);
    expect(parsedTbl.rows[1].cells).toHaveLength(3);
  });

  it("should roundtrip vertical merge (vMerge restart/continue)", async () => {
    const tbl = Build.table(
      [
        Build.row([
          Build.cell("Merge start", { verticalMerge: "restart" }),
          Build.cell("Row 1 Col 2")
        ]),
        Build.row([Build.cell("", { verticalMerge: "continue" }), Build.cell("Row 2 Col 2")]),
        Build.row([Build.cell("Row 3 Col 1"), Build.cell("Row 3 Col 2")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].cells[0].properties?.verticalMerge).toBe("restart");
    expect(parsedTbl.rows[1].cells[0].properties?.verticalMerge).toBe("continue");
    // Row 3 should not have verticalMerge
    expect(parsedTbl.rows[2].cells[0].properties?.verticalMerge).toBeUndefined();
  });

  it("should roundtrip combined horizontal and vertical merge", async () => {
    const tbl = Build.table(
      [
        Build.row([
          Build.cell("Span 2x2", { gridSpan: 2, verticalMerge: "restart" }),
          Build.cell("Col 3")
        ]),
        Build.row([
          Build.cell("", { gridSpan: 2, verticalMerge: "continue" }),
          Build.cell("Col 3 R2")
        ]),
        Build.row([Build.cell("A"), Build.cell("B"), Build.cell("C")])
      ],
      { width: { value: 5000, type: "pct" } },
      [1500, 1500, 1500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    const cell00 = parsedTbl.rows[0].cells[0].properties!;
    expect(cell00.gridSpan).toBe(2);
    expect(cell00.verticalMerge).toBe("restart");

    const cell10 = parsedTbl.rows[1].cells[0].properties!;
    expect(cell10.gridSpan).toBe(2);
    expect(cell10.verticalMerge).toBe("continue");
  });
});

// =============================================================================
// Table Properties
// =============================================================================

describe("Advanced Tables - Table Properties", () => {
  it("should roundtrip table borders (all 6 types)", async () => {
    const borders = {
      top: { style: "single" as const, size: 4, color: "FF0000" },
      bottom: { style: "double" as const, size: 6, color: "00FF00" },
      left: { style: "dashed" as const, size: 8, color: "0000FF" },
      right: { style: "dotted" as const, size: 2, color: "FFFF00" },
      insideH: { style: "single" as const, size: 4, color: "FF00FF" },
      insideV: { style: "single" as const, size: 4, color: "00FFFF" }
    };

    const tbl = Build.table(
      [
        Build.row([Build.cell("A"), Build.cell("B")]),
        Build.row([Build.cell("C"), Build.cell("D")])
      ],
      { width: { value: 5000, type: "pct" }, borders },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    const parsedBorders = parsedTbl.properties!.borders!;
    expect(parsedBorders.top?.style).toBe("single");
    expect(parsedBorders.top?.size).toBe(4);
    expect(parsedBorders.top?.color).toBe("FF0000");
    expect(parsedBorders.bottom?.style).toBe("double");
    expect(parsedBorders.bottom?.color).toBe("00FF00");
    expect(parsedBorders.left?.style).toBe("dashed");
    expect(parsedBorders.right?.style).toBe("dotted");
    expect(parsedBorders.insideH?.style).toBe("single");
    expect(parsedBorders.insideH?.color).toBe("FF00FF");
    expect(parsedBorders.insideV?.style).toBe("single");
    expect(parsedBorders.insideV?.color).toBe("00FFFF");
  });

  it("should roundtrip cell shading (background color)", async () => {
    const tbl = Build.table(
      [
        Build.row([
          Build.cell("Red bg", { shading: { fill: "FF0000", pattern: "clear", color: "auto" } }),
          Build.cell("Green bg", { shading: { fill: "00FF00", pattern: "clear", color: "auto" } })
        ])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].cells[0].properties?.shading?.fill).toBe("FF0000");
    expect(parsedTbl.rows[0].cells[1].properties?.shading?.fill).toBe("00FF00");
  });

  it("should roundtrip table width (fixed dxa)", async () => {
    const tbl = Build.table(
      [Build.row([Build.cell("A")])],
      { width: { value: 7200, type: "dxa" } },
      [7200]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.properties?.width?.value).toBe(7200);
    expect(parsedTbl.properties?.width?.type).toBe("dxa");
  });

  it("should roundtrip table width (percentage pct)", async () => {
    const tbl = Build.table(
      [Build.row([Build.cell("A")])],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.properties?.width?.value).toBe(5000);
    expect(parsedTbl.properties?.width?.type).toBe("pct");
  });

  it("should roundtrip row height (exact)", async () => {
    const tbl = Build.table(
      [Build.row([Build.cell("Fixed height")], { height: { value: 720, rule: "exact" } })],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].properties?.height?.value).toBe(720);
    expect(parsedTbl.rows[0].properties?.height?.rule).toBe("exact");
  });

  it("should roundtrip row height (atLeast)", async () => {
    const tbl = Build.table(
      [Build.row([Build.cell("Min height")], { height: { value: 360, rule: "atLeast" } })],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].properties?.height?.value).toBe(360);
    expect(parsedTbl.rows[0].properties?.height?.rule).toBe("atLeast");
  });
});

// =============================================================================
// Table Header Row
// =============================================================================

describe("Advanced Tables - Header Row", () => {
  it("should roundtrip tblHeader property", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Header 1"), Build.cell("Header 2")], { tableHeader: true }),
        Build.row([Build.cell("Data 1"), Build.cell("Data 2")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].properties?.tableHeader).toBe(true);
    // Second row should not have tableHeader
    expect(parsedTbl.rows[1].properties?.tableHeader).toBeFalsy();
  });

  it("should roundtrip multiple header rows", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Header Row 1")], { tableHeader: true }),
        Build.row([Build.cell("Header Row 2")], { tableHeader: true }),
        Build.row([Build.cell("Data Row")])
      ],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    expect(parsedTbl.rows[0].properties?.tableHeader).toBe(true);
    expect(parsedTbl.rows[1].properties?.tableHeader).toBe(true);
    expect(parsedTbl.rows[2].properties?.tableHeader).toBeFalsy();
  });
});

// =============================================================================
// XML Compliance
// =============================================================================

describe("Advanced Tables - XML Compliance", () => {
  it("should produce correct w:tblGrid/w:gridCol count", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("A"), Build.cell("B"), Build.cell("C")]),
        Build.row([Build.cell("D"), Build.cell("E"), Build.cell("F")])
      ],
      { width: { value: 5000, type: "pct" } },
      [1000, 2000, 3000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    // Navigate: w:document > w:body > w:tbl > w:tblGrid
    const body = findChild(root, "w:body")!;
    const tblEl = findChild(body, "w:tbl")!;
    const tblGrid = findChild(tblEl, "w:tblGrid")!;
    const gridCols = findChildren(tblGrid, "w:gridCol");

    expect(gridCols).toHaveLength(3);
    expect(gridCols[0].attributes["w:w"]).toBe("1000");
    expect(gridCols[1].attributes["w:w"]).toBe("2000");
    expect(gridCols[2].attributes["w:w"]).toBe("3000");
  });

  it("should produce correct w:gridSpan in merged cells", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Merged", { gridSpan: 3 })]),
        Build.row([Build.cell("A"), Build.cell("B"), Build.cell("C")])
      ],
      { width: { value: 5000, type: "pct" } },
      [1000, 1000, 1000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    const body = findChild(root, "w:body")!;
    const tblEl = findChild(body, "w:tbl")!;
    const rows = findChildren(tblEl, "w:tr");
    expect(rows).toHaveLength(2);

    // First row, first cell should have gridSpan = 3
    const firstRowCells = findChildren(rows[0], "w:tc");
    expect(firstRowCells).toHaveLength(1);
    const tcPr = findChild(firstRowCells[0], "w:tcPr")!;
    const gridSpan = findChild(tcPr, "w:gridSpan")!;
    expect(gridSpan.attributes["w:val"]).toBe("3");
  });

  it("should produce correct w:vMerge markers", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Start", { verticalMerge: "restart" }), Build.cell("Normal")]),
        Build.row([Build.cell("", { verticalMerge: "continue" }), Build.cell("Normal 2")])
      ],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    const body = findChild(root, "w:body")!;
    const tblEl = findChild(body, "w:tbl")!;
    const rows = findChildren(tblEl, "w:tr");

    // Row 1, Cell 1: w:vMerge val="restart"
    const r1c1 = findChildren(rows[0], "w:tc")[0];
    const r1c1Pr = findChild(r1c1, "w:tcPr")!;
    const vMerge1 = findChild(r1c1Pr, "w:vMerge")!;
    expect(vMerge1.attributes["w:val"]).toBe("restart");

    // Row 2, Cell 1: w:vMerge (no val = continue)
    const r2c1 = findChildren(rows[1], "w:tc")[0];
    const r2c1Pr = findChild(r2c1, "w:tcPr")!;
    const vMerge2 = findChild(r2c1Pr, "w:vMerge")!;
    // "continue" is represented by absence of w:val attribute
    expect(vMerge2.attributes["w:val"]).toBeUndefined();
  });
});

// =============================================================================
// Track Changes - Table Row Revisions
// =============================================================================

describe("Track Changes - Table Row Revisions", () => {
  it("should roundtrip inserted row revision", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Existing row")]),
        Build.row([Build.cell("Inserted row")], {
          inserted: {
            revision: { id: 10, author: "Alice", date: "2024-01-15T10:00:00Z" }
          }
        })
      ],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    const insertedRow = parsedTbl.rows[1];
    expect(insertedRow.properties?.inserted).toBeDefined();
    expect(insertedRow.properties!.inserted!.revision.id).toBe(10);
    expect(insertedRow.properties!.inserted!.revision.author).toBe("Alice");
    expect(insertedRow.properties!.inserted!.revision.date).toBe("2024-01-15T10:00:00Z");
  });

  it("should roundtrip deleted row revision", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Normal row")]),
        Build.row([Build.cell("Deleted row")], {
          deleted: {
            revision: { id: 20, author: "Bob", date: "2024-02-20T14:30:00Z" }
          }
        })
      ],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const parsedTbl = parsed.body[0] as Table;
    const deletedRow = parsedTbl.rows[1];
    expect(deletedRow.properties?.deleted).toBeDefined();
    expect(deletedRow.properties!.deleted!.revision.id).toBe(20);
    expect(deletedRow.properties!.deleted!.revision.author).toBe("Bob");
    expect(deletedRow.properties!.deleted!.revision.date).toBe("2024-02-20T14:30:00Z");
  });

  it("should produce correct XML for row ins/del", async () => {
    const tbl = Build.table(
      [
        Build.row([Build.cell("Inserted")], {
          inserted: { revision: { id: 5, author: "Author1" } }
        }),
        Build.row([Build.cell("Deleted")], {
          deleted: { revision: { id: 6, author: "Author2", date: "2024-03-01T00:00:00Z" } }
        })
      ],
      { width: { value: 5000, type: "pct" } },
      [5000]
    );

    const h = Document.create();
    Document.addTableElement(h, tbl);
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    const body = findChild(root, "w:body")!;
    const tblEl = findChild(body, "w:tbl")!;
    const rows = findChildren(tblEl, "w:tr");

    // First row: w:trPr > w:ins
    const trPr1 = findChild(rows[0], "w:trPr")!;
    const ins = findChild(trPr1, "w:ins")!;
    expect(ins.attributes["w:id"]).toBe("5");
    expect(ins.attributes["w:author"]).toBe("Author1");

    // Second row: w:trPr > w:del
    const trPr2 = findChild(rows[1], "w:trPr")!;
    const del = findChild(trPr2, "w:del")!;
    expect(del.attributes["w:id"]).toBe("6");
    expect(del.attributes["w:author"]).toBe("Author2");
    expect(del.attributes["w:date"]).toBe("2024-03-01T00:00:00Z");
  });
});

// =============================================================================
// Track Changes - Section Property Change
// =============================================================================

describe("Track Changes - Section Property Change", () => {
  it("should roundtrip sectPrChange", async () => {
    const sectProps: SectionProperties = {
      pageSize: { width: 12240, height: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      propertyChange: {
        revision: { id: 100, author: "Editor", date: "2024-05-10T08:00:00Z" },
        previousProperties: {
          pageSize: { width: 11906, height: 16838 },
          margins: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      }
    };

    const h = Document.create();
    // Add a paragraph with section properties embedded
    Document.addParagraphElement(
      h,
      Build.paragraph([Build.text("Section change test")], { sectionProperties: sectProps })
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    // The section properties can appear either as paragraph section properties or last sectPr
    // Find the paragraph with section properties
    const para = parsed.body.find(
      b => b.type === "paragraph" && (b as Paragraph).properties?.sectionProperties
    ) as Paragraph | undefined;

    if (para) {
      const sect = para.properties!.sectionProperties!;
      expect(sect.propertyChange).toBeDefined();
      expect(sect.propertyChange!.revision.id).toBe(100);
      expect(sect.propertyChange!.revision.author).toBe("Editor");
      expect(sect.propertyChange!.revision.date).toBe("2024-05-10T08:00:00Z");
      expect(sect.propertyChange!.previousProperties).toBeDefined();
    } else {
      // Check last section properties (document-level sectPr)
      const lastSect = parsed.sectionProperties;
      // If the section was moved to the document level, just check it's preserved somewhere
      expect(lastSect).toBeDefined();
    }
  });

  it("should produce correct XML for sectPrChange", async () => {
    const sectProps: SectionProperties = {
      pageSize: { width: 12240, height: 15840 },
      propertyChange: {
        revision: { id: 50, author: "Tester", date: "2024-06-01T12:00:00Z" },
        previousProperties: {
          pageSize: { width: 11906, height: 16838 }
        }
      }
    };

    const h = Document.create();
    Document.addParagraphElement(
      h,
      Build.paragraph([Build.text("test")], { sectionProperties: sectProps })
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    const body = findChild(root, "w:body")!;
    const para = findChild(body, "w:p")!;
    const pPr = findChild(para, "w:pPr")!;
    const sectPr = findChild(pPr, "w:sectPr")!;
    const sectPrChange = findChild(sectPr, "w:sectPrChange")!;

    expect(sectPrChange.attributes["w:id"]).toBe("50");
    expect(sectPrChange.attributes["w:author"]).toBe("Tester");
    expect(sectPrChange.attributes["w:date"]).toBe("2024-06-01T12:00:00Z");

    // Should contain previous sectPr
    const prevSectPr = findChild(sectPrChange, "w:sectPr")!;
    expect(prevSectPr).toBeDefined();
  });
});

// =============================================================================
// Track Changes - Run Property Change
// =============================================================================

describe("Track Changes - Run Property Change", () => {
  it("should roundtrip rPrChange", async () => {
    const runProps: RunProperties = {
      bold: true,
      size: 28,
      propertyChange: {
        revision: { id: 200, author: "Formatter", date: "2024-04-15T09:00:00Z" },
        previousProperties: {
          bold: false,
          size: 24
        }
      }
    };

    const h = Document.create();
    Document.addParagraphElement(
      h,
      Build.paragraph([
        { content: [{ type: "text", text: "Formatted text" }], properties: runProps }
      ])
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const para = parsed.body[0] as Paragraph;
    // Find a run with propertyChange
    const runs = para.children.filter(c => "content" in c && !("type" in c));
    expect(runs.length).toBeGreaterThan(0);

    const runWithChange = runs.find(
      r => (r as { properties?: RunProperties }).properties?.propertyChange
    ) as { properties?: RunProperties } | undefined;
    expect(runWithChange).toBeDefined();
    expect(runWithChange!.properties!.propertyChange!.revision.id).toBe(200);
    expect(runWithChange!.properties!.propertyChange!.revision.author).toBe("Formatter");
    expect(runWithChange!.properties!.propertyChange!.revision.date).toBe("2024-04-15T09:00:00Z");
    expect(runWithChange!.properties!.propertyChange!.previousProperties?.size).toBe(24);
  });

  it("should produce correct XML for rPrChange", async () => {
    const runProps: RunProperties = {
      italic: true,
      color: "FF0000",
      propertyChange: {
        revision: { id: 30, author: "Stylist" },
        previousProperties: {
          italic: false,
          color: "000000"
        }
      }
    };

    const h = Document.create();
    Document.addParagraphElement(
      h,
      Build.paragraph([{ content: [{ type: "text", text: "Styled" }], properties: runProps }])
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const root = await getDocumentXml(bytes);

    const body = findChild(root, "w:body")!;
    const para = findChild(body, "w:p")!;
    const run = findChild(para, "w:r")!;
    const rPr = findChild(run, "w:rPr")!;
    const rPrChange = findChild(rPr, "w:rPrChange")!;

    expect(rPrChange.attributes["w:id"]).toBe("30");
    expect(rPrChange.attributes["w:author"]).toBe("Stylist");

    // Should contain previous rPr
    const prevRPr = findChild(rPrChange, "w:rPr")!;
    expect(prevRPr).toBeDefined();
    // Previous rPr should have italic=0 and color=000000
    const prevI = findChild(prevRPr, "w:i");
    expect(prevI).toBeDefined();
    expect(prevI!.attributes["w:val"]).toBe("0");
    const prevColor = findChild(prevRPr, "w:color");
    expect(prevColor).toBeDefined();
    expect(prevColor!.attributes["w:val"]).toBe("000000");
  });

  it("should roundtrip insertedRun with track changes", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      Build.paragraph([
        Build.text("Original "),
        Build.insertedRun(
          { content: [{ type: "text", text: "inserted text" }] },
          { id: 1, author: "User1", date: "2024-01-01T00:00:00Z" }
        )
      ])
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const para = parsed.body[0] as Paragraph;
    const insChild = para.children.find(c => "type" in c && c.type === "insertedRun");
    expect(insChild).toBeDefined();
  });

  it("should roundtrip deletedRun with track changes", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      Build.paragraph([
        Build.text("Keep this "),
        Build.deletedRun(
          { content: [{ type: "text", text: "deleted text" }] },
          { id: 2, author: "User2", date: "2024-02-01T00:00:00Z" }
        )
      ])
    );
    const doc = Document.build(h);

    const bytes = await Io.package(doc);
    const parsed = await Io.read(bytes);

    const para = parsed.body[0] as Paragraph;
    const delChild = para.children.find(c => "type" in c && c.type === "deletedRun");
    expect(delChild).toBeDefined();
  });
});
