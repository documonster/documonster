/**
 * Integration test: comments subdirectory layout + author round-trip
 *
 * Verifies that excelts correctly reads comments from xlsx files using
 * the subdirectory layout (xl/comments/comment1.xml) with absolute rel
 * targets, as produced by tools like openpyxl.
 */

import { ZipArchive } from "@archive/zip";
import { describe, it, expect } from "vitest";

import { Workbook, Note } from "../../../index";

// =============================================================================
// Helpers
// =============================================================================

/** Build an XLSX buffer with comments from two different authors. */
async function buildCommentsXlsx(): Promise<Uint8Array> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");

  ws.getCell("A1").value = "Hello";
  ws.getCell("A1").comment = new Note({ texts: [{ text: "Comment by Alice" }] }, "Alice");

  ws.getCell("B2").value = "World";
  ws.getCell("B2").comment = new Note({ texts: [{ text: "Comment by Bob" }] }, "Bob");

  return wb.xlsx.writeBuffer();
}

/**
 * Repack an XLSX buffer to use subdirectory layout for comments:
 * - xl/comments1.xml -> xl/comments/comment1.xml
 * - worksheet rels use absolute targets (/xl/comments/comment1.xml)
 * - [Content_Types].xml references updated paths
 */
async function repackToSubdirLayout(buffer: Uint8Array): Promise<Uint8Array> {
  const { extractAll } = await import("@archive/unzip/extract");
  const entries = await extractAll(buffer);
  const archive = new ZipArchive({ level: 0, reproducible: true });

  const decoder = new TextDecoder();
  const encoder = new TextEncoder();

  for (const [name, entry] of entries) {
    // Rename xl/comments1.xml -> xl/comments/comment1.xml
    const commentsMatch = /^xl\/comments(\d+)\.xml$/.exec(name);
    if (commentsMatch) {
      const idx = commentsMatch[1];
      archive.add(`xl/comments/comment${idx}.xml`, entry.data);
      continue;
    }

    // Rewrite [Content_Types].xml
    if (name === "[Content_Types].xml") {
      let xml = decoder.decode(entry.data);
      xml = xml.replace(
        /PartName="\/xl\/comments(\d+)\.xml"/g,
        'PartName="/xl/comments/comment$1.xml"'
      );
      archive.add(name, encoder.encode(xml));
      continue;
    }

    // Rewrite worksheet rels to use absolute targets
    if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(name)) {
      let xml = decoder.decode(entry.data);
      xml = xml.replace(
        /Target="\.\.\/comments(\d+)\.xml"/g,
        'Target="/xl/comments/comment$1.xml"'
      );
      archive.add(name, encoder.encode(xml));
      continue;
    }

    archive.add(name, entry.data);
  }

  return archive.bytes();
}

// =============================================================================
// Tests
// =============================================================================

describe("Comments subdirectory layout", () => {
  it("round-trips comments and authors through flat layout", async () => {
    const buffer = await buildCommentsXlsx();
    const wb = new Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("Sheet1")!;

    const a1 = ws.getCell("A1");
    const b2 = ws.getCell("B2");

    expect(a1.note).toBeDefined();
    expect(b2.note).toBeDefined();
    expect(a1.comment?.author).toBe("Alice");
    expect(b2.comment?.author).toBe("Bob");
  });

  it("reads comments from subdirectory layout with absolute rel targets", async () => {
    const flatBuffer = await buildCommentsXlsx();
    const subdirBuffer = await repackToSubdirLayout(flatBuffer);

    const wb = new Workbook();
    await wb.xlsx.load(subdirBuffer);
    const ws = wb.getWorksheet("Sheet1")!;

    const a1 = ws.getCell("A1");
    const b2 = ws.getCell("B2");

    expect(a1.note).toBeDefined();
    expect(b2.note).toBeDefined();
    expect(a1.comment?.author).toBe("Alice");
    expect(b2.comment?.author).toBe("Bob");
  });

  it("preserves comment text content from subdirectory layout", async () => {
    const flatBuffer = await buildCommentsXlsx();
    const subdirBuffer = await repackToSubdirLayout(flatBuffer);

    const wb = new Workbook();
    await wb.xlsx.load(subdirBuffer);
    const ws = wb.getWorksheet("Sheet1")!;

    // Verify the actual text survived
    const a1Note = ws.getCell("A1").note;
    expect(a1Note).toContain("Comment by Alice");
  });

  it("round-trips comment without explicit author (default author)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Test";
    ws.getCell("A1").comment = new Note({ texts: [{ text: "No author set" }] });

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    expect(ws2.getCell("A1").note).toBeDefined();
    // Default author should be "Author"
    expect(ws2.getCell("A1").comment?.author).toBe("Author");
  });

  it("merges VML metadata correctly even when VML shape order differs from comments", async () => {
    const flatBuffer = await buildCommentsXlsx();

    // Repack: reverse the order of <v:shape> elements in the VML file
    const { extractAll } = await import("@archive/unzip/extract");
    const entries = await extractAll(flatBuffer);
    const archive = new ZipArchive({ level: 0, reproducible: true });

    const decoder = new TextDecoder();
    const encoder = new TextEncoder();

    for (const [name, entry] of entries) {
      if (/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(name)) {
        let vml = decoder.decode(entry.data);
        // Extract all <v:shape ...>...</v:shape> blocks and reverse them
        const shapes: string[] = [];
        const shapeRegex = /<v:shape\b[^]*?<\/v:shape>/g;
        let match: RegExpExecArray | null;
        while ((match = shapeRegex.exec(vml)) !== null) {
          shapes.push(match[0]);
        }
        if (shapes.length >= 2) {
          // Replace shapes in reverse order
          let idx = 0;
          const reversed = shapes.slice().reverse();
          vml = vml.replace(shapeRegex, () => reversed[idx++]);
        }
        archive.add(name, encoder.encode(vml));
        continue;
      }
      archive.add(name, entry.data);
    }

    const repackedBuffer = await archive.bytes();
    const wb = new Workbook();
    await wb.xlsx.load(repackedBuffer);
    const ws = wb.getWorksheet("Sheet1")!;

    // Despite reversed VML order, each cell should get its own note metadata
    const a1 = ws.getCell("A1");
    const b2 = ws.getCell("B2");

    expect(a1.note).toBeDefined();
    expect(b2.note).toBeDefined();
    expect(a1.comment?.author).toBe("Alice");
    expect(b2.comment?.author).toBe("Bob");

    // Verify VML metadata (editAs) is present on both — proves merge happened
    const a1Model = a1.comment?.model;
    const b2Model = b2.comment?.model;
    expect(a1Model?.note?.editAs).toBeDefined();
    expect(b2Model?.note?.editAs).toBeDefined();
  });
});
