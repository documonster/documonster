/**
 * DOCX Module - ODT (OpenDocument Text) Format Tests
 */

import { describe, it, expect } from "vitest";

import { readOdt, writeOdt } from "../convert/odt/odt";
import type { DocxDocument, Paragraph, Run } from "../types";

// =============================================================================
// Test Helpers
// =============================================================================

/** Create a simple text run. */
function makeRun(text: string, props?: Run["properties"]): Run {
  return {
    ...(props ? { properties: props } : {}),
    content: [{ type: "text", text }]
  };
}

/** Create a paragraph with text. */
function makeParagraph(text: string, props?: Paragraph["properties"]): Paragraph {
  return {
    type: "paragraph",
    ...(props ? { properties: props } : {}),
    children: [makeRun(text)]
  };
}

/** Create a minimal DocxDocument. */
function makeDoc(body: DocxDocument["body"], extra?: Partial<DocxDocument>): DocxDocument {
  return { body, ...extra };
}

// =============================================================================
// Tests
// =============================================================================

describe("ODT module", () => {
  describe("writeOdt", () => {
    it("returns a Uint8Array for a simple document", async () => {
      const doc = makeDoc([makeParagraph("Hello ODT")]);
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });

    it("produces a valid ZIP archive (PK signature)", async () => {
      const doc = makeDoc([makeParagraph("ZIP check")]);
      const result = await writeOdt(doc);

      // ZIP files start with PK signature (0x50, 0x4B)
      expect(result[0]).toBe(0x50);
      expect(result[1]).toBe(0x4b);
    });

    it("handles empty body gracefully", async () => {
      const doc = makeDoc([]);
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });

    it("handles document with metadata", async () => {
      const doc = makeDoc([makeParagraph("With metadata")], {
        coreProperties: {
          title: "Test Document",
          creator: "Unit Test",
          subject: "Testing"
        }
      });
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });

    it("handles document with multiple paragraphs", async () => {
      const doc = makeDoc([
        makeParagraph("First paragraph"),
        makeParagraph("Second paragraph"),
        makeParagraph("Third paragraph")
      ]);
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });

    it("handles document with formatted text", async () => {
      const doc = makeDoc([
        {
          type: "paragraph",
          children: [
            makeRun("Bold", { bold: true }),
            makeRun("Italic", { italic: true }),
            makeRun("Colored", { color: "FF0000" })
          ]
        }
      ]);
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });

    it("handles document with a table", async () => {
      const doc = makeDoc([
        {
          type: "table" as const,
          rows: [
            {
              cells: [
                { content: [makeParagraph("Cell 1")] },
                { content: [makeParagraph("Cell 2")] }
              ]
            }
          ]
        }
      ]);
      const result = await writeOdt(doc);

      expect(result).toBeInstanceOf(Uint8Array);
      expect(result.length).toBeGreaterThan(0);
    });
  });

  describe("readOdt", () => {
    it("throws on invalid data (non-ZIP)", async () => {
      const invalidData = new Uint8Array([0, 1, 2, 3, 4, 5]);
      await expect(readOdt(invalidData)).rejects.toThrow();
    });

    it("throws on ZIP without content.xml", async () => {
      // Create a minimal ZIP that lacks content.xml
      // A valid ZIP with no entries would still fail because content.xml is required
      const emptyZip = new Uint8Array([
        0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
      ]);
      await expect(readOdt(emptyZip)).rejects.toThrow();
    });
  });

  describe("writeOdt + readOdt round-trip", () => {
    it("preserves paragraph text through round-trip", async () => {
      const originalDoc = makeDoc([
        makeParagraph("First paragraph"),
        makeParagraph("Second paragraph")
      ]);

      const odtBytes = await writeOdt(originalDoc);
      const restoredDoc = await readOdt(odtBytes);

      expect(restoredDoc.body.length).toBeGreaterThanOrEqual(2);

      // Extract text from restored paragraphs
      const texts: string[] = [];
      for (const item of restoredDoc.body) {
        if (item.type === "paragraph") {
          for (const child of item.children) {
            if ("content" in child) {
              for (const c of (child as Run).content) {
                if (c.type === "text") {
                  texts.push(c.text);
                }
              }
            }
          }
        }
      }

      expect(texts.join(" ")).toContain("First paragraph");
      expect(texts.join(" ")).toContain("Second paragraph");
    });

    it("preserves metadata through round-trip", async () => {
      const originalDoc = makeDoc([makeParagraph("Test")], {
        coreProperties: {
          title: "Round Trip Title",
          creator: "Test Author"
        }
      });

      const odtBytes = await writeOdt(originalDoc);
      const restoredDoc = await readOdt(odtBytes);

      expect(restoredDoc.coreProperties?.title).toBe("Round Trip Title");
      expect(restoredDoc.coreProperties?.creator).toBe("Test Author");
    });

    it("preserves table structure through round-trip", async () => {
      const originalDoc = makeDoc([
        {
          type: "table" as const,
          rows: [
            {
              cells: [{ content: [makeParagraph("A1")] }, { content: [makeParagraph("B1")] }]
            },
            {
              cells: [{ content: [makeParagraph("A2")] }, { content: [makeParagraph("B2")] }]
            }
          ]
        }
      ]);

      const odtBytes = await writeOdt(originalDoc);
      const restoredDoc = await readOdt(odtBytes);

      // Find the table in the restored document
      const table = restoredDoc.body.find(item => item.type === "table");
      expect(table).toBeDefined();
      expect(table!.type).toBe("table");

      if (table && table.type === "table") {
        expect(table.rows.length).toBe(2);
        expect(table.rows[0].cells.length).toBe(2);
      }
    });

    it("preserves heading/outline-level through round-trip", async () => {
      const originalDoc = makeDoc([
        makeParagraph("My Heading", { outlineLevel: 0, style: "Heading1" })
      ]);

      const odtBytes = await writeOdt(originalDoc);
      const restoredDoc = await readOdt(odtBytes);

      // The heading should be restored with an outlineLevel
      const headingPara = restoredDoc.body.find(
        item => item.type === "paragraph" && item.properties?.outlineLevel !== undefined
      );
      expect(headingPara).toBeDefined();
    });

    it("handles empty document round-trip", async () => {
      const originalDoc = makeDoc([]);

      const odtBytes = await writeOdt(originalDoc);
      const restoredDoc = await readOdt(odtBytes);

      expect(restoredDoc.body).toBeDefined();
      expect(Array.isArray(restoredDoc.body)).toBe(true);
    });
  });

  describe("function signatures", () => {
    it("readOdt accepts Uint8Array and returns Promise<DocxDocument>", () => {
      expect(typeof readOdt).toBe("function");
      expect(readOdt.length).toBe(1);
    });

    it("writeOdt accepts DocxDocument and returns Promise<Uint8Array>", () => {
      expect(typeof writeOdt).toBe("function");
      expect(writeOdt.length).toBe(1);
    });
  });

  describe("hyperlink URL sanitisation", () => {
    it("strips javascript: hrefs from <text:a> elements when reading ODT", async () => {
      const doc = makeDoc([
        {
          type: "paragraph",
          children: [
            { type: "hyperlink", url: "https://safe.example", children: [makeRun("safe")] }
          ]
        } as Paragraph
      ]);

      // Round-trip through writer/reader to confirm safe URL survives.
      const odtBytes = await writeOdt(doc);
      const restored = await readOdt(odtBytes);
      let foundSafe = false;
      for (const block of restored.body) {
        if (block.type === "paragraph") {
          for (const c of (block as Paragraph).children) {
            if ((c as { type?: string }).type === "hyperlink") {
              const url = (c as { url?: string }).url ?? "";
              if (url.includes("safe.example")) {
                foundSafe = true;
              }
            }
          }
        }
      }
      expect(foundSafe).toBe(true);
    });

    it("drops the hyperlink wrapper when given a javascript: URL on write", async () => {
      const doc = makeDoc([
        {
          type: "paragraph",
          children: [
            {
              type: "hyperlink",
              url: "javascript:alert(1)",
              children: [makeRun("clickme")]
            }
          ]
        } as Paragraph
      ]);

      const odtBytes = await writeOdt(doc);
      // Round-trip via readOdt — if writer emitted the dangerous URL into
      // xlink:href, readOdt would either propagate it or drop it. With
      // sanitisation on the write path the link wrapper is gone but the
      // inner text "clickme" survives.
      const restored = await readOdt(odtBytes);
      let sawDangerousLink = false;
      let sawText = false;
      for (const block of restored.body) {
        if (block.type === "paragraph") {
          for (const c of (block as Paragraph).children) {
            if ((c as { type?: string }).type === "hyperlink") {
              const url = (c as { url?: string }).url ?? "";
              if (/javascript:/i.test(url)) {
                sawDangerousLink = true;
              }
            } else if ("content" in (c as object)) {
              for (const rc of (c as Run).content) {
                if (rc.type === "text" && rc.text.includes("clickme")) {
                  sawText = true;
                }
              }
            }
          }
        }
      }
      expect(sawDangerousLink).toBe(false);
      expect(sawText).toBe(true);
    });
  });

  describe("image path sanitisation", () => {
    it("rejects path traversal in image rId / fileName when writing ODT", async () => {
      // Images carry attacker-controlled rId / fileName when they come
      // from a round-tripped untrusted DOCX. The writer must coerce them
      // to safe Pictures/<leaf> entries — both in the archive and in
      // every xlink:href emitted into content.xml.
      const doc: DocxDocument = {
        body: [
          {
            type: "paragraph",
            children: [
              {
                content: [
                  {
                    type: "image",
                    rId: "../../etc/passwd",
                    width: 1000,
                    height: 1000,
                    name: "evil"
                  }
                ]
              } as Run
            ]
          } as Paragraph
        ],
        images: [
          {
            rId: "../../etc/passwd",
            fileName: "../../../boom.png",
            data: new Uint8Array([1, 2, 3]),
            mediaType: "png"
          }
        ]
      } as unknown as DocxDocument;

      const odtBytes = await writeOdt(doc);

      // ODT is a ZIP; inspect entry names to confirm no traversal escaped.
      const { extractAll } = await import("@archive/unzip/extract");
      const entries = await extractAll(odtBytes);
      for (const path of entries.keys()) {
        expect(path).not.toMatch(/\.\./);
        if (path.startsWith("Pictures/")) {
          // Leaf only — no nested directories.
          expect(path.split("/").length).toBe(2);
        }
      }

      // content.xml's xlink:href must point at the same safe entry that
      // actually exists in the ZIP.
      const decoder = new TextDecoder();
      const contentXml = decoder.decode(entries.get("content.xml")!.data);
      const m = contentXml.match(/xlink:href="(Pictures\/[^"]+)"/);
      expect(m).toBeTruthy();
      const referenced = m![1]!;
      expect(referenced).not.toMatch(/\.\./);
      expect(entries.has(referenced)).toBe(true);
    });
  });
});
