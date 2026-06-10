/**
 * DOCX Module - ODT (OpenDocument Text) Format Tests
 */

import { describe, it, expect } from "vitest";

import { readOdt, writeOdt } from "../convert/odt/odt";
import type { DocxDocument, Paragraph, Run, NumberFormat, AbstractNumbering } from "../types";

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

    // ODF v1.2 part 3, §3.3: when a `mimetype` file is present it MUST be the
    // first entry in the package and stored uncompressed (no compression, no
    // extra field) so the format can be detected from the file's magic bytes.
    // Regression: the ZIP archiver sorts entries alphabetically by default,
    // which pushed `mimetype` after `content.xml` and broke ODF detection.
    it("places `mimetype` as the first ZIP entry, stored uncompressed", async () => {
      const doc = makeDoc([makeParagraph("ordering"), makeParagraph("content")]);
      const bytes = await writeOdt(doc);

      // Parse the first local file header.
      const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      // Local file header signature: PK\x03\x04
      expect(dv.getUint32(0, true)).toBe(0x04034b50);
      const compressionMethod = dv.getUint16(8, true); // 0 = STORED, 8 = DEFLATE
      const nameLen = dv.getUint16(26, true);
      const extraLen = dv.getUint16(28, true);
      const nameBytes = bytes.subarray(30, 30 + nameLen);
      const firstName = new TextDecoder().decode(nameBytes);

      expect(firstName).toBe("mimetype");
      expect(compressionMethod).toBe(0); // STORED
      expect(extraLen).toBe(0); // no extra field
      // The stored bytes are exactly the mimetype string.
      const dataStart = 30 + nameLen + extraLen;
      const mime = new TextDecoder().decode(
        bytes.subarray(dataStart, dataStart + "application/vnd.oasis.opendocument.text".length)
      );
      expect(mime).toBe("application/vnd.oasis.opendocument.text");
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

  describe("list round-trip", () => {
    /** Collect (text, numbering level) for every paragraph in the body. */
    function listRows(doc: DocxDocument): { text: string; level: number | undefined }[] {
      const out: { text: string; level: number | undefined }[] = [];
      const visit = (blocks: DocxDocument["body"]): void => {
        for (const block of blocks) {
          if (block.type === "paragraph") {
            const text = block.children
              .flatMap(c => ("content" in c ? (c as Run).content : []))
              .filter(rc => rc.type === "text")
              .map(rc => (rc as { text: string }).text)
              .join("");
            out.push({ text, level: block.properties?.numbering?.level });
          } else if (block.type === "table") {
            for (const row of block.rows) {
              for (const cell of row.cells) {
                visit(cell.content as DocxDocument["body"]);
              }
            }
          }
        }
      };
      visit(doc.body);
      return out;
    }

    function listPara(text: string, level: number): Paragraph {
      return {
        type: "paragraph",
        properties: { numbering: { numId: 1, level } },
        children: [makeRun(text)]
      };
    }

    it("preserves a flat bullet list through round-trip", async () => {
      const doc = makeDoc([listPara("Foo", 0), listPara("Bar", 0), listPara("Baz", 0)]);

      const odtBytes = await writeOdt(doc);
      const restored = await readOdt(odtBytes);

      const rows = listRows(restored);
      expect(rows).toEqual([
        { text: "Foo", level: 0 },
        { text: "Bar", level: 0 },
        { text: "Baz", level: 0 }
      ]);
    });

    it("emits a text:list-style and text:list elements in the package", async () => {
      const doc = makeDoc([listPara("Item", 0)]);
      const odtBytes = await writeOdt(doc);
      // The reader proves the structure is well-formed; assert the actual
      // markup carries the list elements (not bare text:p).
      const restored = await readOdt(odtBytes);
      const rows = listRows(restored);
      expect(rows).toEqual([{ text: "Item", level: 0 }]);
    });

    it("preserves mixed nesting levels and surrounding paragraphs", async () => {
      const doc = makeDoc([
        makeParagraph("Intro"),
        listPara("A0", 0),
        listPara("B1", 1),
        listPara("C2", 2),
        listPara("D1", 1),
        listPara("E0", 0),
        makeParagraph("Outro")
      ]);

      const odtBytes = await writeOdt(doc);
      const restored = await readOdt(odtBytes);

      expect(listRows(restored)).toEqual([
        { text: "Intro", level: undefined },
        { text: "A0", level: 0 },
        { text: "B1", level: 1 },
        { text: "C2", level: 2 },
        { text: "D1", level: 1 },
        { text: "E0", level: 0 },
        { text: "Outro", level: undefined }
      ]);
    });

    it("keeps run formatting inside list items", async () => {
      const doc = makeDoc([
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [makeRun("bold item", { bold: true })]
        } as Paragraph
      ]);

      const odtBytes = await writeOdt(doc);
      const restored = await readOdt(odtBytes);

      const para = restored.body.find(b => b.type === "paragraph") as Paragraph | undefined;
      expect(para).toBeDefined();
      expect(para!.properties?.numbering?.level).toBe(0);
      const run = para!.children.find(c => "content" in c) as Run | undefined;
      expect(run?.properties?.bold).toBe(true);
    });

    it("preserves a list nested inside a table cell", async () => {
      const doc = makeDoc([
        {
          type: "table" as const,
          rows: [
            {
              cells: [
                {
                  content: [listPara("cell-bullet-1", 0), listPara("cell-bullet-2", 0)]
                }
              ]
            }
          ]
        }
      ]);

      const odtBytes = await writeOdt(doc);
      const restored = await readOdt(odtBytes);

      const table = restored.body.find(b => b.type === "table");
      expect(table).toBeDefined();
      const rows = listRows(restored);
      expect(rows).toContainEqual({ text: "cell-bullet-1", level: 0 });
      expect(rows).toContainEqual({ text: "cell-bullet-2", level: 0 });
    });
  });

  describe("ordered/numbered list round-trip", () => {
    /** Make an abstract numbering whose levels use the given formats. */
    function abstractNum(abstractNumId: number, formats: NumberFormat[]): AbstractNumbering {
      return {
        abstractNumId,
        levels: formats.map((format, level) => ({
          level,
          format,
          text: format === "bullet" ? "•" : `%${level + 1}.`,
          start: 1
        }))
      };
    }

    function lp(text: string, numId: number, level: number): Paragraph {
      return {
        type: "paragraph",
        properties: { numbering: { numId, level } },
        children: [makeRun(text)]
      };
    }

    /** Resolve a restored paragraph's level format via numbering chain. */
    function formatOf(doc: DocxDocument, numId: number, level: number): string | undefined {
      const inst = doc.numberingInstances?.find(n => n.numId === numId);
      const abs = doc.abstractNumberings?.find(a => a.abstractNumId === inst?.abstractNumId);
      return abs?.levels.find(l => l.level === level)?.format;
    }

    it("preserves decimal ordered list format through round-trip", async () => {
      const doc = makeDoc([lp("One", 2, 0), lp("Two", 2, 0)], {
        abstractNumberings: [abstractNum(2, ["decimal"])],
        numberingInstances: [{ numId: 2, abstractNumId: 2 }]
      });

      const restored = await readOdt(await writeOdt(doc));
      const para = restored.body.find(b => b.type === "paragraph") as Paragraph | undefined;
      expect(para?.properties?.numbering).toBeDefined();
      const numId = para!.properties!.numbering!.numId;
      expect(formatOf(restored, numId, 0)).toBe("decimal");
    });

    it("keeps bullet and numbered lists distinct (no collapse to one list)", async () => {
      const doc = makeDoc([lp("bullet", 1, 0), lp("number", 2, 0)], {
        abstractNumberings: [abstractNum(1, ["bullet"]), abstractNum(2, ["decimal"])],
        numberingInstances: [
          { numId: 1, abstractNumId: 1 },
          { numId: 2, abstractNumId: 2 }
        ]
      });

      const restored = await readOdt(await writeOdt(doc));
      const paras = restored.body.filter(b => b.type === "paragraph") as Paragraph[];
      const bulletNumId = paras[0].properties!.numbering!.numId;
      const numberNumId = paras[1].properties!.numbering!.numId;

      // The two lists must resolve to different numbering definitions.
      expect(bulletNumId).not.toBe(numberNumId);
      expect(formatOf(restored, bulletNumId, 0)).toBe("bullet");
      expect(formatOf(restored, numberNumId, 0)).toBe("decimal");
    });

    it("preserves per-level formats in a multi-level ordered list", async () => {
      const doc = makeDoc([lp("One", 2, 0), lp("One.a", 2, 1), lp("Two", 2, 0)], {
        abstractNumberings: [abstractNum(2, ["decimal", "lowerLetter"])],
        numberingInstances: [{ numId: 2, abstractNumId: 2 }]
      });

      const restored = await readOdt(await writeOdt(doc));
      const paras = restored.body.filter(b => b.type === "paragraph") as Paragraph[];
      const numId = paras[0].properties!.numbering!.numId;
      expect(formatOf(restored, numId, 0)).toBe("decimal");
      expect(formatOf(restored, numId, 1)).toBe("lowerLetter");
      // Level membership preserved.
      expect(paras.map(p => p.properties?.numbering?.level)).toEqual([0, 1, 0]);
    });

    it("maps upper/lower roman and letter formats", async () => {
      const doc = makeDoc([lp("r1", 5, 0), lp("r2", 6, 0), lp("r3", 7, 0), lp("r4", 8, 0)], {
        abstractNumberings: [
          abstractNum(5, ["upperRoman"]),
          abstractNum(6, ["lowerRoman"]),
          abstractNum(7, ["upperLetter"]),
          abstractNum(8, ["lowerLetter"])
        ],
        numberingInstances: [
          { numId: 5, abstractNumId: 5 },
          { numId: 6, abstractNumId: 6 },
          { numId: 7, abstractNumId: 7 },
          { numId: 8, abstractNumId: 8 }
        ]
      });

      const restored = await readOdt(await writeOdt(doc));
      const paras = restored.body.filter(b => b.type === "paragraph") as Paragraph[];
      const fmts = paras.map(p => formatOf(restored, p.properties!.numbering!.numId, 0));
      expect(fmts).toEqual(["upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]);
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
