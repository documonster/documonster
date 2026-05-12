/**
 * DOCX Package Integrity Tests
 *
 * Validates the structural completeness of DOCX ZIP packages
 * produced by packageDocx, including:
 * - ZIP magic bytes
 * - Required parts and relationships
 * - Content types correctness
 * - No mutation of input model
 * - Conditional parts (headers, footers, images, numbering, footnotes, comments)
 * - Floating image position preservation
 */

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Document, packageDocx, readDocx, textParagraph } from "@word/index";
import { describe, it, expect } from "vitest";

// Minimal 1x1 PNG for image tests
const MINI_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
  0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

/** Helper: extract all files from DOCX bytes into a Map<path, string|Uint8Array>. */
async function extractDocx(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const files = await extractAll(bytes);
  const result = new Map<string, Uint8Array>();
  for (const [path, entry] of files) {
    result.set(path, entry.data);
  }
  return result;
}

/** Helper: decode an entry to UTF-8 string. */
function decodeEntry(files: Map<string, Uint8Array>, path: string): string {
  const data = files.get(path);
  if (!data) {
    throw new Error(`Entry not found: ${path}`);
  }
  return new TextDecoder().decode(data);
}

// =============================================================================
// Basic ZIP Structure
// =============================================================================

describe("DOCX Package Integrity", () => {
  it("should produce a valid ZIP with PK magic bytes", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Hello");
    const bytes = await packageDocx(Document.build(h));

    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
    expect(bytes[2]).toBe(0x03);
    expect(bytes[3]).toBe(0x04);
  });

  // ===========================================================================
  // [Content_Types].xml
  // ===========================================================================

  it("should contain [Content_Types].xml with required Default and Override entries", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const ct = decodeEntry(files, "[Content_Types].xml");
    // Default extensions
    expect(ct).toContain('Extension="rels"');
    expect(ct).toContain('Extension="xml"');
    // Override entries for core parts
    expect(ct).toContain('PartName="/word/document.xml"');
    expect(ct).toContain('PartName="/word/styles.xml"');
    expect(ct).toContain('PartName="/word/settings.xml"');
    expect(ct).toContain('PartName="/word/fontTable.xml"');
    expect(ct).toContain('PartName="/word/theme/theme1.xml"');
    expect(ct).toContain('PartName="/docProps/core.xml"');
    expect(ct).toContain('PartName="/docProps/app.xml"');
  });

  // ===========================================================================
  // _rels/.rels
  // ===========================================================================

  it("should contain _rels/.rels with officeDocument, core-properties, and extended-properties relationships", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const rels = decodeEntry(files, "_rels/.rels");
    expect(rels).toContain("officeDocument");
    expect(rels).toContain("core-properties");
    expect(rels).toContain("extended-properties");
    expect(rels).toContain("word/document.xml");
    expect(rels).toContain("docProps/core.xml");
    expect(rels).toContain("docProps/app.xml");
  });

  // ===========================================================================
  // word/_rels/document.xml.rels
  // ===========================================================================

  it("should contain word/_rels/document.xml.rels with styles, settings, fontTable, theme relationships", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const rels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(rels).toContain("styles.xml");
    expect(rels).toContain("settings.xml");
    expect(rels).toContain("fontTable.xml");
    expect(rels).toContain("theme/theme1.xml");
  });

  // ===========================================================================
  // word/document.xml
  // ===========================================================================

  it("should contain word/document.xml with XML declaration and w:document root element", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Hello World");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const doc = decodeEntry(files, "word/document.xml");
    expect(doc).toContain('<?xml version="1.0"');
    expect(doc).toContain("<w:document");
    expect(doc).toContain("</w:document>");
  });

  // ===========================================================================
  // Required auxiliary parts exist
  // ===========================================================================

  it("should contain word/styles.xml", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/styles.xml")).toBe(true);
  });

  it("should contain word/settings.xml", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/settings.xml")).toBe(true);
  });

  it("should contain word/fontTable.xml", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/fontTable.xml")).toBe(true);
  });

  it("should contain word/theme/theme1.xml", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/theme/theme1.xml")).toBe(true);
  });

  it("should contain docProps/core.xml and docProps/app.xml", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Test");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("docProps/core.xml")).toBe(true);
    expect(files.has("docProps/app.xml")).toBe(true);
  });

  // ===========================================================================
  // packageDocx should not pollute the input model
  // ===========================================================================

  it("should not mutate input document images rId", async () => {
    const h = Document.create();
    Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    const doc = Document.build(h);

    // Save the original rId before packaging
    const originalRId = doc.images![0].rId;

    // Package the document
    await packageDocx(doc);

    // The rId on the original doc should remain unchanged
    expect(doc.images![0].rId).toBe(originalRId);
  });

  // ===========================================================================
  // Header/Footer parts
  // ===========================================================================

  it("should include header/footer parts and their rels when headers/footers are present", async () => {
    const h = Document.create();
    Document.setHeader(h, "default", { children: [textParagraph("My Header")] });
    Document.setFooter(h, "default", { children: [textParagraph("My Footer")] });
    Document.setSectionProperties(h, {
      headers: [{ type: "default", rId: "" }],
      footers: [{ type: "default", rId: "" }]
    });
    Document.addParagraph(h, "Body");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    // header1.xml should exist
    const headerPaths = [...files.keys()].filter(p => p.match(/^word\/header\d+\.xml$/));
    expect(headerPaths.length).toBeGreaterThanOrEqual(1);

    // footer should exist
    const footerPaths = [...files.keys()].filter(p => p.match(/^word\/footer\d+\.xml$/));
    expect(footerPaths.length).toBeGreaterThanOrEqual(1);

    // document.xml.rels should reference header/footer
    const rels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(rels).toContain("header");
    expect(rels).toContain("footer");

    // Content types should have Override for header/footer
    const ct = decodeEntry(files, "[Content_Types].xml");
    expect(ct).toContain("header");
    expect(ct).toContain("footer");
  });

  it("should produce header/footer references whose rId resolves in document.xml.rels", async () => {
    const h = Document.create();
    Document.setHeader(h, "default", { children: [textParagraph("My Header")] });
    Document.setFooter(h, "default", { children: [textParagraph("My Footer")] });
    Document.setSectionProperties(h, {
      headers: [{ type: "default", rId: "" }],
      footers: [{ type: "default", rId: "" }]
    });
    Document.addParagraph(h, "Body");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const documentXml = decodeEntry(files, "word/document.xml");
    const documentRels = decodeEntry(files, "word/_rels/document.xml.rels");

    // Pull every r:id used by a headerReference/footerReference.
    const refIdRe = /<w:(?:header|footer)Reference[^>]*r:id="([^"]+)"/g;
    const referencedIds = new Set<string>();
    for (const m of documentXml.matchAll(refIdRe)) {
      referencedIds.add(m[1]!);
    }
    expect(referencedIds.size).toBeGreaterThanOrEqual(2);
    // Every referenced rId must correspond to a relationship entry.
    for (const id of referencedIds) {
      expect(id).not.toBe("");
      expect(documentRels).toContain(`Id="${id}"`);
    }
  });

  it("should auto-fill section header/footer references when builder did not author them", async () => {
    // Builder users typically call setHeader/setFooter without manually
    // adding a HeaderFooterRef in sectionProperties. The packager should
    // synthesize references so the section actually picks them up.
    const h = Document.create();
    Document.setHeader(h, "default", { children: [textParagraph("Auto Header")] });
    Document.setFooter(h, "default", { children: [textParagraph("Auto Footer")] });
    Document.addParagraph(h, "Body");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    const documentXml = decodeEntry(files, "word/document.xml");
    const documentRels = decodeEntry(files, "word/_rels/document.xml.rels");

    expect(documentXml).toMatch(/<w:headerReference[^>]*r:id="[^"]+"/);
    expect(documentXml).toMatch(/<w:footerReference[^>]*r:id="[^"]+"/);

    const refIds = [
      ...documentXml.matchAll(/<w:(?:header|footer)Reference[^>]*r:id="([^"]+)"/g)
    ].map(m => m[1]!);
    for (const id of refIds) {
      expect(documentRels).toContain(`Id="${id}"`);
    }
  });

  it("should round-trip headers/footers with consistent rId references", async () => {
    // Round-trip path: read a generated DOCX (which will carry its own rIds
    // from the original packaging pass), then repackage it. The second pass
    // assigns fresh rIds for header/footer parts, so the section-property
    // references must be rewritten to match.
    const h1 = Document.create();
    Document.setHeader(h1, "default", { children: [textParagraph("Round-trip Header")] });
    Document.setFooter(h1, "default", { children: [textParagraph("Round-trip Footer")] });
    Document.addParagraph(h1, "Body");
    const firstBytes = await packageDocx(Document.build(h1));
    const parsed = await readDocx(firstBytes);

    const secondBytes = await packageDocx(parsed);
    const files = await extractDocx(secondBytes);
    const documentXml = decodeEntry(files, "word/document.xml");
    const documentRels = decodeEntry(files, "word/_rels/document.xml.rels");

    const refIds = [
      ...documentXml.matchAll(/<w:(?:header|footer)Reference[^>]*r:id="([^"]+)"/g)
    ].map(m => m[1]!);
    expect(refIds.length).toBeGreaterThanOrEqual(2);
    for (const id of refIds) {
      expect(documentRels).toContain(`Id="${id}"`);
    }
  });

  // ===========================================================================
  // Image parts
  // ===========================================================================

  it("should include image media files and content type defaults when images are present", async () => {
    const h = Document.create();
    Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraph(h, "With image");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    // media/ directory should have image files
    const mediaPaths = [...files.keys()].filter(p => p.startsWith("word/media/"));
    expect(mediaPaths.length).toBeGreaterThanOrEqual(1);
    expect(mediaPaths.some(p => p.endsWith(".png"))).toBe(true);

    // Content types should include png default
    const ct = decodeEntry(files, "[Content_Types].xml");
    expect(ct).toContain('Extension="png"');
  });

  it("should reject opaque parts whose path collides with a packager-managed part", async () => {
    // Building a doc that already produces word/document.xml and adding an
    // opaque entry under the same path would emit a duplicate ZIP entry.
    // The packager must reject it loudly rather than silently corrupt.
    const h = Document.create();
    Document.addParagraph(h, "Body");
    const doc = Document.build(h);
    const docWithConflict = {
      ...doc,
      opaqueParts: [
        {
          path: "word/document.xml",
          data: new TextEncoder().encode("<bogus/>")
        }
      ]
    };
    await expect(packageDocx(docWithConflict)).rejects.toThrow(/conflicts with a part/);
  });

  it("should reject opaque parts that collide with header parts emitted in this run", async () => {
    const h = Document.create();
    Document.setHeader(h, "default", { children: [textParagraph("Header")] });
    Document.addParagraph(h, "Body");
    const doc = Document.build(h);
    const docWithConflict = {
      ...doc,
      opaqueParts: [
        {
          path: "word/header1.xml",
          data: new TextEncoder().encode("<bogus/>")
        }
      ]
    };
    await expect(packageDocx(docWithConflict)).rejects.toThrow(/conflicts with a part/);
  });

  // ===========================================================================
  // Numbering part
  // ===========================================================================

  it("should include word/numbering.xml and corresponding relationship when numbering is present", async () => {
    const h = Document.create();
    Document.addNumberedList(h, ["First", "Second", "Third"]);
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/numbering.xml")).toBe(true);

    // Relationship in document.xml.rels
    const rels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(rels).toContain("numbering.xml");
  });

  // ===========================================================================
  // Footnotes and Endnotes
  // ===========================================================================

  it("should include footnotes/endnotes parts and relationships when present", async () => {
    const h = Document.create();
    Document.addFootnote(h, "A footnote");
    Document.addEndnote(h, "An endnote");
    Document.addParagraph(h, "Body text");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    // Parts exist
    expect(files.has("word/footnotes.xml")).toBe(true);
    expect(files.has("word/endnotes.xml")).toBe(true);

    // Relationships exist
    const rels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(rels).toContain("footnotes.xml");
    expect(rels).toContain("endnotes.xml");
  });

  // ===========================================================================
  // Comments
  // ===========================================================================

  it("should include word/comments.xml when comments are present", async () => {
    const h = Document.create();
    Document.addComment(h, "Author", "This is a comment");
    Document.addParagraph(h, "Body text");
    const bytes = await packageDocx(Document.build(h));
    const files = await extractDocx(bytes);

    expect(files.has("word/comments.xml")).toBe(true);

    // Relationship should reference comments
    const rels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(rels).toContain("comments.xml");
  });

  it("emits word/_rels/comments.xml.rels for hyperlinks inside comments", async () => {
    // Comments live in their own OPC part — relationships referenced from
    // comment paragraphs must be written into comments.xml.rels rather than
    // document.xml.rels, otherwise readers cannot resolve the URL.
    const link = {
      type: "hyperlink" as const,
      url: "https://example.com/spec",
      children: [{ content: [{ type: "text" as const, text: "spec" }] }]
    };
    const commentPara: any = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "see " }] }, link]
    };
    const seedDoc: any = {
      body: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "body" }] }] }],
      comments: [{ id: 1, author: "X", content: [commentPara] }]
    };
    const bytes = await packageDocx(seedDoc);
    const files = await extractDocx(bytes);

    expect(files.has("word/_rels/comments.xml.rels")).toBe(true);
    const commentsRels = decodeEntry(files, "word/_rels/comments.xml.rels");
    expect(commentsRels).toContain("https://example.com/spec");
    expect(commentsRels).toContain('TargetMode="External"');

    // The same external URL must NOT also leak into document.xml.rels just
    // because of the comment-internal reference.
    const docRels = decodeEntry(files, "word/_rels/document.xml.rels");
    expect(docRels).not.toContain("https://example.com/spec");

    // Round-trip: the comment paragraph reads back with the URL.
    const parsed = await readDocx(bytes);
    const c = parsed.comments?.[0];
    expect(c).toBeDefined();
    const para = c!.content[0] as any;
    const linkChild = para.children.find((ch: any) => ch.type === "hyperlink");
    expect(linkChild?.url).toBe("https://example.com/spec");
  });

  // ===========================================================================
  // Floating image position preservation
  // ===========================================================================

  it("should preserve floating image position between paragraphs through write→read roundtrip", async () => {
    const h = Document.create();
    // Paragraph 1
    Document.addParagraph(h, "First paragraph");
    // Paragraph 2
    Document.addParagraph(h, "Second paragraph");
    // Floating image after paragraph 2
    Document.addFloatingImage(h, MINI_PNG, "png", 914400, 914400, {
      name: "TestFloat"
    });
    // Paragraph 3
    Document.addParagraph(h, "Third paragraph");

    const doc = Document.build(h);
    const bytes = await packageDocx(doc);
    const parsed = await readDocx(bytes);

    // Find the floating image index in body
    const floatingIdx = parsed.body.findIndex(b => b.type === "floatingImage");
    expect(floatingIdx).toBeGreaterThan(-1);

    // Find indices of paragraphs containing specific text.
    // Note: the writer wraps floating images in a <w:p>, which after roundtrip
    // becomes an extra paragraph. We locate paragraphs by their text content.
    let secondParaIdx = -1;
    let thirdParaIdx = -1;
    for (let i = 0; i < parsed.body.length; i++) {
      const item = parsed.body[i];
      if (item.type === "paragraph") {
        const text = (item as any).children
          ?.flatMap((c: any) => c.content ?? [])
          .filter((c: any) => c.type === "text")
          .map((c: any) => c.text)
          .join("");
        if (text?.includes("Second paragraph")) {
          secondParaIdx = i;
        }
        if (text?.includes("Third paragraph")) {
          thirdParaIdx = i;
        }
      }
    }

    expect(secondParaIdx).toBeGreaterThan(-1);
    expect(thirdParaIdx).toBeGreaterThan(-1);

    // Floating image should be after the "Second paragraph" and before "Third paragraph"
    expect(floatingIdx).toBeGreaterThan(secondParaIdx);
    expect(floatingIdx).toBeLessThan(thirdParaIdx);
  });

  // ===========================================================================
  // Image rId integrity & input-model immutability
  // ===========================================================================

  it("inline image references the same rId that the document.xml.rels file declares", async () => {
    const h = Document.create();
    Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraph(h, "after image");
    const doc = Document.build(h);

    const bytes = await packageDocx(doc);
    const files = await extractDocx(bytes);
    const docXml = decodeEntry(files, "word/document.xml");
    const relsXml = decodeEntry(files, "word/_rels/document.xml.rels");

    // Pull the embed rId out of <a:blip r:embed="rIdN"/>.
    const embedMatch = /r:embed="([^"]+)"/.exec(docXml);
    expect(embedMatch).not.toBeNull();
    const embedRId = embedMatch![1];

    // The same rId must appear in the rels file pointing at media/.
    const relRegex = new RegExp(`Id="${embedRId}"[^>]*Target="media/[^"]+\\.png"`, "i");
    expect(relsXml).toMatch(relRegex);
  });

  it("inline image rId is consistent even when builder rId is re-assigned at package time", async () => {
    // Manually craft a doc whose builder-allocated rId clashes with what the
    // packager would otherwise pick (rId1 for styles or anything else).
    const doc: any = {
      body: [
        {
          type: "paragraph",
          children: [
            {
              content: [
                {
                  type: "image",
                  rId: "rId-from-template",
                  width: 914400,
                  height: 914400,
                  drawingId: 1,
                  name: "p"
                }
              ]
            }
          ]
        }
      ],
      images: [
        {
          rId: "rId-from-template",
          fileName: "image1.png",
          mediaType: "png",
          data: MINI_PNG
        }
      ]
    };

    const bytes = await packageDocx(doc);
    const files = await extractDocx(bytes);
    const docXml = decodeEntry(files, "word/document.xml");
    const relsXml = decodeEntry(files, "word/_rels/document.xml.rels");

    const embedMatch = /r:embed="([^"]+)"/.exec(docXml);
    expect(embedMatch).not.toBeNull();
    const embedRId = embedMatch![1];

    const relRegex = new RegExp(`Id="${embedRId}"[^>]*Target="media/image1\\.png"`, "i");
    expect(relsXml).toMatch(relRegex);
  });

  it("packageDocx does not mutate inline image rId on the caller's model", async () => {
    const doc: any = {
      body: [
        {
          type: "paragraph",
          children: [
            {
              content: [
                {
                  type: "image",
                  rId: "rId-from-template",
                  width: 914400,
                  height: 914400,
                  drawingId: 1,
                  name: "p"
                }
              ]
            }
          ]
        }
      ],
      images: [
        {
          rId: "rId-from-template",
          fileName: "image1.png",
          mediaType: "png",
          data: MINI_PNG
        }
      ]
    };

    const originalImageRId = doc.images[0].rId;
    const originalContentRId = doc.body[0].children[0].content[0].rId;

    await packageDocx(doc);

    expect(doc.images[0].rId).toBe(originalImageRId);
    expect(doc.body[0].children[0].content[0].rId).toBe(originalContentRId);
  });

  it("packageDocx is idempotent: a second call produces an equivalent package", async () => {
    const h = Document.create();
    Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraph(h, "Hello");
    const doc = Document.build(h);

    const bytes1 = await packageDocx(doc);
    const bytes2 = await packageDocx(doc);

    const f1 = await extractDocx(bytes1);
    const f2 = await extractDocx(bytes2);

    // Same file set
    expect([...f1.keys()].sort()).toEqual([...f2.keys()].sort());

    // The document.xml.rels and word/document.xml should be byte-equal
    // (rId allocation is deterministic from the input model).
    const a = decodeEntry(f1, "word/_rels/document.xml.rels");
    const b = decodeEntry(f2, "word/_rels/document.xml.rels");
    expect(a).toBe(b);
    const da = decodeEntry(f1, "word/document.xml");
    const db = decodeEntry(f2, "word/document.xml");
    expect(da).toBe(db);
  });

  it("packageDocx preserves a floating image rId that survives multiple packagings", async () => {
    const h = Document.create();
    Document.addParagraph(h, "before");
    Document.addFloatingImage(h, MINI_PNG, "png", 914400, 914400, { name: "fl" });
    Document.addParagraph(h, "after");
    const doc = Document.build(h);

    // Capture floating image rId before packaging.
    const floatingItem: any = doc.body.find(b => b.type === "floatingImage");
    const originalFiRId = floatingItem.rId;

    await packageDocx(doc);

    // The caller's floating image must not have its rId rewritten on the model.
    expect(floatingItem.rId).toBe(originalFiRId);
  });

  it("floating image r:embed in document.xml points at a real relationship", async () => {
    const h = Document.create();
    Document.addParagraph(h, "before");
    Document.addFloatingImage(h, MINI_PNG, "png", 914400, 914400, { name: "fl" });
    Document.addParagraph(h, "after");
    const doc = Document.build(h);

    const bytes = await packageDocx(doc);
    const files = await extractDocx(bytes);
    const docXml = decodeEntry(files, "word/document.xml");
    const relsXml = decodeEntry(files, "word/_rels/document.xml.rels");

    // Find the wp:anchor block and pull its r:embed.
    const anchorMatch = /<wp:anchor[\s\S]*?r:embed="([^"]+)"/.exec(docXml);
    expect(anchorMatch).not.toBeNull();
    const embedRId = anchorMatch![1];

    // That rId must exist in the relationships file with an image target.
    const relRegex = new RegExp(`Id="${embedRId}"[^>]*Target="media/[^"]+\\.png"`, "i");
    expect(relsXml).toMatch(relRegex);
  });

  // ===========================================================================
  // Header/footer rId aliasing for shared media files (Bug #5)
  // ===========================================================================

  it("preserves header rId aliases for media also referenced from the main document", async () => {
    // Construct a synthetic DOCX where the main document references
    // word/media/image1.png as rId10 and header1.xml references the *same*
    // media file as rIdH1 (independent local id space). The reader must keep
    // both rIds alive so that re-packaging produces a valid header1.xml.rels.
    const enc = new TextEncoder();

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`;

    const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;

    // header1.xml references the same image file but under a *different* local id
    const headerRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="P"/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="P"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference r:id="rId20" w:type="default"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:p>
    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="457200" cy="457200"/>
          <wp:docPr id="2" name="HP"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="HP"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdH1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>`;

    const zipBuffer = await createZip([
      { name: "[Content_Types].xml", data: enc.encode(contentTypes) },
      { name: "_rels/.rels", data: enc.encode(packageRels) },
      { name: "word/_rels/document.xml.rels", data: enc.encode(docRels) },
      { name: "word/document.xml", data: enc.encode(documentXml) },
      { name: "word/header1.xml", data: enc.encode(headerXml) },
      { name: "word/_rels/header1.xml.rels", data: enc.encode(headerRels) },
      { name: "word/media/image1.png", data: MINI_PNG }
    ]);

    const parsed = await readDocx(zipBuffer);

    // The image must still be reachable from at least one ImageDef.
    expect(parsed.images?.length).toBeGreaterThan(0);
    const img = parsed.images!.find(i => i.fileName === "image1.png");
    expect(img).toBeDefined();

    // Re-package and verify the header1.xml.rels still resolves rIdH1.
    const repacked = await packageDocx(parsed);
    const files = await extractDocx(repacked);

    const headerRelsKey = [...files.keys()].find(k =>
      /^word\/_rels\/header\d+\.xml\.rels$/.test(k)
    );
    expect(headerRelsKey).toBeDefined();
    const headerRelsXml = decodeEntry(files, headerRelsKey!);

    // The header rId used by header1.xml's r:embed must be registered locally.
    const headerXmlKey = [...files.keys()].find(k => /^word\/header\d+\.xml$/.test(k));
    expect(headerXmlKey).toBeDefined();
    const headerOut = decodeEntry(files, headerXmlKey!);
    const headerEmbedMatch = /<a:blip[^>]*r:embed="([^"]+)"/.exec(headerOut);
    expect(headerEmbedMatch).not.toBeNull();
    const headerEmbedRId = headerEmbedMatch![1];

    const relRegex = new RegExp(`Id="${headerEmbedRId}"[^>]*Target="media/image1\\.png"`, "i");
    expect(headerRelsXml).toMatch(relRegex);
  });

  // ===========================================================================
  // SVG fallback must reach header/footer images, not only doc.body
  // ===========================================================================

  it("auto-injects svgRId for SVG images embedded in header/footer", async () => {
    // Build a doc where a SVG-with-PNG-fallback image is referenced from
    // BOTH the body and the header. The packager auto-allocates a secondary
    // rId for the SVG and must make both the body inline drawing AND the
    // header inline drawing emit asvg:svgBlip pointing at that rId.
    const SVG_DATA = new TextEncoder().encode(
      '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>'
    );
    const inlineImage = (rId: string) => ({
      type: "image" as const,
      rId,
      width: 914400,
      height: 914400,
      drawingId: 1,
      name: "P"
    });

    const doc: any = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [inlineImage("imgSvg1")] }]
        }
      ],
      headers: new Map([
        [
          "default",
          {
            type: "default",
            content: {
              children: [{ type: "paragraph", children: [{ content: [inlineImage("imgSvg1")] }] }]
            }
          }
        ]
      ]),
      sectionProperties: {
        headers: [{ type: "default", rId: "" }]
      },
      images: [
        {
          rId: "imgSvg1",
          fileName: "image1.svg",
          mediaType: "svg",
          data: SVG_DATA,
          fallbackData: MINI_PNG
        }
      ]
    };

    const bytes = await packageDocx(doc);
    const files = await extractDocx(bytes);

    const docXml = decodeEntry(files, "word/document.xml");
    const headerXmlKey = [...files.keys()].find(k => /^word\/header\d+\.xml$/.test(k));
    expect(headerXmlKey).toBeDefined();
    const headerOut = decodeEntry(files, headerXmlKey!);

    // Body must already work (regression guard).
    expect(docXml).toMatch(/asvg:svgBlip[^>]*r:embed="([^"]+)"/);
    // The bug: header inline image must also gain an asvg:svgBlip pointer.
    expect(headerOut).toMatch(/asvg:svgBlip[^>]*r:embed="([^"]+)"/);
  });

  // ===========================================================================
  // altChunk target must not be written twice (was: also kept in opaqueParts)
  // ===========================================================================

  it("does not duplicate the altChunk target file when re-packaging a parsed DOCX", async () => {
    // First, build a DOCX that contains an altChunk using packageDocx.
    const html = "<html><body><p>Hello altChunk</p></body></html>";
    const seed: any = {
      body: [
        {
          type: "altChunk",
          rId: "__altchunk_seed",
          contentType: "text/html",
          fileName: "afchunk1.html",
          data: new TextEncoder().encode(html)
        }
      ]
    };
    const seedBytes = await packageDocx(seed);

    // Read it back. Reader should leave the altChunk content on the body item
    // and NOT additionally retain it as an opaque part.
    const parsed = await readDocx(seedBytes);
    const altItem = parsed.body.find(b => b.type === "altChunk") as any;
    expect(altItem).toBeDefined();

    // Bug guard: the altChunk target must not appear in opaqueParts.
    const altPathsInOpaque = (parsed.opaqueParts ?? []).filter(p =>
      p.path.endsWith("afchunk1.html")
    );
    expect(altPathsInOpaque).toHaveLength(0);

    // Re-packaging: the resulting ZIP must contain exactly ONE copy of
    // word/afchunk1.html. extractAll naturally dedups by name, so we instead
    // scan central directory entries.
    const repacked = await packageDocx(parsed);
    // Quick way: count occurrences of the path in the file.
    const haystack = new TextDecoder("latin1").decode(repacked);
    const matches = haystack.match(/word\/afchunk1\.html/g) ?? [];
    // ZIP local header + central directory header => 2 occurrences for ONE
    // entry; a duplicated entry produces 4. Anything > 2 is a duplicate.
    expect(matches.length).toBeLessThanOrEqual(2);
  });

  // ===========================================================================
  // packageDocx must not mutate caller's hyperlink / altChunk / header / footer rIds
  // ===========================================================================

  it("does not mutate caller's hyperlink rId in body, footnotes or endnotes", async () => {
    const bodyLink: any = {
      type: "hyperlink",
      url: "https://example.com/body",
      children: [{ content: [{ type: "text", text: "body link" }] }]
    };
    const footnoteLink: any = {
      type: "hyperlink",
      url: "https://example.com/footnote",
      children: [{ content: [{ type: "text", text: "fn link" }] }]
    };
    const doc: any = {
      body: [{ type: "paragraph", children: [bodyLink] }],
      footnotes: [
        {
          id: 2,
          content: [{ type: "paragraph", children: [footnoteLink] }]
        }
      ]
    };
    expect(bodyLink.rId).toBeUndefined();
    expect(footnoteLink.rId).toBeUndefined();
    await packageDocx(doc);
    expect(bodyLink.rId).toBeUndefined();
    expect(footnoteLink.rId).toBeUndefined();
  });

  it("does not mutate caller's altChunk rId / fileName", async () => {
    const chunk: any = {
      type: "altChunk",
      contentType: "text/html",
      data: new TextEncoder().encode("<p>x</p>")
    };
    const doc: any = { body: [chunk] };
    expect(chunk.rId).toBeUndefined();
    expect(chunk.fileName).toBeUndefined();
    await packageDocx(doc);
    expect(chunk.rId).toBeUndefined();
    expect(chunk.fileName).toBeUndefined();
  });

  it("does not mutate caller's HeaderDef / FooterDef rId", async () => {
    const headerDef: any = {
      type: "default",
      content: { children: [{ type: "paragraph", children: [] }] }
    };
    const footerDef: any = {
      type: "default",
      content: { children: [{ type: "paragraph", children: [] }] }
    };
    const doc: any = {
      body: [{ type: "paragraph", children: [] }],
      headers: new Map([["default", headerDef]]),
      footers: new Map([["default", footerDef]])
    };
    expect(headerDef.rId).toBeUndefined();
    expect(footerDef.rId).toBeUndefined();
    await packageDocx(doc);
    expect(headerDef.rId).toBeUndefined();
    expect(footerDef.rId).toBeUndefined();
  });

  it("packaging the same doc twice yields equivalent output even with hyperlinks", async () => {
    const link: any = {
      type: "hyperlink",
      url: "https://example.com/idem",
      children: [{ content: [{ type: "text", text: "x" }] }]
    };
    const doc: any = {
      body: [{ type: "paragraph", children: [link] }]
    };
    const a = await packageDocx(doc);
    const b = await packageDocx(doc);
    const fa = await extractDocx(a);
    const fb = await extractDocx(b);
    expect(decodeEntry(fa, "word/document.xml")).toBe(decodeEntry(fb, "word/document.xml"));
    expect(decodeEntry(fa, "word/_rels/document.xml.rels")).toBe(
      decodeEntry(fb, "word/_rels/document.xml.rels")
    );
  });

  // ===========================================================================
  // ZIP bomb / resource-limit guards
  // ===========================================================================

  it("rejects packages exceeding maxPackageSize", async () => {
    const h = Document.create();
    Document.addParagraph(h, "tiny");
    const bytes = await packageDocx(Document.build(h));
    // Force a tiny limit so even this minimal DOCX trips it.
    await expect(readDocx(bytes, { securityPolicy: { maxPackageSize: 100 } })).rejects.toThrow(
      /packageSize limit exceeded/
    );
  });

  it("rejects packages exceeding maxPartCount", async () => {
    const h = Document.create();
    Document.addParagraph(h, "tiny");
    const bytes = await packageDocx(Document.build(h));
    await expect(readDocx(bytes, { securityPolicy: { maxPartCount: 1 } })).rejects.toThrow(
      /partCount limit exceeded/
    );
  });

  it("accepts packages within the configured policy", async () => {
    const h = Document.create();
    Document.addParagraph(h, "ok");
    const bytes = await packageDocx(Document.build(h));
    // Generous limits — should succeed.
    const parsed = await readDocx(bytes, {
      securityPolicy: { maxPackageSize: 50_000_000, maxPartCount: 1000, maxPartSize: 5_000_000 }
    });
    expect(parsed.body.length).toBeGreaterThan(0);
  });

  // ===========================================================================
  // Floating image stability across multiple round-trips
  // ===========================================================================

  it("does not accumulate empty paragraphs across multiple floating-image round-trips", async () => {
    const h = Document.create();
    Document.addParagraph(h, "before");
    Document.addFloatingImage(h, MINI_PNG, "png", 914400, 914400, { name: "fl" });
    Document.addParagraph(h, "after");

    // First round-trip
    const bytes1 = await packageDocx(Document.build(h));
    const parsed1 = await readDocx(bytes1);

    // Second round-trip (re-package whatever the reader produced)
    const bytes2 = await packageDocx(parsed1);
    const parsed2 = await readDocx(bytes2);

    // Third round-trip
    const bytes3 = await packageDocx(parsed2);
    const parsed3 = await readDocx(bytes3);

    // Body composition (paragraph + floatingImage + paragraph + ...) must not
    // grow with each round-trip. We compare body length between rounds 2 and 3.
    expect(parsed3.body.length).toBe(parsed2.body.length);

    // Floating image must still be present.
    expect(parsed3.body.some(b => b.type === "floatingImage")).toBe(true);
  });
});
