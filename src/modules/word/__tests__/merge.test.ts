/**
 * DOCX Module - Merge API Tests
 */

import { describe, it, expect } from "vitest";

import { Build, Io, Transform } from "../index";
import type { DocxDocument, Paragraph, ParagraphChild, Run, StructuredDocumentTag } from "../types";

// Helper to create a minimal document with paragraphs
function createDoc(texts: string[], options?: Partial<DocxDocument>): DocxDocument {
  const body = texts.map(t => ({
    type: "paragraph" as const,
    children: [{ content: [{ type: "text" as const, text: t }] } as Run]
  }));
  return {
    body,
    ...options
  };
}

// Extract all paragraph text from body
function extractTexts(doc: DocxDocument): string[] {
  const result: string[] = [];
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      let t = "";
      for (const child of block.children) {
        if ("content" in child && Array.isArray(child.content)) {
          for (const c of child.content) {
            if (c.type === "text") {
              t += c.text;
            }
          }
        }
      }
      result.push(t);
    }
  }
  return result;
}

describe("mergeDocuments", () => {
  it("returns empty document for empty array", () => {
    const result = Transform.merge([]);
    expect(result.body).toEqual([]);
  });

  it("returns the document itself for single-element array", () => {
    const doc = createDoc(["Hello"]);
    const result = Transform.merge([doc]);
    expect(result.body).toBe(doc.body);
  });

  it("merges two documents with section break between them", () => {
    const doc1 = createDoc(["First doc"]);
    const doc2 = createDoc(["Second doc"]);
    const result = Transform.merge([doc1, doc2]);

    // Should have content from both documents
    const texts = extractTexts(result);
    expect(texts).toContain("First doc");
    expect(texts).toContain("Second doc");
  });

  it("merges three documents", () => {
    const doc1 = createDoc(["A"]);
    const doc2 = createDoc(["B"]);
    const doc3 = createDoc(["C"]);
    const result = Transform.merge([doc1, doc2, doc3]);

    const texts = extractTexts(result);
    expect(texts).toContain("A");
    expect(texts).toContain("B");
    expect(texts).toContain("C");
  });

  it("uses 'continuous' section break when specified", () => {
    const doc1 = createDoc(["First"]);
    const doc2 = createDoc(["Second"]);
    const result = Transform.merge([doc1, doc2], { sectionBreak: "continuous" });

    // The break opens the second document's section.
    expect(result.sectionProperties?.breakType).toBe("continuous");
  });

  it("uses 'nextPage' section break by default", () => {
    const doc1 = createDoc(["First"]);
    const doc2 = createDoc(["Second"]);
    const result = Transform.merge([doc1, doc2]);

    expect(result.sectionProperties?.breakType).toBe("nextPage");
  });

  it("preserves styles from first document", () => {
    const style1 = { styleId: "Heading1", name: "Heading 1", type: "paragraph" };
    const doc1 = createDoc(["First"], { styles: [style1 as any] });
    const doc2 = createDoc(["Second"]);
    const result = Transform.merge([doc1, doc2]);

    expect(result.styles).toBeDefined();
    expect(result.styles!.some((s: any) => s.styleId === "Heading1")).toBe(true);
  });

  it("deduplicates styles by styleId", () => {
    const style = { styleId: "Normal", name: "Normal", type: "paragraph" };
    const doc1 = createDoc(["First"], { styles: [style as any] });
    const doc2 = createDoc(["Second"], { styles: [style as any] });
    const result = Transform.merge([doc1, doc2]);

    const normalStyles = result.styles?.filter((s: any) => s.styleId === "Normal");
    expect(normalStyles?.length).toBe(1);
  });

  it("deduplicates images by fileName", () => {
    const img1 = {
      fileName: "image1.png",
      data: new Uint8Array([1, 2, 3]),
      mediaType: "image/png"
    };
    const doc1 = createDoc(["First"], { images: [img1 as any] });
    const doc2 = createDoc(["Second"], { images: [img1 as any] });
    const result = Transform.merge([doc1, doc2]);

    expect(result.images?.length).toBe(1);
  });

  it("merges different images from both documents", () => {
    const img1 = { fileName: "image1.png", data: new Uint8Array([1]), mediaType: "image/png" };
    const img2 = { fileName: "image2.png", data: new Uint8Array([2]), mediaType: "image/png" };
    const doc1 = createDoc(["First"], { images: [img1 as any] });
    const doc2 = createDoc(["Second"], { images: [img2 as any] });
    const result = Transform.merge([doc1, doc2]);

    expect(result.images?.length).toBe(2);
  });

  // --- Regression: numbering id remapping ---

  it("remaps body paragraph numbering refs when numId conflicts between docs", () => {
    // Both docs use abstractNumId=0 / numId=1 but they describe different lists.
    // After merge, doc2's paragraphs must reference the *new* numId, not the old
    // (which now belongs to doc1's list).
    const doc1: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [{ content: [{ type: "text", text: "doc1 item" }] }]
        }
      ],
      abstractNumberings: [
        { abstractNumId: 0, levels: [{ level: 0, format: "decimal", text: "%1." }] } as any
      ],
      numberingInstances: [{ numId: 1, abstractNumId: 0 }]
    } as unknown as DocxDocument;

    const doc2: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [{ content: [{ type: "text", text: "doc2 item" }] }]
        }
      ],
      abstractNumberings: [
        { abstractNumId: 0, levels: [{ level: 0, format: "bullet", text: "\u2022" }] } as any
      ],
      numberingInstances: [{ numId: 1, abstractNumId: 0 }]
    } as unknown as DocxDocument;

    const merged = Transform.merge([doc1, doc2]);

    // We should now have two distinct numId values and two distinct abstractNumIds.
    expect(merged.numberingInstances?.length).toBe(2);
    const numIds = merged.numberingInstances!.map(i => i.numId);
    expect(new Set(numIds).size).toBe(2);

    // Identify the originally-conflicting numId (doc1 keeps numId=1).
    const originalNumId = 1;
    const remappedInstance = merged.numberingInstances!.find(i => i.numId !== originalNumId);
    expect(remappedInstance).toBeDefined();

    // Find each doc's paragraph in the merged body and check its numId reference.
    const paragraphs = merged.body.filter(b => b.type === "paragraph") as any[];
    const doc1Para = paragraphs.find(p => p.children?.[0]?.content?.[0]?.text === "doc1 item");
    const doc2Para = paragraphs.find(p => p.children?.[0]?.content?.[0]?.text === "doc2 item");
    expect(doc1Para.properties.numbering.numId).toBe(originalNumId);
    expect(doc2Para.properties.numbering.numId).toBe(remappedInstance!.numId);
    // And it must also point to a distinct abstractNumId.
    expect(remappedInstance!.abstractNumId).not.toBe(0);
  });

  it("does not mutate the appended source document", () => {
    const doc1 = createDoc(["a"]);
    const doc2: DocxDocument = {
      body: [
        {
          type: "paragraph",
          properties: { numbering: { numId: 1, level: 0 } },
          children: [{ content: [{ type: "text", text: "x" }] }]
        }
      ],
      abstractNumberings: [
        { abstractNumId: 0, levels: [{ level: 0, format: "decimal", text: "%1." }] } as any
      ],
      numberingInstances: [{ numId: 1, abstractNumId: 0 }]
    } as unknown as DocxDocument;
    const doc1WithList: DocxDocument = {
      ...doc1,
      abstractNumberings: [
        { abstractNumId: 0, levels: [{ level: 0, format: "bullet", text: "\u2022" }] } as any
      ],
      numberingInstances: [{ numId: 1, abstractNumId: 0 }]
    } as DocxDocument;

    Transform.merge([doc1WithList, doc2]);

    // doc2's body paragraph numbering should NOT have been mutated.
    expect((doc2.body[0] as any).properties.numbering.numId).toBe(1);
  });

  it("remaps colliding image rIds and rewrites in-body inline image references", () => {
    // Both docs use rId10 for different images. After merge, doc2's
    // image must get a fresh rId, AND any body run referencing rId10 in
    // doc2 must be rewritten to point at the new rId.
    const img1 = {
      rId: "rId10",
      fileName: "image1.png",
      data: new Uint8Array([1, 2, 3]),
      mediaType: "png"
    };
    const img2 = {
      rId: "rId10",
      fileName: "image2.png",
      data: new Uint8Array([4, 5, 6]),
      mediaType: "png"
    };
    const doc1 = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "image", rId: "rId10", width: 1, height: 1 }] }]
        }
      ],
      images: [img1]
    } as unknown as DocxDocument;
    const doc2 = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "image", rId: "rId10", width: 1, height: 1 }] }]
        }
      ],
      images: [img2]
    } as unknown as DocxDocument;

    const merged = Transform.merge([doc1, doc2]);
    expect(merged.images!.length).toBe(2);

    // Find the rId of the new image (image2.png).
    const img2Final = merged.images!.find(
      i => i.fileName === "image2.png" || i.fileName === "image2_2.png"
    )!;
    expect(img2Final.rId).not.toBe("rId10");

    // Body should now have one paragraph from doc1 (rId10), a section
    // break paragraph, and one paragraph from doc2 referencing the
    // remapped rId.
    const doc2Para = merged.body[merged.body.length - 1] as any;
    const refRId = doc2Para.children[0].content[0].rId;
    expect(refRId).toBe(img2Final.rId);

    // Original docs untouched.
    expect((doc2.body[0] as any).children[0].content[0].rId).toBe("rId10");
  });

  it("merges and remaps colliding footnote ids", () => {
    const doc1 = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "footnoteRef", id: 2 }] }]
        }
      ],
      footnotes: [
        {
          id: 2,
          content: [
            { type: "paragraph", children: [{ content: [{ type: "text", text: "doc1-fn" }] }] }
          ]
        }
      ]
    } as unknown as DocxDocument;
    const doc2 = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "footnoteRef", id: 2 }] }]
        }
      ],
      footnotes: [
        {
          id: 2,
          content: [
            { type: "paragraph", children: [{ content: [{ type: "text", text: "doc2-fn" }] }] }
          ]
        }
      ]
    } as unknown as DocxDocument;

    const merged = Transform.merge([doc1, doc2]);
    // Both footnotes survived.
    expect(merged.footnotes!.length).toBe(2);
    // doc2's footnote got a fresh id (3), and doc2's body ref now points there.
    const fn2 = merged.footnotes![1]!;
    expect(fn2.id).not.toBe(2);

    const doc2BodyPara = merged.body[merged.body.length - 1] as any;
    expect(doc2BodyPara.children[0].content[0].id).toBe(fn2.id);
  });
});

describe("mergeDocuments — section breaks (matches Word)", () => {
  const A4 = { pageSize: { width: 11906, height: 16838 } };
  const LANDSCAPE = { pageSize: { width: 15840, height: 12240 } };

  it("puts the break on the section it opens and keeps each document's page setup", () => {
    const result = Transform.merge(
      [
        createDoc(["First"], { sectionProperties: A4 }),
        createDoc(["Second"], { sectionProperties: LANDSCAPE })
      ],
      { sectionBreak: "oddPage" }
    );
    const closing = result.body[0];
    expect(closing.type === "paragraph" && closing.properties?.sectionProperties).toEqual(A4);
    // The second document's own section opens with the break.
    expect(result.sectionProperties).toEqual({ ...LANDSCAPE, breakType: "oddPage" });
  });

  it("opens an appended document's first inner section with the break", () => {
    const inner = createDoc(["S1", "S2"], { sectionProperties: LANDSCAPE });
    const withBreak = {
      ...inner,
      body: [{ ...inner.body[0], properties: { sectionProperties: A4 } }, inner.body[1]]
    } as DocxDocument;
    const result = Transform.merge([createDoc(["Base"]), withBreak], { sectionBreak: "evenPage" });
    const sects = result.body.flatMap(b =>
      b.type === "paragraph" && b.properties?.sectionProperties
        ? [b.properties.sectionProperties]
        : []
    );
    expect(sects).toEqual([{}, { ...A4, breakType: "evenPage" }]);
    expect(result.sectionProperties).toEqual(LANDSCAPE);
  });

  it("adds a carrier paragraph when the body ends with a section-closing paragraph", () => {
    const ending = createDoc(["Only"]);
    const closed = {
      ...ending,
      body: [{ ...ending.body[0], properties: { sectionProperties: A4 } }],
      sectionProperties: LANDSCAPE
    } as DocxDocument;
    const result = Transform.merge([closed, createDoc(["Next"])]);
    expect(result.body).toHaveLength(3);
    const carrier = result.body[1];
    expect(carrier.type === "paragraph" && carrier.properties?.sectionProperties).toEqual(
      LANDSCAPE
    );
  });
});

describe("mergeDocuments — bookmark names", () => {
  /** A document whose one paragraph is the bookmark `name`, plus `extra` children. */
  function bookmarked(name: string, extra: readonly ParagraphChild[] = []): DocxDocument {
    return {
      body: [
        Build.paragraph([
          Build.bookmarkStart(0, name),
          Build.text(name),
          Build.bookmarkEnd(0),
          ...extra
        ])
      ]
    };
  }

  /** Every bookmark name, hyperlink anchor and field instruction, in order. */
  function references(doc: DocxDocument): string[] {
    const out: string[] = [];
    const visit = (node: unknown): void => {
      if (Array.isArray(node)) {
        node.forEach(visit);
        return;
      }
      if (!node || typeof node !== "object") {
        return;
      }
      const o = node as Record<string, unknown>;
      if (o.type === "bookmarkStart") {
        out.push(`bm:${String(o.name)}`);
      } else if (o.type === "hyperlink" && o.anchor) {
        out.push(`link:${String(o.anchor)}`);
      } else if (o.type === "field") {
        out.push(`field:${String(o.instruction)}`);
      } else if (typeof o.rawXml === "string") {
        out.push(`raw:${o.rawXml}`);
      }
      Object.values(o).forEach(visit);
    };
    visit(doc.body);
    visit(doc.footnotes);
    return out;
  }

  it("renames a name an earlier document already uses, and what refers to it", () => {
    const ref: ParagraphChild = { content: [{ type: "field", instruction: "REF _Toc1 \\h" }] };
    const link = Build.hyperlink("go", { anchor: "_Toc1" });
    const merged = Transform.merge([bookmarked("_Toc1"), bookmarked("_Toc1", [ref, link])]);
    expect(references(merged)).toEqual([
      "bm:_Toc1",
      "bm:_Toc1_2",
      "field:REF _Toc1_2 \\h",
      "link:_Toc1_2"
    ]);
  });

  it("leaves names that do not collide, and never picks one a document already has", () => {
    const merged = Transform.merge([
      bookmarked("a"),
      { body: [...bookmarked("a").body, ...bookmarked("a_2").body] },
      bookmarked("b")
    ]);
    expect(references(merged).filter(r => r.startsWith("bm:"))).toEqual([
      "bm:a",
      "bm:a_3",
      "bm:a_2",
      "bm:b"
    ]);
  });

  it("renames every field kind that names a bookmark, quoted or not", () => {
    const fields = [
      "PAGEREF x \\h",
      'REF "x" \\p',
      "NOTEREF x",
      'HYPERLINK \\l "x"',
      "GOTOBUTTON x Click",
      "REF xy"
    ].map((instruction): ParagraphChild => ({ content: [{ type: "field", instruction }] }));
    const merged = Transform.merge([bookmarked("x"), bookmarked("x", fields)]);
    expect(references(merged).filter(r => r.startsWith("field:"))).toEqual([
      "field:PAGEREF x_2 \\h",
      'field:REF "x_2" \\p',
      "field:NOTEREF x_2",
      'field:HYPERLINK \\l "x_2"',
      "field:GOTOBUTTON x_2 Click",
      "field:REF xy"
    ]);
  });

  it("renames bookmarks in preserved raw XML and in merged notes", () => {
    const raw = {
      type: "opaqueParagraphChild",
      rawXml:
        '<w:bdo w:val="rtl"><w:bookmarkStart w:id="3" w:name="x"/><w:hyperlink w:anchor="x"/></w:bdo>'
    } as ParagraphChild;
    const second: DocxDocument = {
      ...bookmarked("y", [raw]),
      footnotes: [{ id: 1, content: [Build.paragraph([Build.bookmarkStart(5, "x")])] }]
    };
    const merged = Transform.merge([bookmarked("x"), second]);
    const refs = references(merged);
    expect(refs).toContain(
      'raw:<w:bdo w:val="rtl"><w:bookmarkStart w:id="3" w:name="x_2"/><w:hyperlink w:anchor="x_2"/></w:bdo>'
    );
    expect(refs.filter(r => r === "bm:x_2")).toHaveLength(1);
    // The input documents are not modified.
    expect(references(second)).toContain("bm:x");
  });

  it("writes a package whose bookmark names are unique when a document is merged with itself", async () => {
    const doc = bookmarked("_GoBack");
    const parts = await Io.toBuffer(Transform.merge([doc, doc]));
    const read = await Io.read(parts);
    const names = references(read).filter(r => r.startsWith("bm:"));
    expect(new Set(names).size).toBe(names.length);
  });
});

describe("mergeDocuments — comment ids on markers after the content", () => {
  it("remaps a comment range end held after a paragraph's children or an SDT's content", () => {
    const withComment = (): DocxDocument => ({
      body: [
        {
          type: "paragraph",
          children: [{ type: "commentRangeStart", id: 0 }, Build.text("p")],
          rangeMarkersAfter: [{ type: "commentRangeEnd", id: 0 }]
        },
        {
          type: "sdt",
          content: [Build.paragraph([{ type: "commentRangeStart", id: 1 }, Build.text("s")])],
          contentRangeMarkersAfter: [{ type: "commentRangeEnd", id: 1 }]
        }
      ],
      comments: [
        { id: 0, author: "a", content: [] },
        { id: 1, author: "a", content: [] }
      ]
    });
    const merged = Transform.merge([withComment(), withComment()]);
    const [, , , para, sdt] = merged.body as [
      unknown,
      unknown,
      unknown,
      Paragraph,
      StructuredDocumentTag
    ];
    const startId = (para.children[0] as { id: number }).id;
    expect(startId).not.toBe(0);
    expect(para.rangeMarkersAfter).toEqual([{ type: "commentRangeEnd", id: startId }]);
    const sdtStart = ((sdt.content[0] as Paragraph).children[0] as { id: number }).id;
    expect(sdtStart).not.toBe(1);
    expect(sdt.contentRangeMarkersAfter).toEqual([{ type: "commentRangeEnd", id: sdtStart }]);
  });
});
