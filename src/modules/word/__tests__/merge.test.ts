/**
 * DOCX Module - Merge API Tests
 */

import { describe, it, expect } from "vitest";

import { Io } from "../index";
import type { DocxDocument, Run } from "../types";

// Helper to create a minimal document with paragraphs
function createDoc(texts: string[], options?: Partial<DocxDocument>): DocxDocument {
  const body = texts.map(t => ({
    type: "paragraph" as const,
    children: [{ content: [{ type: "text" as const, text: t }] } as Run]
  }));
  return {
    body,
    contentTypes: [],
    ...options
  } as unknown as DocxDocument;
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
    const result = Io.merge([]);
    expect(result.body).toEqual([]);
  });

  it("returns the document itself for single-element array", () => {
    const doc = createDoc(["Hello"]);
    const result = Io.merge([doc]);
    expect(result.body).toBe(doc.body);
  });

  it("merges two documents with section break between them", () => {
    const doc1 = createDoc(["First doc"]);
    const doc2 = createDoc(["Second doc"]);
    const result = Io.merge([doc1, doc2]);

    // Should have content from both documents
    const texts = extractTexts(result);
    expect(texts).toContain("First doc");
    expect(texts).toContain("Second doc");
  });

  it("merges three documents", () => {
    const doc1 = createDoc(["A"]);
    const doc2 = createDoc(["B"]);
    const doc3 = createDoc(["C"]);
    const result = Io.merge([doc1, doc2, doc3]);

    const texts = extractTexts(result);
    expect(texts).toContain("A");
    expect(texts).toContain("B");
    expect(texts).toContain("C");
  });

  it("uses 'continuous' section break when specified", () => {
    const doc1 = createDoc(["First"]);
    const doc2 = createDoc(["Second"]);
    const result = Io.merge([doc1, doc2], { sectionBreak: "continuous" });

    // Verify the section break paragraph has correct section properties
    let hasContinuous = false;
    for (const block of result.body) {
      if (block.type === "paragraph" && block.properties?.sectionProperties) {
        if (block.properties.sectionProperties.breakType === "continuous") {
          hasContinuous = true;
        }
      }
    }
    expect(hasContinuous).toBe(true);
  });

  it("uses 'nextPage' section break by default", () => {
    const doc1 = createDoc(["First"]);
    const doc2 = createDoc(["Second"]);
    const result = Io.merge([doc1, doc2]);

    let hasNextPage = false;
    for (const block of result.body) {
      if (block.type === "paragraph" && block.properties?.sectionProperties) {
        if (block.properties.sectionProperties.breakType === "nextPage") {
          hasNextPage = true;
        }
      }
    }
    expect(hasNextPage).toBe(true);
  });

  it("preserves styles from first document", () => {
    const style1 = { styleId: "Heading1", name: "Heading 1", type: "paragraph" };
    const doc1 = createDoc(["First"], { styles: [style1 as any] });
    const doc2 = createDoc(["Second"]);
    const result = Io.merge([doc1, doc2]);

    expect(result.styles).toBeDefined();
    expect(result.styles!.some((s: any) => s.styleId === "Heading1")).toBe(true);
  });

  it("deduplicates styles by styleId", () => {
    const style = { styleId: "Normal", name: "Normal", type: "paragraph" };
    const doc1 = createDoc(["First"], { styles: [style as any] });
    const doc2 = createDoc(["Second"], { styles: [style as any] });
    const result = Io.merge([doc1, doc2]);

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
    const result = Io.merge([doc1, doc2]);

    expect(result.images?.length).toBe(1);
  });

  it("merges different images from both documents", () => {
    const img1 = { fileName: "image1.png", data: new Uint8Array([1]), mediaType: "image/png" };
    const img2 = { fileName: "image2.png", data: new Uint8Array([2]), mediaType: "image/png" };
    const doc1 = createDoc(["First"], { images: [img1 as any] });
    const doc2 = createDoc(["Second"], { images: [img2 as any] });
    const result = Io.merge([doc1, doc2]);

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

    const merged = Io.merge([doc1, doc2]);

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

    Io.merge([doc1WithList, doc2]);

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

    const merged = Io.merge([doc1, doc2]);
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

    const merged = Io.merge([doc1, doc2]);
    // Both footnotes survived.
    expect(merged.footnotes!.length).toBe(2);
    // doc2's footnote got a fresh id (3), and doc2's body ref now points there.
    const fn2 = merged.footnotes![1]!;
    expect(fn2.id).not.toBe(2);

    const doc2BodyPara = merged.body[merged.body.length - 1] as any;
    expect(doc2BodyPara.children[0].content[0].id).toBe(fn2.id);
  });
});
