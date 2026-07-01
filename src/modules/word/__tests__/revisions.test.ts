/**
 * DOCX Module - Revisions (Accept/Reject) API Tests
 */

import { describe, it, expect } from "vitest";

import { Query } from "../index";
import type {
  DocxDocument,
  Paragraph,
  ParagraphChild,
  Run,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun
} from "../types";

// Helper to create a minimal document
function createDoc(paragraphs: Paragraph[]): DocxDocument {
  return {
    body: paragraphs
  };
}

// Helper to create a run
function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

// Helper to extract text from paragraphs
function extractText(doc: DocxDocument): string {
  const lines: string[] = [];
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      let t = "";
      for (const child of block.children) {
        if ("content" in child && Array.isArray((child as Run).content)) {
          for (const c of (child as Run).content) {
            if (c.type === "text") {
              t += c.text;
            }
          }
        }
      }
      lines.push(t);
    }
  }
  return lines.join("\n");
}

describe("acceptAllRevisions", () => {
  it("accepts inserted runs (keeps content, removes wrapper)", () => {
    const insertedRun: InsertedRun = {
      type: "insertedRun",
      run: textRun("inserted text"),
      revision: { author: "User", id: 1, date: "2024-01-01" }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("before "), insertedRun as ParagraphChild, textRun(" after")]
      } as Paragraph
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).toContain("inserted text");
  });

  it("accepts deleted runs (removes content)", () => {
    const deletedRun: DeletedRun = {
      type: "deletedRun",
      run: textRun("deleted text"),
      revision: { author: "User", id: 2, date: "2024-01-01" }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("before "), deletedRun as ParagraphChild, textRun(" after")]
      } as Paragraph
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).not.toContain("deleted text");
    expect(extractText(doc)).toContain("before ");
    expect(extractText(doc)).toContain(" after");
  });

  it("accepts moved-to runs (keeps content)", () => {
    const movedTo: MovedToRun = {
      type: "movedToRun",
      run: textRun("moved text"),
      revision: { author: "User", id: 3, date: "2024-01-01" }
    };
    const doc = createDoc([
      { type: "paragraph", children: [movedTo as ParagraphChild] } as Paragraph
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).toContain("moved text");
  });

  it("accepts moved-from runs (removes content)", () => {
    const movedFrom: MovedFromRun = {
      type: "movedFromRun",
      run: textRun("original position"),
      revision: { author: "User", id: 4, date: "2024-01-01" }
    };
    const doc = createDoc([
      { type: "paragraph", children: [movedFrom as ParagraphChild] } as Paragraph
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).not.toContain("original position");
  });

  it("returns 0 for document without revisions", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("plain text")] }]);
    const count = Query.acceptAllRevisions(doc);
    expect(count).toBe(0);
  });

  it("handles multiple revisions in same paragraph", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("new"),
      revision: { author: "User", id: 1 }
    };
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("old"),
      revision: { author: "User", id: 2 }
    };
    const doc = createDoc([
      { type: "paragraph", children: [ins as ParagraphChild, del as ParagraphChild] } as Paragraph
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).toContain("new");
    expect(extractText(doc)).not.toContain("old");
  });

  it("handles revisions in table cells", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("cell insert"),
      revision: { author: "User", id: 1 }
    };
    const doc = createDoc([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                content: [{ type: "paragraph", children: [ins as ParagraphChild] }]
              }
            ]
          }
        ]
      } as any
    ]);

    const count = Query.acceptAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
  });
});

describe("rejectAllRevisions", () => {
  it("rejects inserted runs (removes content)", () => {
    const insertedRun: InsertedRun = {
      type: "insertedRun",
      run: textRun("inserted text"),
      revision: { author: "User", id: 1, date: "2024-01-01" }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("before "), insertedRun as ParagraphChild]
      } as Paragraph
    ]);

    const count = Query.rejectAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).not.toContain("inserted text");
    expect(extractText(doc)).toContain("before ");
  });

  it("rejects deleted runs (restores content)", () => {
    const deletedRun: DeletedRun = {
      type: "deletedRun",
      run: textRun("restored text"),
      revision: { author: "User", id: 2, date: "2024-01-01" }
    };
    const doc = createDoc([
      { type: "paragraph", children: [deletedRun as ParagraphChild] } as Paragraph
    ]);

    const count = Query.rejectAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).toContain("restored text");
  });

  it("rejects moved-from runs (restores original position)", () => {
    const movedFrom: MovedFromRun = {
      type: "movedFromRun",
      run: textRun("original"),
      revision: { author: "User", id: 3, date: "2024-01-01" }
    };
    const doc = createDoc([
      { type: "paragraph", children: [movedFrom as ParagraphChild] } as Paragraph
    ]);

    const count = Query.rejectAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).toContain("original");
  });

  it("rejects moved-to runs (removes content)", () => {
    const movedTo: MovedToRun = {
      type: "movedToRun",
      run: textRun("moved here"),
      revision: { author: "User", id: 4, date: "2024-01-01" }
    };
    const doc = createDoc([
      { type: "paragraph", children: [movedTo as ParagraphChild] } as Paragraph
    ]);

    const count = Query.rejectAllRevisions(doc);
    expect(count).toBeGreaterThan(0);
    expect(extractText(doc)).not.toContain("moved here");
  });

  it("returns 0 for document without revisions", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("plain text")] }]);
    const count = Query.rejectAllRevisions(doc);
    expect(count).toBe(0);
  });
});
