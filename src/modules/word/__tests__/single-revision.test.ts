/**
 * DOCX Module - Single Revision API Tests
 */

import { describe, it, expect } from "vitest";

import { listRevisions, acceptRevision, rejectRevision } from "../index";
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

function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

function createDoc(paragraphs: Paragraph[]): DocxDocument {
  return { body: paragraphs } as unknown as DocxDocument;
}

function extractText(doc: DocxDocument): string {
  let text = "";
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      for (const child of block.children) {
        if ("content" in child && Array.isArray((child as Run).content)) {
          for (const c of (child as Run).content) {
            if (c.type === "text") {
              text += c.text;
            }
          }
        }
      }
    }
  }
  return text;
}

describe("listRevisions", () => {
  it("returns empty array for document without revisions", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("plain")] }]);
    expect(listRevisions(doc)).toEqual([]);
  });

  it("lists insertion revisions", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("new"),
      revision: { author: "Alice", id: 1, date: "2024-01-01" }
    };
    const doc = createDoc([{ type: "paragraph", children: [ins as ParagraphChild] } as Paragraph]);
    const revs = listRevisions(doc);
    expect(revs).toHaveLength(1);
    expect(revs[0]).toEqual({
      id: 1,
      type: "insert",
      author: "Alice",
      date: "2024-01-01"
    });
  });

  it("lists deletion revisions", () => {
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("old"),
      revision: { author: "Bob", id: 2 }
    };
    const doc = createDoc([{ type: "paragraph", children: [del as ParagraphChild] } as Paragraph]);
    const revs = listRevisions(doc);
    expect(revs).toHaveLength(1);
    expect(revs[0].type).toBe("delete");
    expect(revs[0].id).toBe(2);
    expect(revs[0].author).toBe("Bob");
  });

  it("lists move revisions", () => {
    const mf: MovedFromRun = {
      type: "movedFromRun",
      run: textRun("from"),
      revision: { author: "Carol", id: 3 }
    };
    const mt: MovedToRun = {
      type: "movedToRun",
      run: textRun("to"),
      revision: { author: "Carol", id: 4 }
    };
    const doc = createDoc([
      { type: "paragraph", children: [mf as ParagraphChild, mt as ParagraphChild] } as Paragraph
    ]);
    const revs = listRevisions(doc);
    expect(revs).toHaveLength(2);
    expect(revs.map(r => r.type).sort()).toEqual(["moveFrom", "moveTo"]);
  });

  it("deduplicates revisions with same type+id+author+date", () => {
    const ins1: InsertedRun = {
      type: "insertedRun",
      run: textRun("a"),
      revision: { author: "Alice", id: 1, date: "2024-01-01" }
    };
    const ins2: InsertedRun = {
      type: "insertedRun",
      run: textRun("b"),
      revision: { author: "Alice", id: 1, date: "2024-01-01" }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [ins1 as ParagraphChild, ins2 as ParagraphChild]
      } as Paragraph
    ]);
    const revs = listRevisions(doc);
    expect(revs).toHaveLength(1);
  });

  it("lists multiple distinct revisions", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("new"),
      revision: { author: "Alice", id: 1 }
    };
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("old"),
      revision: { author: "Bob", id: 2 }
    };
    const doc = createDoc([
      { type: "paragraph", children: [ins as ParagraphChild] } as Paragraph,
      { type: "paragraph", children: [del as ParagraphChild] } as Paragraph
    ]);
    expect(listRevisions(doc)).toHaveLength(2);
  });
});

describe("acceptRevision", () => {
  it("accepts a single insert by id", () => {
    const ins1: InsertedRun = {
      type: "insertedRun",
      run: textRun("first"),
      revision: { author: "A", id: 1 }
    };
    const ins2: InsertedRun = {
      type: "insertedRun",
      run: textRun("second"),
      revision: { author: "A", id: 2 }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [ins1 as ParagraphChild, ins2 as ParagraphChild]
      } as Paragraph
    ]);

    const ok = acceptRevision(doc, 1);
    expect(ok).toBe(true);
    // After accepting id=1, "first" should be a plain run; id=2 still as InsertedRun
    expect(extractText(doc)).toContain("first");
    // The second insert should remain pending
    const remaining = listRevisions(doc);
    expect(remaining).toHaveLength(1);
    expect(remaining[0].id).toBe(2);
  });

  it("accepts a single delete by id (removes content)", () => {
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("removed"),
      revision: { author: "A", id: 5 }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("kept "), del as ParagraphChild]
      } as Paragraph
    ]);

    const ok = acceptRevision(doc, 5);
    expect(ok).toBe(true);
    expect(extractText(doc)).toBe("kept ");
  });

  it("returns false for non-existent revision id", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("text"),
      revision: { author: "A", id: 1 }
    };
    const doc = createDoc([{ type: "paragraph", children: [ins as ParagraphChild] } as Paragraph]);

    expect(acceptRevision(doc, 999)).toBe(false);
  });
});

describe("rejectRevision", () => {
  it("rejects a single insert (removes content)", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("inserted"),
      revision: { author: "A", id: 1 }
    };
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("base "), ins as ParagraphChild]
      } as Paragraph
    ]);

    const ok = rejectRevision(doc, 1);
    expect(ok).toBe(true);
    expect(extractText(doc)).toBe("base ");
  });

  it("rejects a single delete (restores content)", () => {
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("restore"),
      revision: { author: "A", id: 7 }
    };
    const doc = createDoc([{ type: "paragraph", children: [del as ParagraphChild] } as Paragraph]);

    const ok = rejectRevision(doc, 7);
    expect(ok).toBe(true);
    expect(extractText(doc)).toBe("restore");
  });

  it("rejects a moveFrom (keeps original)", () => {
    const mf: MovedFromRun = {
      type: "movedFromRun",
      run: textRun("original"),
      revision: { author: "A", id: 10 }
    };
    const doc = createDoc([{ type: "paragraph", children: [mf as ParagraphChild] } as Paragraph]);

    rejectRevision(doc, 10);
    expect(extractText(doc)).toBe("original");
  });

  it("rejects a moveTo (removes content)", () => {
    const mt: MovedToRun = {
      type: "movedToRun",
      run: textRun("moved"),
      revision: { author: "A", id: 11 }
    };
    const doc = createDoc([{ type: "paragraph", children: [mt as ParagraphChild] } as Paragraph]);

    rejectRevision(doc, 11);
    expect(extractText(doc)).not.toContain("moved");
  });
});
