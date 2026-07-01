/**
 * DOCX Module - Single Revision API Tests
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

function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

function createDoc(paragraphs: Paragraph[]): DocxDocument {
  return { body: paragraphs };
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
    expect(Query.listRevisions(doc)).toEqual([]);
  });

  it("lists insertion revisions", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("new"),
      revision: { author: "Alice", id: 1, date: "2024-01-01" }
    };
    const doc = createDoc([{ type: "paragraph", children: [ins as ParagraphChild] } as Paragraph]);
    const revs = Query.listRevisions(doc);
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
    const revs = Query.listRevisions(doc);
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
    const revs = Query.listRevisions(doc);
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
    const revs = Query.listRevisions(doc);
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
    expect(Query.listRevisions(doc)).toHaveLength(2);
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

    const ok = Query.acceptRevision(doc, 1);
    expect(ok).toBe(true);
    // After accepting id=1, "first" should be a plain run; id=2 still as InsertedRun
    expect(extractText(doc)).toContain("first");
    // The second insert should remain pending
    const remaining = Query.listRevisions(doc);
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

    const ok = Query.acceptRevision(doc, 5);
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

    expect(Query.acceptRevision(doc, 999)).toBe(false);
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

    const ok = Query.rejectRevision(doc, 1);
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

    const ok = Query.rejectRevision(doc, 7);
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

    Query.rejectRevision(doc, 10);
    expect(extractText(doc)).toBe("original");
  });

  it("rejects a moveTo (removes content)", () => {
    const mt: MovedToRun = {
      type: "movedToRun",
      run: textRun("moved"),
      revision: { author: "A", id: 11 }
    };
    const doc = createDoc([{ type: "paragraph", children: [mt as ParagraphChild] } as Paragraph]);

    Query.rejectRevision(doc, 11);
    expect(extractText(doc)).not.toContain("moved");
  });
});

describe("listRevisions / accept / reject — coverage of notes & comments", () => {
  it("lists revisions inside footnotes", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("inserted-fn"),
      revision: { author: "FN", id: 50 }
    };
    const doc = {
      body: [{ type: "paragraph", children: [textRun("body")] } as Paragraph],
      footnotes: [
        { id: 1, content: [{ type: "paragraph", children: [ins as ParagraphChild] } as Paragraph] }
      ]
    } as unknown as DocxDocument;
    const revs = Query.listRevisions(doc);
    expect(revs).toHaveLength(1);
    expect(revs[0].id).toBe(50);
    expect(revs[0].author).toBe("FN");
  });

  it("lists revisions inside comments", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("inserted-cmt"),
      revision: { author: "Z", id: 51 }
    };
    const doc = {
      body: [{ type: "paragraph", children: [textRun("b")] } as Paragraph],
      comments: [
        {
          id: 1,
          author: "X",
          content: [{ type: "paragraph", children: [ins as ParagraphChild] } as Paragraph]
        }
      ]
    } as unknown as DocxDocument;
    const revs = Query.listRevisions(doc);
    expect(revs).toHaveLength(1);
    expect(revs[0].id).toBe(51);
  });

  it("acceptRevision touches footnote content", () => {
    const del: DeletedRun = {
      type: "deletedRun",
      run: textRun("doomed"),
      revision: { author: "FN", id: 52 }
    };
    const fnPara = { type: "paragraph", children: [del as ParagraphChild] } as Paragraph;
    const doc = {
      body: [{ type: "paragraph", children: [textRun("body")] } as Paragraph],
      footnotes: [{ id: 1, content: [fnPara] }]
    } as unknown as DocxDocument;
    const found = Query.acceptRevision(doc, 52);
    expect(found).toBe(true);
    // After accept on a deletion the run is gone
    expect(fnPara.children).toHaveLength(0);
  });

  it("rejectRevision touches comment content", () => {
    const ins: InsertedRun = {
      type: "insertedRun",
      run: textRun("temp"),
      revision: { author: "X", id: 53 }
    };
    const cmtPara = { type: "paragraph", children: [ins as ParagraphChild] } as Paragraph;
    const doc = {
      body: [{ type: "paragraph", children: [textRun("b")] } as Paragraph],
      comments: [{ id: 1, author: "X", content: [cmtPara] }]
    } as unknown as DocxDocument;
    const found = Query.rejectRevision(doc, 53);
    expect(found).toBe(true);
    // Reject on an insertion removes the run
    expect(cmtPara.children).toHaveLength(0);
  });
});
