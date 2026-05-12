/**
 * DOCX Module - Walker Tests (stop/skip/section tracking)
 */

import { describe, it, expect } from "vitest";

import { mapDocument } from "../core/mapper";
import { walkDocument, collectParagraphs, collectRuns } from "../core/walker";
import type { DocxDocument, Paragraph, Run, Table, BodyContent } from "../types";

// Helpers
function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

function para(t: string, props?: Paragraph["properties"]): Paragraph {
  return { type: "paragraph", children: [textRun(t)], properties: props };
}

function createDoc(body: BodyContent[]): DocxDocument {
  return { body, contentTypes: [] } as unknown as DocxDocument;
}

describe("walkDocument", () => {
  describe("stop action", () => {
    it("stops traversal when enterParagraph returns 'stop'", () => {
      const doc = createDoc([para("first"), para("second"), para("third")]);

      const visited: string[] = [];
      walkDocument(doc, {
        enterParagraph(p) {
          const text =
            p.children[0] && "content" in p.children[0]
              ? (p.children[0] as Run).content[0]?.type === "text"
                ? ((p.children[0] as Run).content[0] as any).text
                : ""
              : "";
          visited.push(text);
          if (text === "second") {
            return "stop";
          }
          return "continue";
        }
      });

      expect(visited).toEqual(["first", "second"]);
    });

    it("stops traversal when enterRun returns 'stop'", () => {
      const doc = createDoc([
        { type: "paragraph", children: [textRun("a"), textRun("b"), textRun("c")] }
      ]);

      let runCount = 0;
      walkDocument(doc, {
        enterRun() {
          runCount++;
          if (runCount === 2) {
            return "stop";
          }
          return "continue";
        }
      });

      expect(runCount).toBe(2);
    });
  });

  describe("skip action", () => {
    it("skips paragraph children when enterParagraph returns 'skip'", () => {
      const doc = createDoc([
        { type: "paragraph", children: [textRun("skipped")] },
        { type: "paragraph", children: [textRun("visited")] }
      ]);

      const runs: string[] = [];
      let paraCount = 0;
      walkDocument(doc, {
        enterParagraph() {
          paraCount++;
          if (paraCount === 1) {
            return "skip";
          }
          return "continue";
        },
        enterRun(run) {
          if (run.content[0]?.type === "text") {
            runs.push((run.content[0] as any).text);
          }
          return "continue";
        }
      });

      expect(runs).toEqual(["visited"]);
    });

    it("skips table content when enterTable returns 'skip'", () => {
      const table: Table = {
        type: "table",
        rows: [
          {
            cells: [
              {
                content: [para("inside table")]
              }
            ]
          }
        ]
      } as any;

      const doc = createDoc([table, para("after table")]);
      const visited: string[] = [];

      walkDocument(doc, {
        enterTable() {
          return "skip";
        },
        enterParagraph(p) {
          const text =
            p.children[0] && "content" in p.children[0]
              ? (((p.children[0] as Run).content[0] as any)?.text ?? "")
              : "";
          visited.push(text);
          return "continue";
        }
      });

      expect(visited).toEqual(["after table"]);
    });
  });

  describe("section tracking", () => {
    it("increments section when paragraph has sectionProperties", () => {
      const doc = createDoc([
        para("section 0 content"),
        para("section break", {
          sectionProperties: { breakType: "nextPage" }
        }),
        para("section 1 content"),
        para("another break", {
          sectionProperties: { breakType: "continuous" }
        }),
        para("section 2 content")
      ]);

      const sections: number[] = [];
      walkDocument(doc, {
        enterParagraph(_p, path) {
          sections.push(path.section);
          return "continue";
        }
      });

      expect(sections).toEqual([0, 0, 1, 1, 2]);
    });

    it("starts at section 0 for document without section breaks", () => {
      const doc = createDoc([para("a"), para("b"), para("c")]);

      const sections: number[] = [];
      walkDocument(doc, {
        enterParagraph(_p, path) {
          sections.push(path.section);
          return "continue";
        }
      });

      expect(sections).toEqual([0, 0, 0]);
    });
  });

  describe("options", () => {
    it("skips headers when includeHeaders is false", () => {
      const doc = createDoc([para("body")]);
      (doc as any).headers = new Map([
        [
          "rId1",
          {
            type: "default",
            content: { children: [para("header")] }
          }
        ]
      ]);

      const visited: string[] = [];
      walkDocument(
        doc,
        {
          enterParagraph(p) {
            const text =
              p.children[0] && "content" in p.children[0]
                ? (((p.children[0] as Run).content[0] as any)?.text ?? "")
                : "";
            visited.push(text);
            return "continue";
          }
        },
        { includeHeaders: false }
      );

      expect(visited).toEqual(["body"]);
    });

    it("includes comments when includeComments is true", () => {
      const doc = createDoc([para("body")]);
      (doc as any).comments = [{ id: 1, author: "User", content: [para("comment text")] }];

      const visited: string[] = [];
      walkDocument(
        doc,
        {
          enterParagraph(p) {
            const text =
              p.children[0] && "content" in p.children[0]
                ? (((p.children[0] as Run).content[0] as any)?.text ?? "")
                : "";
            visited.push(text);
            return "continue";
          }
        },
        { includeComments: true }
      );

      expect(visited).toContain("comment text");
    });
  });
});

describe("collectParagraphs", () => {
  it("collects all paragraphs including nested in tables", () => {
    const table: Table = {
      type: "table",
      rows: [
        {
          cells: [
            {
              content: [para("in cell")]
            }
          ]
        }
      ]
    } as any;

    const doc = createDoc([para("top"), table, para("bottom")]);
    const paragraphs = collectParagraphs(doc);

    expect(paragraphs.length).toBe(3);
  });
});

describe("collectRuns", () => {
  it("collects all runs", () => {
    const doc = createDoc([
      { type: "paragraph", children: [textRun("a"), textRun("b")] },
      { type: "paragraph", children: [textRun("c")] }
    ]);

    const runs = collectRuns(doc);
    expect(runs.length).toBe(3);
  });
});

describe("mapDocument", () => {
  it("transforms paragraphs", () => {
    const doc = createDoc([para("hello"), para("world")]);

    const result = mapDocument(doc, {
      transformParagraph(p) {
        const text =
          p.children[0] && "content" in p.children[0]
            ? (((p.children[0] as Run).content[0] as any)?.text ?? "")
            : "";
        if (text === "hello") {
          return null; // Remove
        }
        return p;
      }
    });

    const paragraphs = collectParagraphs(result);
    expect(paragraphs.length).toBe(1);
  });

  it("transforms runs", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("keep"), textRun("remove")] }]);

    const result = mapDocument(doc, {
      transformRun(run) {
        if (run.content[0]?.type === "text" && (run.content[0] as any).text === "remove") {
          return null;
        }
        return run;
      }
    });

    const runs = collectRuns(result);
    expect(runs.length).toBe(1);
  });

  it("never mutates the input document", () => {
    const doc = createDoc([para("original")]);
    const originalBody = doc.body;

    mapDocument(doc, {
      transformParagraph() {
        return null;
      }
    });

    expect(doc.body).toBe(originalBody);
    expect(doc.body.length).toBe(1);
  });

  it("transformBodyContent: replaces a single block with a single block", () => {
    const doc = createDoc([para("a"), para("b")]);

    const result = mapDocument(doc, {
      transformBodyContent(block) {
        if (block.type === "paragraph") {
          const first = block.children[0] as Run;
          const t = (first.content[0] as { text: string }).text;
          if (t === "a") {
            return para("A-replaced");
          }
        }
        return block;
      }
    });

    const paragraphs = collectParagraphs(result);
    expect(paragraphs.length).toBe(2);
    const t0 = ((paragraphs[0].children[0] as Run).content[0] as { text: string }).text;
    expect(t0).toBe("A-replaced");
  });

  it("transformBodyContent: flat-maps a single block into multiple", () => {
    const doc = createDoc([para("split-me"), para("keep")]);

    const result = mapDocument(doc, {
      transformBodyContent(block) {
        if (block.type === "paragraph") {
          const first = block.children[0] as Run;
          const t = (first.content[0] as { text: string }).text;
          if (t === "split-me") {
            return [para("part-1"), para("part-2"), para("part-3")];
          }
        }
        return block;
      }
    });

    const paragraphs = collectParagraphs(result);
    expect(paragraphs.length).toBe(4);
    const texts = paragraphs.map(p => ((p.children[0] as Run).content[0] as { text: string }).text);
    expect(texts).toEqual(["part-1", "part-2", "part-3", "keep"]);
  });

  it("transformBodyContent: returning null removes the block", () => {
    const doc = createDoc([para("a"), para("b"), para("c")]);

    const result = mapDocument(doc, {
      transformBodyContent(block) {
        if (block.type === "paragraph") {
          const first = block.children[0] as Run;
          const t = (first.content[0] as { text: string }).text;
          if (t === "b") {
            return null;
          }
        }
        return block;
      }
    });

    const paragraphs = collectParagraphs(result);
    const texts = paragraphs.map(p => ((p.children[0] as Run).content[0] as { text: string }).text);
    expect(texts).toEqual(["a", "c"]);
  });

  it("transformTable: recurses into cells regardless of return value", () => {
    const tbl: Table = {
      type: "table",
      rows: [
        {
          cells: [{ content: [para("cell-text")] }, { content: [para("other-text")] }]
        }
      ]
    };
    const doc = createDoc([tbl]);

    let runsSeen = 0;
    const result = mapDocument(doc, {
      transformTable(t) {
        // Return as-is; mapper should still recurse into cells.
        return t;
      },
      transformRun(run) {
        runsSeen++;
        return run;
      }
    });

    expect(runsSeen).toBe(2);
    expect(collectParagraphs(result).length).toBe(2);
  });

  it("transformRunContent: filters individual run content nodes", () => {
    const para1: Paragraph = {
      type: "paragraph",
      children: [
        {
          content: [{ type: "text", text: "keep" }, { type: "tab" }, { type: "text", text: "drop" }]
        } as Run
      ]
    };
    const doc = createDoc([para1]);

    const result = mapDocument(doc, {
      transformRunContent(rc) {
        if (rc.type === "tab") {
          return null;
        }
        return rc;
      }
    });

    const runs = collectRuns(result);
    expect(runs.length).toBe(1);
    expect(runs[0].content.length).toBe(2);
    expect(runs[0].content.every(c => c.type === "text")).toBe(true);
  });

  it("recurses into nested tables", () => {
    const innerTable: Table = {
      type: "table",
      rows: [{ cells: [{ content: [para("inner")] }] }]
    };
    const outerTable: Table = {
      type: "table",
      rows: [{ cells: [{ content: [para("outer-cell"), innerTable] }] }]
    };
    const doc = createDoc([outerTable]);

    const seen: string[] = [];
    mapDocument(doc, {
      transformParagraph(p) {
        const t = ((p.children[0] as Run).content[0] as { text: string }).text;
        seen.push(t);
        return p;
      }
    });

    expect(seen).toEqual(["outer-cell", "inner"]);
  });
});

describe("walker coverage of mixed body content", () => {
  it("walks Run children of an inline (run-only) SDT", () => {
    // SDT.content allows (Paragraph | Run | Table)[]. Inline content
    // controls put a Run directly inside the SDT. Previously the walker
    // dropped these.
    const sdt = {
      type: "sdt",
      content: [textRun("inside-sdt")],
      properties: {}
    } as unknown as BodyContent;
    const doc = createDoc([sdt]);
    const runs: string[] = [];
    walkDocument(doc, {
      enterRun(r) {
        const t = ((r as Run).content[0] as { text: string }).text;
        runs.push(t);
      }
    });
    expect(runs).toContain("inside-sdt");
  });

  it("walks paragraphs cached on a TableOfContents node", () => {
    // TOC.cachedParagraphs is the rendered fallback content; text-aware
    // visitors should see it so search/replace doesn't lose them.
    const toc = {
      type: "tableOfContents",
      cachedParagraphs: [para("toc-cached")]
    } as unknown as BodyContent;
    const doc = createDoc([toc]);
    const seen: string[] = [];
    walkDocument(doc, {
      enterParagraph(p) {
        const t = ((p.children[0] as Run).content[0] as { text: string }).text;
        seen.push(t);
      }
    });
    expect(seen).toContain("toc-cached");
  });

  it("walks DrawingShape.textContent paragraphs", () => {
    const shape = {
      type: "drawingShape",
      textContent: [para("shape-text")]
    } as unknown as BodyContent;
    const doc = createDoc([shape]);
    const seen: string[] = [];
    walkDocument(doc, {
      enterParagraph(p) {
        const t = ((p.children[0] as Run).content[0] as { text: string }).text;
        seen.push(t);
      }
    });
    expect(seen).toContain("shape-text");
  });
});

describe("mapDocument coverage of textBox / TOC / DrawingShape", () => {
  it("transforms paragraphs inside a TextBox", () => {
    const tb = {
      type: "textBox",
      content: [para("orig-tb")]
    } as unknown as BodyContent;
    const doc = createDoc([tb]);
    const result = mapDocument(doc, {
      transformParagraph(p) {
        const newRun: Run = { content: [{ type: "text", text: "MAPPED" }] };
        return { ...p, children: [newRun] };
      }
    });
    const newTb = result.body[0] as unknown as { content: Paragraph[] };
    const txt = ((newTb.content[0].children[0] as Run).content[0] as { text: string }).text;
    expect(txt).toBe("MAPPED");
  });

  it("transforms paragraphs cached on a TableOfContents", () => {
    const toc = {
      type: "tableOfContents",
      cachedParagraphs: [para("orig-toc")]
    } as unknown as BodyContent;
    const doc = createDoc([toc]);
    const result = mapDocument(doc, {
      transformParagraph(p) {
        const newRun: Run = { content: [{ type: "text", text: "MAPPED" }] };
        return { ...p, children: [newRun] };
      }
    });
    const newToc = result.body[0] as unknown as { cachedParagraphs: Paragraph[] };
    const txt = ((newToc.cachedParagraphs[0].children[0] as Run).content[0] as { text: string })
      .text;
    expect(txt).toBe("MAPPED");
  });

  it("transforms paragraphs inside a DrawingShape.textContent", () => {
    const shape = {
      type: "drawingShape",
      textContent: [para("orig-shape")]
    } as unknown as BodyContent;
    const doc = createDoc([shape]);
    const result = mapDocument(doc, {
      transformParagraph(p) {
        const newRun: Run = { content: [{ type: "text", text: "MAPPED" }] };
        return { ...p, children: [newRun] };
      }
    });
    const newShape = result.body[0] as unknown as { textContent: Paragraph[] };
    const txt = ((newShape.textContent[0].children[0] as Run).content[0] as { text: string }).text;
    expect(txt).toBe("MAPPED");
  });
});
