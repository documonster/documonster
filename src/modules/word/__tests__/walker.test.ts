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
});
