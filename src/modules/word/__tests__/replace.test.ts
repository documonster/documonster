/**
 * DOCX Module - Replace API Tests
 */

import { describe, it, expect } from "vitest";

import { replaceText } from "../index";
import type { DocxDocument, Paragraph, Run } from "../types";

// Helper to create a minimal document
function createDoc(paragraphs: Paragraph[]): DocxDocument {
  return {
    body: paragraphs,
    contentTypes: []
  } as unknown as DocxDocument;
}

// Helper to create a run with text
function textRun(t: string, props?: Run["properties"]): Run {
  return { content: [{ type: "text", text: t }], properties: props };
}

// Helper to extract all text from a document
function extractText(doc: DocxDocument): string {
  const lines: string[] = [];
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
      lines.push(t);
    }
  }
  return lines.join("\n");
}

describe("replaceText", () => {
  it("replaces simple text in a single run", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Hello World")] }]);
    const count = replaceText(doc, "World", "Earth");
    expect(count).toBe(1);
    expect(extractText(doc)).toBe("Hello Earth");
  });

  it("replaces multiple occurrences", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("foo bar foo baz foo")] }]);
    const count = replaceText(doc, "foo", "x");
    expect(count).toBe(3);
    expect(extractText(doc)).toBe("x bar x baz x");
  });

  it("returns 0 when no match found", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Hello World")] }]);
    const count = replaceText(doc, "xyz", "abc");
    expect(count).toBe(0);
    expect(extractText(doc)).toBe("Hello World");
  });

  it("replaces text across multiple runs (cross-run)", () => {
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("Hel"), textRun("lo Wor"), textRun("ld")]
      }
    ]);
    const count = replaceText(doc, "Hello World", "Hi Earth");
    expect(count).toBe(1);
    expect(extractText(doc)).toBe("Hi Earth");
  });

  it("supports regex replacement", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Date: 2024-01-15")] }]);
    const count = replaceText(doc, /\d{4}-\d{2}-\d{2}/, "REDACTED");
    expect(count).toBe(1);
    expect(extractText(doc)).toBe("Date: REDACTED");
  });

  it("supports regex with global flag", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("a1 b2 c3")] }]);
    const count = replaceText(doc, /\d/g, "X");
    expect(count).toBeGreaterThanOrEqual(1);
    expect(extractText(doc)).toBe("aX bX cX");
  });

  it("replaces text in multiple paragraphs", () => {
    const doc = createDoc([
      { type: "paragraph", children: [textRun("Hello World")] },
      { type: "paragraph", children: [textRun("Hello Again")] }
    ]);
    const count = replaceText(doc, "Hello", "Hi");
    expect(count).toBe(2);
    expect(extractText(doc)).toBe("Hi World\nHi Again");
  });

  it("replaces text in headers", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Body text")] }]);
    (doc as any).headers = new Map([
      [
        "rId1",
        {
          type: "default",
          content: {
            children: [{ type: "paragraph", children: [textRun("Header: {{name}}")] }]
          }
        }
      ]
    ]);
    const count = replaceText(doc, "{{name}}", "John");
    expect(count).toBe(1);
  });

  it("replaces text in footers", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Body text")] }]);
    (doc as any).footers = new Map([
      [
        "rId2",
        {
          type: "default",
          content: {
            children: [{ type: "paragraph", children: [textRun("Footer: {{page}}")] }]
          }
        }
      ]
    ]);
    const count = replaceText(doc, "{{page}}", "1");
    expect(count).toBe(1);
  });

  it("replaces text inside table cells", () => {
    const doc = createDoc([
      {
        type: "table",
        rows: [
          {
            cells: [
              {
                content: [{ type: "paragraph", children: [textRun("cell text old")] }]
              }
            ]
          }
        ]
      } as any
    ]);
    const count = replaceText(doc, "old", "new");
    expect(count).toBe(1);
  });

  it("handles empty document gracefully", () => {
    const doc = createDoc([]);
    const count = replaceText(doc, "foo", "bar");
    expect(count).toBe(0);
  });

  it("handles empty string search", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Hello")] }]);
    // Empty string should match nothing or handle gracefully
    const count = replaceText(doc, "", "x");
    expect(count).toBe(0);
  });

  it("replaces with empty string (deletion)", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("Hello World")] }]);
    const count = replaceText(doc, " World", "");
    expect(count).toBe(1);
    expect(extractText(doc)).toBe("Hello");
  });

  // --- Regression: regex semantics & cross-run mixing ---

  it("regex without global flag replaces ALL occurrences and counts them accurately", () => {
    // Documented behavior: replaceText is a document-wide replacement, so a non-global
    // regex must still replace every match (mirroring the string-mode behavior) and
    // the returned count must equal the number of replacements actually performed.
    const doc = createDoc([{ type: "paragraph", children: [textRun("a1 b2 c3 d4")] }]);
    const count = replaceText(doc, /\d/, "X");
    expect(count).toBe(4);
    expect(extractText(doc)).toBe("aX bX cX dX");
  });

  it("regex with global flag returns the exact number of matches", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("aaa bbb aaa")] }]);
    const count = replaceText(doc, /aaa/g, "X");
    expect(count).toBe(2);
    expect(extractText(doc)).toBe("X bbb X");
  });

  it("regex with capture groups expands $1 in replacement", () => {
    const doc = createDoc([{ type: "paragraph", children: [textRun("v1.0 and v2.5 and v3.10")] }]);
    const count = replaceText(doc, /v(\d+)\.(\d+)/g, "$1-$2");
    expect(count).toBe(3);
    expect(extractText(doc)).toBe("1-0 and 2-5 and 3-10");
  });

  it("handles same paragraph with both per-run and cross-run matches", () => {
    // First match falls inside a single run ("foo"), second match spans two runs ("ba" + "r").
    const doc = createDoc([
      {
        type: "paragraph",
        children: [textRun("foo "), textRun("ba"), textRun("r baz")]
      }
    ]);
    const count = replaceText(doc, /foo|bar/g, "X");
    expect(count).toBe(2);
    expect(extractText(doc)).toBe("X X baz");
  });
});
