/**
 * Template engine — block vs inline directive classification & error handling.
 *
 * Covers regressions that previously slipped through:
 *  - paragraphs containing INLINE `{{#if}}...{{/if}}` were misclassified as
 *    block-level directives because the detector used `text.includes(...)`.
 *  - unclosed inline `{{#if}}` / `{{#each}}` silently produced wrong output
 *    instead of throwing in strict mode.
 */

import { describe, it, expect } from "vitest";

import { TemplateError } from "../errors";
import { fillTemplate } from "../template/template-engine";
import type { DocxDocument, Paragraph, Run } from "../types";

function paraWithText(text: string): Paragraph {
  return {
    type: "paragraph",
    children: [{ content: [{ type: "text", text }] } as Run]
  };
}

function readPara(p: Paragraph): string {
  return p.children
    .map(c => {
      const r = c as Run;
      return r.content.map(rc => (rc.type === "text" ? rc.text : "")).join("");
    })
    .join("");
}

function makeDoc(text: string): DocxDocument {
  return {
    body: [paraWithText(text)],
    contentTypes: []
  } as unknown as DocxDocument;
}

describe("template engine — inline directives", () => {
  it("evaluates inline {{#if}}...{{/if}} inside a normal paragraph", () => {
    const doc = makeDoc("Hello {{#if show}}World{{/if}}!");
    const result = fillTemplate(doc, { show: true });
    expect(readPara(result.body[0] as Paragraph)).toBe("Hello World!");
  });

  it("inline {{#if}} renders falsy branch correctly", () => {
    const doc = makeDoc("Status: {{#if active}}on{{else}}off{{/if}}");
    const result = fillTemplate(doc, { active: false });
    expect(readPara(result.body[0] as Paragraph)).toBe("Status: off");
  });

  it("evaluates inline {{#each}} inside a normal paragraph", () => {
    const doc = makeDoc("Tags: {{#each tags}}[{{.}}]{{/each}}");
    const result = fillTemplate(doc, { tags: ["a", "b", "c"] });
    expect(readPara(result.body[0] as Paragraph)).toBe("Tags: [a][b][c]");
  });

  it("does not misclassify inline directive as block-level", () => {
    // Previously this triggered "Unclosed {{#if show}}" because
    // hasBlockDirective() matched `text.includes(...)`.
    const doc: DocxDocument = {
      body: [paraWithText("Pre {{#if show}}X{{/if}} Post"), paraWithText("After")],
      contentTypes: []
    } as unknown as DocxDocument;
    const result = fillTemplate(doc, { show: true });
    expect(readPara(result.body[0] as Paragraph)).toBe("Pre X Post");
    expect(readPara(result.body[1] as Paragraph)).toBe("After");
  });

  it("throws on unclosed inline {{#if}} in strict mode", () => {
    const doc = makeDoc("Hello {{#if x}}A");
    expect(() => fillTemplate(doc, { x: true })).toThrow(TemplateError);
  });

  it("throws on unclosed inline {{#each}} in strict mode", () => {
    const doc = makeDoc("List {{#each items}}A");
    expect(() => fillTemplate(doc, { items: [1, 2] })).toThrow(TemplateError);
  });
});

describe("template engine — block directive whole-paragraph rule", () => {
  it("treats a paragraph that is *only* a block directive as block-level", () => {
    const doc: DocxDocument = {
      body: [
        paraWithText("intro"),
        paraWithText("{{#each xs}}"),
        paraWithText("- {{.}}"),
        paraWithText("{{/each}}"),
        paraWithText("done")
      ],
      contentTypes: []
    } as unknown as DocxDocument;
    const result = fillTemplate(doc, { xs: [1, 2] });
    const texts = (result.body as Paragraph[]).map(readPara);
    // intro + 2 iterations of "- {value}" + done
    expect(texts).toEqual(["intro", "- 1", "- 2", "done"]);
  });
});
