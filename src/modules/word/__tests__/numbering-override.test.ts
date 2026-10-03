/**
 * Numbering resolution through `w:lvlOverride` and paragraph styles.
 *
 * Every converter resolves numId + ilvl through one helper
 * (`resolveNumberingLevelDef`) and a paragraph's numbering through another
 * (`resolveParagraphNumbering`). These tests pin that each surface honours a
 * level replacement, a `w:startOverride`, style-inherited numbering and
 * `numId` 0.
 */

import { docxToSemantic } from "@word/convert/docx-to-semantic";
import { renderToHtml } from "@word/convert/html/html-renderer";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { readOdt, writeOdt } from "@word/convert/odt/odt";
import { resolveNumberingLevelDef, resolveParagraphNumbering } from "@word/query/style-resolve";
import type { DocxDocument, NumberingRef, Paragraph, Run } from "@word/types";
import { describe, it, expect } from "vitest";

function para(text: string, numbering?: NumberingRef, style?: string): Paragraph {
  return {
    type: "paragraph",
    properties: { ...(numbering ? { numbering } : {}), ...(style ? { style } : {}) },
    children: [{ content: [{ type: "text", text }] } as Run]
  };
}

/**
 * abstractNum 1 is a bullet list. numId 1 uses it as-is; numId 2 replaces
 * level 0 with a decimal level and restarts it at 5 via startOverride.
 */
function makeDoc(body: DocxDocument["body"], styles: DocxDocument["styles"] = []): DocxDocument {
  return {
    body,
    styles,
    abstractNumberings: [
      { abstractNumId: 1, levels: [{ level: 0, format: "bullet", text: "•", start: 1 }] }
    ],
    numberingInstances: [
      { numId: 1, abstractNumId: 1 },
      {
        numId: 2,
        abstractNumId: 1,
        overrides: [
          {
            level: 0,
            startOverride: 5,
            levelDef: { level: 0, format: "decimal", text: "%1.", start: 1 }
          }
        ]
      }
    ],
    headers: new Map(),
    footers: new Map(),
    footnotes: [],
    endnotes: [],
    comments: [],
    images: [],
    fonts: [],
    embeddedFonts: [],
    customXmlParts: [],
    customProperties: [],
    opaqueParts: []
  } as unknown as DocxDocument;
}

const overridden = (): DocxDocument =>
  makeDoc([para("five", { numId: 2, level: 0 }), para("six", { numId: 2, level: 0 })]);

describe("numbering resolution helpers", () => {
  it("applies the level replacement and startOverride", () => {
    const def = resolveNumberingLevelDef(overridden(), 2, 0);
    expect(def?.format).toBe("decimal");
    expect(def?.start).toBe(5);
    expect(resolveNumberingLevelDef(overridden(), 1, 0)?.format).toBe("bullet");
    expect(resolveNumberingLevelDef(overridden(), 0, 0)).toBeUndefined();
  });

  it("inherits numbering from the style chain and treats numId 0 as none", () => {
    const doc = makeDoc([], [
      {
        type: "paragraph",
        styleId: "Base",
        name: "Base",
        paragraphProperties: { numbering: { numId: 2, level: 0 } }
      },
      { type: "paragraph", styleId: "Child", name: "Child", basedOn: "Base" }
    ] as unknown as DocxDocument["styles"]);
    expect(resolveParagraphNumbering(doc, para("x", undefined, "Child"))).toEqual({
      numId: 2,
      level: 0
    });
    expect(resolveParagraphNumbering(doc, para("x", { numId: 0, level: 0 }, "Child"))).toBe(
      undefined
    );
  });
});

describe("converters honour w:lvlOverride", () => {
  it("HTML renders an ordered list starting at the override", () => {
    const { html } = renderToHtml(overridden());
    expect(html).toContain('<ol start="5">');
    expect(html).not.toContain("<ul>");
  });

  it("Markdown renders ordered markers starting at the override", () => {
    const md = renderToMarkdown(overridden());
    expect(md).toContain("5. five");
    expect(md).toContain("6. six");
    expect(md).not.toMatch(/^- /m);
  });

  it("semantic IR emits an ordered list with the start value", () => {
    const { document } = docxToSemantic(overridden());
    const list = document.blocks.find(b => b.type === "list");
    expect(list).toMatchObject({ type: "list", ordered: true, start: 5 });
  });

  it("ODT writes a numbered list style with the start value", async () => {
    const bytes = await writeOdt(overridden());
    const { extractAll } = await import("@archive/unzip/extract");
    const entries = await extractAll(bytes);
    const decoder = new TextDecoder();
    const xml = ["content.xml", "styles.xml"]
      .map(n => (entries.get(n) ? decoder.decode(entries.get(n)!.data) : ""))
      .join("");
    expect(xml).toMatch(
      /<text:list-level-style-number[^>]*text:level="1"[^>]*text:start-value="5"/
    );

    const restored = await readOdt(bytes);
    const p = restored.body.find(b => b.type === "paragraph") as Paragraph;
    const num = p.properties!.numbering!;
    expect(resolveNumberingLevelDef(restored, num.numId, num.level)?.format).toBe("decimal");
  });

  it("style-inherited numbering makes a list item in every converter", async () => {
    const doc = makeDoc([para("styled", undefined, "ListNum")], [
      {
        type: "paragraph",
        styleId: "ListNum",
        name: "ListNum",
        paragraphProperties: { numbering: { numId: 2, level: 0 } }
      }
    ] as unknown as DocxDocument["styles"]);
    expect(renderToHtml(doc).html).toContain('<ol start="5">');
    expect(renderToMarkdown(doc)).toContain("5. styled");
    expect(docxToSemantic(doc).document.blocks[0]).toMatchObject({ type: "list", ordered: true });
    const restored = await readOdt(await writeOdt(doc));
    const p = restored.body.find(b => b.type === "paragraph") as Paragraph;
    expect(p.properties?.numbering).toBeDefined();
  });
});
