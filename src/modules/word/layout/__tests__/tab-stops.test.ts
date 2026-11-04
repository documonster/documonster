import { Build, Document } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { PositionedRun } from "@word/layout/layout-model";
import { createStyleResolver } from "@word/query/style-resolver";
import type { DocxDocument, ParagraphProperties } from "@word/types";
import { describe, expect, it } from "vitest";

/**
 * `w:tab` advances to the paragraph's next stop (pages2 R7j); it used to be
 * laid out as four spaces, so a right-aligned page number in a table of
 * contents sat just after its title.
 */

function firstLine(props: ParagraphProperties, ...texts: string[]): PositionedRun[] {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  const children = texts.flatMap((t, i) =>
    i === 0 ? [Build.text(t)] : [Build.tab(), Build.text(t)]
  );
  Document.addParagraph(doc, Build.paragraph(children, props));
  const block = layoutDocumentFull(doc).pages[0]!.content[0]!;
  if (block.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  return block.lines[0]!.runs as PositionedRun[];
}

describe("tab stops", () => {
  it("ends right-tab text at the stop and draws its leader up to it", () => {
    const runs = firstLine(
      { tabs: [{ type: "right", position: 7200, leader: "dot" }] },
      "Title",
      "12"
    );
    const leader = runs.find(r => r.text.startsWith("."))!;
    const page = runs.at(-1)!;
    expect(page.text).toBe("12");
    expect(page.x + page.width).toBeCloseTo(360, 3);
    expect(leader.x + leader.width).toBeCloseTo(page.x, 3);
  });

  it("centres on a centre stop and lines the point up on a decimal stop", () => {
    const runs = firstLine(
      {
        tabs: [
          { type: "center", position: 2880 },
          { type: "decimal", position: 5760 }
        ]
      },
      "a",
      "mid",
      "12.5"
    );
    const mid = runs.find(r => r.text === "mid")!;
    expect(mid.x + mid.width / 2).toBeCloseTo(144, 3);
    const num = runs.find(r => r.text === "12.5")!;
    expect(num.x + num.width - num.width * (2 / 4)).toBeGreaterThan(280);
    expect(num.x).toBeLessThan(288);
  });

  it("uses the default stops when none is set", () => {
    const runs = firstLine({}, "a", "b");
    expect(runs.find(r => r.text === "b")!.x).toBeCloseTo(36, 3);
  });

  it("moves a tab with no stop left on its line to the next one", () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    Document.addParagraph(
      doc,
      Build.paragraph(
        [Build.text("a"), Build.tab(), Build.text("b"), Build.tab(), Build.text("c")],
        { tabs: [{ type: "right", position: 9360 }] }
      )
    );
    const block = layoutDocumentFull(doc).pages[0]!.content[0]!;
    expect(block.type === "paragraph" && block.lines.length).toBe(2);
  });
});

describe("tab stops in the style hierarchy", () => {
  it("add to the style's stops, and a clear removes one", () => {
    const doc: DocxDocument = Document.create();
    Document.useDefaultStyles(doc);
    Document.addStyle(doc, {
      type: "paragraph",
      styleId: "Tabbed",
      name: "Tabbed",
      paragraphProperties: {
        tabs: [
          { type: "center", position: 4680 },
          { type: "right", position: 9360 }
        ]
      }
    } as never);
    const para = Build.paragraph([], {
      style: "Tabbed",
      tabs: [
        { type: "left", position: 1440 },
        { type: "clear", position: 4680 }
      ]
    });
    const tabs = createStyleResolver(doc).paragraph(para).paragraphProperties.tabs;
    expect(tabs).toEqual([
      { type: "left", position: 1440 },
      { type: "right", position: 9360 }
    ]);
  });
});
