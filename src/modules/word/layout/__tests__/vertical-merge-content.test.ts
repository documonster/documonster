import { Build, Document } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { PageContent } from "@word/layout/layout-model";
import { describe, expect, it } from "vitest";

/**
 * Word draws nothing a vertical merge's continuing cell holds (pages2 R4j,
 * R7i); its bookmarks still sit where the merged cell's text begins.
 */
describe("vertical merge continuing cell", () => {
  it("draws none of its text but keeps its bookmark", () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    const id = Document.allocateBookmarkId(doc);
    Document.addTable(
      doc,
      Build.table(
        [
          Build.row([
            Build.cell([Build.paragraph([Build.text("shown")])], { verticalMerge: "restart" }),
            Build.cell([Build.paragraph([Build.text("a")])])
          ]),
          Build.row([
            Build.cell(
              [
                Build.paragraph([
                  Build.bookmarkStart(id, "HIDDEN"),
                  Build.bookmarkEnd(id),
                  Build.text("hidden")
                ])
              ],
              { verticalMerge: "continue" }
            ),
            Build.cell([Build.paragraph([Build.text("b")])])
          ])
        ],
        { borders: Build.gridBorders() },
        [3000, 3000]
      )
    );
    const texts: string[] = [];
    const marks: string[] = [];
    const visit = (blocks: readonly PageContent[]): void => {
      for (const b of blocks) {
        if (b.type === "paragraph") {
          for (const line of b.lines) {
            for (const run of line.runs) {
              if (run.type !== "image") {
                texts.push(run.text);
              }
            }
          }
          marks.push(...(b.lineBookmarks ?? []).flat());
        } else if (b.type === "table") {
          for (const c of b.cells) {
            visit(c.content as readonly PageContent[]);
          }
        }
      }
    };
    visit(layoutDocumentFull(doc).pages[0]!.content as readonly PageContent[]);
    expect(texts).toContain("shown");
    expect(texts.join("")).not.toContain("hidden");
    expect(marks).toContain("HIDDEN");
  });
});
