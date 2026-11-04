/**
 * An internal hyperlink names a bookmark, so the HTML anchor it points at is
 * keyed by that name. It was keyed by the numeric id, so `#bookmark-Intro`
 * pointed at nothing — and two bookmarks sharing an id produced duplicate
 * element ids.
 */
import { describe, it, expect } from "vitest";

import { WordHtml } from "../html";
import { Build, Document } from "../index";

describe("HTML bookmark anchors", () => {
  it("are keyed by bookmark name, which internal links target", () => {
    const doc = Document.create();
    Document.addParagraph(
      doc,
      Build.paragraph([Build.bookmarkStart(0, "Intro"), Build.text("Intro"), Build.bookmarkEnd(0)])
    );
    Document.addParagraph(
      doc,
      Build.paragraph([Build.bookmarkStart(0, "Other"), Build.text("Other"), Build.bookmarkEnd(0)])
    );
    Document.addParagraph(doc, Build.paragraph([Build.hyperlink("go", { anchor: "Intro" })]));
    const { html } = WordHtml.fromDocx(doc);
    expect(html).toContain('id="bookmark-Intro"');
    expect(html).toContain('id="bookmark-Other"');
    expect(html).toContain('href="#bookmark-Intro"');
  });
});
