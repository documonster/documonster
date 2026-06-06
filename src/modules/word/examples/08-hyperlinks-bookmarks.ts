/**
 * Word Example 08 — Hyperlinks & Bookmarks
 *
 * Covers:
 *   - External hyperlink (URL) — direct w:hyperlink element
 *   - External hyperlink — HYPERLINK field form
 *   - Email mailto: link
 *   - Internal anchor (#bookmark) hyperlink
 *   - Bookmark range (start + end) wrapping a paragraph
 *   - Cross-reference: REF (text), PAGEREF (page), with hyperlink option
 *   - Tooltip / target frame / "history" attribute
 *   - Inline images as the visible link content (image link)
 *   - Edge cases: empty bookmark name (skipped), hyperlink with no rId/url
 *     (rendered as plain text), nested hyperlink inside a list item.
 *
 * Output: tmp/word-examples/08-hyperlinks-bookmarks.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  hyperlink,
  hyperlinkField,
  pageBreak,
  bookmarkStart,
  bookmarkEnd,
  refField,
  pageRefField,
  toBuffer
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word — Hyperlinks & Bookmarks", 1);

// ---------------------------------------------------------------------------
// 1. External link via w:hyperlink element (the most common form)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. External link", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Visit "),
    hyperlink("the project site", {
      url: "https://example.com",
      tooltip: "Open example.com in your browser"
    }),
    text(" for details.")
  ])
);

// External link via the HYPERLINK field form (alternative representation)
Document.addParagraphElement(
  doc,
  paragraph([
    text("(field form) "),
    hyperlinkField("https://example.com/docs", {
      displayText: "documentation",
      newWindow: true,
      tooltip: "Open in a new window"
    })
  ])
);

// Email link
Document.addParagraphElement(
  doc,
  paragraph([
    text("Contact: "),
    hyperlink("support@example.com", { url: "mailto:support@example.com" })
  ])
);

// "history" attribute → visited-link styling. With history:true the builder
// references the FollowedHyperlink character style (and emits the standard
// visited purple as direct formatting), so the link renders purple, in
// contrast to the blue unvisited links above.
Document.addParagraphElement(
  doc,
  paragraph([
    text("This link is styled as visited (purple): "),
    hyperlink("visited", { url: "https://example.com/visited", history: true })
  ])
);

// ---------------------------------------------------------------------------
// 2. Internal bookmark + anchor hyperlink + cross-reference
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Internal navigation", 2);

// Allocate stable bookmark IDs from the document handle so they don't clash
// with auto-assigned IDs the writer might emit elsewhere.
const introId = Document.nextBookmarkId(doc);
const detailsId = Document.nextBookmarkId(doc);

// Cross-references via REF / PAGEREF fields. ECMA-376 §17.16.5.57 says
// the \h flag tells Word to render the result as a hyperlink — but the
// cached value Word ships in the file is *not* automatically painted blue
// on first open; that styling is applied during the F9 update pass. To
// match what Word writes after F9 (and thus give a correctly-styled link
// even before update) we wrap the field in an explicit <w:hyperlink>
// element, which is the runtime form Word produces internally for \h.
const wrapAsLink = (anchor: string, run: ReturnType<typeof refField>) => ({
  type: "hyperlink" as const,
  anchor,
  children: [run]
});

Document.addParagraphElement(
  doc,
  paragraph([
    text("Jump to: "),
    hyperlink("Introduction", { anchor: "intro" }),
    text(" · "),
    hyperlink("Details", { anchor: "details" }),
    text(" · "),
    text("see also: "),
    wrapAsLink("intro", refField("intro", { hyperlink: true, cachedValue: "Introduction" })),
    text(" (page "),
    wrapAsLink("intro", pageRefField("intro", { hyperlink: true, cachedValue: "1" })),
    text(").")
  ])
);

// Force a page break so PAGEREF resolves to something different at runtime
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// "Introduction" — wrapped in a bookmark range
Document.addParagraphElement(
  doc,
  paragraph([bookmarkStart(introId, "intro"), text("Introduction"), bookmarkEnd(introId)], {
    style: "Heading2"
  })
);
Document.addParagraph(doc, "Lorem ipsum… ".repeat(40));

// Force another page break
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// "Details" — bookmark wrapping a heading
Document.addParagraphElement(
  doc,
  paragraph([bookmarkStart(detailsId, "details"), text("Details"), bookmarkEnd(detailsId)], {
    style: "Heading2"
  })
);
Document.addParagraph(doc, "Sed ut perspiciatis… ".repeat(40));

// ---------------------------------------------------------------------------
// 3. Hyperlink with rich formatting (bold, italic, custom color)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Rich-formatted link content", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("A "),
    hyperlink("bold red link", {
      url: "https://example.com",
      properties: { bold: true, color: "C00000", underline: "single" }
    }),
    text(" and an "),
    hyperlink("italic underlined link", {
      url: "https://example.com",
      properties: { italic: true, underline: "double" }
    }),
    text(".")
  ])
);

// ---------------------------------------------------------------------------
// 4. Tooltip / target frame
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("With target frame _blank: "),
    hyperlink("open in new tab", {
      url: "https://example.com",
      tgtFrame: "_blank",
      tooltip: "_blank target"
    })
  ])
);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Hyperlink with neither url nor anchor — degraded to plain text by writers
Document.addParagraphElement(
  doc,
  paragraph([text("Broken link (no url/anchor): "), hyperlink("plain text", {})])
);

// Empty link text
Document.addParagraphElement(doc, paragraph([hyperlink("", { url: "https://example.com" })]));

// Bookmark with empty name — written but flagged at validation time
const emptyBookmarkId = Document.nextBookmarkId(doc);
Document.addParagraphElement(
  doc,
  paragraph([
    bookmarkStart(emptyBookmarkId, ""),
    text("(bookmark with empty name)"),
    bookmarkEnd(emptyBookmarkId)
  ])
);

// Bookmark wrapping multiple paragraphs (range straddles paragraph boundary)
const rangeId = Document.nextBookmarkId(doc);
Document.addParagraphElement(
  doc,
  paragraph([bookmarkStart(rangeId, "multi"), text("First paragraph in range.")])
);
Document.addParagraph(doc, "Second paragraph in range.");
Document.addParagraphElement(
  doc,
  paragraph([text("Last paragraph; bookmark ends here. "), bookmarkEnd(rangeId)])
);
// REF \h fields are how Word usually represents a clickable cross-reference
// to a bookmark, but Word does not paint the cached value with hyperlink
// styling on first open — only after F9 / Update Fields. For a demo that
// always renders correctly we use the equivalent <w:hyperlink w:anchor>
// element instead. (The REF/PAGEREF field examples in section 2 above
// already cover the field-level form.)
Document.addParagraphElement(
  doc,
  paragraph([
    text("Reference into the multi-paragraph bookmark: "),
    hyperlink("[multi paragraph block]", { anchor: "multi" })
  ])
);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "08-hyperlinks-bookmarks.docx"), buf);
console.log(`  → 08-hyperlinks-bookmarks.docx (${buf.length} bytes)`);
