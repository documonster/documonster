/**
 * Word Example 37 — SVG rendering & layout
 *
 * Covers:
 *   - Layout.getPageMap — analyse the doc into pages and remember which body
 *     element falls on which page; build a bookmark → page lookup.
 *   - Layout.compute — produce a full LayoutDocument tree with concrete
 *     line boxes, positioned runs and floats (used internally by SVG/PDF).
 *   - Layout.pageToSvg — render a single page as standalone SVG markup.
 *   - Layout.toSvg — render every page (returns one SVG per page).
 *   - Layout.computedPageToSvg — when the caller already has the LayoutDocument
 *     and wants to render only one page (faster than re-laying-out for SVG).
 *
 * Output: tmp/word-examples/37-svg/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Layout } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/37-svg"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Build a multi-page document with a bookmark and explicit page break
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addHeading(d, "Layout & SVG render demo", 1);
Document.addParagraph(d, "Page 1 body — short.");

// Page break, then a bookmark on page 2
Document.addParagraph(d, Build.paragraph([Build.pageBreak()]));
Document.addParagraph(
  d,
  Build.paragraph([
    Build.bookmarkStart(0, "section-two"),
    Build.bold("Section two"),
    Build.bookmarkEnd(0)
  ])
);
for (let i = 0; i < 25; i++) {
  Document.addParagraph(d, `Page 2 paragraph ${i + 1} — Lorem ipsum dolor sit amet.`);
}

// Page break + bookmark on page 3
Document.addParagraph(d, Build.paragraph([Build.pageBreak()]));
Document.addParagraph(
  d,
  Build.paragraph([
    Build.bookmarkStart(1, "appendix"),
    Build.bold("Appendix"),
    Build.bookmarkEnd(1)
  ])
);
Document.addParagraph(d, "End of document.");

// ---------------------------------------------------------------------------
// 1. Layout.getPageMap — high-level page mapping
// ---------------------------------------------------------------------------
const layout = Layout.getPageMap(d);
console.log(
  `  layoutDocument: ${layout.pageCount} pages, ${layout.sectionPageCounts.length} section(s)`
);
console.log(`  bookmark → page mapping:`);
for (const [name, pageNum] of layout.bookmarkPages) {
  console.log(`    ${name} → page ${pageNum}`);
}
console.log(
  `  contentPages (first 8): ${[...layout.contentPages.slice(0, 8)].join(", ")} (length=${layout.contentPages.length})`
);

// Save the source .docx so we can compare with the SVG visually
const docxBuf = await Io.toBuffer(d);
fs.writeFileSync(path.join(outDir, "00-source.docx"), docxBuf);

// ---------------------------------------------------------------------------
// 2. Layout.pageToSvg — one page at a time
// ---------------------------------------------------------------------------
for (let p = 1; p <= layout.pageCount; p++) {
  const svg = Layout.pageToSvg(d, p);
  fs.writeFileSync(path.join(outDir, `page-${p}.svg`), svg);
  console.log(`  → page-${p}.svg (${svg.length} chars)`);
}

// ---------------------------------------------------------------------------
// 3. Layout.toSvg — every page in one call
// ---------------------------------------------------------------------------
const allPages = Layout.toSvg(d);
console.log(`  Layout.toSvg → ${allPages.length} page(s)`);

// ---------------------------------------------------------------------------
// 4. Layout.compute → Layout.computedPageToSvg (re-use the layout)
// ---------------------------------------------------------------------------
const fullLayout = Layout.compute(d);
console.log(
  `  layoutDocumentFull: ${fullLayout.pages.length} pages; first page geometry = ${JSON.stringify(fullLayout.pages[0]?.geometry)}, content blocks = ${fullLayout.pages[0]?.content.length ?? 0}`
);
const svgPage1 = Layout.computedPageToSvg(fullLayout, 1);
fs.writeFileSync(path.join(outDir, "page-1-from-layout.svg"), svgPage1);
console.log(`  → page-1-from-layout.svg (${svgPage1.length} chars)`);

// ---------------------------------------------------------------------------
// 5. Edge: Layout.pageToSvg out-of-range page should throw
// ---------------------------------------------------------------------------
try {
  Layout.pageToSvg(d, 999);
  console.log("  ERROR: expected RangeError");
} catch (err) {
  console.log(`  out-of-range → ${(err as Error).constructor.name}: ${(err as Error).message}`);
}
