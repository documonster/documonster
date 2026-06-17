/**
 * Word Example 25 — HTML ↔ DOCX conversion
 *
 * Covers:
 *   - htmlToDocxBody — parses an HTML string into BodyContent[]
 *     · headings, paragraphs, divs
 *     · ordered & unordered lists (incl. task lists)
 *     · tables with colspan/rowspan, borders, alignment, header rows
 *     · inline: strong/em/u/s/sub/sup, code, span (with style), a, br
 *     · block-level CSS: text-align, font-family, font-size, color
 *     · class-based styling via classStyles
 *     · embedded <style> rules
 *     · pre/code blocks
 *     · base64 inline images (data: URL)
 *     · page-break div
 *   - renderToHtml — DOCX → HTML, with inline styles, images as data URLs,
 *     comments rendered, footnotes, custom style map
 *   - Edge cases: malformed tags, mixed-script content, very long tables.
 *
 * Output:
 *   - 25-html-imported.docx
 *   - 25-html-imported.html (round-trip output)
 *   - 25-html-roundtrip.html (DOCX → HTML again)
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { renderToHtml, htmlToDocx } from "../html";
import { Document, Io } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/25-html"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. HTML → DOCX
// ---------------------------------------------------------------------------
const inputHtml = /* html */ `
<!doctype html>
<html>
  <head>
    <style>
      .highlight { background-color: #FFFFCC; font-weight: bold; }
      .quote     { font-style: italic; color: #555555; }
    </style>
  </head>
  <body>
    <h1>HTML import demo</h1>
    <h2 style="color: #1F4E79">A subsection</h2>
    <p>A paragraph with <strong>bold</strong>, <em>italic</em>, <u>underline</u>,
       <s>strike</s>, <code>code</code>, x<sup>2</sup>, H<sub>2</sub>O, and
       <a href="https://example.com">a link</a>.</p>
    <p>CSS-styled span: <span class="highlight">highlighted bold text</span>.</p>
    <blockquote class="quote">"The single biggest problem in communication
      is the illusion that it has taken place." — G.B. Shaw</blockquote>

    <h2>Lists</h2>
    <ul>
      <li>First bullet</li>
      <li>Second bullet
        <ul>
          <li>Nested A</li>
          <li>Nested B</li>
        </ul>
      </li>
      <li>Third bullet</li>
    </ul>
    <ol>
      <li>Step one</li>
      <li>Step two</li>
      <li>Step three</li>
    </ol>

    <h2>Tables</h2>
    <table border="1" style="border-collapse: collapse">
      <thead><tr><th>Country</th><th>Capital</th><th>Population</th></tr></thead>
      <tbody>
        <tr><td>Australia</td><td>Canberra</td><td>26M</td></tr>
        <tr><td>Canada</td><td>Ottawa</td><td>40M</td></tr>
        <tr><td colspan="2" style="text-align: right">Total</td><td>66M</td></tr>
      </tbody>
    </table>

    <h2>Images</h2>
    <p>Inline base64 image:
      <img alt="checkerboard"
           src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAwElEQVR4nO3QMRGAQBTE0HOCD0TgB8e4ABG/eHNDiq2TzXrO853suO7RNH9pAc0vgBbQ/AJoAc0vgBbQ/AJoAc0vgBbQ/AJoAc0vgBbQ/LX7gSm/AFpA8wugBTS/AFpA8wugBTS/AFpA8wugBTS/AFpA8wuw+4EpvwBaQPMLoAU0vwBaQPMLoAU0vwBaQPMLoAU0vwBaQPMLsPuBKb8AWkDzC6AFNL8AWkDzC6AFNL8AWkDzC6AFNL8AWkDzfx/gAwMMAf/OkDDpAAAAAElFTkSuQmCC" /></p>

    <h2>Code block</h2>
    <pre>function hello() {
  return "world";
}</pre>

    <h2>Edge cases</h2>
    <p>Special chars: 1 &lt; 2 &amp;&amp; 3 &gt; 2; "smart" 'quotes';
       NBSP&#160;here; CJK 你好世界; emoji 🎉; Arabic مرحبا.</p>
    <p>Unclosed-tag-tolerated: <strong>bold then... nothing else.</p>
    <div style="page-break-before: always"></div>
    <p>This paragraph follows a hard page break.</p>
  </body>
</html>`;

const { body, images: htmlImages } = htmlToDocx(inputHtml, {
  defaultFont: "Calibri",
  defaultFontSize: 22,
  classStyles: {
    quote: "border-left: 3px solid #1F4E79; padding-left: 12px;"
  }
});
console.log(`  htmlToDocx → ${body.length} body content blocks, ${htmlImages.length} images`);

// Wrap into a full document. htmlToDocxBody emits paragraphs that reference
// numId=1 (bullets) and numId=2 (ordered), but it does NOT define those
// abstractNumberings — the caller owns numbering. Inject the matching
// definitions so the bullets/numbers render.
const doc = Document.create();
Document.useDefaultStyles(doc);
for (const item of body) {
  Document.addContent(doc, item);
}
const built = Document.build(doc);
const finalDoc = {
  ...built,
  images: [...(built.images ?? []), ...htmlImages],
  abstractNumberings: [
    ...(built.abstractNumberings ?? []),
    {
      abstractNumId: 1,
      multiLevelType: "hybridMultilevel" as const,
      levels: [
        {
          level: 0,
          start: 1,
          format: "bullet" as const,
          text: "\u2022",
          justification: "left" as const,
          paragraphProperties: { indent: { left: 720, hanging: 360 } },
          runProperties: { font: { ascii: "Symbol", hAnsi: "Symbol" } }
        },
        {
          level: 1,
          start: 1,
          format: "bullet" as const,
          text: "\u25E6",
          justification: "left" as const,
          paragraphProperties: { indent: { left: 1440, hanging: 360 } },
          runProperties: { font: { ascii: "Courier New", hAnsi: "Courier New" } }
        }
      ]
    },
    {
      abstractNumId: 2,
      multiLevelType: "hybridMultilevel" as const,
      levels: [
        {
          level: 0,
          start: 1,
          format: "decimal" as const,
          text: "%1.",
          justification: "left" as const,
          paragraphProperties: { indent: { left: 720, hanging: 360 } }
        }
      ]
    }
  ],
  numberingInstances: [
    ...(built.numberingInstances ?? []),
    { numId: 1, abstractNumId: 1 },
    { numId: 2, abstractNumId: 2 }
  ]
};
const docxBytes = await Io.toBuffer(finalDoc);
fs.writeFileSync(path.join(outDir, "25-html-imported.docx"), docxBytes);
console.log(`  → 25-html-imported.docx (${docxBytes.length} bytes)`);

// ---------------------------------------------------------------------------
// 2. DOCX → HTML (round-trip)
// ---------------------------------------------------------------------------
const reread = await Io.read(docxBytes);
const htmlOut = renderToHtml(reread, {
  fullDocument: true,
  includeStyles: true,
  imageMode: "dataUrl",
  includeNotes: true,
  classPrefix: "doc-",
  styleMap: {
    Heading1: "h1.section-title",
    Heading2: "h2.subsection-title",
    MyQuote: "blockquote.fancy"
  }
});
console.log(
  `  renderToHtml → ${htmlOut.html.length} chars, ${htmlOut.images.size} images, ${htmlOut.warnings.length} warnings`
);
fs.writeFileSync(path.join(outDir, "25-html-roundtrip.html"), htmlOut.html);
console.log("  → 25-html-roundtrip.html");

// ---------------------------------------------------------------------------
// 3. Fragment mode — produce only the body fragment (useful for emails)
// ---------------------------------------------------------------------------
const frag = renderToHtml(reread, {
  fullDocument: false,
  includeStyles: false,
  imageMode: "filename"
});
fs.writeFileSync(path.join(outDir, "25-html-fragment.html"), frag.html);
console.log("  → 25-html-fragment.html");

// In "filename" image mode the <img> tags reference external files by name,
// so the HTML is only usable if those files sit next to it. Write each image
// the fragment references out to disk using the raw bytes from the re-read
// document model.
const fileNames = new Set(frag.images.keys());
for (const img of reread.images ?? []) {
  if (fileNames.has(img.fileName)) {
    fs.writeFileSync(path.join(outDir, img.fileName), img.data);
    console.log(`  → ${img.fileName} (${img.data.length} bytes)`);
  }
}
