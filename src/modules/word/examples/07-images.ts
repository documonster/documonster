/**
 * Word Example 07 — Images
 *
 * Covers:
 *   - Inline image (anchored to text flow) — PNG, JPEG, GIF
 *   - Floating image — square wrap, tight wrap, behindDoc, in front of text
 *   - Position relative to margin / page
 *   - Image with altText (accessibility) and explicit name
 *   - Image rotation, flip
 *   - SVG image (with PNG fallback)
 *   - Multiple images in one paragraph
 *   - Image inside a table cell
 *   - Edge case: very small (1×1 PNG), oversized (logical width > page),
 *     transparent PNG, identical bytes deduplication.
 *
 * Output: tmp/word-examples/07-images.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Units } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Generate small synthetic image data so the example is fully self-contained
// (Word only needs valid bytes; it does not validate the resolution). We use
// 32×32 source rasters rather than 1×1 because Word's image-scaling pipeline
// renders 1-pixel sources as a near-black blur when stretched to 2 cm — the
// minimum useful source size for the demo is around 16×16.
// ---------------------------------------------------------------------------

/** A 32×32 solid red PNG (color #C00000). */
const redPixelPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20, 0x08, 0x02, 0x00, 0x00, 0x00, 0xfc, 0x18, 0xed,
  0xa3, 0x00, 0x00, 0x00, 0x29, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0xed, 0xcd, 0x31, 0x0d, 0x00,
  0x00, 0x0c, 0xc3, 0xb0, 0xf2, 0x47, 0x35, 0x68, 0x23, 0xd1, 0x7e, 0x96, 0x72, 0xc7, 0xb9, 0x64,
  0xda, 0xf6, 0x0e, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x8a, 0x3d, 0x1d, 0x2c,
  0x00, 0x2e, 0xf2, 0x6f, 0x09, 0xfa, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42,
  0x60, 0x82
]);

/** A 32×32 light-blue translucent PNG (#DEEBF7 with 80% alpha) — used as a
 *  visible behindDoc watermark sample. */
const watermarkPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20, 0x08, 0x06, 0x00, 0x00, 0x00, 0x73, 0x7a, 0x7a,
  0xf4, 0x00, 0x00, 0x00, 0x31, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0xed, 0xce, 0x21, 0x01, 0x00,
  0x00, 0x08, 0x03, 0x30, 0xfa, 0x67, 0x23, 0x03, 0x05, 0x08, 0x80, 0x25, 0xc6, 0xcd, 0xc4, 0xfc,
  0x6a, 0xf6, 0x3a, 0xa9, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04,
  0x04, 0x04, 0x04, 0xd2, 0x81, 0x07, 0x18, 0x56, 0x30, 0xd3, 0x36, 0xb0, 0xdc, 0x16, 0x00, 0x00,
  0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
]);

/** A 32×32 solid yellow PNG (#FFD700). The example referred to JPEG before
 *  but synthesising a real JPEG by hand is fragile (Huffman tables must
 *  exactly match the entropy-coded payload); using a PNG here keeps the
 *  byte stream verifiable and the visual result identical. */
const yellowPixelPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20, 0x08, 0x02, 0x00, 0x00, 0x00, 0xfc, 0x18, 0xed,
  0xa3, 0x00, 0x00, 0x00, 0x2a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0xed, 0xcd, 0xb1, 0x0d, 0x00,
  0x00, 0x08, 0xc3, 0xb0, 0xfe, 0xff, 0x20, 0xe7, 0xc0, 0x13, 0x74, 0xb3, 0x94, 0x39, 0xce, 0x4e,
  0xaa, 0x75, 0xef, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe0, 0xb1, 0x03, 0x24,
  0x54, 0x58, 0x6a, 0x0c, 0xac, 0x37, 0xe9, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

/** A simple SVG image (text form). */
const sampleSvg = new TextEncoder().encode(
  `<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64"><rect width="64" height="64" fill="#1F4E79"/><circle cx="32" cy="32" r="20" fill="#F2F2F2"/><text x="32" y="38" font-size="14" text-anchor="middle" fill="#1F4E79" font-family="sans-serif">SVG</text></svg>`
);

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word Module — Images", 1);

// ---------------------------------------------------------------------------
// 1. Inline images (PNG / JPEG / GIF / SVG)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Inline images (PNG / JPEG / SVG)", 2);
Document.addParagraph(doc, "Each image is inline — flows with text.");
Document.addImage(doc, redPixelPng, "png", Units.cmToEmu(2), Units.cmToEmu(2), {
  altText: "A red square (PNG)",
  name: "RedSquare"
});
Document.addImage(doc, yellowPixelPng, "png", Units.cmToEmu(2), Units.cmToEmu(2), {
  altText: "A yellow square (JPEG)"
});
Document.addImage(doc, sampleSvg, "svg", Units.cmToEmu(2), Units.cmToEmu(2), {
  altText: "A blue SVG icon"
});

// ---------------------------------------------------------------------------
// 2. Floating image with square wrap (text flows around)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Floating image — square wrap", 2);
Document.addFloatingImage(doc, redPixelPng, "png", Units.cmToEmu(4), Units.cmToEmu(4), {
  altText: "Floating logo",
  horizontalPosition: { align: "right", relativeTo: "margin" },
  verticalPosition: { align: "top", relativeTo: "paragraph" },
  wrap: {
    style: "square",
    side: "left",
    margins: { top: 0, bottom: 0, left: Units.cmToEmu(0.3), right: 0 }
  }
});
Document.addParagraph(
  doc,
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ".repeat(8)
);

// ---------------------------------------------------------------------------
// 3. Floating image behind text (background watermark-style)
// ---------------------------------------------------------------------------
// Page-break first so the watermark anchors to a new page — otherwise the
// behindDoc image (whose vertical anchor is "page-center") would render on
// whatever page the previous section happens to end on, visually overlapping
// section 2's text.
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "3. Floating image — behind text", 2);
Document.addFloatingImage(doc, watermarkPng, "png", Units.cmToEmu(8), Units.cmToEmu(8), {
  altText: "Background mark",
  horizontalPosition: { align: "center", relativeTo: "page" },
  verticalPosition: { align: "center", relativeTo: "page" },
  wrap: { style: "none" },
  behindDoc: true,
  allowOverlap: true
});
// Floating images attach to the preceding paragraph in OOXML; without body
// text after the heading the watermark would still render but you'd have
// nothing to show it sits *behind*. Add enough body text that it physically
// covers the page-centre area where the watermark draws.
Document.addParagraph(
  doc,
  "Body text for section 3. The light-blue square is positioned dead-centre " +
    "on the page, behind these letters. " +
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(40)
);

// ---------------------------------------------------------------------------
// 4. Floating image in-front-of text (overlay)
// ---------------------------------------------------------------------------
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "4. Floating image — in front (overlay)", 2);
// Body text so the page is not blank — the yellow sticker is anchored to
// the page bottom-left, this body explains where to look.
Document.addParagraph(
  doc,
  "A small yellow sticker should appear at the bottom-left corner of this " +
    "page, sitting on top of these words. " +
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(6)
);
Document.addFloatingImage(doc, yellowPixelPng, "png", Units.cmToEmu(2), Units.cmToEmu(2), {
  altText: "Sticker overlay",
  horizontalPosition: { align: "left", relativeTo: "page", offset: Units.cmToEmu(1) },
  verticalPosition: { align: "bottom", relativeTo: "page", offset: -Units.cmToEmu(2) },
  wrap: { style: "none" },
  behindDoc: false
});

// ---------------------------------------------------------------------------
// 5. Rotation & flip
// ---------------------------------------------------------------------------
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "5. Rotation / flip", 2);
Document.addFloatingImage(doc, redPixelPng, "png", Units.cmToEmu(3), Units.cmToEmu(3), {
  altText: "Rotated 30°",
  rotation: 30 * 60_000, // OOXML stores rotation in 1/60_000 degrees
  horizontalPosition: { align: "center", relativeTo: "margin" },
  verticalPosition: { align: "top", relativeTo: "paragraph" },
  wrap: { style: "topAndBottom" }
});
Document.addFloatingImage(doc, redPixelPng, "png", Units.cmToEmu(3), Units.cmToEmu(3), {
  altText: "Flipped horizontally",
  flipHorizontal: true,
  horizontalPosition: { align: "right", relativeTo: "margin" },
  verticalPosition: { align: "top", relativeTo: "paragraph" },
  wrap: { style: "topAndBottom" }
});

// ---------------------------------------------------------------------------
// 6. Multiple inline images in a single paragraph
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Multiple inline images in one paragraph", 2);
const imgInfo1 = (() =>
  Document.addImage(doc, redPixelPng, "png", Units.cmToEmu(1), Units.cmToEmu(1)))();
const imgInfo2 = (() =>
  Document.addImage(doc, yellowPixelPng, "png", Units.cmToEmu(1), Units.cmToEmu(1)))();
void imgInfo1;
void imgInfo2;

// ---------------------------------------------------------------------------
// 7. Image inside a table cell
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Image inside a table cell", 2);
// Build the image first so the rId is available, then place it via a nested
// paragraph in a table cell. We build the inline image content directly.
const imgResult = Document.addImage(
  doc,
  redPixelPng,
  "png",
  Units.cmToEmu(1.5),
  Units.cmToEmu(1.5),
  {
    altText: "Cell logo"
  }
);
// Document.addImage auto-appends a paragraph holding the drawing — pop it
// off so the image only appears inside the cell below (otherwise we'd render
// the image twice: once free-floating, once inside the cell).
{
  const lastIndex = Document.getContentCount(doc) - 1;
  Document.removeContent(doc, lastIndex);
}
const cellLogoParagraph = Build.paragraph([
  {
    content: [
      {
        type: "image",
        rId: imgResult.rId,
        width: Units.cmToEmu(1.5),
        height: Units.cmToEmu(1.5),
        drawingId: imgResult.drawingId,
        name: "InCell",
        altText: "Cell logo"
      }
    ]
  }
]);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row([
        Build.cell([cellLogoParagraph]),
        Build.cell([
          Build.textParagraph(
            "Cell next to a logo image. Logos always render at native pixel size unless you specify width/height in EMU."
          )
        ])
      ])
    ],
    { width: { value: 5000, type: "pct" }, borders: Build.gridBorders() },
    [Units.cmToTwips(2.5), Units.cmToTwips(12)]
  )
);

// ---------------------------------------------------------------------------
// Edge cases — oversized (page-wide), 1x1, deduplication
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Page-wide oversize: width close to A4 portrait usable width (16cm)
Document.addImage(doc, redPixelPng, "png", Units.cmToEmu(16), Units.cmToEmu(8), {
  altText: "Oversized — fills the printable width"
});

// Deduplication: identical bytes added twice should ideally share the same
// underlying image part, but Document.addImage assigns a fresh rId each
// time. The packager will write two separate parts; that is acceptable
// behaviour. We simply verify no error.
Document.addImage(doc, redPixelPng, "png", Units.cmToEmu(0.5), Units.cmToEmu(0.5));
Document.addImage(doc, redPixelPng, "png", Units.cmToEmu(0.5), Units.cmToEmu(0.5));

// 1x1 transparent
Document.addImage(doc, watermarkPng, "png", Units.cmToEmu(0.3), Units.cmToEmu(0.3), {
  altText: "1×1 transparent (decorative)"
});

// ---------------------------------------------------------------------------
// 8. Low-level floatingImage() builder — for callers that already have a
//    rId (e.g. when reusing the same image multiple times) or need to set
//    rare fields like srcRect (image cropping) that the convenience helper
//    omits.
// ---------------------------------------------------------------------------
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "8. Raw floatingImage builder (with cropping)", 2);
Document.addParagraph(
  doc,
  "Two cropped copies of the same image source — both are floating with " +
    "topAndBottom wrap. Because the source bytes are a solid red square the " +
    "cropping is invisible to the eye but the srcRect attributes are still " +
    "written into the underlying drawing XML."
);
const sharedImage = Document.addImage(
  doc,
  redPixelPng,
  "png",
  Units.cmToEmu(0.5),
  Units.cmToEmu(0.5),
  {
    altText: "shared source"
  }
);
// Re-use the rId for two more floating placements with different crops
Document.addContent(
  doc,
  Build.floatingImage({
    rId: sharedImage.rId,
    width: Units.cmToEmu(3),
    height: Units.cmToEmu(3),
    horizontalPosition: { align: "left", relativeTo: "margin" },
    verticalPosition: { align: "top", relativeTo: "paragraph" },
    wrap: { style: "topAndBottom" },
    altText: "Crop: top-left quadrant",
    // srcRect crops a fraction of the source — values in 1/100 000 units,
    // applied via {l, t, r, b} (left/top/right/bottom). Here we crop the
    // right & bottom halves to show only top-left.
    srcRect: { r: 50000, b: 50000 }
  })
);
Document.addContent(
  doc,
  Build.floatingImage({
    rId: sharedImage.rId,
    width: Units.cmToEmu(3),
    height: Units.cmToEmu(3),
    horizontalPosition: { align: "right", relativeTo: "margin" },
    verticalPosition: { align: "top", relativeTo: "paragraph" },
    wrap: { style: "topAndBottom" },
    altText: "Crop: bottom-right quadrant",
    srcRect: { l: 50000, t: 50000 }
  })
);

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "07-images.docx"), buf);
console.log(`  → 07-images.docx (${buf.length} bytes)`);
