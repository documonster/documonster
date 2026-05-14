/**
 * Word Example 24 — Watermark, page background, theme
 *
 * Covers:
 *   - Text watermark (semi-transparent, rotated)
 *   - Image watermark (washout)
 *   - Page background colour
 *   - Custom theme (colour scheme + font scheme)
 *   - Edge case: watermark + background simultaneously
 *
 * Output: tmp/word-examples/24-watermark-bg/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, paragraph, text, ptToHalfPoint, cmToTwips, toBuffer } from "../index";
import type { DocumentTheme, Watermark } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/24-watermark-bg"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Text watermark
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Text watermark", 1);
  Document.addParagraph(d, "A diagonal CONFIDENTIAL stamp will appear behind the body text.");
  Document.addParagraph(d, "Lorem ipsum dolor sit amet… ".repeat(40));

  const wm: Watermark = {
    type: "text",
    text: "CONFIDENTIAL",
    color: "C0C0C0",
    fontSize: ptToHalfPoint(72),
    rotation: -45,
    semiTransparent: true,
    font: "Arial"
  };
  Document.setWatermark(d, wm);
  fs.writeFileSync(path.join(outDir, "01-text-watermark.docx"), await toBuffer(Document.build(d)));
  console.log("  → 01-text-watermark.docx");
}

// ---------------------------------------------------------------------------
// 2. Image watermark — needs an image to be already added so we can grab
//    its rId. Document.addImage adds it inline; for a watermark we want the
//    image present in the package without rendering it inline. The
//    watermark image rId points to a header-anchored image that the
//    packager auto-builds when we set type:"image" on the watermark.
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Image watermark", 1);
  Document.addParagraph(d, "Logo printed faintly behind every page.");
  Document.addParagraph(d, "Filler text… ".repeat(60));

  // Embed a 1×1 PNG as an image so the packager has bytes for the rId.
  const tinyPng = Uint8Array.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82
  ]);
  // Store directly in the model so the watermark can reference it.
  const handle = d as unknown as {
    _state: {
      images: Array<{ data: Uint8Array; mediaType: string; fileName: string; rId: string }>;
    };
  };
  // We use the public addImage path which assigns a stable rId then strip
  // the inserted body paragraph: only the image part survives in the package.
  const imgInfo = Document.addImage(d, tinyPng, "png", cmToTwips(2), cmToTwips(2), {
    altText: "Watermark logo",
    name: "WatermarkSource"
  });
  // Pop the auto-inserted body paragraph (it will be replaced by the watermark)
  Document.removeContent(d, Document.getContentCount(d) - 1);
  void handle;

  const wm: Watermark = { type: "image", rId: imgInfo.rId, scale: 100, washout: true };
  Document.setWatermark(d, wm);
  fs.writeFileSync(path.join(outDir, "02-image-watermark.docx"), await toBuffer(Document.build(d)));
  console.log("  → 02-image-watermark.docx");
}

// ---------------------------------------------------------------------------
// 3. Page background colour
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Page background", 1);
  Document.addParagraph(d, "The page is pale blue.");
  Document.setBackground(d, { color: "DEEBF7" });
  fs.writeFileSync(path.join(outDir, "03-bg-color.docx"), await toBuffer(Document.build(d)));
  console.log("  → 03-bg-color.docx");
}

// ---------------------------------------------------------------------------
// 4. Custom theme (referenced by run color "themeColor: accent1")
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Custom theme", 1);
  Document.addParagraphElement(
    d,
    paragraph([
      text("Accent1: ", { color: { val: "auto", themeColor: "accent1" } }),
      text("Accent2 ", { color: { val: "auto", themeColor: "accent2" } }),
      text("Accent3 ", { color: { val: "auto", themeColor: "accent3" } })
    ])
  );

  const theme: DocumentTheme = {
    name: "Sunset",
    colorScheme: {
      name: "Sunset",
      colors: {
        dk1: "1F1F1F",
        lt1: "FFFFFF",
        dk2: "303030",
        lt2: "F2F2F2",
        accent1: "C0392B",
        accent2: "E67E22",
        accent3: "F1C40F",
        accent4: "27AE60",
        accent5: "2980B9",
        accent6: "8E44AD",
        hlink: "0563C1",
        folHlink: "954F72"
      }
    },
    fontScheme: {
      name: "Sunset",
      majorFont: "Cambria",
      minorFont: "Calibri",
      major: { latin: "Cambria", eastAsia: "SimSun" },
      minor: { latin: "Calibri", eastAsia: "Microsoft YaHei" }
    }
  };
  // Theme is attached via the build()-state mutator surface
  const built = Document.build(d);
  const themed = { ...built, theme };
  fs.writeFileSync(path.join(outDir, "04-custom-theme.docx"), await toBuffer(themed));
  console.log("  → 04-custom-theme.docx");
}

// ---------------------------------------------------------------------------
// 5. Edge case: watermark + background together
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Combined", 1);
  Document.addParagraph(d, "DRAFT watermark + cream-coloured page.");
  Document.setBackground(d, { color: "FFFBE6" });
  Document.setWatermark(d, {
    type: "text",
    text: "DRAFT",
    color: "FFCC00",
    fontSize: ptToHalfPoint(96),
    rotation: -30,
    semiTransparent: true
  });
  fs.writeFileSync(path.join(outDir, "05-combined.docx"), await toBuffer(Document.build(d)));
  console.log("  → 05-combined.docx");
}
