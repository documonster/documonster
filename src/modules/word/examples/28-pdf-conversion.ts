/**
 * Word Example 28 — DOCX → PDF
 *
 * Covers:
 *   - Basic conversion using Document.build() → docxToPdf
 *   - Custom page size, margins, default font
 *   - Reading an existing .docx and converting it to PDF
 *   - Using a custom chart renderer (the built-in renderer is intentionally
 *     simple; consumers can plug in the high-quality Excel chart renderer
 *     for publication-grade charts)
 *
 * Output: tmp/word-examples/28-pdf/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { docxToPdf } from "../../pdf";
import {
  Document,
  paragraph,
  text,
  bold,
  toBuffer,
  readDocx,
  cmToEmu,
  ptToHalfPoint
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/28-pdf"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build a moderately rich document and convert
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "DOCX → PDF", 1);
  Document.addParagraph(d, "Plain paragraph rendered through the PDF flow renderer.");
  Document.addParagraphElement(
    d,
    paragraph([
      text("Mixed run: "),
      bold("bold"),
      text(", "),
      text("colored", { color: "C00000" }),
      text(", "),
      text("LARGE", { size: ptToHalfPoint(18) })
    ])
  );
  Document.addBulletList(d, ["First", "Second", "Third"]);
  Document.addNumberedList(d, ["Step 1", "Step 2", "Step 3"]);
  Document.addTable(
    d,
    [
      ["Quarter", "Revenue", "Profit"],
      ["Q1", "1.2M", "0.2M"],
      ["Q2", "1.5M", "0.3M"],
      ["Q3", "1.8M", "0.4M"]
    ],
    { headerRow: true, borders: true }
  );

  // Add an inline image
  const tinyPng = Uint8Array.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82
  ]);
  Document.addImage(d, tinyPng, "png", cmToEmu(1.5), cmToEmu(1.5), { altText: "logo" });

  const pdfBytes = await docxToPdf(Document.build(d));
  fs.writeFileSync(path.join(outDir, "01-basic.pdf"), pdfBytes);
  console.log(`  → 01-basic.pdf (${pdfBytes.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 2. Custom page size & margins
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Custom A5 portrait", 1);
  Document.addParagraph(d, "This PDF is rendered onto an A5 page (148 × 210 mm).");
  Document.addParagraph(d, "Lorem ipsum… ".repeat(80));

  const pdfBytes = await docxToPdf(Document.build(d), {
    // A5 portrait: 148mm × 210mm = 419.5 × 595.3 points
    pageWidth: 419.5,
    pageHeight: 595.3,
    marginTop: 36,
    marginBottom: 36,
    marginLeft: 36,
    marginRight: 36,
    defaultFont: "Helvetica",
    defaultFontSize: 10
  });
  fs.writeFileSync(path.join(outDir, "02-a5.pdf"), pdfBytes);
  console.log(`  → 02-a5.pdf (${pdfBytes.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 3. Read an existing .docx and convert it
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "From file", 1);
  Document.addParagraph(d, "This document was first written to disk and then re-read.");
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "03-source.docx"), buf);

  const reread = await readDocx(buf);
  const pdfBytes = await docxToPdf(reread);
  fs.writeFileSync(path.join(outDir, "03-from-file.pdf"), pdfBytes);
  console.log(`  → 03-from-file.pdf (${pdfBytes.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 4. Custom chart renderer — opt-in publication-grade rendering. Here we
//    register a stub that draws a coloured rectangle as a placeholder so
//    the output looks different from the built-in renderer.
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Chart renderer demo", 1);
  Document.addContent(d, {
    type: "chart",
    chart: {
      type: "column",
      title: "Stub chart",
      series: [
        {
          name: "MAU",
          categories: ["Jan", "Feb", "Mar"],
          values: [10, 20, 30],
          color: "4472C4"
        }
      ],
      legend: "b",
      width: cmToEmu(12),
      height: cmToEmu(7)
    }
  });

  const pdfBytes = await docxToPdf(Document.build(d), {
    chartRenderer: (chart, page, rect) => {
      // Draw a coloured filled rectangle then a label so the output is
      // visibly different from the default chart renderer.
      page.drawRect({
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        fill: { r: 0xde / 255, g: 0xeb / 255, b: 0xf7 / 255 }
      });
      page.drawText(`[Custom renderer: ${chart.title ?? "chart"}]`, {
        x: rect.x + 10,
        y: rect.y + rect.height / 2,
        fontFamily: "Helvetica",
        fontSize: 12,
        color: { r: 0x1f / 255, g: 0x4e / 255, b: 0x79 / 255 }
      });
    }
  });
  fs.writeFileSync(path.join(outDir, "04-chart-renderer.pdf"), pdfBytes);
  console.log(`  → 04-chart-renderer.pdf (${pdfBytes.length} bytes)`);
}
