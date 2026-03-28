/**
 * Example: PDF Export with Images
 *
 * Demonstrates embedding JPEG and PNG images in PDF output,
 * including PNG alpha transparency support.
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-images.ts
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, excelToPdf } from "../../../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../output/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const excelDataDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../excel/examples/data"
);

console.log("=== PDF Image Examples ===\n");

// =============================================================================
// 1. JPEG image embedded in PDF
// =============================================================================

const wb1 = new Workbook();
const ws1 = wb1.addWorksheet("JPEG Example");

ws1.columns = [
  { header: "Product", key: "product", width: 25 },
  { header: "Status", key: "status", width: 15 },
  { header: "Notes", key: "notes", width: 30 }
];

ws1.addRows([
  { product: "Widget A", status: "Active", notes: "Best seller" },
  { product: "Widget B", status: "Pending", notes: "New launch" },
  { product: "Widget C", status: "Active", notes: "Updated design" }
]);

// Add a JPEG image
const jpegPath = path.join(excelDataDir, "bubbles.jpg");
if (fs.existsSync(jpegPath)) {
  const jpegId = wb1.addImage({ filename: jpegPath, extension: "jpeg" });
  ws1.addImage(jpegId, {
    tl: { col: 0, row: 4 },
    ext: { width: 300, height: 200 }
  });

  const pdf1 = excelToPdf(wb1, { showGridLines: true });
  fs.writeFileSync(path.join(outDir, "images-jpeg.pdf"), pdf1);
  console.log("1. images-jpeg.pdf — table with embedded JPEG");
} else {
  console.log("1. SKIPPED — bubbles.jpg not found");
}

// =============================================================================
// 2. PNG image with alpha transparency
// =============================================================================

const wb2 = new Workbook();
const ws2 = wb2.addWorksheet("PNG Alpha");

ws2.getCell("A1").value = "PNG with transparency";
ws2.getCell("A1").font = { bold: true, size: 14 };
ws2.getCell("A2").value = "The image below has an alpha channel.";

for (let r = 3; r <= 12; r++) {
  ws2.getCell(`A${r}`).value = `Background row ${r}`;
  ws2.getCell(`B${r}`).value = r * 100;
  ws2.getCell(`C${r}`).value = "filler";
}

const pngPath = path.join(excelDataDir, "image2.png");
if (fs.existsSync(pngPath)) {
  const pngId = wb2.addImage({ filename: pngPath, extension: "png" });
  ws2.addImage(pngId, {
    tl: { col: 1, row: 2 },
    ext: { width: 150, height: 150 }
  });

  const pdf2 = excelToPdf(wb2, { showGridLines: true });
  fs.writeFileSync(path.join(outDir, "images-png-alpha.pdf"), pdf2);
  console.log("2. images-png-alpha.pdf — PNG image with alpha transparency");
} else {
  console.log("2. SKIPPED — image2.png not found");
}

// =============================================================================
// 3. Multiple images on the same sheet
// =============================================================================

const wb3 = new Workbook();
const ws3 = wb3.addWorksheet("Multi-Image");

ws3.getCell("A1").value = "Multiple Images";
ws3.getCell("A1").font = { bold: true, size: 14 };
for (let r = 2; r <= 20; r++) {
  ws3.getCell(`A${r}`).value = `Data row ${r}`;
  ws3.getCell(`B${r}`).value = r;
  ws3.getCell(`C${r}`).value = `Item ${r}`;
  ws3.getCell(`D${r}`).value = r * 1.5;
}

if (fs.existsSync(jpegPath) && fs.existsSync(pngPath)) {
  const img1 = wb3.addImage({ filename: jpegPath, extension: "jpeg" });
  const img2 = wb3.addImage({ filename: pngPath, extension: "png" });

  // Place images at different positions
  ws3.addImage(img1, {
    tl: { col: 4, row: 1 },
    ext: { width: 200, height: 150 }
  });
  ws3.addImage(img2, {
    tl: { col: 4, row: 10 },
    ext: { width: 150, height: 150 }
  });

  const pdf3 = excelToPdf(wb3, {
    showGridLines: true,
    showPageNumbers: true,
    title: "Multi-Image Report"
  });
  fs.writeFileSync(path.join(outDir, "images-multi.pdf"), pdf3);
  console.log("3. images-multi.pdf — multiple images on one sheet");
} else {
  console.log("3. SKIPPED — image files not found");
}

console.log("\nAll image examples generated.");
