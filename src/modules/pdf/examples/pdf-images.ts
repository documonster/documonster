/**
 * Example: PDF Export with Images
 *
 * Demonstrates embedding JPEG and PNG images via:
 * 1. excelToPdf() — Excel workbook with images → PDF
 * 2. pdf() — Standalone PDF with images (no Excel)
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-images.ts
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, excelToPdf } from "../../../index";
import { pdf } from "../pdf";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const excelDataDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../excel/examples/data"
);

const jpegPath = path.join(excelDataDir, "bubbles.jpg");
const pngPath = path.join(excelDataDir, "image2.png");
const hasJpeg = fs.existsSync(jpegPath);
const hasPng = fs.existsSync(pngPath);

console.log("=== PDF Image Examples ===\n");
console.log("--- Part A: excelToPdf (Excel workbook with images) ---\n");

// =============================================================================
// A1. Excel → PDF: JPEG image
// =============================================================================

if (hasJpeg) {
  const wb = new Workbook();
  const ws = wb.addWorksheet("JPEG Example");

  ws.columns = [
    { header: "Product", key: "product", width: 25 },
    { header: "Status", key: "status", width: 15 },
    { header: "Notes", key: "notes", width: 30 }
  ];
  ws.addRows([
    { product: "Widget A", status: "Active", notes: "Best seller" },
    { product: "Widget B", status: "Pending", notes: "New launch" },
    { product: "Widget C", status: "Active", notes: "Updated design" }
  ]);

  const jpegId = wb.addImage({ buffer: fs.readFileSync(jpegPath), extension: "jpeg" });
  ws.addImage(jpegId, { tl: { col: 0, row: 4 }, ext: { width: 300, height: 200 } });

  fs.writeFileSync(
    path.join(outDir, "excel-images-jpeg.pdf"),
    excelToPdf(wb, { showGridLines: true })
  );
  console.log("A1. excel-images-jpeg.pdf — table with embedded JPEG");
} else {
  console.log("A1. SKIPPED — bubbles.jpg not found");
}

// =============================================================================
// A2. Excel → PDF: PNG with alpha
// =============================================================================

if (hasPng) {
  const wb = new Workbook();
  const ws = wb.addWorksheet("PNG Alpha");

  ws.getCell("A1").value = "PNG with transparency";
  ws.getCell("A1").font = { bold: true, size: 14 };
  for (let r = 2; r <= 12; r++) {
    ws.getCell(`A${r}`).value = `Row ${r}`;
    ws.getCell(`B${r}`).value = r * 100;
    ws.getCell(`C${r}`).value = "filler";
  }

  const pngId = wb.addImage({ buffer: fs.readFileSync(pngPath), extension: "png" });
  ws.addImage(pngId, { tl: { col: 1, row: 2 }, ext: { width: 150, height: 150 } });

  fs.writeFileSync(
    path.join(outDir, "excel-images-png.pdf"),
    excelToPdf(wb, { showGridLines: true })
  );
  console.log("A2. excel-images-png.pdf — PNG with alpha transparency");
} else {
  console.log("A2. SKIPPED — image2.png not found");
}

// =============================================================================
// A3. Excel → PDF: Multiple images
// =============================================================================

if (hasJpeg && hasPng) {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Multi-Image");

  ws.getCell("A1").value = "Multiple Images";
  ws.getCell("A1").font = { bold: true, size: 14 };
  for (let r = 2; r <= 20; r++) {
    ws.getCell(`A${r}`).value = `Data ${r}`;
    ws.getCell(`B${r}`).value = r;
    ws.getCell(`C${r}`).value = `Item ${r}`;
    ws.getCell(`D${r}`).value = r * 1.5;
    ws.getCell(`E${r}`).value = " ";
    ws.getCell(`F${r}`).value = " ";
    ws.getCell(`G${r}`).value = " ";
  }

  const img1 = wb.addImage({ buffer: fs.readFileSync(jpegPath), extension: "jpeg" });
  const img2 = wb.addImage({ buffer: fs.readFileSync(pngPath), extension: "png" });
  ws.addImage(img1, { tl: { col: 4, row: 1 }, ext: { width: 200, height: 150 } });
  ws.addImage(img2, { tl: { col: 4, row: 10 }, ext: { width: 150, height: 150 } });

  fs.writeFileSync(
    path.join(outDir, "excel-images-multi.pdf"),
    excelToPdf(wb, { showGridLines: true, showPageNumbers: true, title: "Multi-Image Report" })
  );
  console.log("A3. excel-images-multi.pdf — multiple images on one sheet");
} else {
  console.log("A3. SKIPPED — image files not found");
}

// =============================================================================
// Part B: Standalone pdf() with images
// =============================================================================

console.log("\n--- Part B: Standalone pdf() with images ---\n");

// =============================================================================
// B1. Standalone: JPEG image
// =============================================================================

if (hasJpeg) {
  const result = pdf(
    {
      columns: [
        { width: 20, header: "Product" },
        { width: 15, header: "Price" }
      ],
      data: [
        ["Widget A", "$10"],
        ["Widget B", "$25"],
        ["Widget C", "$15"]
      ],
      images: [
        {
          data: fs.readFileSync(jpegPath),
          format: "jpeg",
          col: 0,
          row: 4,
          width: 300,
          height: 200
        }
      ]
    },
    { showGridLines: true }
  );

  fs.writeFileSync(path.join(outDir, "standalone-images-jpeg.pdf"), result);
  console.log("B1. standalone-images-jpeg.pdf — standalone PDF with JPEG");
} else {
  console.log("B1. SKIPPED — bubbles.jpg not found");
}

// =============================================================================
// B2. Standalone: PNG with alpha
// =============================================================================

if (hasPng) {
  const rows: (string | number)[][] = [];
  for (let r = 1; r <= 10; r++) {
    rows.push([`Row ${r}`, r * 100, "filler"]);
  }

  const result = pdf(
    {
      name: "PNG Test",
      data: rows,
      images: [
        {
          data: fs.readFileSync(pngPath),
          format: "png",
          col: 1,
          row: 2,
          width: 150,
          height: 150
        }
      ]
    },
    { showGridLines: true }
  );

  fs.writeFileSync(path.join(outDir, "standalone-images-png.pdf"), result);
  console.log("B2. standalone-images-png.pdf — standalone PDF with PNG alpha");
} else {
  console.log("B2. SKIPPED — image2.png not found");
}

console.log("\nAll image examples generated.");
