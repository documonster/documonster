/**
 * Word Example 13 — Drawing shapes & text boxes
 *
 * Covers:
 *   - Rectangle, rounded rect, ellipse with solid fill / no-fill
 *   - Star, polygon, callout, flowchart shape, arrow, line
 *   - Gradient fill, pattern fill
 *   - Outline customization (color, width, no-line)
 *   - Effects: shadow, glow, reflection, soft edges, 3D bevel
 *   - Rotation / flip
 *   - Shape with text inside
 *   - Plain text box (Document.addTextBox)
 *   - Edge case: very small / very large shape, "no-fill, no-outline" (invisible)
 *
 * Output: tmp/word-examples/13-shapes.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  textParagraph,
  paragraph,
  text,
  bold,
  createShape,
  createRect,
  createRoundRect,
  createEllipse,
  createLine,
  createArrow,
  createCallout,
  createFlowchartShape,
  createStar,
  drawingShape,
  cmToEmu,
  cmToTwips,
  toBuffer
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word — Drawing shapes & text boxes", 1);

// ---------------------------------------------------------------------------
// 1. Basic shapes
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Basic shapes (solid fill)", 2);
Document.addContent(
  doc,
  createRect(cmToEmu(6), cmToEmu(2), { fill: { type: "solid", color: "4472C4" } })
);
Document.addContent(
  doc,
  createRoundRect(cmToEmu(6), cmToEmu(2), { fill: { type: "solid", color: "70AD47" } })
);
Document.addContent(
  doc,
  createEllipse(cmToEmu(4), cmToEmu(4), { fill: { type: "solid", color: "ED7D31" } })
);

// ---------------------------------------------------------------------------
// 2. Gradient & pattern fill
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Gradient / pattern fills", 2);
Document.addContent(
  doc,
  createRect(cmToEmu(6), cmToEmu(2), {
    fill: {
      type: "gradient",
      angle: 5400000, // 90° in 1/60_000ths
      stops: [
        { position: 0, color: "FF6B6B" },
        { position: 100000, color: "6B5BFF" }
      ]
    }
  })
);
Document.addContent(
  doc,
  createRect(cmToEmu(6), cmToEmu(2), {
    fill: {
      type: "pattern",
      preset: "ltUpDiag",
      foregroundColor: "1F4E79",
      backgroundColor: "DEEBF7"
    }
  })
);

// ---------------------------------------------------------------------------
// 3. Outline / line / arrow
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Outline & lines", 2);
Document.addContent(
  doc,
  createRect(cmToEmu(6), cmToEmu(2), {
    fill: { type: "none" },
    outline: { color: "C00000", width: cmToEmu(0.1), dash: "dash" }
  })
);
Document.addContent(doc, createLine(cmToEmu(8), 0));
Document.addContent(doc, createArrow("right", cmToEmu(4), cmToEmu(1)));
Document.addContent(doc, createArrow("up", cmToEmu(2), cmToEmu(4)));

// ---------------------------------------------------------------------------
// 4. Stars, callouts, flowchart shapes
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Stars, callouts, flowchart", 2);
Document.addContent(
  doc,
  createStar(5, cmToEmu(4), cmToEmu(4), { fill: { type: "solid", color: "FFD700" } })
);
Document.addContent(
  doc,
  createCallout("rect", cmToEmu(6), cmToEmu(3), {
    fill: { type: "solid", color: "FFFFCC" },
    textBody: { paragraphs: [textParagraph("Speech bubble")] }
  })
);
Document.addContent(
  doc,
  createFlowchartShape("decision", cmToEmu(5), cmToEmu(3), {
    fill: { type: "solid", color: "F4B084" }
  })
);

// ---------------------------------------------------------------------------
// 5. Effects: shadow / glow / reflection / soft edges / 3D bevel
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Effects", 2);
Document.addContent(
  doc,
  createRect(cmToEmu(8), cmToEmu(3), {
    fill: { type: "solid", color: "4472C4" },
    effects: {
      shadow: {
        type: "outer",
        color: "000000",
        blurRadius: 50800,
        distance: 38100,
        direction: 2700000,
        transparency: 60
      }
    }
  })
);
Document.addContent(
  doc,
  createEllipse(cmToEmu(5), cmToEmu(5), {
    fill: {
      type: "gradient",
      stops: [
        { position: 0, color: "FF6B6B" },
        { position: 100000, color: "6B5BFF" }
      ]
    },
    effects: {
      glow: { color: "FFDD00", radius: 101600 },
      reflection: { startOpacity: 50, endOpacity: 0, distance: 25400 }
    }
  })
);
Document.addContent(
  doc,
  createShape({
    shapeType: "roundRect",
    width: cmToEmu(10),
    height: cmToEmu(4),
    fill: { type: "solid", color: "70AD47" },
    effects: {
      softEdges: 63500,
      effect3d: {
        camera: "perspectiveFront",
        bevelTop: { width: 127000, height: 63500, preset: "circle" }
      }
    }
  })
);

// ---------------------------------------------------------------------------
// 6. Rotation & flip
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Rotation & flip", 2);
Document.addContent(
  doc,
  createRect(cmToEmu(6), cmToEmu(2), {
    fill: { type: "solid", color: "1F4E79" },
    rotation: 30 * 60_000
  })
);
Document.addContent(
  doc,
  createArrow("right", cmToEmu(4), cmToEmu(1), {
    flipH: true,
    fill: { type: "solid", color: "C00000" }
  })
);

// ---------------------------------------------------------------------------
// 7. Shape with rich text inside
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Shape with text body", 2);
Document.addContent(
  doc,
  createRoundRect(cmToEmu(10), cmToEmu(4), {
    fill: { type: "solid", color: "DEEBF7" },
    outline: { color: "1F4E79", width: cmToEmu(0.06) },
    textBody: {
      anchor: "ctr",
      paragraphs: [
        paragraph([bold("Notice:")], { alignment: "center" }),
        paragraph([text("Multi-line text inside a shape.")], { alignment: "center" })
      ]
    }
  })
);

// ---------------------------------------------------------------------------
// 8. Plain text box (separate API: Document.addTextBox)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "8. Text box", 2);
Document.addTextBox(
  doc,
  [
    paragraph([bold("Sidebar")], { alignment: "center" }),
    textParagraph("A simple text box with a stroke and fill.")
  ],
  { width: cmToTwips(8), height: cmToTwips(4), stroke: true, fill: true }
);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// Tiny shape
Document.addContent(
  doc,
  createRect(cmToEmu(0.2), cmToEmu(0.2), { fill: { type: "solid", color: "C00000" } })
);

// Huge shape (page-wide)
Document.addContent(
  doc,
  createRect(cmToEmu(16), cmToEmu(0.5), { fill: { type: "solid", color: "1F4E79" } })
);

// Invisible shape (no fill, no outline) — still produces valid drawing XML
Document.addContent(
  doc,
  createRect(cmToEmu(4), cmToEmu(2), {
    fill: { type: "none" },
    outline: { noLine: true }
  })
);

// ---------------------------------------------------------------------------
// 9. Low-level drawingShape() builder — bypass createRect/createEllipse and
//    construct the DrawingShape value directly. Useful when you need an
//    OOXML preset that doesn't have a dedicated helper (e.g. "wave",
//    "doubleWave", "leftBracket", any of the 187 preset geometries).
// ---------------------------------------------------------------------------
Document.addHeading(doc, "9. Raw drawingShape builder", 2);
Document.addContent(
  doc,
  drawingShape({
    shapeType: "wave",
    width: cmToEmu(8),
    height: cmToEmu(2),
    fillColor: "F4B084",
    outlineColor: "1F4E79",
    outlineWidth: cmToEmu(0.06),
    altText: "wave shape via raw builder"
  })
);
Document.addContent(
  doc,
  drawingShape({
    shapeType: "leftBracket",
    width: cmToEmu(2),
    height: cmToEmu(4),
    noFill: true,
    outlineColor: "C00000",
    outlineWidth: cmToEmu(0.1)
  })
);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "13-shapes.docx"), buf);
console.log(`  → 13-shapes.docx (${buf.length} bytes)`);
