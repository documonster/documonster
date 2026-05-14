/**
 * Word Example 22 — Streaming writer
 *
 * Covers:
 *   - createDocxStream() — incrementally write a large DOCX without
 *     keeping the whole body in memory.
 *   - addText / add (raw BodyContent) / addMany
 *   - Progress callback every N elements
 *   - Combining streaming with auxiliary parts (styles, headers/footers)
 *   - Edge case: finalising an empty stream produces a minimal valid DOCX.
 *
 * Output:
 *   - 22-streaming-large.docx — 5 000 paragraphs
 *   - 22-streaming-empty.docx — no body content
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createDocxStream, textParagraph, paragraph, bold, cmToTwips } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Large streaming write — 5 000 paragraphs, progress every 500
// ---------------------------------------------------------------------------
{
  const t0 = performance.now();
  const stream = createDocxStream({
    compressionLevel: 6,
    chunkSize: 500,
    sectionProperties: {
      pageSize: { width: cmToTwips(21), height: cmToTwips(29.7) },
      margins: { top: cmToTwips(2), bottom: cmToTwips(2), left: cmToTwips(2), right: cmToTwips(2) }
    },
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true, qFormat: true },
      // The body below uses Heading1 / Heading2 — register them so Word
      // doesn't complain about undefined pStyle references.
      {
        type: "paragraph",
        styleId: "Heading1",
        name: "heading 1",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        runProperties: { font: "Calibri Light", color: "2F5496", size: 32 }
      },
      {
        type: "paragraph",
        styleId: "Heading2",
        name: "heading 2",
        basedOn: "Normal",
        next: "Normal",
        qFormat: true,
        uiPriority: 9,
        runProperties: { font: "Calibri Light", color: "2F5496", size: 26 }
      }
    ],
    coreProperties: { title: "Streamed report", creator: "OpenCode" }
  });

  // Title
  stream.add(paragraph([bold("Streamed Report")], { style: "Heading1" }));

  let lastProgress = 0;
  for (let i = 1; i <= 5000; i++) {
    stream.addText(`Paragraph ${i}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.`);
    if (i - lastProgress >= 1000) {
      lastProgress = i;
      // No public progress hook on createDocxStream — but elementCount works
      console.log(`  ... ${stream.elementCount} elements written`);
    }
  }

  // A trailing element of a non-paragraph type (a TOC entry placeholder)
  stream.add(textParagraph("End of report.", { style: "Heading2" }));

  const buf = await stream.finalize();
  const t1 = performance.now();
  fs.writeFileSync(path.join(outDir, "22-streaming-large.docx"), buf);
  console.log(
    `  → 22-streaming-large.docx (${(buf.length / 1024).toFixed(1)} KB) — ${(t1 - t0).toFixed(0)} ms`
  );
}

// ---------------------------------------------------------------------------
// 2. Empty stream — finalize immediately
// ---------------------------------------------------------------------------
{
  const stream = createDocxStream();
  const buf = await stream.finalize();
  fs.writeFileSync(path.join(outDir, "22-streaming-empty.docx"), buf);
  console.log(`  → 22-streaming-empty.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 3. addMany convenience: pre-batched elements
// ---------------------------------------------------------------------------
{
  const stream = createDocxStream();
  stream.addMany(
    Array.from({ length: 100 }, (_, i) => ({
      type: "paragraph",
      children: [{ content: [{ type: "text", text: `Batch row ${i}` }] }]
    }))
  );
  const buf = await stream.finalize();
  fs.writeFileSync(path.join(outDir, "22-streaming-batch.docx"), buf);
  console.log(`  → 22-streaming-batch.docx (${buf.length} bytes)`);
}
