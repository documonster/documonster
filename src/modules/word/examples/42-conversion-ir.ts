/**
 * Word Example 42 — Conversion intermediate representation
 *
 * The conversion IR is the format-agnostic semantic model that lives between
 * DOCX and downstream targets (HTML, Markdown, ePub, etc.).  Building your
 * own exporter against this IR is much easier than threading the DOCX model
 * directly.
 *
 * Covers:
 *   - createConversionContext — collects warnings + asset registry
 *   - docxToSemantic — turn a DocxDocument into a SemanticDocument
 *   - Walking the SemanticDocument: blocks, inlines, headings, lists, tables,
 *     code, images
 *   - Custom export: render the SemanticDocument as plain text and as
 *     a tiny DSL ("our HTML" — single-line per block).
 *
 * Output: tmp/word-examples/42-ir/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Convert, Io, Units } from "../index";
import type { SemanticBlock, SemanticInline } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/42-ir"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Build a moderately rich source document
// ---------------------------------------------------------------------------
const tinyPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

const d = Document.create();
Document.useDefaultStyles(d);
Document.setCoreProperties(d, { title: "IR demo", creator: "OpenCode" });
Document.addHeading(d, "IR demo", 1);
Document.addHeading(d, "Section A", 2);
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.text("A paragraph with "),
    Build.bold("bold"),
    Build.text(", "),
    Build.italic("italic"),
    Build.text(", and "),
    Build.hyperlink("a link", { url: "https://example.com" }),
    Build.text(".")
  ])
);
Document.addBulletList(d, ["First", "Second", "Third"]);
Document.addNumberedList(d, ["Step 1", "Step 2"]);
Document.addTable(
  d,
  [
    ["Region", "Sales"],
    ["North", "120"],
    ["South", "90"]
  ],
  { headerRow: true, borders: true }
);
Document.addImage(d, tinyPng, "png", Units.cmToEmu(2), Units.cmToEmu(2), { altText: "tiny dot" });
Document.addFootnote(d, "Footnote at the end of section A.");
Document.addParagraph(d, "End of section A.");

const docModel = Document.build(d);
fs.writeFileSync(path.join(outDir, "00-source.docx"), await Io.toBuffer(docModel));

// ---------------------------------------------------------------------------
// 1. createConversionContext + docxToSemantic
//    docxToSemantic creates its own internal context and returns
//    { document, context }. createConversionContext is the building block
//    used by custom converters that drive the IR pipeline manually.
// ---------------------------------------------------------------------------
const standalone = Convert.createConversionContext();
standalone.addWarning("info", "demo", "createConversionContext usable on its own", "/example");
console.log(`  standalone context: ${standalone.warnings.length} warning(s) collected`);

const result = Convert.docxToSemantic(docModel);
const ir = result.document;
console.log(
  `  IR: ${ir.blocks.length} blocks, ${ir.assets.length} asset(s), ${ir.footnotes.length} footnote(s)`
);
console.log(`  metadata: ${JSON.stringify(ir.metadata)}`);
console.log(`  conversion context warnings: ${result.context.warnings.length}`);
for (const w of result.context.warnings) {
  console.log(`    [${w.severity}] ${w.code}: ${w.message}`);
}

// ---------------------------------------------------------------------------
// 2. Render the IR as plain text
// ---------------------------------------------------------------------------
function renderInline(node: SemanticInline): string {
  switch (node.type) {
    case "text": {
      let txt = node.text;
      const f = node.format;
      if (!f) {
        return txt;
      }
      // GFM emphasis tokens: stack from innermost outwards.
      if (f.code) {
        txt = "`" + txt + "`";
      }
      if (f.strikethrough) {
        txt = "~~" + txt + "~~";
      }
      if (f.bold && f.italic) {
        txt = "***" + txt + "***";
      } else if (f.bold) {
        txt = "**" + txt + "**";
      } else if (f.italic) {
        txt = "*" + txt + "*";
      }
      return txt;
    }
    case "lineBreak":
      return "  \n";
    case "link":
      return `[${node.children.map(renderInline).join("")}](${node.href})`;
    case "image":
      return `[image:${node.assetId}]`;
    case "code":
      return "`" + node.text + "`";
    case "math":
      return node.latex ? `$${node.latex}$` : node.text;
    case "footnoteRef":
      return `[^${node.id}]`;
    case "endnoteRef":
      return `[en:${node.id}]`;
    default:
      return "";
  }
}
function renderBlock(node: SemanticBlock, indent = ""): string {
  switch (node.type) {
    case "paragraph":
      return indent + node.children.map(renderInline).join("");
    case "heading":
      return indent + "#".repeat(node.level) + " " + node.children.map(renderInline).join("");
    case "list":
      return node.items
        .map(
          (item, i) =>
            indent +
            (node.ordered ? `${i + 1}. ` : "- ") +
            item.children.map(renderInline).join("") +
            (item.subList ? "\n" + renderBlock(item.subList, indent + "  ") : "")
        )
        .join("\n");
    case "table": {
      // GFM tables require a separator row of dashes after the header. Our
      // simple renderer assumes the first row is the header.
      const lines = node.rows.map(
        r =>
          indent +
          "| " +
          r.cells.map(c => c.children.map(b => renderBlock(b)).join(" ")).join(" | ") +
          " |"
      );
      if (lines.length > 0 && node.rows[0]) {
        const sep = indent + "| " + node.rows[0].cells.map(() => "---").join(" | ") + " |";
        lines.splice(1, 0, sep);
      }
      return lines.join("\n");
    }
    case "codeBlock":
      return `${indent}\`\`\`${node.language ?? ""}\n${node.text}\n${indent}\`\`\``;
    case "blockquote":
      return node.children.map(b => renderBlock(b, indent + "> ")).join("\n");
    case "horizontalRule":
      return indent + "---";
    case "image":
      return `${indent}[image:${node.assetId}]`;
    default:
      return "";
  }
}
const plain = ir.blocks.map(b => renderBlock(b)).join("\n\n");
fs.writeFileSync(path.join(outDir, "01-rendered.md"), plain);
console.log(`  → 01-rendered.md (${plain.length} chars)`);

// ---------------------------------------------------------------------------
// 3. Per-block summary
// ---------------------------------------------------------------------------
const summary: Record<string, number> = {};
for (const block of ir.blocks) {
  summary[block.type] = (summary[block.type] ?? 0) + 1;
}
console.log(`  block-type counts: ${JSON.stringify(summary)}`);

// ---------------------------------------------------------------------------
// 4. Edge: empty doc still produces a SemanticDocument
// ---------------------------------------------------------------------------
const empty = Convert.docxToSemantic(Document.build(Document.create()));
console.log(`  empty doc → ${empty.document.blocks.length} block(s)`);
