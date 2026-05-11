/**
 * Example: Word Module — Complete Guide
 *
 * Covers:
 * - Document creation (paragraphs, headings, tables)
 * - Text formatting (bold, italic, underline, color, font size)
 * - Page layout (margins, orientation)
 * - Lists (bullet, numbered)
 * - Hyperlinks and bookmarks
 * - Reading and modifying existing .docx
 * - Search and replace
 * - Template engine (variables, conditions, loops)
 * - Form fields (extract and fill)
 * - Data binding (CustomXML / OpenDoPE)
 * - Drawing shapes with effects (shadow, glow, reflection, 3D)
 * - Document merge and diff
 * - Streaming writer
 * - Document protection
 * - Validation
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  textParagraph,
  paragraph,
  heading,
  text,
  bold,
  italic,
  underline,
  hyperlink,
  formTextField,
  formCheckboxField,
  formDropdownField,
  toBuffer,
  readDocx,
  resolveDataBindings,
  extractFormFields,
  fillFormFields,
  createRect,
  createEllipse,
  createShape,
  fillTemplate,
  mergeDocuments,
  diffDocuments,
  replaceText,
  extractText,
  protectDocument,
  isDocumentProtected,
  validateDocument,
  createDocxStream,
  cmToTwips,
  cmToEmu,
  ptToHalfPoint
} from "../index";
import type { DocxDocument } from "../types";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Basic Document Creation
// =============================================================================

console.log("=== 1. Basic Document Creation ===");

const doc1 = Document.create();
Document.addHeading(doc1, "Document Title", 1);
Document.addParagraph(doc1, "A simple paragraph with plain text.");
Document.addHeading(doc1, "Section A", 2);
Document.addParagraph(doc1, "Content under section A.");

const buf1 = await toBuffer(Document.build(doc1));
fs.writeFileSync(path.join(outDir, "01-basic.docx"), buf1);
console.log("  → 01-basic.docx");

// =============================================================================
// 2. Text Formatting
// =============================================================================

console.log("\n=== 2. Text Formatting ===");

const doc2 = Document.create();
Document.addHeading(doc2, "Formatted Text", 1);
Document.addParagraphElement(
  doc2,
  paragraph([text("Normal, "), bold("bold, "), italic("italic, "), underline("underlined.")])
);
Document.addParagraphElement(
  doc2,
  paragraph([
    {
      properties: { size: ptToHalfPoint(24), color: { val: "FF0000" } },
      content: [{ type: "text", text: "Red 24pt text" }]
    }
  ])
);

const buf2 = await toBuffer(Document.build(doc2));
fs.writeFileSync(path.join(outDir, "02-formatting.docx"), buf2);
console.log("  → 02-formatting.docx");

// =============================================================================
// 3. Tables
// =============================================================================

console.log("\n=== 3. Tables ===");

const doc3 = Document.create();
Document.addHeading(doc3, "Table Examples", 1);
Document.addTable(
  doc3,
  [
    ["Product", "Price", "Qty"],
    ["Widget A", "$10.00", "100"],
    ["Widget B", "$25.00", "50"]
  ],
  { headerRow: true, borders: true }
);

const buf3 = await toBuffer(Document.build(doc3));
fs.writeFileSync(path.join(outDir, "03-tables.docx"), buf3);
console.log("  → 03-tables.docx");

// =============================================================================
// 4. Page Layout
// =============================================================================

console.log("\n=== 4. Page Layout ===");

const doc4 = Document.create();
Document.addHeading(doc4, "A4 Portrait Page", 1);
Document.addParagraph(doc4, "This page uses A4 portrait with custom margins.");
Document.setSectionProperties(doc4, {
  pageSize: { width: cmToTwips(21), height: cmToTwips(29.7) },
  margins: {
    top: cmToTwips(2.54),
    bottom: cmToTwips(2.54),
    left: cmToTwips(3.17),
    right: cmToTwips(3.17)
  }
});

const buf4 = await toBuffer(Document.build(doc4));
fs.writeFileSync(path.join(outDir, "04-layout.docx"), buf4);
console.log("  → 04-layout.docx");

// =============================================================================
// 5. Lists
// =============================================================================

console.log("\n=== 5. Lists ===");

const doc5 = Document.create();
Document.addHeading(doc5, "Lists", 1);
Document.addBulletList(doc5, ["First bullet", "Second bullet", "Third bullet"]);
Document.addNumberedList(doc5, ["Step one", "Step two", "Step three"]);

const buf5 = await toBuffer(Document.build(doc5));
fs.writeFileSync(path.join(outDir, "05-lists.docx"), buf5);
console.log("  → 05-lists.docx");

// =============================================================================
// 6. Hyperlinks
// =============================================================================

console.log("\n=== 6. Hyperlinks ===");

const doc6 = Document.create();
Document.addHeading(doc6, "Links", 1);
Document.addParagraphElement(
  doc6,
  paragraph([
    text("Visit "),
    hyperlink("GitHub", { url: "https://github.com" }),
    text(" for code.")
  ])
);

const buf6 = await toBuffer(Document.build(doc6));
fs.writeFileSync(path.join(outDir, "06-links.docx"), buf6);
console.log("  → 06-links.docx");

// =============================================================================
// 7. Read & Modify
// =============================================================================

console.log("\n=== 7. Read & Modify ===");

const parsed = await readDocx(buf1);
console.log(`  Read doc: ${parsed.body.length} body elements`);
replaceText(parsed, "plain text", "MODIFIED text");
const buf7 = await toBuffer(parsed);
fs.writeFileSync(path.join(outDir, "07-modified.docx"), buf7);
console.log("  → 07-modified.docx");

// =============================================================================
// 8. Search & Replace
// =============================================================================

console.log("\n=== 8. Search & Replace ===");

const doc8 = Document.create();
Document.addParagraph(doc8, "Dear CUSTOMER, your order ORDER_ID is ready.");
const model8 = Document.build(doc8);
replaceText(model8, "CUSTOMER", "Alice");
replaceText(model8, "ORDER_ID", "#12345");
const buf8 = await toBuffer(model8);
fs.writeFileSync(path.join(outDir, "08-search-replace.docx"), buf8);
console.log(`  → 08-search-replace.docx ("${extractText(model8).trim()}")`);

// =============================================================================
// 9. Template Engine
// =============================================================================

console.log("\n=== 9. Template Engine ===");

const doc9 = Document.create();
Document.addHeading(doc9, "Report: {{company}}", 1);
Document.addParagraph(doc9, "Date: {{date}}");
Document.addParagraph(doc9, "{{#if active}}Status: Active{{/if}}");
Document.addParagraph(doc9, "{{#each items}}• {{this}}");
const model9 = Document.build(doc9);

const filled9 = fillTemplate(model9, {
  company: "Acme Corp",
  date: "2025-05-09",
  active: true,
  items: ["Widget A", "Widget B"]
});
const buf9 = await toBuffer(filled9);
fs.writeFileSync(path.join(outDir, "09-template.docx"), buf9);
console.log("  → 09-template.docx");

// =============================================================================
// 10. Form Fields
// =============================================================================

console.log("\n=== 10. Form Fields ===");

const doc10 = Document.create();
Document.addHeading(doc10, "Form", 1);
Document.addParagraphElement(
  doc10,
  paragraph([text("Name: "), formTextField({ name: "Name", default: "" })])
);
Document.addParagraphElement(
  doc10,
  paragraph([text("Agree: "), formCheckboxField({ name: "Agree", checked: false })])
);
Document.addParagraphElement(
  doc10,
  paragraph([
    text("Plan: "),
    formDropdownField({ name: "Plan", entries: ["Free", "Pro"], default: 0 })
  ])
);
const model10 = Document.build(doc10);

const fields = extractFormFields(model10);
console.log(`  Fields: ${fields.map(f => f.name).join(", ")}`);

const filledForm = fillFormFields(
  model10,
  new Map<string, string | boolean | number>([
    ["Name", "Jane"],
    ["Agree", true],
    ["Plan", 1]
  ])
);
const buf10 = await toBuffer(filledForm);
fs.writeFileSync(path.join(outDir, "10-forms.docx"), buf10);
console.log("  → 10-forms.docx");

// =============================================================================
// 11. Data Binding
// =============================================================================

console.log("\n=== 11. Data Binding ===");

const doc11: DocxDocument = {
  body: [
    heading("Invoice", 1),
    {
      type: "sdt",
      properties: {
        dataBinding: {
          xpath: "/invoice/customer",
          storeItemId: "{11111111-2222-3333-4444-555555555555}"
        }
      },
      content: [textParagraph("[Customer]")]
    }
  ],
  customXmlParts: [
    {
      itemId: "{11111111-2222-3333-4444-555555555555}",
      xmlContent: "<invoice><customer>John Smith</customer></invoice>",
      fileName: "item1.xml"
    }
  ]
};
const resolved11 = resolveDataBindings(doc11);
const buf11 = await toBuffer(resolved11);
fs.writeFileSync(path.join(outDir, "11-databinding.docx"), buf11);
console.log("  → 11-databinding.docx");

// =============================================================================
// 12. Drawing Shapes with Effects
// =============================================================================

console.log("\n=== 12. Shapes with Effects ===");

const doc12 = Document.create();
Document.addHeading(doc12, "Shapes", 1);
Document.addContent(
  doc12,
  createRect(cmToEmu(8), cmToEmu(3), {
    fill: { type: "solid", color: "4472C4" },
    effects: {
      shadow: {
        type: "outer",
        color: "000000",
        blurRadius: 50800,
        distance: 38100,
        direction: 2700000
      },
      softEdges: 63500
    }
  })
);
Document.addContent(
  doc12,
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
  doc12,
  createShape({
    shapeType: "roundRect",
    width: cmToEmu(10),
    height: cmToEmu(4),
    fill: { type: "solid", color: "70AD47" },
    effects: {
      effect3d: {
        camera: "perspectiveFront",
        bevelTop: { width: 127000, height: 63500, preset: "circle" }
      }
    }
  })
);
const buf12 = await toBuffer(Document.build(doc12));
fs.writeFileSync(path.join(outDir, "12-shapes.docx"), buf12);
console.log("  → 12-shapes.docx");

// =============================================================================
// 13. Document Merge
// =============================================================================

console.log("\n=== 13. Merge ===");

const docA = Document.create();
Document.addHeading(docA, "Chapter 1", 1);
Document.addParagraph(docA, "Content A.");
const docB = Document.create();
Document.addHeading(docB, "Chapter 2", 1);
Document.addParagraph(docB, "Content B.");

const merged = mergeDocuments([Document.build(docA), Document.build(docB)], {
  sectionBreak: "nextPage"
});
const buf13 = await toBuffer(merged);
fs.writeFileSync(path.join(outDir, "13-merged.docx"), buf13);
console.log("  → 13-merged.docx");

// =============================================================================
// 14. Diff
// =============================================================================

console.log("\n=== 14. Diff ===");

const oldDoc = Document.create();
Document.addParagraph(oldDoc, "Original line.");
const newDoc = Document.create();
Document.addParagraph(newDoc, "Modified line.");
Document.addParagraph(newDoc, "Added line.");

const diff = diffDocuments(Document.build(oldDoc), Document.build(newDoc));
console.log(
  `  Summary: ${diff.summary.added} added, ${diff.summary.deleted} deleted, ${diff.summary.modified} modified`
);

// =============================================================================
// 15. Streaming Writer
// =============================================================================

console.log("\n=== 15. Streaming Writer ===");

const stream = createDocxStream();
stream.addText("Large Document");
for (let i = 1; i <= 200; i++) {
  stream.addText(`Paragraph ${i}: Lorem ipsum.`);
}
const buf15 = await stream.finalize();
fs.writeFileSync(path.join(outDir, "15-streaming.docx"), buf15);
console.log(`  → 15-streaming.docx (${(buf15.length / 1024).toFixed(1)} KB)`);

// =============================================================================
// 16. Protection
// =============================================================================

console.log("\n=== 16. Protection ===");

const doc16 = Document.create();
Document.addParagraph(doc16, "Protected document.");
const model16 = Document.build(doc16);
const protected16 = await protectDocument(model16, { edit: "readOnly", password: "secret" });
console.log(`  Protected: ${isDocumentProtected(protected16)}`);
const buf16 = await toBuffer(protected16);
fs.writeFileSync(path.join(outDir, "16-protected.docx"), buf16);
console.log("  → 16-protected.docx");

// =============================================================================
// 17. Validation
// =============================================================================

console.log("\n=== 17. Validation ===");

const doc17 = Document.create();
Document.addParagraph(doc17, "Valid doc.");
const validation = validateDocument(Document.build(doc17));
console.log(`  Valid: ${validation.valid}, Issues: ${validation.issues.length}`);

// =============================================================================
// Done
// =============================================================================

console.log(`\n=== All examples written to ${outDir} ===`);
