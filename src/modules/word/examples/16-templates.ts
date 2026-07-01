/**
 * Word Example 16 — Template engine
 *
 * Covers:
 *   - Plain `{{variable}}` substitution (with strict / non-strict mode)
 *   - `{{#if cond}} … {{else}} … {{/if}}` blocks
 *   - `{{#each items}} … {{/each}}` loops with `{{this}}` and item paths
 *   - Custom delimiters
 *   - Image / rich-text / sub-document / chart / HTML chunk placeholders
 *   - listTemplateTags() — introspection
 *   - fillTemplateFromSource using createJsonDataSource / createXmlDataSource / createCsvDataSource
 *   - patchDocument (the lower-level placeholder API)
 *   - Edge case: missing variable in non-strict mode, nested loops, deep paths
 *
 * Output: tmp/word-examples/16-templates/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Template, Units } from "../index";
import type { DocxDocument } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/16-templates"
);
fs.mkdirSync(outDir, { recursive: true });

// 1x1 png to use as image substitution
const redPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

// ---------------------------------------------------------------------------
// Helper: build the template document used by every test below
// ---------------------------------------------------------------------------
function buildTemplate(): DocxDocument {
  const d = Document.create();
  Document.useDefaultStyles(d);

  Document.addHeading(d, "Invoice for {{customer.name}}", 1);
  Document.addParagraph(d, "Date: {{date}}");
  Document.addParagraph(d, "Project: {{project}}");
  Document.addParagraph(d, "Status: {{#if active}}ACTIVE{{else}}archived{{/if}}");

  Document.addHeading(d, "Items", 2);
  // The Word template engine uses {{...}} for placeholders; the dollar sign
  // here is a literal currency marker.  We wrap the dollar in a separate
  // string concatenation so lint doesn't flag the literal "$" + "{{x}}"
  // sequence as a misplaced template literal.
  Document.addParagraph(
    d,
    "{{#each items}}• {{name}} — {{qty}} × $" + "{{price}} = $" + "{{total}}{{/each}}"
  );

  Document.addParagraph(d, "Sub-total: $" + "{{subtotal}}");
  Document.addParagraph(d, "Tax: $" + "{{tax}}");
  Document.addParagraph(d, "TOTAL: $" + "{{grandTotal}}");

  Document.addHeading(d, "Notes", 2);
  // Image, rich-text and HTML-chunk placeholders are recognised by
  // fillTemplateEnhanced only when they occupy a paragraph on their own.
  //   {{%path}} → image (TemplateImage)
  //   {{&path}} → rich text (Run[])
  //   {{!path}} → HTML chunk (TemplateHtmlChunk)
  Document.addParagraph(d, "Logo:");
  Document.addParagraph(d, "{{%logo}}");
  Document.addParagraph(d, "Custom rich text:");
  Document.addParagraph(d, "{{&greeting}}");
  Document.addParagraph(d, "{{? missing }}"); // non-strict resolves to empty
  return Document.build(d);
}

// ---------------------------------------------------------------------------
// 1. Strict fill — every placeholder must resolve
// ---------------------------------------------------------------------------
{
  const tpl = buildTemplate();
  // listTemplateTags requires the template to use simple {{...}}; placeholders
  // with leading '%' / '!' / '?' are still tags, so the listing helps when
  // designing the data shape.
  const tags = Template.listTemplateTags(tpl);
  console.log(
    `  detected ${tags.length} template tags, types: ${[...new Set(tags.map(t => t.type))].join(", ")}`
  );

  // Build the input data
  const items = [
    { name: "Widget", qty: 2, price: 12.5, total: 25 },
    { name: "Gadget", qty: 5, price: 4, total: 20 },
    { name: "Sprocket", qty: 3, price: 7.5, total: 22.5 }
  ];
  const subtotal = items.reduce((s, x) => s + x.total, 0);
  const tax = +(subtotal * 0.1).toFixed(2);
  const grandTotal = +(subtotal + tax).toFixed(2);

  const filled = Template.fillTemplate(
    tpl,
    {
      customer: { name: "Acme Corp" },
      date: "2026-05-09",
      project: "Q2 hardware order",
      active: true,
      items,
      subtotal,
      tax,
      grandTotal
      // %logo and !greeting are richer placeholder kinds only supported by
      // fillTemplateEnhanced (see test 2 below). The basic fillTemplate
      // engine does plain text substitution only — passing them here would
      // serialise the object as JSON into the document body.
    },
    { strict: false }
  );

  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "01-fillTemplate.docx"), buf);
  console.log(`  → 01-fillTemplate.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 2. fillTemplateEnhanced — supports image / richText / subDocument /
//    chart / HTML chunk values that the basic fillTemplate doesn't.
// ---------------------------------------------------------------------------
{
  const tpl = buildTemplate();
  const result = Template.fillTemplateEnhanced(
    tpl,
    {
      customer: { name: "Sparse Co." },
      date: "—",
      // Note: data keys do NOT carry the placeholder prefix — the engine
      // strips `%` / `!` from the placeholder, so `{{%logo}}` resolves to
      // data["logo"] (not data["%logo"]).
      logo: {
        image: { data: redPng, fileName: "enh-logo.png", mediaType: "png" },
        width: Units.cmToEmu(2),
        height: Units.cmToEmu(2)
      },
      greeting: [
        Build.bold("Dear customer, "),
        Build.italic("thank you "),
        Build.text("for your business.")
      ]
      // intentionally missing many tags — non-strict mode keeps going
    },
    { strict: false }
  );

  const buf = await Io.toBuffer(result);
  fs.writeFileSync(path.join(outDir, "02-enhanced.docx"), buf);
  console.log(`  → 02-enhanced.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 3. Custom delimiters: <% var %>
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Hello <% who %>", 1);
  Document.addParagraph(d, "Today is <% date %>.");
  const filled = Template.fillTemplate(
    Document.build(d),
    { who: "World", date: "Friday" },
    { delimiters: ["<%", "%>"] }
  );
  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "03-custom-delimiters.docx"), buf);
  console.log(`  → 03-custom-delimiters.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 4. JSON / XML / CSV data sources
// ---------------------------------------------------------------------------
{
  const tpl = buildTemplate();
  const json = Template.createJsonDataSource({
    customer: { name: "JSON Co." },
    date: "json-date",
    project: "from JSON",
    active: true,
    items: [{ name: "JSON-A", qty: 1, price: 10, total: 10 }],
    subtotal: 10,
    tax: 1,
    grandTotal: 11
  });
  const filled = Template.fillTemplateFromSource(tpl, json, { strict: false });
  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "04-json-source.docx"), buf);
  console.log(`  → 04-json-source.docx (${buf.length} bytes)`);
}

{
  const tpl = buildTemplate();
  const xml = Template.createXmlDataSource(`<?xml version="1.0"?>
    <root>
      <customer><name>XML Co.</name></customer>
      <date>xml-date</date>
      <project>from XML</project>
      <active>true</active>
      <subtotal>20</subtotal>
      <tax>2</tax>
      <grandTotal>22</grandTotal>
    </root>`);
  const filled = Template.fillTemplateFromSource(tpl, xml, { strict: false });
  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "05-xml-source.docx"), buf);
  console.log(`  → 05-xml-source.docx (${buf.length} bytes)`);
}

{
  // CSV gives an array under a key (default "rows", here "items" via rowsKey);
  // useful for #each loops.
  const csvSource = Template.createCsvDataSource(
    "name,qty,price,total\nWidget,2,12.5,25\nGadget,5,4,20\n",
    {
      rowsKey: "items"
    }
  );
  const data = csvSource.getData();
  const tpl = buildTemplate();
  const filled = Template.fillTemplate(
    tpl,
    {
      customer: { name: "CSV Co." },
      date: "csv-date",
      project: "from CSV",
      active: false,
      ...data,
      subtotal: 45,
      tax: 4.5,
      grandTotal: 49.5
    },
    { strict: false }
  );
  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "06-csv-source.docx"), buf);
  console.log(`  → 06-csv-source.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 5. Round-trip: package the template, then patchDocument with the
//    lower-level (non-template) placeholder API to replace simple text
//    placeholders. This shows the alternative pathway used by some
//    integrations that need a flat string-replacement model.
// ---------------------------------------------------------------------------
{
  // Use a template whose ONLY placeholders are the ones we patch — patchDocument
  // is a flat literal-string replacement API and does not understand template
  // control syntax ({{#if}}, {{#each}}, …). Feeding it the full invoice template
  // would leave all the control/loop tags unresolved in the output (they would
  // show up verbatim in Word). A focused template keeps the produced document
  // clean while still demonstrating the round-trip + patch workflow.
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Invoice for {{customer.name}}", 1);
  Document.addParagraph(d, "Date: {{date}}");
  Document.addParagraph(d, "Project: {{project}}");
  const tpl = Document.build(d);
  const tplBuf = await Io.toBuffer(tpl);
  // Round-trip: read it back, then patch with named placeholders that
  // happen to match the {{...}} tokens. patchDocument doesn't understand
  // template syntax — we treat each `{{tag}}` as a literal string. This
  // demonstrates that the same .docx is a valid input to both workflows.
  const round = await Io.read(tplBuf);
  void round; // unused; kept to demonstrate readDocx works on the template

  const out = await Io.patchDocument(tplBuf, [
    { placeholder: "{{customer.name}}", content: { type: "text", text: "Patch Co." } },
    { placeholder: "{{date}}", content: { type: "text", text: "patch-date" } },
    { placeholder: "{{project}}", content: { type: "text", text: "patched project" } }
  ]);
  fs.writeFileSync(path.join(outDir, "07-patchDocument.docx"), out);
  console.log(`  → 07-patchDocument.docx (${out.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 6. Edge case: nested {{#each}} with deep object paths.  Block-level each
//    blocks may straddle multiple paragraphs as long as the open/close tags
//    each live in their own paragraph (the engine pairs them by depth).
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Sales report", 1);
  Document.addParagraph(d, "{{#each regions}}");
  Document.addParagraph(d, "Region {{name}}:");
  Document.addParagraph(d, "{{#each products}}");
  Document.addParagraph(d, "  - {{name}}: {{units}} units");
  Document.addParagraph(d, "{{/each}}");
  Document.addParagraph(d, "{{/each}}");

  const filled = Template.fillTemplate(Document.build(d), {
    regions: [
      {
        name: "North",
        products: [
          { name: "Widget", units: 100 },
          { name: "Gadget", units: 50 }
        ]
      },
      {
        name: "South",
        products: [
          { name: "Widget", units: 80 },
          { name: "Gadget", units: 60 }
        ]
      }
    ]
  });
  const buf = await Io.toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "08-nested-each.docx"), buf);
  console.log(`  → 08-nested-each.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 7. Strict mode → TemplateError when a placeholder is missing
// ---------------------------------------------------------------------------
{
  const tpl = buildTemplate();
  try {
    Template.fillTemplate(tpl, { customer: { name: "Strict Co." } /* most fields missing */ });
    console.log("  ERROR: expected TemplateError to be thrown in strict mode");
  } catch (err) {
    if (err instanceof Template.TemplateError) {
      console.log(
        `  TemplateError caught: placeholder=${JSON.stringify(err.placeholder)}, location=${err.location}, tagName=${err.tagName ?? "(none)"}`
      );
    } else {
      throw err;
    }
  }
}

// ---------------------------------------------------------------------------
// 8. isTemplateChart guard — when feeding fillTemplateEnhanced an opaque
//    `unknown` map, this guard distinguishes a chart-shaped value from
//    other rich values.
// ---------------------------------------------------------------------------
{
  const candidates: unknown[] = [
    { chart: { type: "column", series: [], title: "x" } },
    { html: "<p>foo</p>" },
    [{ type: "paragraph", children: [] }], // sub-document
    "plain string"
  ];
  for (const c of candidates) {
    console.log(
      `  isTemplateChart(${JSON.stringify(c).slice(0, 50)}…) → ${Template.isTemplateChart(c)}`
    );
  }
}
