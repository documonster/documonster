/**
 * Word Example 17 — Forms & Structured Document Tags
 *
 * Covers:
 *   - Legacy form fields (FORMTEXT / FORMCHECKBOX / FORMDROPDOWN)
 *   - extractFormFields / fillFormFields round-trip
 *   - Modern Structured Document Tags (SDTs)
 *     · plain-text / rich-text controls
 *     · checkbox SDT (w14:checkbox)
 *     · dropdown / combo-box
 *     · date picker
 *     · group, repeating section
 *     · data binding to a CustomXML part (OpenDoPE)
 *     · resolveDataBindings to populate from Custom XML
 *   - Edge cases: empty default value, an SDT with placeholder text,
 *     dropdown with one item, repeating section with no items.
 *
 * Output:
 *   - 17-forms-legacy.docx — original
 *   - 17-forms-legacy-filled.docx — programmatically filled
 *   - 17-sdt.docx — content controls
 *   - 17-sdt-resolved.docx — data bindings resolved
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  text,
  bold,
  formTextField,
  formCheckboxField,
  formDropdownField,
  structuredDocumentTag,
  extractFormFields,
  fillFormFields,
  resolveDataBindings,
  toBuffer
} from "../index";
import type { CustomXmlPart } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ===========================================================================
// Part 1 — Legacy form fields
// ===========================================================================
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Legacy form fields", 1);

  Document.addParagraphElement(
    doc,
    paragraph([bold("Full name: "), formTextField({ name: "FullName", default: "(your name)" })])
  );
  Document.addParagraphElement(
    doc,
    paragraph([bold("Agree to terms? "), formCheckboxField({ name: "AgreeTerms", checked: false })])
  );
  Document.addParagraphElement(
    doc,
    paragraph([
      bold("Country: "),
      formDropdownField({
        name: "Country",
        entries: ["", "Australia", "Canada", "Japan"],
        default: 0
      })
    ])
  );
  // Edge: dropdown with a single item
  Document.addParagraphElement(
    doc,
    paragraph([
      bold("Region (single option): "),
      formDropdownField({ name: "Region", entries: ["APAC"], default: 0 })
    ])
  );

  const built = Document.build(doc);

  // Extract → render as a list
  const fields = extractFormFields(built);
  console.log(`  legacy form has ${fields.length} fields:`);
  for (const f of fields) {
    console.log(`    · ${f.name} [${f.type}] = ${JSON.stringify(f.value)}`);
  }

  // Save the empty form
  const buf1 = await toBuffer(built);
  fs.writeFileSync(path.join(outDir, "17-forms-legacy.docx"), buf1);
  console.log(`  → 17-forms-legacy.docx (${buf1.length} bytes)`);

  // Fill programmatically
  const filled = fillFormFields(
    built,
    new Map<string, string | boolean | number>([
      ["FullName", "Jane Q. Public"],
      ["AgreeTerms", true],
      ["Country", 2], // Canada
      ["Region", 0]
    ])
  );
  const buf2 = await toBuffer(filled);
  fs.writeFileSync(path.join(outDir, "17-forms-legacy-filled.docx"), buf2);
  console.log(`  → 17-forms-legacy-filled.docx (${buf2.length} bytes)`);
}

// ===========================================================================
// Part 2 — Structured Document Tags (modern content controls)
// ===========================================================================
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Content controls (SDTs)", 1);

  // Plain-text control
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("Click here to enter your name")], {
      tag: "name",
      alias: "Customer Name",
      plainText: true,
      showingPlaceholder: true,
      appearance: "boundingBox"
    })
  );

  // Rich-text control
  Document.addContent(
    doc,
    structuredDocumentTag(
      [paragraph([bold("Notes "), text("(rich text — bold/italic allowed)")])],
      { tag: "notes", alias: "Notes", richText: true, appearance: "boundingBox" }
    )
  );

  // Date picker
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("YYYY-MM-DD")], {
      tag: "deadline",
      alias: "Deadline",
      date: { dateFormat: "yyyy-MM-dd", lid: "en-US", storeMappedDataAs: "dateTime" }
    })
  );

  // Dropdown list
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("Choose a department")], {
      tag: "dept",
      alias: "Department",
      dropdownList: [
        { displayText: "Engineering", value: "eng" },
        { displayText: "Sales", value: "sales" },
        { displayText: "Operations", value: "ops" }
      ]
    })
  );

  // Combo box (free-text + suggestions)
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("Type or pick a status")], {
      tag: "status",
      alias: "Status",
      comboBox: [
        { displayText: "Draft", value: "draft" },
        { displayText: "In review", value: "review" },
        { displayText: "Final", value: "final" }
      ]
    })
  );

  // Checkbox SDT (w14:checkbox). The visible content of the SDT is a paragraph
  // showing ☒ / ☐; the SDT properties carry the boolean state.
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("☒")], {
      tag: "shipped",
      alias: "Shipped",
      checkbox: { checked: true }
    })
  );
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("☐")], {
      tag: "received",
      alias: "Received",
      checkbox: { checked: false }
    })
  );

  // Repeating section: a placeholder for an arbitrary number of "Item" rows.
  // We seed it with one example item (the inner SDT must be a block-level
  // child of the outer SDT — both are wrapped via structuredDocumentTag).
  const itemTemplate = structuredDocumentTag([textParagraph("• Item placeholder")], {
    tag: "item",
    alias: "Item",
    repeatingSectionItem: true
  });
  Document.addContent(
    doc,
    structuredDocumentTag(
      [itemTemplate as unknown as Parameters<typeof structuredDocumentTag>[0][0]],
      {
        tag: "items",
        alias: "Items",
        repeatingSection: { sectionTitle: "Item", allowInsertDelete: true }
      }
    )
  );

  // Group SDT — read-only wrapper around a paragraph
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("This text is grouped (read-only) inside an SDT.")], {
      tag: "readonly",
      alias: "Read-only block",
      group: true,
      lockContent: true,
      lockSdt: true
    })
  );

  const built = Document.build(doc);
  const buf = await toBuffer(built);
  fs.writeFileSync(path.join(outDir, "17-sdt.docx"), buf);
  console.log(`  → 17-sdt.docx (${buf.length} bytes)`);
}

// ===========================================================================
// Part 3 — Data binding (OpenDoPE) — SDT bound to a CustomXML part
// ===========================================================================
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Data binding (OpenDoPE)", 1);

  const guid = "{11111111-2222-3333-4444-555555555555}";

  Document.addParagraph(doc, "Two SDTs below are bound to a CustomXML part:");

  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("[Customer]")], {
      tag: "customer",
      alias: "Customer",
      dataBinding: { xpath: "/invoice/customer", storeItemId: guid }
    })
  );
  Document.addContent(
    doc,
    structuredDocumentTag([textParagraph("[Total]")], {
      tag: "total",
      alias: "Total",
      dataBinding: { xpath: "/invoice/total", storeItemId: guid }
    })
  );

  // Attach the Custom XML part — the source of the bound values
  const customXml: CustomXmlPart = {
    itemId: guid,
    fileName: "item1.xml",
    xmlContent: "<invoice><customer>John Smith</customer><total>$1,234.56</total></invoice>"
  };

  // We can't use the public Document API to push customXmlParts; the
  // build() result includes a customXmlParts slot only when the handle has
  // already been seeded. Instead we attach the part to the built model.
  const built = Document.build(doc);
  const final = { ...built, customXmlParts: [customXml] };

  // First save: bindings present but unresolved (Word resolves them on open)
  const buf1 = await toBuffer(final);
  fs.writeFileSync(path.join(outDir, "17-sdt-bound.docx"), buf1);
  console.log(`  → 17-sdt-bound.docx (${buf1.length} bytes)`);

  // Resolve bindings ahead of time and save again — useful for tools
  // that don't know how to update bindings on open.
  const resolved = resolveDataBindings(final);
  const buf2 = await toBuffer(resolved);
  fs.writeFileSync(path.join(outDir, "17-sdt-resolved.docx"), buf2);
  console.log(`  → 17-sdt-resolved.docx (${buf2.length} bytes)`);
}
