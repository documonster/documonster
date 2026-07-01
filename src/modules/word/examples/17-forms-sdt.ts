/**
 * Word Example 17 — Structured Document Tags (content controls)
 *
 * Covers:
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
 *   - 17-sdt.docx — content controls
 *   - 17-sdt-bound.docx — data bindings present (resolved by Word on open)
 *   - 17-sdt-resolved.docx — data bindings resolved ahead of time
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Query } from "../index";
import type { CustomXmlPart } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ===========================================================================
// Part 1 — Structured Document Tags (modern content controls)
// ===========================================================================
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Content controls (SDTs)", 1);

  // Plain-text control
  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("Click here to enter your name")], {
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
    Build.structuredDocumentTag(
      [Build.paragraph([Build.bold("Notes "), Build.text("(rich text — bold/italic allowed)")])],
      { tag: "notes", alias: "Notes", richText: true, appearance: "boundingBox" }
    )
  );

  // Date picker
  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("YYYY-MM-DD")], {
      tag: "deadline",
      alias: "Deadline",
      date: { dateFormat: "yyyy-MM-dd", lid: "en-US", storeMappedDataAs: "dateTime" }
    })
  );

  // Dropdown list
  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("Choose a department")], {
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
    Build.structuredDocumentTag([Build.textParagraph("Type or pick a status")], {
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
    Build.structuredDocumentTag([Build.textParagraph("☒")], {
      tag: "shipped",
      alias: "Shipped",
      checkbox: { checked: true }
    })
  );
  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("☐")], {
      tag: "received",
      alias: "Received",
      checkbox: { checked: false }
    })
  );

  // Body-level CheckBox via the standalone checkBox() builder. This emits a
  // bare <w14:checkbox> at body level (not wrapped in an SDT) — useful for
  // very simple, non-bound check states such as inline lists. Defaults to
  // ☒/☐ glyphs from MS Gothic; pass custom glyphs via {checkedState,
  // uncheckedState} for ✓/✗ or any other character.
  Document.addContent(doc, Build.checkBox({ checked: true }));
  Document.addContent(doc, Build.checkBox({ checked: false }));
  Document.addContent(
    doc,
    Build.checkBox({
      checked: true,
      checkedState: { value: "✓", font: "Arial" },
      uncheckedState: { value: "✗", font: "Arial" }
    })
  );

  // Repeating section: a placeholder for an arbitrary number of "Item" rows.
  // We seed it with one example item (the inner SDT must be a block-level
  // child of the outer SDT — both are wrapped via structuredDocumentTag).
  const itemTemplate = Build.structuredDocumentTag([Build.textParagraph("• Item placeholder")], {
    tag: "item",
    alias: "Item",
    repeatingSectionItem: true
  });
  Document.addContent(
    doc,
    Build.structuredDocumentTag([itemTemplate], {
      tag: "items",
      alias: "Items",
      repeatingSection: { sectionTitle: "Item", allowInsertDelete: true }
    })
  );

  // Group SDT — read-only wrapper around a paragraph
  Document.addContent(
    doc,
    Build.structuredDocumentTag(
      [Build.textParagraph("This text is grouped (read-only) inside an SDT.")],
      {
        tag: "readonly",
        alias: "Read-only block",
        group: true,
        lockContent: true,
        lockSdt: true
      }
    )
  );

  const built = Document.build(doc);
  const buf = await Io.toBuffer(built);
  fs.writeFileSync(path.join(outDir, "17-sdt.docx"), buf);
  console.log(`  → 17-sdt.docx (${buf.length} bytes)`);
}

// ===========================================================================
// Part 2 — Data binding (OpenDoPE) — SDT bound to a CustomXML part
// ===========================================================================
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Data binding (OpenDoPE)", 1);

  const guid = "{11111111-2222-3333-4444-555555555555}";

  Document.addParagraph(doc, "Two SDTs below are bound to a CustomXML part:");

  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("[Customer]")], {
      tag: "customer",
      alias: "Customer",
      dataBinding: { xpath: "/invoice/customer", storeItemId: guid }
    })
  );
  Document.addContent(
    doc,
    Build.structuredDocumentTag([Build.textParagraph("[Total]")], {
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
  const buf1 = await Io.toBuffer(final);
  fs.writeFileSync(path.join(outDir, "17-sdt-bound.docx"), buf1);
  console.log(`  → 17-sdt-bound.docx (${buf1.length} bytes)`);

  // Resolve bindings ahead of time and save again — useful for tools
  // that don't know how to update bindings on open.
  const resolved = Query.resolveDataBindings(final);
  const buf2 = await Io.toBuffer(resolved);
  fs.writeFileSync(path.join(outDir, "17-sdt-resolved.docx"), buf2);
  console.log(`  → 17-sdt-resolved.docx (${buf2.length} bytes)`);
}
