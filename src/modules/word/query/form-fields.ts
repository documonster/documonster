/**
 * Form Field Data API
 *
 * Extract and fill form field values in a document.
 */

import { walkDocument } from "../core/walker";
import type { DocxVisitor } from "../core/walker";
import type { DocxDocument, BodyContent, Paragraph, Run, Table, FormField } from "../types";

// =============================================================================
// Types
// =============================================================================

/** Result of extracting form field values from a document. */
export interface FormFieldEntry {
  /** Field name (from ffData). */
  readonly name: string;
  /** Field type. */
  readonly type: "text" | "checkBox" | "dropDown";
  /** Current value (text content for text fields, checked state for checkbox, selected index for dropdown). */
  readonly value: string | boolean | number;
  /** For dropdown: the available entries. */
  readonly entries?: readonly string[];
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract all form field values from a document.
 *
 * Traverses the document body (paragraphs, tables, headers, footers)
 * and collects all legacy form field data.
 *
 * @param doc - The document to extract from.
 * @returns Array of form field entries with their current values.
 */
export function extractFormFields(doc: DocxDocument): FormFieldEntry[] {
  const results: FormFieldEntry[] = [];

  const visitor: DocxVisitor = {
    enterRun(run: Run) {
      if (!run.content) {
        return "continue";
      }
      for (const rc of run.content) {
        if (rc.type === "field" && rc.formField) {
          const ff = rc.formField;
          let value: string | boolean | number;
          if (ff.type === "text") {
            value = ff.default ?? "";
          } else if (ff.type === "checkBox") {
            value = ff.checked ?? ff.default ?? false;
          } else {
            // dropdown
            value = ff.default ?? 0;
          }
          results.push({
            name: ff.name ?? "",
            type: ff.type,
            value,
            entries: ff.type === "dropDown" ? ff.entries : undefined
          });
        }
      }
      return "continue";
    }
  };

  walkDocument(doc, visitor, {
    includeHeaders: true,
    includeFooters: true,
    includeFootnotes: false,
    includeEndnotes: false,
    includeComments: false
  });

  return results;
}

/**
 * Fill form field values in a document.
 *
 * @param doc - The document to fill (returns a new copy with filled values).
 * @param values - Map of field name → new value.
 * @returns A new document with form fields populated.
 */
export function fillFormFields(
  doc: DocxDocument,
  values: ReadonlyMap<string, string | boolean | number>
): DocxDocument {
  const newBody = doc.body.map(block => fillFieldsInBlock(block, values));
  return { ...doc, body: newBody };
}

// =============================================================================
// Internal helpers
// =============================================================================

function fillFieldsInBlock(
  block: BodyContent,
  values: ReadonlyMap<string, string | boolean | number>
): BodyContent {
  if (block.type === "paragraph") {
    return fillFieldsInParagraph(block, values);
  }
  if (block.type === "table") {
    return {
      ...block,
      rows: block.rows.map(r => ({
        ...r,
        cells: r.cells.map(c => ({
          ...c,
          content: c.content.map(inner =>
            fillFieldsInBlock(inner as BodyContent, values)
          ) as readonly (Paragraph | Table)[]
        }))
      }))
    };
  }
  if (block.type === "sdt") {
    return {
      ...block,
      content: block.content.map(c => {
        if ("type" in c && (c.type === "paragraph" || c.type === "table")) {
          return fillFieldsInBlock(c as Paragraph | Table, values) as Paragraph | Table;
        }
        return c;
      })
    };
  }
  return block;
}

function fillFieldsInParagraph(
  para: Paragraph,
  values: ReadonlyMap<string, string | boolean | number>
): Paragraph {
  let modified = false;
  const newChildren = para.children.map(child => {
    if (!("content" in child)) {
      return child;
    }
    const run = child as Run;
    if (!run.content) {
      return child;
    }
    let runModified = false;
    const newContent = run.content.map(rc => {
      if (rc.type !== "field" || !rc.formField) {
        return rc;
      }
      const ff = rc.formField;
      const name = ff.name ?? "";
      if (!values.has(name)) {
        return rc;
      }
      const newVal = values.get(name)!;
      runModified = true;
      let newFF: FormField;
      if (ff.type === "text") {
        newFF = { ...ff, default: String(newVal) };
      } else if (ff.type === "checkBox") {
        newFF = { ...ff, checked: Boolean(newVal) };
      } else {
        // dropdown: value is the selected index
        newFF = {
          ...ff,
          default: typeof newVal === "number" ? newVal : parseInt(String(newVal), 10) || 0
        };
      }
      return { ...rc, formField: newFF };
    });
    if (runModified) {
      modified = true;
      return { ...run, content: newContent };
    }
    return child;
  });
  return modified ? { ...para, children: newChildren } : para;
}
