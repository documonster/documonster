/**
 * Form Field Data API
 *
 * Extract and fill form field values in a document.
 */

import { isHyperlink, isRun } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type { DocxVisitor } from "@word/core/walker";
import type {
  DocxDocument,
  BodyContent,
  CommentDef,
  EndnoteDef,
  FootnoteDef,
  FooterDef,
  FormField,
  HeaderDef,
  Paragraph,
  ParagraphChild,
  Run,
  StructuredDocumentTag,
  Table
} from "@word/types";

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
 * Fill form field values throughout a document.
 *
 * The body, headers, footers, footnotes, endnotes and comments are all
 * processed so a form field placed inside a header (a common case) still
 * gets filled. Inside paragraphs the visitor descends into hyperlinks and
 * track-change wrappers so a form field referenced from a tracked
 * insertion is also reachable.
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

  const newHeaders = doc.headers
    ? new Map(
        Array.from(doc.headers, ([k, h]) => [
          k,
          {
            ...h,
            content: {
              ...h.content,
              children: h.content.children.map(
                c => fillFieldsInBlock(c as BodyContent, values) as Paragraph | Table
              )
            }
          } as HeaderDef
        ])
      )
    : undefined;
  const newFooters = doc.footers
    ? new Map(
        Array.from(doc.footers, ([k, f]) => [
          k,
          {
            ...f,
            content: {
              ...f.content,
              children: f.content.children.map(
                c => fillFieldsInBlock(c as BodyContent, values) as Paragraph | Table
              )
            }
          } as FooterDef
        ])
      )
    : undefined;
  const newFootnotes = doc.footnotes
    ? doc.footnotes.map(
        fn =>
          ({
            ...fn,
            content: fn.content.map(p => fillFieldsInParagraph(p, values))
          }) as FootnoteDef
      )
    : undefined;
  const newEndnotes = doc.endnotes
    ? doc.endnotes.map(
        en =>
          ({
            ...en,
            content: en.content.map(p => fillFieldsInParagraph(p, values))
          }) as EndnoteDef
      )
    : undefined;
  const newComments = doc.comments
    ? doc.comments.map(
        cm =>
          ({
            ...cm,
            content: cm.content.map(p => fillFieldsInParagraph(p, values))
          }) as CommentDef
      )
    : undefined;

  return {
    ...doc,
    body: newBody,
    ...(newHeaders ? { headers: newHeaders } : {}),
    ...(newFooters ? { footers: newFooters } : {}),
    ...(newFootnotes ? { footnotes: newFootnotes } : {}),
    ...(newEndnotes ? { endnotes: newEndnotes } : {}),
    ...(newComments ? { comments: newComments } : {})
  };
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
        if (
          c &&
          typeof c === "object" &&
          !("type" in c) &&
          "content" in c &&
          Array.isArray((c as { content?: unknown }).content)
        ) {
          // Inline (run-only) SDT child — fill any field run-content nodes
          // it carries directly.
          return fillFieldsInRun(c as Run, values);
        }
        return c;
      })
    } as StructuredDocumentTag;
  }
  if (block.type === "textBox") {
    return {
      ...block,
      content: block.content.map(p => fillFieldsInParagraph(p, values))
    };
  }
  if (block.type === "drawingShape") {
    if (!block.textContent || block.textContent.length === 0) {
      return block;
    }
    return {
      ...block,
      textContent: block.textContent.map(p => fillFieldsInParagraph(p, values))
    };
  }
  if (block.type === "tableOfContents") {
    if (!block.cachedParagraphs || block.cachedParagraphs.length === 0) {
      return block;
    }
    return {
      ...block,
      cachedParagraphs: block.cachedParagraphs.map(p => fillFieldsInParagraph(p, values))
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
    // Hyperlink: descend into its run children.
    if (isHyperlink(child)) {
      let hlModified = false;
      const newRuns = child.children.map(r => {
        const filled = fillFieldsInRun(r, values);
        if (filled !== r) {
          hlModified = true;
        }
        return filled;
      });
      if (hlModified) {
        modified = true;
        return { ...child, children: newRuns } as ParagraphChild;
      }
      return child;
    }
    // Track-change wrappers around a run.
    if (
      "type" in child &&
      ((child as { type?: string }).type === "insertedRun" ||
        (child as { type?: string }).type === "movedToRun")
    ) {
      const wrapper = child as { run: Run; type: string } & ParagraphChild;
      const filled = fillFieldsInRun(wrapper.run, values);
      if (filled !== wrapper.run) {
        modified = true;
        return { ...wrapper, run: filled } as ParagraphChild;
      }
      return child;
    }
    if (!isRun(child)) {
      return child;
    }
    const filled = fillFieldsInRun(child, values);
    if (filled !== child) {
      modified = true;
    }
    return filled;
  });
  return modified ? { ...para, children: newChildren } : para;
}

function fillFieldsInRun(run: Run, values: ReadonlyMap<string, string | boolean | number>): Run {
  if (!run.content) {
    return run;
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
  return runModified ? { ...run, content: newContent } : run;
}
