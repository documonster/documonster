/**
 * Form Field Data API
 *
 * Extract and fill form field values in a document.
 */

import {
  checkHandle,
  checkNamedValues,
  invalidArgument,
  namedValueField
} from "@word/core/argument-checks";
import { markDerived } from "@word/core/ownership";
import { isHyperlink, isRun } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type { DocxVisitor } from "@word/core/walker";
import { DocxUnsupportedFeatureError } from "@word/errors";
import type {
  DocxDocument,
  BodyContent,
  CommentDef,
  DropDownFormField,
  EndnoteDef,
  FieldContent,
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
 * Traverses the body (paragraphs, tables, content controls, text boxes,
 * shapes), headers, footers, footnotes, endnotes and comments — every part
 * {@link fillFormFields} fills — and collects all legacy form field data.
 *
 * A value is what the field shows: a text field's result (what was typed into
 * it), falling back to its default text; a checkbox's checked state, falling
 * back to its default; a drop-down's selected index (`w:result`), falling back
 * to its default selection.
 *
 * @param doc - The document to extract from.
 * @returns Array of form field entries with their current values.
 */
export function extractFormFields(doc: DocxDocument): FormFieldEntry[] {
  checkHandle(doc, "document", "doc", "Query.listFormFields");
  return collectFormFields(doc).map(field => {
    const ff = field.formField!;
    let value: string | boolean | number;
    if (ff.type === "text") {
      value = field.cachedValue ?? ff.default ?? "";
    } else if (ff.type === "checkBox") {
      value = ff.checked ?? ff.default ?? false;
    } else {
      value = ff.result ?? ff.default ?? 0;
    }
    return {
      name: ff.name ?? "",
      type: ff.type,
      value,
      entries: ff.type === "dropDown" ? ff.entries : undefined
    };
  });
}

/** Every field carrying form field data, in document order, across the parts form fields are filled in. */
function collectFormFields(doc: DocxDocument): FieldContent[] {
  const fields: FieldContent[] = [];
  const visitor: DocxVisitor = {
    enterRun(run: Run) {
      for (const rc of run.content ?? []) {
        if (rc.type === "field" && rc.formField) {
          fields.push(rc);
        }
      }
      return "continue";
    }
  };
  walkDocument(doc, visitor, {
    includeHeaders: true,
    includeFooters: true,
    includeFootnotes: true,
    includeEndnotes: true,
    includeComments: true
  });
  return fields;
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
 * Each value sets what the field shows, as typing into it in Word does:
 * - a text field takes a string (or a finite number, written as text) as its
 *   result; its default text is left as it is, and a value longer than its
 *   `maxLength` is refused;
 * - a checkbox takes a boolean;
 * - a drop-down takes the index of an entry, or the entry's text, and selects
 *   it (`w:result`).
 *
 * A name no field has is skipped, so one set of values can be applied to forms
 * that lack some of the fields. Every value is checked against the fields of
 * its name before anything is filled.
 *
 * Does not modify `doc`; the result may share unchanged nodes with it; `Document.*`
 * copies what it edits, so an edit made through one never reaches the other.
 *
 * @param doc - The document to fill (returns a new copy with filled values).
 * @param values - Field name → new value, as a `Map` or a plain object.
 * @returns A new document with form fields populated.
 * @throws {DocxArgumentError} If a value is one its field cannot hold.
 * @throws {DocxUnsupportedFeatureError} If a text field to fill is held as
 * field pieces (its result is not in one run of runs the reader could join).
 */
export function fillFormFields(
  doc: DocxDocument,
  values:
    | ReadonlyMap<string, string | boolean | number>
    | Readonly<Record<string, string | boolean | number>>
): DocxDocument {
  checkHandle(doc, "document", "doc", "Transform.fillFormFields");
  const where = "Transform.fillFormFields";
  const entries = checkNamedValues(values, "values", where);
  for (const [name, value] of entries) {
    if (
      typeof value !== "string" &&
      typeof value !== "boolean" &&
      !(typeof value === "number" && Number.isFinite(value))
    ) {
      invalidArgument(
        where,
        namedValueField("values", name),
        "a string, boolean or finite number",
        value
      );
    }
  }
  const byField = new Map(entries as [string, string | boolean | number][]);
  const byName = new Map<string, FieldContent[]>();
  for (const field of collectFormFields(doc)) {
    const name = field.formField!.name ?? "";
    let list = byName.get(name);
    if (!list) {
      list = [];
      byName.set(name, list);
    }
    list.push(field);
  }
  for (const [name, value] of byField) {
    // A name no field has is skipped, so one set of values can be applied to
    // forms that lack some of the fields.
    for (const field of byName.get(name) ?? []) {
      checkFormFieldValue(field, name, value, where);
    }
  }
  const newBody = doc.body.map(block => fillFieldsInBlock(block, byField));

  const newHeaders = doc.headers
    ? new Map(
        Array.from(doc.headers, ([k, h]) => [
          k,
          {
            ...h,
            content: {
              ...h.content,
              children: h.content.children.map(
                c => fillFieldsInBlock(c as BodyContent, byField) as Paragraph | Table
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
                c => fillFieldsInBlock(c as BodyContent, byField) as Paragraph | Table
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
            content: fn.content.map(p => fillFieldsInParagraph(p, byField))
          }) as FootnoteDef
      )
    : undefined;
  const newEndnotes = doc.endnotes
    ? doc.endnotes.map(
        en =>
          ({
            ...en,
            content: en.content.map(p => fillFieldsInParagraph(p, byField))
          }) as EndnoteDef
      )
    : undefined;
  const newComments = doc.comments
    ? doc.comments.map(
        cm =>
          ({
            ...cm,
            content: cm.content.map(p => fillFieldsInParagraph(p, byField))
          }) as CommentDef
      )
    : undefined;

  return markDerived(
    {
      ...doc,
      body: newBody,
      ...(newHeaders ? { headers: newHeaders } : {}),
      ...(newFooters ? { footers: newFooters } : {}),
      ...(newFootnotes ? { footnotes: newFootnotes } : {}),
      ...(newEndnotes ? { endnotes: newEndnotes } : {}),
      ...(newComments ? { comments: newComments } : {})
    },
    doc
  );
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
    if (ff.type === "text") {
      const text = String(newVal);
      // A marker inside the old result keeps its place, up to the new end.
      const rangeMarkers = rc.rangeMarkers?.map(m =>
        m.part === "result" && m.offset > text.length ? { ...m, offset: text.length } : m
      );
      return { ...rc, cachedValue: text, ...(rangeMarkers ? { rangeMarkers } : {}) };
    }
    let newFF: FormField;
    if (ff.type === "checkBox") {
      newFF = { ...ff, checked: newVal as boolean };
    } else {
      newFF = { ...ff, result: dropDownIndex(ff, newVal) };
    }
    return { ...rc, formField: newFF };
  });
  return runModified ? { ...run, content: newContent } : run;
}

/** `value` checked against the field it fills, in the name of `where`. */
function checkFormFieldValue(
  field: FieldContent,
  name: string,
  value: string | boolean | number,
  where: string
): void {
  const ff = field.formField!;
  const at = namedValueField("values", name);
  if (ff.type === "text") {
    if (typeof value === "boolean") {
      invalidArgument(where, at, "a string or number for text field", value);
    }
    if (ff.maxLength !== undefined && ff.maxLength > 0 && String(value).length > ff.maxLength) {
      invalidArgument(
        where,
        at,
        `at most ${ff.maxLength} characters (the field's maxLength)`,
        value
      );
    }
    if (field.fieldChar !== undefined) {
      throw new DocxUnsupportedFeatureError(
        `${where}: text form field ${JSON.stringify(name)} whose result is held as field pieces`
      );
    }
  } else if (ff.type === "checkBox") {
    if (typeof value !== "boolean") {
      invalidArgument(where, at, "a boolean for a checkbox", value);
    }
  } else if (dropDownIndex(ff, value) === undefined) {
    const entries = ff.entries ?? [];
    invalidArgument(
      where,
      at,
      `an entry of the drop-down (${entries.map(e => JSON.stringify(e)).join(", ")}) or its index 0–${entries.length - 1}`,
      value
    );
  }
}

/** The index `value` selects in `ff`, or `undefined` when it names no entry. */
function dropDownIndex(
  ff: DropDownFormField,
  value: string | boolean | number
): number | undefined {
  const entries = ff.entries ?? [];
  if (typeof value === "number") {
    return Number.isInteger(value) && value >= 0 && value < entries.length ? value : undefined;
  }
  if (typeof value === "string") {
    const index = entries.indexOf(value);
    return index >= 0 ? index : undefined;
  }
  return undefined;
}
