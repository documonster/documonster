/**
 * Field listing: every field of a document as it reads, whole or held in
 * pieces (`FieldContent.fieldChar`), with the instruction and the result it
 * reads as.
 */

import { checkHandle } from "@word/core/argument-checks";
import { fieldCodeView, forEachField } from "@word/core/field-view";
import type { DocxDocument } from "@word/types";

/** A field of a document, as {@link listFields} reports it. */
export interface FieldEntry {
  /**
   * The field's whole instruction, e.g. `" PAGE "` or `" MERGEFIELD Name "`,
   * read across every run it is written in. A field nested in it appears as
   * Word's field codes view (Toggle Field Codes) shows it — its own code in
   * braces, `" IF { MERGEFIELD Vip } = \"Y\" \"VIP\" \"std\" "` — never by
   * the result it last showed; the nested field is listed after it.
   */
  readonly instruction: string;
  /** The instruction's first word, upper-cased (`"PAGE"`, `"MERGEFIELD"`, `"="`). */
  readonly type: string;
  /** What the field shows: its result as last computed (a paragraph break as `"\n"`). */
  readonly result: string;
  /**
   * How the document model stores it, not how many runs the file spread it
   * over: `false` when it is one `FieldContent` holding the instruction and
   * result (`w:fldSimple`, and a complex field the reader could gather into
   * one — its instruction may still have spanned several `w:instrText`
   * runs); `true` when it is kept as its `w:fldChar` begin, separate and end
   * pieces (`FieldContent.fieldChar`) at the places they are, because its
   * result spans paragraphs, it nests another field, or its field characters
   * share a run with other content.
   */
  readonly split: boolean;
  /** Whether it lies in a tracked deletion. */
  readonly deleted: boolean;
  /** The part of the document holding it. */
  readonly part: "body" | "header" | "footer" | "footnote" | "endnote" | "comment";
}

function fieldType(instruction: string): string {
  const trimmed = instruction.trim();
  return trimmed.startsWith("=") ? "=" : (trimmed.split(/\s+/)[0] ?? "").toUpperCase();
}

/**
 * Every field in `doc` — in the body, then the headers, footers, footnotes,
 * endnotes and comments — in document order, nested fields after the field
 * they are in. A field held in pieces is one entry, its instruction and
 * result read across the runs and paragraphs it spans.
 */
export function listFields(doc: DocxDocument): FieldEntry[] {
  checkHandle(doc, "document", "doc", "Query.listFields");
  const out: FieldEntry[] = [];
  const visit = (blocks: readonly object[], part: FieldEntry["part"]): void => {
    forEachField(blocks, field => {
      out.push({
        instruction: field.kind === "whole" ? field.instruction : fieldCodeView(field.field),
        type: fieldType(field.instruction),
        result: field.kind === "whole" ? (field.content.cachedValue ?? "") : field.field.result,
        split: field.kind === "split",
        deleted: field.deleted,
        part
      });
    });
  };
  visit(doc.body, "body");
  for (const header of doc.headers?.values() ?? []) {
    visit(header.content.children, "header");
  }
  for (const footer of doc.footers?.values() ?? []) {
    visit(footer.content.children, "footer");
  }
  for (const note of doc.footnotes ?? []) {
    visit(note.content, "footnote");
  }
  for (const note of doc.endnotes ?? []) {
    visit(note.content, "endnote");
  }
  for (const comment of doc.comments ?? []) {
    visit(comment.content, "comment");
  }
  return out;
}
