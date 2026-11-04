/**
 * Mail Merge API
 *
 * Replace MERGEFIELD fields in a document with values from a data map.
 * Uses the unified walker from core/walker.ts for consistent traversal
 * across body, headers, footers, footnotes, endnotes, tables, and SDTs.
 *
 * Note: This API mutates the document in place for backward compatibility.
 */

import {
  checkBoolean,
  checkHandle,
  checkNamedValues,
  checkRecord,
  checkString,
  namedValueField
} from "@word/core/argument-checks";
import { fieldSwitchArgument, formatFieldResult, tokenizeFieldCode } from "@word/core/field-format";
import { evaluateIfField } from "@word/core/field-if";
import type { SplitField } from "@word/core/field-view";
import {
  evaluateSplitFields,
  mapDocumentLists,
  resolveSplitFields,
  rewriteSplitFields
} from "@word/core/field-view";
import { hyperlinkContent, isHyperlink, isRun } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type {
  DocxDocument,
  FieldContent,
  Paragraph,
  ParagraphChild,
  RunContent
} from "@word/types";

// =============================================================================
// Public API
// =============================================================================

/**
 * Execute a mail merge: replace all MERGEFIELD fields in the document with values from the data map.
 *
 * Fields not found in the data map are left unchanged (or optionally cleared).
 * Traverses body, headers, footers, footnotes, endnotes, tables, SDTs, and comments.
 *
 * A value is shown as the field's switches say, as Word merges it: the
 * general `\*` (`Upper`, `Lower`, `Caps`, `FirstCap`, number formats), `\#`
 * (numeric picture) and `\@` (date picture) — see `formatFieldResult` — then
 * `\b "text"` before and `\f "text"` after it, only when it is not empty.
 * An `IF` field testing a merged field is evaluated again with the merged
 * value, inner fields first; its result is replaced, its code keeps the
 * values merged into it.
 *
 * @param doc - The document to modify (mutated in place).
 * @param data - Field name → replacement value, as a plain object or a `Map`.
 * @param options - Optional settings.
 * @returns The number of fields replaced.
 */
export function mailMerge(
  doc: DocxDocument,
  data: Readonly<Record<string, string>> | ReadonlyMap<string, string>,
  options?: {
    /** If true, remove fields not found in data. Default: false (leave unchanged). */
    removeUnmatched?: boolean;
  }
): number {
  checkHandle(doc, "document", "doc", "Document.mailMerge");
  const where = "Document.mailMerge";
  const entries = checkNamedValues(data, "data", where);
  for (const [name, value] of entries) {
    checkString(value, namedValueField("data", name), where);
  }
  const values: Record<string, string> =
    data instanceof Map
      ? Object.fromEntries(entries as [string, string][])
      : (data as Record<string, string>);
  checkRecord(options, "options", where, { optional: true });
  checkBoolean(options?.removeUnmatched, "options.removeUnmatched", where, { optional: true });
  let count = 0;
  const removeUnmatched = options?.removeUnmatched ?? false;

  // The fields to evaluate again once merged — found first, while the merge
  // fields they test are still fields.
  const dependents = new Set<FieldContent>();
  const merges = (field: FieldContent | SplitField): boolean =>
    mergeValue(field.instruction, values, removeUnmatched) !== undefined ||
    ("code" in field && field.code.some(part => typeof part !== "string" && merges(part)));
  mapDocumentLists(doc, list => {
    for (const field of resolveSplitFields(list)?.fields ?? []) {
      if (/^\s*IF\b/i.test(field.instruction) && merges(field)) {
        dependents.add(field.begin);
      }
    }
    return list;
  });

  walkDocument(
    doc,
    {
      enterParagraph(para: Paragraph) {
        count += mergeFieldsInParagraph(para, values, removeUnmatched);
        return "skip"; // We handle run content ourselves
      }
    },
    {
      includeHeaders: true,
      includeFooters: true,
      includeFootnotes: true,
      includeEndnotes: true,
      includeComments: true
    }
  );

  // Fields held in pieces span runs, wrappers and paragraphs: each is
  // replaced with its value across all of them, in the list of blocks
  // holding it.
  const merge = <T extends object>(list: readonly T[]): readonly T[] =>
    rewriteSplitFields(list, field => {
      if (field.deleted) {
        return undefined;
      }
      const value = mergeValue(field.instruction, values, removeUnmatched);
      if (value === undefined) {
        return undefined;
      }
      count++;
      return { replace: "field", text: value };
    });
  replaceDocumentLists(doc, mapDocumentLists(doc, merge));

  if (dependents.size > 0) {
    const evaluate = <T extends object>(list: readonly T[]): readonly T[] => {
      const resolved = resolveSplitFields(list);
      if (!resolved) {
        return list;
      }
      const values = evaluateSplitFields(
        resolved,
        (field, instruction) =>
          dependents.has(field.begin) && field.end && !field.deleted
            ? evaluateIfField(instruction)
            : undefined,
        content => content.cachedValue ?? ""
      );
      return rewriteSplitFields(
        list,
        field => {
          const value = values.get(field);
          return value === undefined || value === field.result
            ? undefined
            : { replace: "result", text: value };
        },
        resolved
      );
    };
    replaceDocumentLists(doc, mapDocumentLists(doc, evaluate));
  }

  return count;
}

// =============================================================================
// Internal helpers
// =============================================================================

/**
 * The text a merge field with `instruction` is replaced with: its value, the
 * empty string when unmatched and `removeUnmatched`, otherwise `undefined`.
 */
function mergeValue(
  instruction: string,
  data: Record<string, string>,
  removeUnmatched: boolean
): string | undefined {
  const tokens = tokenizeFieldCode(instruction);
  const [head, name] = tokens;
  if (
    !head ||
    head.quoted ||
    head.text.toUpperCase() !== "MERGEFIELD" ||
    !name ||
    (!name.quoted && name.text.startsWith("\\"))
  ) {
    return undefined;
  }
  const fieldName = name.text;
  // Use Object.hasOwn so prototype-chain entries (`__proto__`,
  // `toString`, `constructor`, …) cannot bind to merge fields. With
  // `in`, an unsanitised CSV header like `__proto__` would resolve
  // to `Object.prototype.__proto__` and inject a non-string value
  // (or even a function reference) into the document.
  if (Object.hasOwn(data, fieldName)) {
    const raw = (data as Record<string, unknown>)[fieldName];
    // Coerce to string defensively — extractMergeFields advertises
    // `Record<string, string>`, but callers commonly pass values
    // through unfettered JSON / CSV parsers.
    const value = typeof raw === "string" ? raw : raw == null ? "" : String(raw);
    if (value === "") {
      return "";
    }
    const before = fieldSwitchArgument(tokens, "\\b") ?? "";
    const after = fieldSwitchArgument(tokens, "\\f") ?? "";
    return before + formatFieldResult(value, instruction) + after;
  }
  return removeUnmatched ? "" : undefined;
}

/**
 * Put the lists of blocks of `next` — `doc` with some rewritten — into `doc`
 * itself, in place, as the rest of the merge works: the arrays and objects
 * left alone keep their identity.
 */
function replaceDocumentLists(doc: DocxDocument, next: DocxDocument): void {
  if (next === doc) {
    return;
  }
  const into = (target: readonly object[], source: readonly object[]): void => {
    if (target === source) {
      return;
    }
    (target as object[]).splice(0, target.length, ...source);
  };
  into(doc.body, next.body);
  for (const key of ["headers", "footers"] as const) {
    const parts = doc[key];
    const nextParts = next[key];
    if (!parts || !nextParts || parts === nextParts) {
      continue;
    }
    for (const [id, part] of parts) {
      const nextPart = nextParts.get(id);
      if (nextPart) {
        into(part.content.children, nextPart.content.children);
      }
    }
  }
  for (const key of ["footnotes", "endnotes", "comments"] as const) {
    const notes = doc[key] as readonly { readonly content: readonly object[] }[] | undefined;
    const nextNotes = next[key] as readonly { readonly content: readonly object[] }[] | undefined;
    if (!notes || !nextNotes || notes === nextNotes) {
      continue;
    }
    notes.forEach((note, i) => into(note.content, nextNotes[i]!.content));
  }
}

function mergeFieldsInParagraph(
  para: Paragraph,
  data: Record<string, string>,
  removeUnmatched: boolean
): number {
  let count = 0;

  // Recurse into hyperlinks too — MERGEFIELDs are commonly placed inside
  // hyperlink text (e.g. for personalized URLs). Also descend through
  // tracked-insert and moved-to wrappers so merge fields inside accepted
  // tracked changes still get filled. (Deleted / movedFromRun wrappers
  // represent removals and are skipped, matching replaceText's
  // convention.)
  const visit = (children: readonly ParagraphChild[]): void => {
    for (const child of children) {
      if (isHyperlink(child)) {
        visit(hyperlinkContent(child));
        continue;
      }
      if (
        "type" in child &&
        ((child as { type?: string }).type === "insertedRun" ||
          (child as { type?: string }).type === "movedToRun")
      ) {
        const inner = (child as { run?: ParagraphChild }).run;
        if (inner) {
          visit([inner]);
        }
        continue;
      }
      if (!isRun(child)) {
        continue;
      }
      const content = child.content as RunContent[];
      for (let j = 0; j < content.length; j++) {
        const c = content[j];
        // A field read in pieces is merged across the runs it spans, below.
        if (c.type !== "field" || c.fieldChar) {
          continue;
        }
        const text = mergeValue(c.instruction, data, removeUnmatched);
        if (text !== undefined) {
          content[j] = { type: "text", text } as RunContent;
          count++;
        }
      }
    }
  };
  visit(para.children);
  return count;
}
