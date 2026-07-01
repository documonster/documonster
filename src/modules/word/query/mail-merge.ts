/**
 * Mail Merge API
 *
 * Replace MERGEFIELD fields in a document with values from a data map.
 * Uses the unified walker from core/walker.ts for consistent traversal
 * across body, headers, footers, footnotes, endnotes, tables, and SDTs.
 *
 * Note: This API mutates the document in place for backward compatibility.
 */

import { isHyperlink, isRun } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type { DocxDocument, Paragraph, ParagraphChild, RunContent } from "@word/types";

// =============================================================================
// Constants
// =============================================================================

/** Regex to parse MERGEFIELD instruction: MERGEFIELD "FieldName" or MERGEFIELD FieldName */
const MERGEFIELD_RE = /^\s*MERGEFIELD\s+(?:"([^"]+)"|(\S+))/i;

// =============================================================================
// Public API
// =============================================================================

/**
 * Execute a mail merge: replace all MERGEFIELD fields in the document with values from the data map.
 *
 * Fields not found in the data map are left unchanged (or optionally cleared).
 * Traverses body, headers, footers, footnotes, endnotes, tables, SDTs, and comments.
 *
 * @param doc - The document to modify (mutated in place).
 * @param data - Map of field names to replacement values.
 * @param options - Optional settings.
 * @returns The number of fields replaced.
 */
export function mailMerge(
  doc: DocxDocument,
  data: Record<string, string>,
  options?: {
    /** If true, remove fields not found in data. Default: false (leave unchanged). */
    removeUnmatched?: boolean;
  }
): number {
  let count = 0;
  const removeUnmatched = options?.removeUnmatched ?? false;

  walkDocument(
    doc,
    {
      enterParagraph(para: Paragraph) {
        count += mergeFieldsInParagraph(para, data, removeUnmatched);
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

  return count;
}

// =============================================================================
// Internal helpers
// =============================================================================

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
        visit(child.children as readonly ParagraphChild[]);
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
        if (c.type !== "field") {
          continue;
        }
        const match = MERGEFIELD_RE.exec(c.instruction);
        if (!match) {
          continue;
        }
        const fieldName = match[1] ?? match[2];
        // Use Object.hasOwn so prototype-chain entries (`__proto__`,
        // `toString`, `constructor`, …) cannot bind to merge fields. With
        // `in`, an unsanitised CSV header like `__proto__` would resolve
        // to `Object.prototype.__proto__` and inject a non-string value
        // (or even a function reference) into the document.
        if (fieldName !== undefined && Object.hasOwn(data, fieldName)) {
          const raw = (data as Record<string, unknown>)[fieldName];
          // Coerce to string defensively — extractMergeFields advertises
          // `Record<string, string>`, but callers commonly pass values
          // through unfettered JSON / CSV parsers.
          const text = typeof raw === "string" ? raw : raw == null ? "" : String(raw);
          content[j] = { type: "text", text } as RunContent;
          count++;
        } else if (removeUnmatched) {
          content[j] = { type: "text", text: "" } as RunContent;
          count++;
        }
      }
    }
  };
  visit(para.children);
  return count;
}
