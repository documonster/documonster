/**
 * Mail Merge API
 *
 * Replace MERGEFIELD fields in a document with values from a data map.
 * Uses the unified walker from core/walker.ts for consistent traversal
 * across body, headers, footers, footnotes, endnotes, tables, and SDTs.
 *
 * Note: This API mutates the document in place for backward compatibility.
 */

import { walkDocument } from "../core/walker";
import type { DocxDocument, Paragraph, Run, RunContent } from "../types";

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
  // hyperlink text (e.g. for personalized URLs).
  const visit = (children: readonly unknown[]): void => {
    for (const child of children) {
      if (
        child &&
        typeof child === "object" &&
        "type" in (child as Record<string, unknown>) &&
        (child as { type: unknown }).type === "hyperlink" &&
        Array.isArray((child as { children?: unknown[] }).children)
      ) {
        visit((child as { children: unknown[] }).children);
        continue;
      }
      if (
        !child ||
        typeof child !== "object" ||
        !("content" in (child as Record<string, unknown>))
      ) {
        continue;
      }
      const runLike = child as Run;
      if (!Array.isArray(runLike.content)) {
        continue;
      }
      const content = runLike.content as RunContent[];
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
        if (fieldName in data) {
          content[j] = { type: "text", text: data[fieldName] } as RunContent;
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
