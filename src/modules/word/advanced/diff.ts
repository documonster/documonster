/**
 * DOCX Module - Document Comparison (Diff)
 *
 * Compares two DocxDocument models and returns a list of differences.
 * This is a text-level comparison — it compares paragraphs by their text content
 * and reports additions, deletions, and modifications.
 *
 * @example
 * ```ts
 * import { diffDocuments } from "excelts/word";
 *
 * const changes = diffDocuments(oldDoc, newDoc);
 * console.log(changes.filter(c => c.type === "modified"));
 * ```
 */

import { extractParagraphText } from "../core/text-utils";
import { walkBlocks } from "../core/walker";
import type { DocxDocument, BodyContent } from "../types";

/** Type of change detected between two documents. */
export type DiffChangeType = "added" | "deleted" | "modified" | "unchanged";

/** A single diff entry representing a change between two documents. */
export interface DiffEntry {
  /** Type of change. */
  readonly type: DiffChangeType;
  /** Paragraph index in the old document (undefined for "added"). */
  readonly oldIndex?: number;
  /** Paragraph index in the new document (undefined for "deleted"). */
  readonly newIndex?: number;
  /** Text content from old document. */
  readonly oldText?: string;
  /** Text content from new document. */
  readonly newText?: string;
}

/** Summary statistics for a diff result. */
export interface DiffSummary {
  /** Total number of paragraphs compared. */
  readonly totalParagraphs: number;
  /** Number of unchanged paragraphs. */
  readonly unchanged: number;
  /** Number of added paragraphs. */
  readonly added: number;
  /** Number of deleted paragraphs. */
  readonly deleted: number;
  /** Number of modified paragraphs. */
  readonly modified: number;
}

/** Result of a document comparison. */
export interface DiffResult {
  /** Individual change entries. */
  readonly entries: readonly DiffEntry[];
  /** Summary statistics. */
  readonly summary: DiffSummary;
}

/**
 * Compare two DocxDocuments and return a diff of their text content.
 *
 * Uses the Myers diff algorithm (LCS-based) for optimal minimal edit sequence.
 *
 * @param oldDoc - The original document.
 * @param newDoc - The modified document.
 * @returns DiffResult with entries and summary.
 */
export function diffDocuments(oldDoc: DocxDocument, newDoc: DocxDocument): DiffResult {
  const oldTexts = extractParagraphTexts(oldDoc.body);
  const newTexts = extractParagraphTexts(newDoc.body);

  const entries = computeDiff(oldTexts, newTexts);

  let added = 0;
  let deleted = 0;
  let modified = 0;
  let unchanged = 0;
  for (const e of entries) {
    switch (e.type) {
      case "added":
        added++;
        break;
      case "deleted":
        deleted++;
        break;
      case "modified":
        modified++;
        break;
      case "unchanged":
        unchanged++;
        break;
    }
  }

  return {
    entries,
    summary: {
      totalParagraphs: Math.max(oldTexts.length, newTexts.length),
      unchanged,
      added,
      deleted,
      modified
    }
  };
}

/** Extract all paragraph texts from body content (flattens tables). */
function extractParagraphTexts(body: readonly BodyContent[]): string[] {
  const texts: string[] = [];
  walkBlocks(body, {
    enterParagraph(para) {
      texts.push(extractParagraphText(para));
      return "skip";
    }
  });
  return texts;
}

/**
 * Compute diff using LCS (Longest Common Subsequence) approach.
 * This produces an optimal edit sequence.
 */
function computeDiff(oldTexts: string[], newTexts: string[]): DiffEntry[] {
  const m = oldTexts.length;
  const n = newTexts.length;

  // Build LCS table
  const dp: number[][] = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (oldTexts[i - 1] === newTexts[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  // Backtrack to produce diff entries
  const entries: DiffEntry[] = [];
  let i = m;
  let j = n;

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && oldTexts[i - 1] === newTexts[j - 1]) {
      entries.push({
        type: "unchanged",
        oldIndex: i - 1,
        newIndex: j - 1,
        oldText: oldTexts[i - 1],
        newText: newTexts[j - 1]
      });
      i--;
      j--;
    } else if (i > 0 && (j === 0 || dp[i - 1][j] >= dp[i][j - 1])) {
      entries.push({
        type: "deleted",
        oldIndex: i - 1,
        oldText: oldTexts[i - 1]
      });
      i--;
    } else {
      entries.push({
        type: "added",
        newIndex: j - 1,
        newText: newTexts[j - 1]
      });
      j--;
    }
  }

  // Reverse since we built it backwards
  entries.reverse();

  // Post-process: pair adjacent delete+add as "modified" when they're at the same position
  const result: DiffEntry[] = [];
  for (let k = 0; k < entries.length; k++) {
    const curr = entries[k];
    const next = k + 1 < entries.length ? entries[k + 1] : undefined;

    if (curr.type === "deleted" && next?.type === "added") {
      result.push({
        type: "modified",
        oldIndex: curr.oldIndex,
        newIndex: next.newIndex,
        oldText: curr.oldText,
        newText: next.newText
      });
      k++; // skip next
    } else {
      result.push(curr);
    }
  }

  return result;
}
