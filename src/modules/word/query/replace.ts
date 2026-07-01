/**
 * Document Replace API
 *
 * Uses the unified walker from core/walker.ts for consistent traversal
 * across body, headers, footers, footnotes, endnotes, comments, and SDTs.
 *
 * Note: This API mutates the document in place for backward compatibility.
 *
 * Implementation notes:
 *   We always operate on the paragraph's concatenated plain text and the
 *   ordered list of text-bearing run-content nodes ("text segments").
 *   This guarantees consistent semantics regardless of whether matches sit
 *   inside a single run or straddle multiple runs:
 *     1. Build segment table [seg0, seg1, ...] with absolute offsets.
 *     2. Compute every match using a global regex (string searches are
 *        promoted to a global regex with the literal escaped).
 *     3. Walk segments and matches together. For each segment we emit:
 *        - the literal (un-matched) prefix/middle/suffix slices, and
 *        - the replacement string the first time we see each match.
 *      Subsequent segments overlapping the same match contribute the empty
 *      string (so the replacement is not duplicated) but otherwise keep
 *      their original formatting (RunProperties remain on the parent run).
 *   The returned count equals the number of matches actually replaced.
 */

import { isHyperlink, isRun } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type {
  DocxDocument,
  Paragraph,
  ParagraphChild,
  Run,
  StructuredDocumentTag
} from "@word/types";

// =============================================================================
// Public API
// =============================================================================

/**
 * Replace text in a document (mutates in place).
 *
 * Performs document-wide text replacement within run content. Matches are
 * applied uniformly whether they fall inside a single run or span several
 * runs; the formatting of each run that survives the replacement is preserved.
 *
 * Traverses body, headers, footers, footnotes, endnotes, comments, tables, and
 * SDTs using the unified document walker. SDTs may carry inline runs without
 * an enclosing paragraph (content controls wrapping a single Run); those
 * inline runs are treated as a single virtual paragraph for replacement.
 *
 * @param doc - The document model to modify (mutated in place).
 * @param search - String or RegExp to find. Both are treated as global
 *                 (every occurrence is replaced); the `g` flag on a RegExp
 *                 is therefore optional.
 * @param replacement - Replacement string. When `search` is a RegExp the
 *                      replacement supports the standard `$1`/`$&`/`$$` etc.
 *                      backreferences.
 * @returns The exact number of replacements made.
 */
export function replaceText(
  doc: DocxDocument,
  search: string | RegExp,
  replacement: string
): number {
  let totalCount = 0;

  walkDocument(
    doc,
    {
      enterParagraph(para: Paragraph) {
        totalCount += replaceInParagraph(para, search, replacement);
        return "skip"; // We handle children ourselves
      },
      enterSdt(sdt: StructuredDocumentTag) {
        // Inline SDTs may store Run-only content directly. The enclosing
        // walker visits them, but the paragraph-oriented replaceInParagraph
        // handler above never fires because there's no paragraph wrapper.
        // Pull the inline runs out and treat them as a synthetic paragraph
        // so replacement semantics match: a match spanning multiple runs
        // inside one inline SDT still gets stitched correctly.
        const inlineRuns: Run[] = [];
        for (const c of sdt.content) {
          if (
            c &&
            typeof c === "object" &&
            !("type" in c) &&
            "content" in c &&
            Array.isArray((c as { content?: unknown }).content)
          ) {
            inlineRuns.push(c as Run);
          }
        }
        if (inlineRuns.length === 0) {
          return; // Paragraph/Table children are handled by the walker
        }
        const synthetic: Paragraph = {
          type: "paragraph",
          children: inlineRuns as ParagraphChild[]
        };
        totalCount += replaceInParagraph(synthetic, search, replacement);
        // Don't `return "skip"` — the walker still needs to descend into
        // any Paragraph / Table children inside the same SDT.
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

  return totalCount;
}

// =============================================================================
// Internal helpers
// =============================================================================

interface TextSegment {
  /** The text-content node we will mutate. */
  readonly node: { text: string };
  /** Absolute start offset within the paragraph plain text. */
  readonly start: number;
  /** Absolute end offset (exclusive). */
  readonly end: number;
}

interface MatchSpan {
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

function replaceInParagraph(para: Paragraph, search: string | RegExp, replacement: string): number {
  // 1) Build segment table from text-bearing run-content nodes (in order).
  //
  // The visit walks anything that contributes visible text:
  //   - Run (the obvious case)
  //   - Hyperlink (children are runs too)
  //   - InsertedRun / MovedToRun: these wrap a Run that DOES surface in the
  //     rendered document. searchText / extractParagraphText already treat
  //     them as visible, so replaceText must descend into them — otherwise
  //     a search hit in a tracked-insert is not replaced and counts skew.
  //   - DeletedRun / MovedFromRun: do NOT contribute visible text by
  //     convention (they are pending deletions / moves), so they're skipped.
  const segments: TextSegment[] = [];
  let fullText = "";
  const visit = (children: readonly ParagraphChild[]): void => {
    for (const child of children) {
      // Hyperlink: recurse into its children.
      if (isHyperlink(child)) {
        visit(child.children as readonly ParagraphChild[]);
        continue;
      }
      // Track-change wrappers around a single run: descend through the
      // wrapper into the inner run so visible text is still mutated.
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
      for (const c of child.content) {
        if (c.type !== "text") {
          continue;
        }
        const node = c as { text: string };
        const start = fullText.length;
        fullText += node.text;
        segments.push({ node, start, end: fullText.length });
      }
    }
  };
  visit(para.children);

  if (segments.length === 0 || fullText.length === 0) {
    return 0;
  }

  // 2) Resolve all matches (with replacement strings).
  const matches = collectMatches(fullText, search, replacement);
  if (matches.length === 0) {
    return 0;
  }

  // 3) For every segment, rebuild its text by stitching together:
  //    - literal slice before/between/after any matches that overlap it,
  //    - the replacement (only on the first segment that touches a match).
  const claimed = new Set<number>(); // index into matches
  for (const seg of segments) {
    const newText = rewriteSegment(seg, matches, claimed);
    seg.node.text = newText;
  }

  return matches.length;
}

/** Collect non-overlapping matches with their (already-substituted) replacement strings. */
function collectMatches(
  fullText: string,
  search: string | RegExp,
  replacement: string
): MatchSpan[] {
  const matches: MatchSpan[] = [];

  if (typeof search === "string") {
    if (search === "") {
      return matches;
    }
    let idx = 0;
    while (idx <= fullText.length) {
      const found = fullText.indexOf(search, idx);
      if (found === -1) {
        break;
      }
      matches.push({ start: found, end: found + search.length, replacement });
      idx = found + search.length;
    }
    return matches;
  }

  // RegExp: always treat as global, regardless of caller's flags.
  const flags = search.flags.includes("g") ? search.flags : search.flags + "g";
  const re = new RegExp(search.source, flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(fullText)) !== null) {
    if (m[0].length === 0) {
      // Avoid infinite loops on zero-width matches.
      re.lastIndex++;
      continue;
    }
    const expanded = expandReplacement(replacement, m, fullText);
    matches.push({ start: m.index, end: m.index + m[0].length, replacement: expanded });
  }
  return matches;
}

/**
 * Compute the new text content of a single segment, given the global match table.
 * Marks each match in `claimed` the first time a segment "owns" it so the
 * replacement is emitted exactly once.
 */
function rewriteSegment(
  seg: TextSegment,
  matches: readonly MatchSpan[],
  claimed: Set<number>
): string {
  const segText = seg.node.text;
  let out = "";
  let cursor = seg.start;

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    if (m.end <= seg.start) {
      continue; // match is entirely before this segment
    }
    if (m.start >= seg.end) {
      break; // matches are sorted; remaining are after this segment
    }

    // Emit the literal slice before this match (if it falls in this segment).
    const literalEnd = Math.max(seg.start, Math.min(seg.end, m.start));
    if (literalEnd > cursor) {
      out += segText.slice(cursor - seg.start, literalEnd - seg.start);
    }

    // Emit the replacement only on the first segment that touches this match.
    if (!claimed.has(i)) {
      claimed.add(i);
      out += m.replacement;
    }

    // Advance the cursor past the part of the match that lies in this segment.
    cursor = Math.min(seg.end, m.end);
  }

  // Tail literal slice.
  if (cursor < seg.end) {
    out += segText.slice(cursor - seg.start);
  }

  return out;
}

/**
 * Mimic `String.prototype.replace` substitution semantics for a single match:
 * supports $$, $&, $`, $', $1..$9 and $<name> back-references.
 */
function expandReplacement(template: string, match: RegExpExecArray, fullText: string): string {
  let result = "";
  for (let i = 0; i < template.length; i++) {
    const ch = template[i];
    if (ch !== "$" || i === template.length - 1) {
      result += ch;
      continue;
    }
    const next = template[i + 1];
    if (next === "$") {
      result += "$";
      i++;
    } else if (next === "&") {
      result += match[0];
      i++;
    } else if (next === "`") {
      result += fullText.slice(0, match.index);
      i++;
    } else if (next === "'") {
      result += fullText.slice(match.index + match[0].length);
      i++;
    } else if (next >= "0" && next <= "9") {
      // $N or $NN — prefer two digits when valid.
      let groupIdx = next.charCodeAt(0) - 48;
      let consumed = 1;
      const after = template[i + 2];
      if (after && after >= "0" && after <= "9") {
        const twoDigit = groupIdx * 10 + (after.charCodeAt(0) - 48);
        if (twoDigit < match.length) {
          groupIdx = twoDigit;
          consumed = 2;
        }
      }
      if (groupIdx > 0 && groupIdx < match.length) {
        result += match[groupIdx] ?? "";
        i += consumed;
      } else if (groupIdx === 0) {
        // $0 is not a valid back-reference in JS replace semantics; keep literal.
        result += "$" + next;
        i++;
      } else {
        result += "$" + next;
        i++;
      }
    } else if (next === "<") {
      const close = template.indexOf(">", i + 2);
      const groups = match.groups;
      if (close !== -1 && groups) {
        const name = template.slice(i + 2, close);
        result += groups[name] ?? "";
        i = close;
      } else {
        result += ch;
      }
    } else {
      result += ch;
    }
  }
  return result;
}
