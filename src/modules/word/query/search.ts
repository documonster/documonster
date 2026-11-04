/**
 * Document Search & Query API (read-only helpers)
 *
 * Functions for searching, counting, and extracting text from a DocxDocument.
 */

import {
  checkHandle,
  checkInteger,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import { extractParagraphText } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import type { DocxVisitor, WalkOptions } from "@word/core/walker";
import { resolveHeadingLevel } from "@word/query/heading";
import { indexStyles } from "@word/query/style-resolve";
import type {
  DocxDocument,
  Paragraph,
  Table,
  Hyperlink,
  BookmarkStart,
  RangeMarker,
  CommentDef,
  ImageDef,
  SectionProperties
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** Result of a text search in a document. */
export interface SearchResult {
  /**
   * The paragraph's visit order across the entire document, counting
   * paragraphs reachable from the body (including those nested in tables,
   * SDTs, text boxes, headers, footers, footnotes, endnotes and TOC
   * caches) in walk order.
   *
   * This is **not** an index into `doc.body`: nested paragraphs are
   * counted too. Use it as a stable ordinal for ordering results, not
   * for direct array access.
   */
  readonly paragraphIndex: number;
  /** The matched text. */
  readonly match: string;
  /** Character offset within the paragraph's concatenated text. */
  readonly offset: number;
}

/** A heading extracted from a document. */
export interface DocumentHeading {
  /** Heading level (1-9). */
  readonly level: number;
  /** Plain text of the heading. */
  readonly text: string;
  /** Index into doc.body where the paragraph resides. */
  readonly paragraphIndex: number;
  /** Style ID used (e.g. "Heading1"). */
  readonly style?: string;
}

/** A section definition found in the document. */
export interface DocumentSection {
  /** The section properties. */
  readonly properties: SectionProperties;
  /** Index of the paragraph containing this section break (or -1 for the final section). */
  readonly paragraphIndex: number;
  /** Whether this is the final section (from doc.sectionProperties). */
  readonly isFinal: boolean;
}

// =============================================================================
// Helpers
// =============================================================================

/** Extract concatenated plain text from a paragraph's runs. */
export function paragraphText(para: Paragraph): string {
  return extractParagraphText(para);
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Count all top-level paragraphs in the document body.
 */
export function paragraphCount(doc: DocxDocument): number {
  checkHandle(doc, "document", "doc", "Query.getParagraphCount");
  let count = 0;
  for (const block of doc.body) {
    if (block.type === "paragraph") {
      count++;
    }
  }
  return count;
}

/**
 * Count words across all paragraphs in the document body.
 * Uses simple whitespace splitting; for East Asian text, each CJK character
 * is counted as one "word" to approximate meaningful unit count.
 */
export function countWords(doc: DocxDocument): number {
  checkHandle(doc, "document", "doc", "Query.getWordCount");
  let count = 0;
  const cjkRe = /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g;
  const visitor: DocxVisitor = {
    enterParagraph(para: Paragraph) {
      const text = paragraphText(para);
      const cjkCount = (text.match(cjkRe) ?? []).length;
      const latin = text.replace(cjkRe, " ").trim();
      const latinCount = latin ? latin.split(/\s+/).length : 0;
      count += cjkCount + latinCount;
      return "skip";
    }
  };
  walkDocument(doc, visitor, {
    includeHeaders: false,
    includeFooters: false,
    includeFootnotes: false,
    includeEndnotes: false,
    includeComments: false
  });
  return count;
}

/**
 * Extract the heading outline from a document.
 *
 * Uses the same rule as the converters, layout and the TOC field
 * ({@link resolveHeadingLevel}): the paragraph's effective outline level —
 * direct `w:outlineLvl`, else inherited through the style's `basedOn` chain,
 * with level 9 meaning body text — falling back to the built-in
 * `Heading1`…`Heading9` styles only when nothing in the chain sets one.
 *
 * The built-in `Title` style is not part of the outline and is excluded, as
 * Word's own TOC (`\o`) excludes it.
 */
export function getHeadings(doc: DocxDocument): DocumentHeading[] {
  checkHandle(doc, "document", "doc", "Query.listHeadings");
  const out: DocumentHeading[] = [];
  const styles = indexStyles(doc);
  doc.body.forEach((block, i) => {
    if (block.type !== "paragraph") {
      return;
    }
    const heading = resolveHeadingLevel(doc, block, styles);
    if (heading?.kind === "heading") {
      out.push({
        level: heading.level,
        text: paragraphText(block),
        paragraphIndex: i,
        style: block.properties?.style
      });
    }
  });
  return out;
}

/**
 * Where {@link findBookmark} found a bookmark. For a bookmark inside a
 * paragraph, `paragraphIndex` is the body index of that paragraph and
 * `childIndex` the index of the start among its children.
 *
 * Word also starts a bookmark that spans whole blocks (a table, say) between
 * them. Such a bookmark is reported with `between` set: `"before"` when it is
 * in `doc.body[paragraphIndex].rangeMarkersBefore` — the range starts at the
 * beginning of that block, which may be a table rather than a paragraph — and
 * `"after"` when it is in `doc.bodyRangeMarkersAfter`, after the last block
 * (`paragraphIndex` is then that block's index). `childIndex` is `0`, the
 * start of the block, and `markerIndex` the index within that marker list.
 */
export interface BookmarkLocation {
  readonly bookmark: BookmarkStart;
  readonly paragraphIndex: number;
  readonly childIndex: number;
  /** Set when the bookmark starts between body blocks rather than inside a paragraph. */
  readonly between?: "before" | "after";
  /** With `between`: the index of the start within its marker list. */
  readonly markerIndex?: number;
}

/**
 * Find a bookmark by name, in the body's paragraphs or between its blocks
 * (see {@link BookmarkLocation}).
 *
 * @returns The bookmark start marker + its location, or `undefined` if not found.
 */
export function findBookmark(doc: DocxDocument, name: string): BookmarkLocation | undefined {
  checkHandle(doc, "document", "doc", "Query.getBookmark");
  checkString(name, "name", "Query.getBookmark");
  const inMarkers = (
    markers: readonly RangeMarker[] | undefined,
    paragraphIndex: number,
    between: "before" | "after"
  ): BookmarkLocation | undefined => {
    const markerIndex = (markers ?? []).findIndex(
      m => m.type === "bookmarkStart" && m.name === name
    );
    return markerIndex < 0
      ? undefined
      : {
          bookmark: markers![markerIndex] as BookmarkStart,
          paragraphIndex,
          childIndex: 0,
          between,
          markerIndex
        };
  };
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    const before = inMarkers(block.rangeMarkersBefore, i, "before");
    if (before) {
      return before;
    }
    if (block.type !== "paragraph") {
      continue;
    }
    for (let j = 0; j < block.children.length; j++) {
      const ch = block.children[j];
      if ("type" in ch && ch.type === "bookmarkStart" && ch.name === name) {
        return { bookmark: ch, paragraphIndex: i, childIndex: j };
      }
    }
  }
  return inMarkers(doc.bodyRangeMarkersAfter, Math.max(0, doc.body.length - 1), "after");
}

/**
 * Find a comment by its ID.
 */
export function findComment(doc: DocxDocument, id: number): CommentDef | undefined {
  checkHandle(doc, "document", "doc", "Query.getComment");
  checkInteger(id, "id", "Query.getComment");
  return doc.comments?.find(c => c.id === id);
}

/**
 * List all images registered in the document.
 */
export function listImages(doc: DocxDocument): readonly ImageDef[] {
  checkHandle(doc, "document", "doc", "Query.listImages");
  return doc.images ?? [];
}

/**
 * Options for {@link listTables}. Every `include*` flag defaults to `false`, so
 * by default only the document body is searched.
 */
export interface ListTablesOptions extends WalkOptions {
  /** Only direct children of the body — no nested tables, no other parts. */
  readonly topLevelOnly?: boolean;
}

/**
 * List all tables in the document.
 *
 * By default this returns **all** tables in the document body (including
 * tables nested inside other tables, SDTs, text boxes, and TOC cached
 * paragraphs). For top-level only behavior pass `{ topLevelOnly: true }`.
 *
 * @example
 * listTables(doc)                          // all body tables (consistent with countTables)
 * listTables(doc, { includeHeaders: true, includeFooters: true }) // also headers/footers
 * listTables(doc, { topLevelOnly: true })  // direct children of body only
 */
export function listTables(doc: DocxDocument, options?: ListTablesOptions): readonly Table[] {
  checkHandle(doc, "document", "doc", "Query.listTables");
  checkRecord(options, "options", "Query.listTables", { optional: true });
  if (options?.topLevelOnly) {
    return doc.body.filter((b): b is Table => b.type === "table");
  }
  const out: Table[] = [];
  walkDocument(
    doc,
    {
      enterTable(t) {
        out.push(t);
        return "continue";
      }
    },
    {
      includeHeaders: options?.includeHeaders ?? false,
      includeFooters: options?.includeFooters ?? false,
      includeFootnotes: options?.includeFootnotes ?? false,
      includeEndnotes: options?.includeEndnotes ?? false,
      includeComments: options?.includeComments ?? false
    }
  );
  return out;
}

/**
 * Collect all hyperlinks in the document body.
 */
export function listHyperlinks(doc: DocxDocument): readonly Hyperlink[] {
  checkHandle(doc, "document", "doc", "Query.listHyperlinks");
  const out: Hyperlink[] = [];
  walkDocument(
    doc,
    {
      enterHyperlink(hl) {
        out.push(hl);
        return "continue";
      }
    },
    {
      includeHeaders: false,
      includeFooters: false,
      includeFootnotes: false,
      includeEndnotes: false
    }
  );
  return out;
}

/**
 * Get the total number of tables (top-level) and nested tables.
 */
export function tableCount(doc: DocxDocument): number {
  checkHandle(doc, "document", "doc", "Query.getTableCount");
  let count = 0;
  const visitor: DocxVisitor = {
    enterTable() {
      count++;
      return "continue";
    }
  };
  walkDocument(doc, visitor, {
    includeHeaders: false,
    includeFooters: false,
    includeFootnotes: false,
    includeEndnotes: false,
    includeComments: false
  });
  return count;
}

/**
 * List all sections in a document.
 *
 * Sections are defined by section breaks within paragraph properties
 * and the final section at the document level.
 *
 * @param doc - The document to inspect.
 * @returns Array of section definitions in document order.
 */
export function listSections(doc: DocxDocument): DocumentSection[] {
  checkHandle(doc, "document", "doc", "Query.listSections");
  const sections: DocumentSection[] = [];
  for (let i = 0; i < doc.body.length; i++) {
    const block = doc.body[i];
    if (block.type === "paragraph" && block.properties?.sectionProperties) {
      sections.push({
        properties: block.properties.sectionProperties,
        paragraphIndex: i,
        isFinal: false
      });
    }
  }
  // Add final section
  if (doc.sectionProperties) {
    sections.push({
      properties: doc.sectionProperties,
      paragraphIndex: -1,
      isFinal: true
    });
  }
  return sections;
}

/**
 * Extract plain text from the entire document: body (with tables, content
 * controls and text boxes), headers, footers, footnotes, and endnotes.
 * Comments are not part of the document's text and are left out — read them
 * from `doc.comments`, or with `Query.listParagraphs(doc, { includeComments: true })`.
 *
 * Paragraphs are separated by `\n`. Tables render as tab-separated cell text.
 */
export function extractText(doc: DocxDocument): string {
  checkHandle(doc, "document", "doc", "Query.getText");
  const lines: string[] = [];

  // State stack for nested table handling.
  // Each entry represents the context of a table being processed.
  interface TableFrame {
    rowCellTexts: string[];
    cellLines: string[];
    tableLines: string[];
  }
  const stack: TableFrame[] = [];

  const visitor: DocxVisitor = {
    enterTable() {
      stack.push({ rowCellTexts: [], cellLines: [], tableLines: [] });
      return "continue";
    },
    leaveTable() {
      const frame = stack.pop()!;
      // Contribute this table's collected lines to parent context
      if (stack.length > 0) {
        // Inside an outer cell — add table lines to the outer cell
        const parent = stack[stack.length - 1];
        for (const line of frame.tableLines) {
          parent.cellLines.push(line);
        }
      } else {
        // Top-level: add directly to output lines
        for (const line of frame.tableLines) {
          lines.push(line);
        }
      }
    },
    enterTableRow() {
      if (stack.length > 0) {
        stack[stack.length - 1].rowCellTexts = [];
      }
      return "continue";
    },
    leaveTableRow() {
      if (stack.length > 0) {
        const frame = stack[stack.length - 1];
        frame.tableLines.push(frame.rowCellTexts.join("\t"));
      }
    },
    enterTableCell() {
      if (stack.length > 0) {
        stack[stack.length - 1].cellLines = [];
      }
      return "continue";
    },
    leaveTableCell() {
      if (stack.length > 0) {
        const frame = stack[stack.length - 1];
        frame.rowCellTexts.push(frame.cellLines.join(" "));
      }
    },
    enterParagraph(para: Paragraph) {
      const text = paragraphText(para);
      if (stack.length > 0) {
        stack[stack.length - 1].cellLines.push(text);
      } else {
        lines.push(text);
      }
      return "skip";
    }
  };

  walkDocument(doc, visitor, {
    includeHeaders: true,
    includeFooters: true,
    includeFootnotes: true,
    includeEndnotes: true,
    includeComments: false
  });

  return lines.join("\n");
}

/**
 * Search for text occurrences across the entire document: body, tables, SDTs,
 * text boxes, headers, footers, footnotes, and endnotes — the text
 * {@link extractText} returns. Comments are not searched, unlike
 * `Document.replaceText`, which also replaces inside comments: a match a
 * search did not report may still be replaced there.
 *
 * @param doc - The document model to search.
 * @param query - String or RegExp to search for.
 * @returns Array of search results.
 */
export function searchText(doc: DocxDocument, query: string | RegExp): SearchResult[] {
  checkHandle(doc, "document", "doc", "Query.searchText");
  if (typeof query !== "string" && !(query instanceof RegExp)) {
    invalidArgument("Query.searchText", "query", "a string or a RegExp", query);
  }
  const results: SearchResult[] = [];
  let idx = 0;

  const visitor: DocxVisitor = {
    enterParagraph(para: Paragraph) {
      const text = paragraphText(para);
      if (typeof query === "string") {
        // Empty query would loop forever; treat it as "no match".
        if (query.length === 0) {
          idx++;
          return "skip";
        }
        let pos = text.indexOf(query);
        while (pos !== -1) {
          results.push({ paragraphIndex: idx, match: query, offset: pos });
          pos = text.indexOf(query, pos + query.length);
        }
      } else {
        // Strip the sticky flag — it would constrain matching to lastIndex
        // and silently miss occurrences elsewhere in the paragraph text.
        const baseFlags = query.flags.replace(/y/g, "");
        const re = new RegExp(query.source, baseFlags.includes("g") ? baseFlags : baseFlags + "g");
        let m: RegExpExecArray | null;
        while ((m = re.exec(text)) !== null) {
          results.push({ paragraphIndex: idx, match: m[0], offset: m.index });
          // Advance past zero-width matches to prevent infinite loops.
          if (m[0].length === 0) {
            re.lastIndex++;
          }
        }
      }
      idx++;
      return "skip";
    }
  };

  walkDocument(doc, visitor, {
    includeHeaders: true,
    includeFooters: true,
    includeFootnotes: true,
    includeEndnotes: true,
    includeComments: false
  });

  return results;
}
