/**
 * Paragraph-level builder functions for DOCX documents.
 *
 * Includes paragraph, textParagraph, heading, hyperlink, bookmarks,
 * comment markers, and track-change run wrappers.
 */

import type {
  Paragraph,
  ParagraphProperties,
  ParagraphChild,
  Run,
  RunProperties,
  CommentRangeStart,
  CommentRangeEnd,
  CommentReference,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  MoveRangeMarker,
  RevisionInfo
} from "../types";
import { text } from "./run-builders";

// =============================================================================
// Paragraph Builders
// =============================================================================

/** Create a paragraph. */
export function paragraph(children: ParagraphChild[], properties?: ParagraphProperties): Paragraph {
  return { type: "paragraph", properties, children };
}

/** Create a simple text paragraph. */
export function textParagraph(
  content: string,
  properties?: ParagraphProperties & { run?: RunProperties }
): Paragraph {
  const { run: runProps, ...pProps } = properties ?? {};
  return paragraph([text(content, runProps)], Object.keys(pProps).length > 0 ? pProps : undefined);
}

/** Create a heading paragraph. Accepts plain text or an array of runs for mixed formatting. */
export function heading(
  content: string | ParagraphChild[],
  level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9
): Paragraph {
  const children: ParagraphChild[] = typeof content === "string" ? [text(content)] : content;
  return paragraph(children, { style: `Heading${level}` });
}

/** Create a hyperlink. */
export function hyperlink(
  linkText: string,
  options: {
    rId?: string;
    url?: string;
    anchor?: string;
    tooltip?: string;
    properties?: RunProperties;
  }
): ParagraphChild {
  return {
    type: "hyperlink",
    rId: options.rId,
    url: options.url,
    anchor: options.anchor,
    tooltip: options.tooltip,
    children: [text(linkText, options.properties ?? { color: "0563C1", underline: "single" })]
  };
}

/** Create a bookmark start. */
export function bookmarkStart(id: number, name: string): ParagraphChild {
  return { type: "bookmarkStart", id, name };
}

/** Create a bookmark end. */
export function bookmarkEnd(id: number): ParagraphChild {
  return { type: "bookmarkEnd", id };
}

/** Create a comment range start marker. */
export function commentRangeStart(id: number): CommentRangeStart {
  return { type: "commentRangeStart", id };
}

/** Create a comment range end marker. */
export function commentRangeEnd(id: number): CommentRangeEnd {
  return { type: "commentRangeEnd", id };
}

/** Create a comment reference (inside paragraph children). */
export function commentReference(id: number): CommentReference {
  return { type: "commentReference", id };
}

/** Create an inserted run (track changes). */
export function insertedRun(run: Run, revision: RevisionInfo): InsertedRun {
  return { type: "insertedRun", revision, run };
}

/** Create a deleted run (track changes). */
export function deletedRun(run: Run, revision: RevisionInfo): DeletedRun {
  return { type: "deletedRun", revision, run };
}

/** Create a moved-from run (track changes — source of a move). */
export function movedFromRun(run: Run, revision: RevisionInfo): MovedFromRun {
  return { type: "movedFromRun", revision, run };
}

/** Create a moved-to run (track changes — destination of a move). */
export function movedToRun(run: Run, revision: RevisionInfo): MovedToRun {
  return { type: "movedToRun", revision, run };
}

/** Create a move range start marker. */
export function moveFromRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  return { type: "moveFromRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move range end marker. */
export function moveFromRangeEnd(id: number): MoveRangeMarker {
  return { type: "moveFromRangeEnd", id };
}

/** Create a move-to range start marker. */
export function moveToRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  return { type: "moveToRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move-to range end marker. */
export function moveToRangeEnd(id: number): MoveRangeMarker {
  return { type: "moveToRangeEnd", id };
}
