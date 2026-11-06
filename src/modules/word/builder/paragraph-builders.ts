/**
 * Paragraph-level builder functions for DOCX documents.
 *
 * Includes paragraph, textParagraph, heading, hyperlink, bookmarks,
 * comment markers, and track-change run wrappers.
 */

import { checkedTextRun } from "@word/builder/run-builders";
import {
  checkDateTime,
  checkHeadingLevel,
  checkId,
  checkNodeArray,
  checkParagraphProperties,
  checkRecord,
  checkRevision,
  checkRunProperties,
  checkString
} from "@word/core/argument-checks";
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
} from "@word/types";

// =============================================================================
// Paragraph Builders
// =============================================================================

/** Create a paragraph. */
export function paragraph(
  children: readonly ParagraphChild[],
  properties?: ParagraphProperties
): Paragraph {
  checkNodeArray(children, "children", "Build.paragraph");
  checkParagraphProperties(properties, "properties", "Build.paragraph");
  return { type: "paragraph", properties, children };
}

/**
 * A text paragraph, its arguments checked in the name of `where` — the public
 * function that was called.
 */
export function checkedTextParagraph(
  where: string,
  content: string,
  properties: (ParagraphProperties & { run?: RunProperties }) | undefined
): Paragraph {
  checkString(content, "content", where);
  checkParagraphProperties(properties, "properties", where);
  const { run: runProps, ...pProps } = properties ?? {};
  return {
    type: "paragraph",
    properties: Object.keys(pProps).length > 0 ? pProps : undefined,
    children: [checkedTextRun(where, content, runProps)]
  };
}

/** Create a simple text paragraph. */
export function textParagraph(
  content: string,
  properties?: ParagraphProperties & { run?: RunProperties }
): Paragraph {
  return checkedTextParagraph("Build.textParagraph", content, properties);
}

/**
 * A heading paragraph, its arguments checked in the name of `where`. A level
 * outside 1–9 names a style no document defines (`Heading12`), so it is
 * refused rather than written.
 */
export function checkedHeading(
  where: string,
  content: string | readonly ParagraphChild[],
  level: unknown
): Paragraph {
  checkHeadingLevel(level, "level", where);
  let children: readonly ParagraphChild[];
  if (typeof content === "string") {
    children = [checkedTextRun(where, content, undefined)];
  } else {
    checkNodeArray(content, "content", where);
    children = content;
  }
  return { type: "paragraph", properties: { style: `Heading${level as number}` }, children };
}

/**
 * Create a heading paragraph. Accepts plain text or an array of runs for mixed formatting.
 * `level` defaults to 1, as in `Document.addHeading`.
 */
export function heading(
  content: string | readonly ParagraphChild[],
  level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 = 1
): Paragraph {
  return checkedHeading("Build.heading", content, level);
}

/** Create a hyperlink. */
export function hyperlink(
  linkText: string,
  options: {
    rId?: string;
    url?: string;
    anchor?: string;
    tooltip?: string;
    /** Document location bookmark / fragment id (`w:docLocation`). */
    docLocation?: string;
    /** Target frame name for web views (`w:tgtFrame`). */
    tgtFrame?: string;
    /** Whether the hyperlink has been visited (`w:history`). */
    history?: boolean;
    properties?: RunProperties;
  }
): ParagraphChild {
  // Reference the built-in character style so the colour is governed by the
  // style table (and follows the theme) — Hyperlink for unvisited links,
  // FollowedHyperlink for visited ones (`history: true`). We also emit the
  // matching colour + underline as direct formatting so the link still renders
  // correctly when the document has no style table (Word does the same).
  const where = "Build.hyperlink";
  checkString(linkText, "linkText", where);
  checkRecord(options, "options", where);
  for (const key of ["rId", "url", "anchor", "tooltip", "docLocation", "tgtFrame"] as const) {
    checkString(options[key], `options.${key}`, where, { optional: true });
  }
  checkRunProperties(options.properties, "options.properties", where);
  const visited = options.history === true;
  const defaultProps: RunProperties = visited
    ? { style: "FollowedHyperlink", color: "954F72", underline: "single" }
    : { style: "Hyperlink", color: "0563C1", underline: "single" };
  return {
    type: "hyperlink",
    rId: options.rId,
    url: options.url,
    anchor: options.anchor,
    tooltip: options.tooltip,
    docLocation: options.docLocation,
    tgtFrame: options.tgtFrame,
    history: options.history,
    children: [checkedTextRun(where, linkText, options.properties ?? defaultProps)]
  };
}

/** Create a bookmark start. */
export function bookmarkStart(id: number, name: string): ParagraphChild {
  checkId(id, "id", "Build.bookmarkStart");
  checkString(name, "name", "Build.bookmarkStart");
  return { type: "bookmarkStart", id, name };
}

/**
 * Create a bookmark end. Pass the bookmark's `name` when ranges that share an
 * id overlap, so the writer can tell which start this end closes.
 */
export function bookmarkEnd(id: number, options?: { name?: string }): ParagraphChild {
  checkId(id, "id", "Build.bookmarkEnd");
  checkString(options?.name, "options.name", "Build.bookmarkEnd", { optional: true });
  return options?.name === undefined
    ? { type: "bookmarkEnd", id }
    : { type: "bookmarkEnd", id, name: options.name };
}

/** Create a comment range start marker. */
export function commentRangeStart(id: number): CommentRangeStart {
  checkId(id, "id", "Build.commentRangeStart");
  return { type: "commentRangeStart", id };
}

/** Create a comment range end marker. */
export function commentRangeEnd(id: number): CommentRangeEnd {
  checkId(id, "id", "Build.commentRangeEnd");
  return { type: "commentRangeEnd", id };
}

/** Create a comment reference (inside paragraph children). */
export function commentReference(id: number): CommentReference {
  checkId(id, "id", "Build.commentReference");
  return { type: "commentReference", id };
}

/** Create an inserted run (track changes). */
export function insertedRun(run: Run, revision: RevisionInfo): InsertedRun {
  checkRecord(run, "run", "Build.insertedRun");
  checkRevision(revision, "revision", "Build.insertedRun");
  return { type: "insertedRun", revision, run };
}

/** Create a deleted run (track changes). */
export function deletedRun(run: Run, revision: RevisionInfo): DeletedRun {
  checkRecord(run, "run", "Build.deletedRun");
  checkRevision(revision, "revision", "Build.deletedRun");
  return { type: "deletedRun", revision, run };
}

/** Create a moved-from run (track changes — source of a move). */
export function movedFromRun(run: Run, revision: RevisionInfo): MovedFromRun {
  checkRecord(run, "run", "Build.movedFromRun");
  checkRevision(revision, "revision", "Build.movedFromRun");
  return { type: "movedFromRun", revision, run };
}

/** Create a moved-to run (track changes — destination of a move). */
export function movedToRun(run: Run, revision: RevisionInfo): MovedToRun {
  checkRecord(run, "run", "Build.movedToRun");
  checkRevision(revision, "revision", "Build.movedToRun");
  return { type: "movedToRun", revision, run };
}

/** Create a move range start marker. */
export function moveFromRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  checkId(id, "id", "Build.moveFromRangeStart");
  checkString(author, "author", "Build.moveFromRangeStart");
  checkDateTime(options?.date, "options.date", "Build.moveFromRangeStart", { optional: true });
  checkString(options?.name, "options.name", "Build.moveFromRangeStart", { optional: true });
  return { type: "moveFromRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move range end marker. */
export function moveFromRangeEnd(id: number): MoveRangeMarker {
  checkId(id, "id", "Build.moveFromRangeEnd");
  return { type: "moveFromRangeEnd", id };
}

/** Create a move-to range start marker. */
export function moveToRangeStart(
  id: number,
  author: string,
  options?: { date?: string; name?: string }
): MoveRangeMarker {
  checkId(id, "id", "Build.moveToRangeStart");
  checkString(author, "author", "Build.moveToRangeStart");
  checkDateTime(options?.date, "options.date", "Build.moveToRangeStart", { optional: true });
  checkString(options?.name, "options.name", "Build.moveToRangeStart", { optional: true });
  return { type: "moveToRangeStart", id, author, date: options?.date, name: options?.name };
}

/** Create a move-to range end marker. */
export function moveToRangeEnd(id: number): MoveRangeMarker {
  checkId(id, "id", "Build.moveToRangeEnd");
  return { type: "moveToRangeEnd", id };
}
