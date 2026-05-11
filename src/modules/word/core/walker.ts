/**
 * DOCX Module - Document Walker (Visitor Pattern)
 *
 * Unified traversal engine for the DocxDocument model. Provides a single,
 * tested implementation of document tree walking that all consumers
 * (renderers, validators, field engines, template engines, etc.) can use
 * instead of each implementing their own recursive traversal.
 *
 * Design principles:
 * - Visitor receives enter/leave callbacks for structural elements
 * - Return "skip" to prune subtrees, "stop" to abort early
 * - WalkPath provides full context (location, nesting depth, container)
 * - Options control which parts of the document are visited
 */

import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  Table,
  TableRow,
  TableCell,
  MathBlock,
  MathContent,
  StructuredDocumentTag,
  Hyperlink,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  FloatingImage,
  BookmarkStart,
  CommentRangeStart,
  CommentRangeEnd,
  TextBox,
  TableOfContents
} from "../types";
import { isRun } from "./text-utils";

// =============================================================================
// Types
// =============================================================================

/** Action returned by visitor callbacks to control traversal. */
export type VisitAction = "continue" | "skip" | "stop";

/** Location context passed to visitor callbacks. */
export interface WalkPath {
  /** Current section index (0-based). */
  readonly section: number;
  /** Nesting depth (0 = top-level body content). */
  readonly depth: number;
  /** Whether currently inside a header. */
  readonly inHeader: boolean;
  /** Whether currently inside a footer. */
  readonly inFooter: boolean;
  /** Whether currently inside a footnote. */
  readonly inFootnote: boolean;
  /** Whether currently inside an endnote. */
  readonly inEndnote: boolean;
  /** Whether currently inside a comment. */
  readonly inComment: boolean;
  /** Index within the parent container. */
  readonly index: number;
}

/** Options controlling which parts of the document to walk. */
export interface WalkOptions {
  /** Walk header content. Default: true. */
  readonly includeHeaders?: boolean;
  /** Walk footer content. Default: true. */
  readonly includeFooters?: boolean;
  /** Walk footnote content. Default: true. */
  readonly includeFootnotes?: boolean;
  /** Walk endnote content. Default: true. */
  readonly includeEndnotes?: boolean;
  /** Walk comment content. Default: false. */
  readonly includeComments?: boolean;
}

/**
 * Visitor interface for document traversal.
 * All methods are optional — implement only what you need.
 */
export interface DocxVisitor {
  // — Block-level —
  enterParagraph?(para: Paragraph, path: WalkPath): VisitAction | void;
  leaveParagraph?(para: Paragraph, path: WalkPath): void;
  enterTable?(table: Table, path: WalkPath): VisitAction | void;
  leaveTable?(table: Table, path: WalkPath): void;
  enterTableRow?(row: TableRow, path: WalkPath): VisitAction | void;
  leaveTableRow?(row: TableRow, path: WalkPath): void;
  enterTableCell?(cell: TableCell, path: WalkPath): VisitAction | void;
  leaveTableCell?(cell: TableCell, path: WalkPath): void;
  enterSdt?(sdt: StructuredDocumentTag, path: WalkPath): VisitAction | void;
  leaveSdt?(sdt: StructuredDocumentTag, path: WalkPath): void;
  enterMathBlock?(math: MathBlock, path: WalkPath): VisitAction | void;
  leaveMathBlock?(math: MathBlock, path: WalkPath): void;
  visitFloatingImage?(image: FloatingImage, path: WalkPath): void;
  visitTextBox?(textBox: TextBox, path: WalkPath): VisitAction | void;
  visitTableOfContents?(toc: TableOfContents, path: WalkPath): void;

  // — Paragraph-level —
  enterRun?(run: Run, path: WalkPath): VisitAction | void;
  leaveRun?(run: Run, path: WalkPath): void;
  enterHyperlink?(hyperlink: Hyperlink, path: WalkPath): VisitAction | void;
  leaveHyperlink?(hyperlink: Hyperlink, path: WalkPath): void;
  visitInsertedRun?(inserted: InsertedRun, path: WalkPath): VisitAction | void;
  visitDeletedRun?(deleted: DeletedRun, path: WalkPath): void;
  visitMovedFromRun?(moved: MovedFromRun, path: WalkPath): void;
  visitMovedToRun?(moved: MovedToRun, path: WalkPath): VisitAction | void;
  visitBookmarkStart?(bookmark: BookmarkStart, path: WalkPath): void;
  visitBookmarkEnd?(id: number, path: WalkPath): void;
  visitCommentRangeStart?(comment: CommentRangeStart, path: WalkPath): void;
  visitCommentRangeEnd?(comment: CommentRangeEnd, path: WalkPath): void;

  // — Run-level —
  visitRunContent?(content: RunContent, run: Run, path: WalkPath): void;

  // — Math —
  visitMathContent?(content: MathContent, path: WalkPath): void;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Walk an entire DocxDocument, calling visitor methods for each element.
 *
 * @param doc - The document to walk.
 * @param visitor - Visitor with callbacks for elements of interest.
 * @param options - Controls which document parts to include.
 */
export function walkDocument(doc: DocxDocument, visitor: DocxVisitor, options?: WalkOptions): void {
  const opts: Required<WalkOptions> = {
    includeHeaders: options?.includeHeaders ?? true,
    includeFooters: options?.includeFooters ?? true,
    includeFootnotes: options?.includeFootnotes ?? true,
    includeEndnotes: options?.includeEndnotes ?? true,
    includeComments: options?.includeComments ?? false
  };

  const basePath: WalkPath = {
    section: 0,
    depth: 0,
    inHeader: false,
    inFooter: false,
    inFootnote: false,
    inEndnote: false,
    inComment: false,
    index: 0
  };

  // Walk body
  if (walkBlocks(doc.body, visitor, basePath) === "stop") {
    return;
  }

  // Walk headers
  if (opts.includeHeaders && doc.headers) {
    for (const [, headerDef] of doc.headers) {
      const headerPath: WalkPath = { ...basePath, inHeader: true };
      if (
        walkBlocks(headerDef.content.children as readonly BodyContent[], visitor, headerPath) ===
        "stop"
      ) {
        return;
      }
    }
  }

  // Walk footers
  if (opts.includeFooters && doc.footers) {
    for (const [, footerDef] of doc.footers) {
      const footerPath: WalkPath = { ...basePath, inFooter: true };
      if (
        walkBlocks(footerDef.content.children as readonly BodyContent[], visitor, footerPath) ===
        "stop"
      ) {
        return;
      }
    }
  }

  // Walk footnotes
  if (opts.includeFootnotes && doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      } // Skip separator footnotes
      const notePath: WalkPath = { ...basePath, inFootnote: true };
      if (walkBlocks(note.content as readonly BodyContent[], visitor, notePath) === "stop") {
        return;
      }
    }
  }

  // Walk endnotes
  if (opts.includeEndnotes && doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      const notePath: WalkPath = { ...basePath, inEndnote: true };
      if (walkBlocks(note.content as readonly BodyContent[], visitor, notePath) === "stop") {
        return;
      }
    }
  }

  // Walk comments
  if (opts.includeComments && doc.comments) {
    for (const comment of doc.comments) {
      const commentPath: WalkPath = { ...basePath, inComment: true };
      if (walkBlocks(comment.content as readonly BodyContent[], visitor, commentPath) === "stop") {
        return;
      }
    }
  }
}

/**
 * Walk a list of body content blocks (useful for walking sub-documents).
 *
 * @param blocks - Body content array to walk.
 * @param visitor - Visitor with callbacks.
 * @param basePath - Optional initial path context.
 * @returns "stop" if traversal was aborted, "continue" otherwise.
 */
export function walkBlocks(
  blocks: readonly BodyContent[],
  visitor: DocxVisitor,
  basePath?: WalkPath
): "stop" | "continue" {
  const path = basePath ?? {
    section: 0,
    depth: 0,
    inHeader: false,
    inFooter: false,
    inFootnote: false,
    inEndnote: false,
    inComment: false,
    index: 0
  };

  let section = path.section;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const currentPath: WalkPath = { ...path, section, index: i };
    const result = walkBodyContent(block, visitor, currentPath);
    if (result === "stop") {
      return "stop";
    }
    // A paragraph with sectionProperties marks the end of the current section
    if (block.type === "paragraph" && block.properties?.sectionProperties && path.depth === 0) {
      section++;
    }
  }
  return "continue";
}

// =============================================================================
// Internal Traversal
// =============================================================================

function walkBodyContent(
  block: BodyContent,
  visitor: DocxVisitor,
  path: WalkPath
): "stop" | "continue" {
  switch (block.type) {
    case "paragraph":
      return walkParagraph(block, visitor, path);
    case "table":
      return walkTable(block, visitor, path);
    case "sdt":
      return walkSdt(block as StructuredDocumentTag, visitor, path);
    case "math":
      return walkMathBlock(block as MathBlock, visitor, path);
    case "floatingImage":
      visitor.visitFloatingImage?.(block as FloatingImage, path);
      return "continue";
    case "textBox": {
      const tb = block as TextBox;
      const action = visitor.visitTextBox?.(tb, path);
      if (action === "stop") {
        return "stop";
      }
      if (action === "skip") {
        return "continue";
      }
      if (tb.content) {
        const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
        if (walkBlocks(tb.content as readonly BodyContent[], visitor, innerPath) === "stop") {
          return "stop";
        }
      }
      return "continue";
    }
    case "tableOfContents":
      visitor.visitTableOfContents?.(block as TableOfContents, path);
      return "continue";
    default:
      return "continue";
  }
}

function walkParagraph(para: Paragraph, visitor: DocxVisitor, path: WalkPath): "stop" | "continue" {
  const action = visitor.enterParagraph?.(para, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveParagraph?.(para, path);
    return "continue";
  }

  // Walk children
  for (let i = 0; i < para.children.length; i++) {
    const child = para.children[i];
    const childPath: WalkPath = { ...path, index: i, depth: path.depth + 1 };
    const result = walkParagraphChild(child, visitor, childPath);
    if (result === "stop") {
      return "stop";
    }
  }

  visitor.leaveParagraph?.(para, path);
  return "continue";
}

function walkParagraphChild(
  child: ParagraphChild,
  visitor: DocxVisitor,
  path: WalkPath
): "stop" | "continue" {
  // Type-discriminated dispatch
  if ("type" in child) {
    const typed = child as { type: string };
    switch (typed.type) {
      case "hyperlink":
        return walkHyperlink(child as Hyperlink, visitor, path);
      case "insertedRun": {
        const ins = child as InsertedRun;
        const action = visitor.visitInsertedRun?.(ins, path);
        if (action === "stop") {
          return "stop";
        }
        if (action === "skip") {
          return "continue";
        }
        // Walk the inner run
        return walkRun(ins.run, visitor, path);
      }
      case "deletedRun":
        visitor.visitDeletedRun?.(child as DeletedRun, path);
        return "continue";
      case "movedFromRun":
        visitor.visitMovedFromRun?.(child as MovedFromRun, path);
        return "continue";
      case "movedToRun": {
        const moved = child as MovedToRun;
        const action = visitor.visitMovedToRun?.(moved, path);
        if (action === "stop") {
          return "stop";
        }
        if (action === "skip") {
          return "continue";
        }
        return walkRun(moved.run, visitor, path);
      }
      case "bookmarkStart":
        visitor.visitBookmarkStart?.(child as BookmarkStart, path);
        return "continue";
      case "bookmarkEnd":
        visitor.visitBookmarkEnd?.((child as { id: number }).id, path);
        return "continue";
      case "commentRangeStart":
        visitor.visitCommentRangeStart?.(child as CommentRangeStart, path);
        return "continue";
      case "commentRangeEnd":
        visitor.visitCommentRangeEnd?.(child as CommentRangeEnd, path);
        return "continue";
      default:
        // Unknown typed child — might be a Run with a `type` property on content
        break;
    }
  }

  // Default: treat as Run if it has `content` array
  if (isRun(child)) {
    return walkRun(child, visitor, path);
  }

  return "continue";
}

function walkRun(run: Run, visitor: DocxVisitor, path: WalkPath): "stop" | "continue" {
  const action = visitor.enterRun?.(run, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveRun?.(run, path);
    return "continue";
  }

  // Walk run content
  if (visitor.visitRunContent) {
    for (const content of run.content) {
      visitor.visitRunContent(content, run, path);
    }
  }

  visitor.leaveRun?.(run, path);
  return "continue";
}

function walkHyperlink(hl: Hyperlink, visitor: DocxVisitor, path: WalkPath): "stop" | "continue" {
  const action = visitor.enterHyperlink?.(hl, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveHyperlink?.(hl, path);
    return "continue";
  }

  for (let i = 0; i < hl.children.length; i++) {
    const child = hl.children[i] as ParagraphChild;
    const childPath: WalkPath = { ...path, index: i, depth: path.depth + 1 };
    const result = walkParagraphChild(child, visitor, childPath);
    if (result === "stop") {
      return "stop";
    }
  }

  visitor.leaveHyperlink?.(hl, path);
  return "continue";
}

function walkTable(table: Table, visitor: DocxVisitor, path: WalkPath): "stop" | "continue" {
  const action = visitor.enterTable?.(table, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveTable?.(table, path);
    return "continue";
  }

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowPath: WalkPath = { ...path, index: ri, depth: path.depth + 1 };

    const rowAction = visitor.enterTableRow?.(row, rowPath);
    if (rowAction === "stop") {
      return "stop";
    }
    if (rowAction === "skip") {
      visitor.leaveTableRow?.(row, rowPath);
      continue;
    }

    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const cellPath: WalkPath = { ...rowPath, index: ci, depth: rowPath.depth + 1 };

      const cellAction = visitor.enterTableCell?.(cell, cellPath);
      if (cellAction === "stop") {
        return "stop";
      }
      if (cellAction === "skip") {
        visitor.leaveTableCell?.(cell, cellPath);
        continue;
      }

      // Walk cell content (paragraphs and nested tables)
      if (
        walkBlocks(cell.content as readonly BodyContent[], visitor, {
          ...cellPath,
          depth: cellPath.depth + 1
        }) === "stop"
      ) {
        return "stop";
      }

      visitor.leaveTableCell?.(cell, cellPath);
    }

    visitor.leaveTableRow?.(row, rowPath);
  }

  visitor.leaveTable?.(table, path);
  return "continue";
}

function walkSdt(
  sdt: StructuredDocumentTag,
  visitor: DocxVisitor,
  path: WalkPath
): "stop" | "continue" {
  const action = visitor.enterSdt?.(sdt, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveSdt?.(sdt, path);
    return "continue";
  }

  const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
  const filtered = sdt.content.filter(
    (c): c is Paragraph | Table => "type" in c && (c.type === "paragraph" || c.type === "table")
  );
  if (walkBlocks(filtered as readonly BodyContent[], visitor, innerPath) === "stop") {
    return "stop";
  }

  visitor.leaveSdt?.(sdt, path);
  return "continue";
}

function walkMathBlock(math: MathBlock, visitor: DocxVisitor, path: WalkPath): "stop" | "continue" {
  const action = visitor.enterMathBlock?.(math, path);
  if (action === "stop") {
    return "stop";
  }
  if (action === "skip") {
    visitor.leaveMathBlock?.(math, path);
    return "continue";
  }

  if (visitor.visitMathContent) {
    walkMathContentArray(math.content, visitor, path);
  }

  visitor.leaveMathBlock?.(math, path);
  return "continue";
}

function walkMathContentArray(
  contents: readonly MathContent[],
  visitor: DocxVisitor,
  path: WalkPath
): void {
  for (const content of contents) {
    visitor.visitMathContent!(content, path);

    // Recursively walk sub-content
    switch (content.type) {
      case "mathFraction":
        walkMathContentArray(content.numerator, visitor, path);
        walkMathContentArray(content.denominator, visitor, path);
        break;
      case "mathRadical":
        if (content.degree) {
          walkMathContentArray(content.degree, visitor, path);
        }
        walkMathContentArray(content.content, visitor, path);
        break;
      case "mathSuperScript":
        walkMathContentArray(content.base, visitor, path);
        walkMathContentArray(content.superScript, visitor, path);
        break;
      case "mathSubScript":
        walkMathContentArray(content.base, visitor, path);
        walkMathContentArray(content.subScript, visitor, path);
        break;
      case "mathSubSuperScript":
        walkMathContentArray(content.base, visitor, path);
        walkMathContentArray(content.subScript, visitor, path);
        walkMathContentArray(content.superScript, visitor, path);
        break;
      case "mathPreSubSuperScript":
        walkMathContentArray(content.base, visitor, path);
        walkMathContentArray(content.preSubScript, visitor, path);
        walkMathContentArray(content.preSuperScript, visitor, path);
        break;
      case "mathDelimiter":
        for (const group of content.content) {
          walkMathContentArray(group, visitor, path);
        }
        break;
      case "mathNary":
        if (content.sub) {
          walkMathContentArray(content.sub, visitor, path);
        }
        if (content.sup) {
          walkMathContentArray(content.sup, visitor, path);
        }
        walkMathContentArray(content.content, visitor, path);
        break;
      case "mathFunction":
        walkMathContentArray(content.name, visitor, path);
        walkMathContentArray(content.content, visitor, path);
        break;
      case "mathLimit":
        walkMathContentArray(content.base, visitor, path);
        walkMathContentArray(content.limit, visitor, path);
        break;
      case "mathMatrix":
        for (const row of content.rows) {
          for (const cell of row) {
            walkMathContentArray(cell, visitor, path);
          }
        }
        break;
      case "mathAccent":
      case "mathBar":
      case "mathBox":
      case "mathPhantom":
      case "mathBorderBox":
        walkMathContentArray(content.content, visitor, path);
        break;
      case "mathGroupChar":
        walkMathContentArray(content.base, visitor, path);
        break;
      case "mathEquationArray":
        for (const row of content.rows) {
          walkMathContentArray(row, visitor, path);
        }
        break;
      // mathRun has no children
    }
  }
}
// =============================================================================
// Convenience Collectors
// =============================================================================

/**
 * Collect all paragraphs from a document (including nested in tables/SDTs).
 */
export function collectParagraphs(doc: DocxDocument, options?: WalkOptions): Paragraph[] {
  const result: Paragraph[] = [];
  walkDocument(
    doc,
    {
      enterParagraph(para) {
        result.push(para);
        return "continue";
      }
    },
    options
  );
  return result;
}

/**
 * Collect all runs from a document (including nested).
 */
export function collectRuns(doc: DocxDocument, options?: WalkOptions): Run[] {
  const result: Run[] = [];
  walkDocument(
    doc,
    {
      enterRun(run) {
        result.push(run);
        return "continue";
      }
    },
    options
  );
  return result;
}

/**
 * Collect all tables from a document (including nested).
 */
export function collectTables(doc: DocxDocument, options?: WalkOptions): Table[] {
  const result: Table[] = [];
  walkDocument(
    doc,
    {
      enterTable(table) {
        result.push(table);
        return "continue";
      }
    },
    options
  );
  return result;
}
