/**
 * DOCX Module - Immutable Document Mapper
 *
 * Provides `mapDocument()` and the `DocxTransformer` interface for rebuilding
 * a document tree with transformations applied to specific node types.
 *
 * Unlike the read-only `walkDocument()` (in walker.ts), this mapper produces
 * a new document without mutating the input.
 */

import type {
  BodyContent,
  DocxDocument,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  StructuredDocumentTag,
  Hyperlink,
  InsertedRun,
  MovedToRun,
  Table,
  TableRow,
  TableCell
} from "../types";
import { isRun } from "./text-utils";
import type { WalkPath } from "./walker";

// =============================================================================
// Types
// =============================================================================

/**
 * Transformer interface for immutable document tree rebuilding.
 * Return the node to keep/replace it, or null to remove it.
 * For flatMapBodyContent, return an array to expand one node into multiple.
 */
export interface DocxTransformer {
  transformParagraph?(para: Paragraph, path: WalkPath): Paragraph | null;
  transformRun?(run: Run, path: WalkPath): Run | null;
  transformTable?(table: Table, path: WalkPath): Table | null;
  transformSdt?(sdt: StructuredDocumentTag, path: WalkPath): StructuredDocumentTag | null;
  transformRunContent?(content: RunContent, run: Run, path: WalkPath): RunContent | null;
  /** Transform a body content item. Return null to remove, a single item to replace, or an array to expand. */
  transformBodyContent?(content: BodyContent, path: WalkPath): BodyContent | BodyContent[] | null;
}

/** Options controlling which parts of the document to map. */
export interface MapOptions {
  readonly includeHeaders?: boolean;
  readonly includeFooters?: boolean;
  readonly includeFootnotes?: boolean;
  readonly includeEndnotes?: boolean;
  readonly includeComments?: boolean;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Immutable document tree transformation.
 * Rebuilds the document tree, allowing callers to replace or filter nodes.
 * Never mutates the input document.
 *
 * @param doc - The document to transform.
 * @param transformer - Callbacks to transform each node type.
 * @param options - Controls which document parts to include.
 * @returns A new DocxDocument with transformations applied.
 */
export function mapDocument(
  doc: DocxDocument,
  transformer: DocxTransformer,
  options?: MapOptions
): DocxDocument {
  const opts: Required<MapOptions> = {
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

  // Map body
  const body = mapBlocks(doc.body, transformer, basePath);

  // Map headers
  let headers = doc.headers;
  if (opts.includeHeaders && doc.headers) {
    const newHeaders = new Map<
      string,
      typeof doc.headers extends ReadonlyMap<string, infer V> ? V : never
    >();
    for (const [key, headerDef] of doc.headers) {
      const headerPath: WalkPath = { ...basePath, inHeader: true };
      const mappedChildren = mapBlocks(
        headerDef.content.children as readonly BodyContent[],
        transformer,
        headerPath
      ) as readonly (Paragraph | Table)[];
      newHeaders.set(key, {
        ...headerDef,
        content: { ...headerDef.content, children: mappedChildren }
      });
    }
    headers = newHeaders;
  }

  // Map footers
  let footers = doc.footers;
  if (opts.includeFooters && doc.footers) {
    const newFooters = new Map<
      string,
      typeof doc.footers extends ReadonlyMap<string, infer V> ? V : never
    >();
    for (const [key, footerDef] of doc.footers) {
      const footerPath: WalkPath = { ...basePath, inFooter: true };
      const mappedChildren = mapBlocks(
        footerDef.content.children as readonly BodyContent[],
        transformer,
        footerPath
      ) as readonly (Paragraph | Table)[];
      newFooters.set(key, {
        ...footerDef,
        content: { ...footerDef.content, children: mappedChildren }
      });
    }
    footers = newFooters;
  }

  // Map footnotes
  let footnotes = doc.footnotes;
  if (opts.includeFootnotes && doc.footnotes) {
    footnotes = doc.footnotes.map(note => {
      if (note.id <= 0) {
        return note;
      }
      const notePath: WalkPath = { ...basePath, inFootnote: true };
      const mappedContent = mapBlocks(
        note.content as readonly BodyContent[],
        transformer,
        notePath
      ) as readonly Paragraph[];
      return { ...note, content: mappedContent };
    });
  }

  // Map endnotes
  let endnotes = doc.endnotes;
  if (opts.includeEndnotes && doc.endnotes) {
    endnotes = doc.endnotes.map(note => {
      if (note.id <= 0) {
        return note;
      }
      const notePath: WalkPath = { ...basePath, inEndnote: true };
      const mappedContent = mapBlocks(
        note.content as readonly BodyContent[],
        transformer,
        notePath
      ) as readonly Paragraph[];
      return { ...note, content: mappedContent };
    });
  }

  // Map comments
  let comments = doc.comments;
  if (opts.includeComments && doc.comments) {
    comments = doc.comments.map(comment => {
      const commentPath: WalkPath = { ...basePath, inComment: true };
      const mappedContent = mapBlocks(
        comment.content as readonly BodyContent[],
        transformer,
        commentPath
      ) as readonly Paragraph[];
      return { ...comment, content: mappedContent };
    });
  }

  return {
    ...doc,
    body,
    headers,
    footers,
    footnotes,
    endnotes,
    comments
  };
}

// =============================================================================
// Internal Mapper Helpers
// =============================================================================

function mapBlocks(
  blocks: readonly BodyContent[],
  transformer: DocxTransformer,
  basePath: WalkPath
): readonly BodyContent[] {
  const result: BodyContent[] = [];

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const currentPath: WalkPath = { ...basePath, index: i };

    // Call transformBodyContent first (supports null, single, or array return)
    if (transformer.transformBodyContent) {
      const transformed = transformer.transformBodyContent(block, currentPath);
      if (transformed === null) {
        continue;
      }
      if (Array.isArray(transformed)) {
        // Flat-map: one block becomes multiple blocks
        for (const item of transformed) {
          const mapped = mapBodyContent(item, transformer, currentPath);
          if (mapped !== null) {
            result.push(mapped);
          }
        }
        continue;
      }
      // Single replacement
      const mapped = mapBodyContent(transformed, transformer, currentPath);
      if (mapped !== null) {
        result.push(mapped);
      }
    } else {
      // No transformBodyContent — just recurse
      const mapped = mapBodyContent(block, transformer, currentPath);
      if (mapped !== null) {
        result.push(mapped);
      }
    }
  }

  return result;
}

/**
 * Maximum nesting depth permitted during transformation. See
 * `core/walker.ts` for the rationale; mirrored here so map and walk
 * share the same abuse-resistance posture.
 */
const MAX_MAP_DEPTH = 1000;

function mapBodyContent(
  block: BodyContent,
  transformer: DocxTransformer,
  path: WalkPath
): BodyContent | null {
  if (path.depth > MAX_MAP_DEPTH) {
    // Drop the offending subtree rather than blow the stack. Returning
    // null removes the block from its parent collection in mapBlocks.
    return null;
  }
  // Per-type transformer hooks fire only for paragraph / table / sdt.
  // Other BodyContent variants don't have a dedicated hook — callers that
  // want to replace them whole-cloth use `transformBodyContent` in
  // mapBlocks. We still need to recurse into their nested paragraph
  // structures so transformParagraph / transformRun / transformRunContent
  // visit text inside text boxes, drawing shape text, and TOC cached
  // paragraphs. Without that, e.g. replaceText silently misses content in
  // those structures.
  switch (block.type) {
    case "paragraph":
      return mapParagraph(block, transformer, path);
    case "table":
      return mapTable(block, transformer, path);
    case "sdt":
      return mapSdt(block as StructuredDocumentTag, transformer, path);
    case "textBox": {
      const tb = block as { content?: readonly Paragraph[] } & BodyContent;
      if (!tb.content || tb.content.length === 0) {
        return block;
      }
      const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
      const mappedContent = mapBlocks(tb.content as readonly BodyContent[], transformer, innerPath);
      if (mappedContent === (tb.content as readonly BodyContent[])) {
        return block;
      }
      // textBox.content is typed as `readonly Paragraph[]`. If a
      // `transformBodyContent` produced anything other than a paragraph
      // (a table, an image, etc.) we'd be silently violating the type
      // invariant — and downstream writers would crash trying to render
      // the unexpected node. Throw loudly instead of dropping the data.
      const filtered: Paragraph[] = [];
      for (const b of mappedContent) {
        if (b.type === "paragraph") {
          filtered.push(b);
        } else {
          throw new Error(
            `mapDocument: textBox.content must remain Paragraph[] but a ` +
              `transform produced "${b.type}". Either rewrite the transform ` +
              `to return only paragraphs here, or restructure the document ` +
              `at the body level instead.`
          );
        }
      }
      return { ...block, content: filtered } as BodyContent;
    }
    case "drawingShape": {
      const shape = block as { textContent?: readonly Paragraph[] } & BodyContent;
      if (!shape.textContent || shape.textContent.length === 0) {
        return block;
      }
      const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
      const mappedText = mapBlocks(
        shape.textContent as readonly BodyContent[],
        transformer,
        innerPath
      );
      if (mappedText === (shape.textContent as readonly BodyContent[])) {
        return block;
      }
      const filtered: Paragraph[] = [];
      for (const b of mappedText) {
        if (b.type === "paragraph") {
          filtered.push(b);
        } else {
          throw new Error(
            `mapDocument: drawingShape.textContent must remain Paragraph[] ` +
              `but a transform produced "${b.type}".`
          );
        }
      }
      return { ...block, textContent: filtered } as BodyContent;
    }
    case "tableOfContents": {
      const toc = block as { cachedParagraphs?: readonly Paragraph[] } & BodyContent;
      if (!toc.cachedParagraphs || toc.cachedParagraphs.length === 0) {
        return block;
      }
      const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
      const mappedCache = mapBlocks(
        toc.cachedParagraphs as readonly BodyContent[],
        transformer,
        innerPath
      );
      if (mappedCache === (toc.cachedParagraphs as readonly BodyContent[])) {
        return block;
      }
      const filtered: Paragraph[] = [];
      for (const b of mappedCache) {
        if (b.type === "paragraph") {
          filtered.push(b);
        } else {
          throw new Error(
            `mapDocument: tableOfContents.cachedParagraphs must remain ` +
              `Paragraph[] but a transform produced "${b.type}".`
          );
        }
      }
      return { ...block, cachedParagraphs: filtered } as BodyContent;
    }
    default:
      return block;
  }
}

function mapParagraph(
  para: Paragraph,
  transformer: DocxTransformer,
  path: WalkPath
): Paragraph | null {
  // Call transformParagraph
  let current: Paragraph | null = para;
  if (transformer.transformParagraph) {
    current = transformer.transformParagraph(para, path);
    if (current === null) {
      return null;
    }
  }

  // Map children (runs)
  if (transformer.transformRun || transformer.transformRunContent) {
    const newChildren: ParagraphChild[] = [];
    for (let i = 0; i < current.children.length; i++) {
      const child = current.children[i];
      const childPath: WalkPath = { ...path, index: i, depth: path.depth + 1 };
      const mapped = mapParagraphChild(child, transformer, childPath);
      if (mapped !== null) {
        newChildren.push(mapped);
      }
    }
    current = { ...current, children: newChildren };
  }

  return current;
}

function mapParagraphChild(
  child: ParagraphChild,
  transformer: DocxTransformer,
  path: WalkPath
): ParagraphChild | null {
  if (isRun(child)) {
    return mapRun(child, transformer, path);
  }

  if ("type" in child) {
    const typed = child as { type: string };
    switch (typed.type) {
      case "hyperlink": {
        const hl = child as Hyperlink;
        if (transformer.transformRun || transformer.transformRunContent) {
          const newChildren: Run[] = [];
          for (let i = 0; i < hl.children.length; i++) {
            const run = hl.children[i];
            const runPath: WalkPath = { ...path, index: i, depth: path.depth + 1 };
            const mapped = mapRun(run, transformer, runPath);
            if (mapped !== null) {
              newChildren.push(mapped);
            }
          }
          return { ...hl, children: newChildren } as ParagraphChild;
        }
        return child;
      }
      case "insertedRun": {
        const ins = child as InsertedRun;
        const mappedRun = mapRun(ins.run, transformer, path);
        if (mappedRun === null) {
          return null;
        }
        return { ...ins, run: mappedRun } as ParagraphChild;
      }
      case "movedToRun": {
        const moved = child as MovedToRun;
        const mappedRun = mapRun(moved.run, transformer, path);
        if (mappedRun === null) {
          return null;
        }
        return { ...moved, run: mappedRun } as ParagraphChild;
      }
      default:
        return child;
    }
  }

  // Default: treat as Run if isRun matches.
  if (isRun(child)) {
    return mapRun(child, transformer, path);
  }

  return child;
}

function mapRun(run: Run, transformer: DocxTransformer, path: WalkPath): Run | null {
  let current: Run | null = run;

  // Call transformRun
  if (transformer.transformRun) {
    current = transformer.transformRun(run, path);
    if (current === null) {
      return null;
    }
  }

  // Map run content
  if (transformer.transformRunContent) {
    const newContent: RunContent[] = [];
    for (const content of current.content) {
      const mapped = transformer.transformRunContent(content, current, path);
      if (mapped !== null) {
        newContent.push(mapped);
      }
    }
    current = { ...current, content: newContent };
  }

  return current;
}

function mapTable(table: Table, transformer: DocxTransformer, path: WalkPath): Table | null {
  // Call transformTable
  let current: Table | null = table;
  if (transformer.transformTable) {
    current = transformer.transformTable(table, path);
    if (current === null) {
      return null;
    }
  }

  // Always recurse into cells (even if transformTable returned as-is)
  const newRows: TableRow[] = [];
  for (let ri = 0; ri < current.rows.length; ri++) {
    const row = current.rows[ri];
    const rowPath: WalkPath = { ...path, index: ri, depth: path.depth + 1 };
    const newCells: TableCell[] = [];
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const cellPath: WalkPath = { ...rowPath, index: ci, depth: rowPath.depth + 1 };
      const mappedContent = mapBlocks(cell.content as readonly BodyContent[], transformer, {
        ...cellPath,
        depth: cellPath.depth + 1
      }) as readonly (Paragraph | Table)[];
      newCells.push({ ...cell, content: mappedContent });
    }
    newRows.push({ ...row, cells: newCells });
  }

  return { ...current, rows: newRows };
}

function mapSdt(
  sdt: StructuredDocumentTag,
  transformer: DocxTransformer,
  path: WalkPath
): StructuredDocumentTag | null {
  // Call transformSdt
  let current: StructuredDocumentTag | null = sdt;
  if (transformer.transformSdt) {
    current = transformer.transformSdt(sdt, path);
    if (current === null) {
      return null;
    }
  }

  // Recurse into content
  const innerPath: WalkPath = { ...path, depth: path.depth + 1 };
  const newContent: (Paragraph | Run | Table)[] = [];
  for (let i = 0; i < current.content.length; i++) {
    const item = current.content[i];
    const itemPath: WalkPath = { ...innerPath, index: i };
    if ("type" in item && item.type === "paragraph") {
      const mapped = mapParagraph(item as Paragraph, transformer, itemPath);
      if (mapped !== null) {
        newContent.push(mapped);
      }
    } else if ("type" in item && item.type === "table") {
      const mapped = mapTable(item as Table, transformer, itemPath);
      if (mapped !== null) {
        newContent.push(mapped);
      }
    } else if ("content" in item && Array.isArray((item as Run).content)) {
      const mapped = mapRun(item as Run, transformer, innerPath);
      if (mapped !== null) {
        newContent.push(mapped);
      }
    } else {
      newContent.push(item);
    }
  }

  return { ...current, content: newContent };
}
