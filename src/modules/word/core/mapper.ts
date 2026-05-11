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

function mapBodyContent(
  block: BodyContent,
  transformer: DocxTransformer,
  path: WalkPath
): BodyContent | null {
  switch (block.type) {
    case "paragraph":
      return mapParagraph(block, transformer, path);
    case "table":
      return mapTable(block, transformer, path);
    case "sdt":
      return mapSdt(block as StructuredDocumentTag, transformer, path);
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
  // Check if it's a Run (has content array, no discriminating type)
  if ("content" in child && Array.isArray((child as Run).content) && !("type" in child)) {
    return mapRun(child as Run, transformer, path);
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

  // Default: treat as Run if it has `content` array
  if ("content" in child && Array.isArray((child as Run).content)) {
    return mapRun(child as Run, transformer, path);
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
