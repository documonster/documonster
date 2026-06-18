/**
 * Track Changes — Accept / Reject API
 *
 * Functions for accepting or rejecting tracked changes in a DocxDocument.
 */

import { type Mutable } from "@word/core/internal-utils";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Run,
  Table,
  TableRow,
  TableCell,
  TableRowProperties,
  TableCellProperties,
  TableProperties,
  ParagraphProperties,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun
} from "@word/types";

// =============================================================================
// Public API
// =============================================================================

/**
 * Accept all tracked changes in a document (mutates in place).
 *
 * - Inserted runs: their content is kept, the `InsertedRun` wrapper is removed.
 * - Deleted runs: their content is removed entirely.
 * - Moved content: moved-to content is kept, moved-from is removed.
 * - Property changes: the new properties are kept, change info is removed.
 *
 * @param doc - The document to modify (mutated in place).
 * @returns Number of revisions accepted.
 */
export function acceptAllRevisions(doc: DocxDocument): number {
  let count = 0;
  const processBody = (body: BodyContent[]): void => {
    for (const block of body) {
      if (block.type === "paragraph") {
        count += acceptRevisionsInParagraph(block);
      } else if (block.type === "table") {
        // Table property change: accept keeps current, removes change
        if (block.properties?.propertyChange) {
          const props: Mutable<TableProperties> = { ...block.properties };
          delete props.propertyChange;
          (block as Mutable<Table>).properties = props;
          count++;
        }

        // Filter rows based on row-level insertion/deletion revisions
        const newRows: TableRow[] = [];
        for (const row of block.rows) {
          if (row.properties?.deleted) {
            // Accept deletion: remove the row
            count++;
            continue;
          }
          if (row.properties?.inserted) {
            // Accept insertion: keep the row, remove the revision marker
            const props: Mutable<TableRowProperties> = { ...row.properties };
            delete props.inserted;
            (row as Mutable<TableRow>).properties = props;
            count++;
          }
          // Row property change: accept keeps current, removes change
          if (row.properties?.propertyChange) {
            const props: Mutable<TableRowProperties> = { ...row.properties };
            delete props.propertyChange;
            (row as Mutable<TableRow>).properties = props;
            count++;
          }
          // Process cells
          for (const cell of row.cells) {
            // Cell property change: accept keeps current, removes change
            if (cell.properties?.propertyChange) {
              const props: Mutable<TableCellProperties> = { ...cell.properties };
              delete props.propertyChange;
              (cell as Mutable<TableCell>).properties = props;
              count++;
            }
            // Cell merge revision: accept removes the marker
            if (cell.properties?.cellMerge) {
              const props: Mutable<TableCellProperties> = { ...cell.properties };
              delete props.cellMerge;
              (cell as Mutable<TableCell>).properties = props;
              count++;
            }
            processBody(cell.content as BodyContent[]);
          }
          newRows.push(row);
        }
        (block as Mutable<Table>).rows = newRows;
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table")
        );
        processBody(filtered as BodyContent[]);
      } else if (block.type === "textBox") {
        // textBox.content is `readonly Paragraph[]`
        processBody(block.content as BodyContent[]);
      } else if (block.type === "drawingShape") {
        // shape text is `readonly Paragraph[]` (optional)
        if (block.textContent && block.textContent.length > 0) {
          processBody(block.textContent as BodyContent[]);
        }
      } else if (block.type === "tableOfContents") {
        if (block.cachedParagraphs && block.cachedParagraphs.length > 0) {
          processBody(block.cachedParagraphs as BodyContent[]);
        }
      }
    }
  };
  processBody(doc.body as BodyContent[]);

  // Process headers/footers
  if (doc.headers) {
    for (const [, header] of doc.headers) {
      processBody(header.content.children as BodyContent[]);
    }
  }
  if (doc.footers) {
    for (const [, footer] of doc.footers) {
      processBody(footer.content.children as BodyContent[]);
    }
  }
  // Process footnotes/endnotes — their content is paragraph[] which is a
  // subset of BodyContent[].
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      processBody(note.content as BodyContent[]);
    }
  }
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      processBody(note.content as BodyContent[]);
    }
  }
  // Process comments
  if (doc.comments) {
    for (const comment of doc.comments) {
      processBody(comment.content as BodyContent[]);
    }
  }

  return count;
}

/**
 * Reject all tracked changes in a document (mutates in place).
 *
 * - Inserted runs: removed entirely (content was not in original).
 * - Deleted runs: their content is kept (restoring original text).
 * - Moved content: moved-from content is kept, moved-to is removed.
 *
 * @param doc - The document to modify (mutated in place).
 * @returns Number of revisions rejected.
 */
export function rejectAllRevisions(doc: DocxDocument): number {
  let count = 0;
  const processBody = (body: BodyContent[]): void => {
    for (const block of body) {
      if (block.type === "paragraph") {
        count += rejectRevisionsInParagraph(block);
      } else if (block.type === "table") {
        // Table property change: reject restores previous properties
        if (block.properties?.propertyChange) {
          const change = block.properties.propertyChange;
          const props: Mutable<TableProperties> = {
            ...block.properties,
            ...change.previousProperties
          };
          delete props.propertyChange;
          (block as Mutable<Table>).properties = props;
          count++;
        }

        // Filter rows based on row-level insertion/deletion revisions
        const newRows: TableRow[] = [];
        for (const row of block.rows) {
          if (row.properties?.inserted) {
            // Reject insertion: remove the row (it wasn't in original)
            count++;
            continue;
          }
          if (row.properties?.deleted) {
            // Reject deletion: keep the row, remove the revision marker
            const props: Mutable<TableRowProperties> = { ...row.properties };
            delete props.deleted;
            (row as Mutable<TableRow>).properties = props;
            count++;
          }
          // Row property change: reject restores previous properties
          if (row.properties?.propertyChange) {
            const change = row.properties.propertyChange;
            const props: Mutable<TableRowProperties> = {
              ...row.properties,
              ...change.previousProperties
            };
            delete props.propertyChange;
            (row as Mutable<TableRow>).properties = props;
            count++;
          }
          // Process cells
          for (const cell of row.cells) {
            // Cell property change: reject restores previous properties
            if (cell.properties?.propertyChange) {
              const change = cell.properties.propertyChange;
              const props: Mutable<TableCellProperties> = {
                ...cell.properties,
                ...change.previousProperties
              };
              delete props.propertyChange;
              (cell as Mutable<TableCell>).properties = props;
              count++;
            }
            // Cell merge revision: reject removes the marker
            if (cell.properties?.cellMerge) {
              const props: Mutable<TableCellProperties> = { ...cell.properties };
              delete props.cellMerge;
              (cell as Mutable<TableCell>).properties = props;
              count++;
            }
            processBody(cell.content as BodyContent[]);
          }
          newRows.push(row);
        }
        (block as Mutable<Table>).rows = newRows;
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table")
        );
        processBody(filtered as BodyContent[]);
      } else if (block.type === "textBox") {
        // textBox.content is `readonly Paragraph[]`
        processBody(block.content as BodyContent[]);
      } else if (block.type === "drawingShape") {
        // shape text is `readonly Paragraph[]` (optional)
        if (block.textContent && block.textContent.length > 0) {
          processBody(block.textContent as BodyContent[]);
        }
      } else if (block.type === "tableOfContents") {
        if (block.cachedParagraphs && block.cachedParagraphs.length > 0) {
          processBody(block.cachedParagraphs as BodyContent[]);
        }
      }
    }
  };
  processBody(doc.body as BodyContent[]);

  // Process headers/footers
  if (doc.headers) {
    for (const [, header] of doc.headers) {
      processBody(header.content.children as BodyContent[]);
    }
  }
  if (doc.footers) {
    for (const [, footer] of doc.footers) {
      processBody(footer.content.children as BodyContent[]);
    }
  }
  // Process footnotes/endnotes/comments — symmetrical with acceptAllRevisions.
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      processBody(note.content as BodyContent[]);
    }
  }
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      processBody(note.content as BodyContent[]);
    }
  }
  if (doc.comments) {
    for (const comment of doc.comments) {
      processBody(comment.content as BodyContent[]);
    }
  }

  return count;
}

// =============================================================================
// Internal helpers
// =============================================================================

/** Accept revisions in a single paragraph — returns count of revisions accepted. */
function acceptRevisionsInParagraph(para: Paragraph): number {
  let count = 0;
  const children = para.children as ParagraphChild[];
  const newChildren: ParagraphChild[] = [];

  for (const child of children) {
    if ("type" in child) {
      switch ((child as { type: string }).type) {
        case "insertedRun": {
          // Accept insertion: keep the run's content
          const ins = child as InsertedRun;
          newChildren.push(ins.run as ParagraphChild);
          count++;
          break;
        }
        case "deletedRun": {
          // Accept deletion: remove the content
          count++;
          break;
        }
        case "movedFromRun": {
          // Accept move: moved-from content is removed
          count++;
          break;
        }
        case "movedToRun": {
          // Accept move: moved-to content is kept
          const moved = child as MovedToRun;
          newChildren.push(moved.run as ParagraphChild);
          count++;
          break;
        }
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          // Remove range markers
          count++;
          break;
        case "hyperlink": {
          // Hyperlink children are runs that may themselves carry tracked
          // changes — recurse via a synthetic paragraph so we re-use the
          // same accept logic. Other revision APIs (replaceText, fillFormFields)
          // already descend into hyperlinks; revisions must agree.
          const hl = child as { type: "hyperlink"; children: readonly ParagraphChild[] } & object;
          const synth = {
            type: "paragraph",
            children: [...hl.children]
          } as unknown as Paragraph;
          count += acceptRevisionsInParagraph(synth);
          newChildren.push({
            ...(hl as object),
            children: synth.children
          } as ParagraphChild);
          break;
        }
        default:
          newChildren.push(child);
          break;
      }
    } else {
      newChildren.push(child);
    }
  }

  // Replace children array
  (para as Mutable<Paragraph>).children = newChildren;

  // Remove paragraph property change if present
  if (para.properties?.propertyChange) {
    const props: Mutable<ParagraphProperties> = { ...para.properties };
    delete props.propertyChange;
    (para as Mutable<Paragraph>).properties = props;
    count++;
  }

  return count;
}

/** Reject revisions in a single paragraph — returns count of revisions rejected. */
function rejectRevisionsInParagraph(para: Paragraph): number {
  let count = 0;
  const children = para.children as ParagraphChild[];
  const newChildren: ParagraphChild[] = [];

  for (const child of children) {
    if ("type" in child) {
      switch ((child as { type: string }).type) {
        case "insertedRun": {
          // Reject insertion: remove the content (it wasn't in original)
          count++;
          break;
        }
        case "deletedRun": {
          // Reject deletion: keep the content (restore original)
          const del = child as DeletedRun;
          newChildren.push(del.run as ParagraphChild);
          count++;
          break;
        }
        case "movedFromRun": {
          // Reject move: keep moved-from content (restore to original position)
          const movedFrom = child as MovedFromRun;
          newChildren.push(movedFrom.run as ParagraphChild);
          count++;
          break;
        }
        case "movedToRun": {
          // Reject move: remove moved-to content
          count++;
          break;
        }
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          count++;
          break;
        case "hyperlink": {
          const hl = child as { type: "hyperlink"; children: readonly ParagraphChild[] } & object;
          const synth = {
            type: "paragraph",
            children: [...hl.children]
          } as unknown as Paragraph;
          count += rejectRevisionsInParagraph(synth);
          newChildren.push({
            ...(hl as object),
            children: synth.children
          } as ParagraphChild);
          break;
        }
        default:
          newChildren.push(child);
          break;
      }
    } else {
      newChildren.push(child);
    }
  }

  (para as Mutable<Paragraph>).children = newChildren;

  // Restore previous paragraph properties on rejection.
  if (para.properties?.propertyChange) {
    const change = para.properties.propertyChange;
    const props: Mutable<ParagraphProperties> = {
      ...para.properties,
      ...change.previousProperties
    };
    delete props.propertyChange;
    (para as Mutable<Paragraph>).properties = props;
    count++;
  }

  return count;
}

// =============================================================================
// Single-Revision API
// =============================================================================

/** A revision found in the document, with its location and metadata. */
export interface RevisionEntry {
  /** Unique revision id (from w:ins/w:del/w:moveFrom/w:moveTo @id). */
  readonly id: number;
  /** Revision type. */
  readonly type:
    | "insert"
    | "delete"
    | "moveFrom"
    | "moveTo"
    | "rowInsert"
    | "rowDelete"
    | "cellMerge"
    | "tablePropertyChange"
    | "rowPropertyChange"
    | "cellPropertyChange"
    | "paragraphPropertyChange"
    | "runPropertyChange";
  /** Author of the revision. */
  readonly author?: string;
  /** Date of the revision (ISO 8601 string). */
  readonly date?: string;
}

/**
 * List all tracked changes in a document.
 *
 * Walks the document body, headers, footers, footnotes, endnotes, and comments
 * collecting all `RevisionInfo`-bearing nodes (insertions, deletions, moves,
 * property changes).
 *
 * @param doc - The document to scan.
 * @returns Array of revision entries.
 */
export function listRevisions(doc: DocxDocument): RevisionEntry[] {
  const result: RevisionEntry[] = [];
  const seenIds = new Set<string>();

  const push = (entry: RevisionEntry): void => {
    const key = `${entry.type}:${entry.id}:${entry.author ?? ""}:${entry.date ?? ""}`;
    if (seenIds.has(key)) {
      return;
    }
    seenIds.add(key);
    result.push(entry);
  };

  const visitParagraph = (para: Paragraph): void => {
    if (para.properties?.propertyChange) {
      const rev = para.properties.propertyChange.revision;
      push({ id: rev.id, type: "paragraphPropertyChange", author: rev.author, date: rev.date });
    }
    for (const child of para.children) {
      if (!("type" in child)) {
        // Run — check for run property change
        if (
          (
            child as {
              properties?: {
                propertyChange?: { revision: { id: number; author: string; date?: string } };
              };
            }
          ).properties?.propertyChange
        ) {
          const rev = (
            child as {
              properties: {
                propertyChange: { revision: { id: number; author: string; date?: string } };
              };
            }
          ).properties.propertyChange.revision;
          push({ id: rev.id, type: "runPropertyChange", author: rev.author, date: rev.date });
        }
        continue;
      }
      const typed = child as { type: string };
      if (typed.type === "insertedRun") {
        const ins = child as InsertedRun;
        push({
          id: ins.revision.id,
          type: "insert",
          author: ins.revision.author,
          date: ins.revision.date
        });
      } else if (typed.type === "deletedRun") {
        const del = child as DeletedRun;
        push({
          id: del.revision.id,
          type: "delete",
          author: del.revision.author,
          date: del.revision.date
        });
      } else if (typed.type === "movedFromRun") {
        const mf = child as MovedFromRun;
        push({
          id: mf.revision.id,
          type: "moveFrom",
          author: mf.revision.author,
          date: mf.revision.date
        });
      } else if (typed.type === "movedToRun") {
        const mt = child as MovedToRun;
        push({
          id: mt.revision.id,
          type: "moveTo",
          author: mt.revision.author,
          date: mt.revision.date
        });
      } else if (typed.type === "hyperlink") {
        // Recurse into hyperlink — its run children may carry tracked
        // changes that the bulk acceptAll/rejectAll already process.
        const hl = child as { type: "hyperlink"; children: readonly ParagraphChild[] };
        visitParagraph({ type: "paragraph", children: [...hl.children] } as Paragraph);
      }
    }
  };

  const visitBlocks = (blocks: readonly BodyContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        visitParagraph(block);
      } else if (block.type === "table") {
        if (block.properties?.propertyChange) {
          const rev = block.properties.propertyChange.revision;
          push({
            id: rev.id,
            type: "tablePropertyChange",
            author: rev.author,
            date: rev.date
          });
        }
        for (const row of block.rows) {
          if (row.properties?.inserted) {
            push({
              id: row.properties.inserted.revision.id,
              type: "rowInsert",
              author: row.properties.inserted.revision.author,
              date: row.properties.inserted.revision.date
            });
          }
          if (row.properties?.deleted) {
            push({
              id: row.properties.deleted.revision.id,
              type: "rowDelete",
              author: row.properties.deleted.revision.author,
              date: row.properties.deleted.revision.date
            });
          }
          if (row.properties?.propertyChange) {
            const rev = row.properties.propertyChange.revision;
            push({
              id: rev.id,
              type: "rowPropertyChange",
              author: rev.author,
              date: rev.date
            });
          }
          for (const cell of row.cells) {
            if (cell.properties?.cellMerge) {
              push({
                id: cell.properties.cellMerge.revision.id,
                type: "cellMerge",
                author: cell.properties.cellMerge.revision.author,
                date: cell.properties.cellMerge.revision.date
              });
            }
            if (cell.properties?.propertyChange) {
              const rev = cell.properties.propertyChange.revision;
              push({
                id: rev.id,
                type: "cellPropertyChange",
                author: rev.author,
                date: rev.date
              });
            }
            visitBlocks(cell.content as readonly BodyContent[]);
          }
        }
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table")
        );
        visitBlocks(filtered as readonly BodyContent[]);
        // Inline (Run-only) SDT children may also wrap inserted/deleted
        // runs. We don't have a synthetic-paragraph helper here, but
        // check each Run directly for revision wrappers via its content
        // — the same fields paragraph runs would carry.
        for (const c of block.content) {
          if (
            c &&
            typeof c === "object" &&
            !("type" in c) &&
            "content" in c &&
            Array.isArray((c as { content?: unknown }).content)
          ) {
            // Synthesise a single-run paragraph so existing recursion
            // (visitParagraph) collects its revisions.
            visitParagraph({
              type: "paragraph",
              children: [c as Run]
            } as Paragraph);
          }
        }
      } else if (block.type === "textBox") {
        visitBlocks(block.content as readonly BodyContent[]);
      } else if (block.type === "drawingShape") {
        if (block.textContent && block.textContent.length > 0) {
          visitBlocks(block.textContent as readonly BodyContent[]);
        }
      } else if (block.type === "tableOfContents") {
        if (block.cachedParagraphs && block.cachedParagraphs.length > 0) {
          visitBlocks(block.cachedParagraphs as readonly BodyContent[]);
        }
      }
    }
  };

  visitBlocks(doc.body as readonly BodyContent[]);

  if (doc.headers) {
    for (const [, header] of doc.headers) {
      visitBlocks(header.content.children as readonly BodyContent[]);
    }
  }
  if (doc.footers) {
    for (const [, footer] of doc.footers) {
      visitBlocks(footer.content.children as readonly BodyContent[]);
    }
  }
  // Footnotes/endnotes/comments may also contain tracked changes — keeping
  // them in sync with acceptAllRevisions/rejectAllRevisions.
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      visitBlocks(note.content as readonly BodyContent[]);
    }
  }
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      visitBlocks(note.content as readonly BodyContent[]);
    }
  }
  if (doc.comments) {
    for (const c of doc.comments) {
      visitBlocks(c.content as readonly BodyContent[]);
    }
  }

  return result;
}

/**
 * Accept a single revision by id (mutates in place).
 *
 * @param doc - The document to modify.
 * @param revisionId - The revision id to accept.
 * @returns true if the revision was found and accepted, false otherwise.
 */
export function acceptRevision(doc: DocxDocument, revisionId: number): boolean {
  return processSingleRevision(doc, revisionId, "accept");
}

/**
 * Reject a single revision by id (mutates in place).
 *
 * @param doc - The document to modify.
 * @param revisionId - The revision id to reject.
 * @returns true if the revision was found and rejected, false otherwise.
 */
export function rejectRevision(doc: DocxDocument, revisionId: number): boolean {
  return processSingleRevision(doc, revisionId, "reject");
}

function processSingleRevision(
  doc: DocxDocument,
  revisionId: number,
  mode: "accept" | "reject"
): boolean {
  let found = false;

  const processParagraph = (para: Paragraph): void => {
    const children = para.children as ParagraphChild[];
    const newChildren: ParagraphChild[] = [];
    for (const child of children) {
      if (!("type" in child)) {
        newChildren.push(child);
        continue;
      }
      const typed = child as { type: string };
      switch (typed.type) {
        case "insertedRun": {
          const ins = child as InsertedRun;
          if (ins.revision.id === revisionId) {
            found = true;
            // Accept: keep run; Reject: remove
            if (mode === "accept") {
              newChildren.push(ins.run as ParagraphChild);
            }
          } else {
            newChildren.push(child);
          }
          break;
        }
        case "deletedRun": {
          const del = child as DeletedRun;
          if (del.revision.id === revisionId) {
            found = true;
            // Accept: remove; Reject: restore
            if (mode === "reject") {
              newChildren.push(del.run as ParagraphChild);
            }
          } else {
            newChildren.push(child);
          }
          break;
        }
        case "movedFromRun": {
          const mf = child as MovedFromRun;
          if (mf.revision.id === revisionId) {
            found = true;
            // Accept: remove; Reject: keep
            if (mode === "reject") {
              newChildren.push(mf.run as ParagraphChild);
            }
          } else {
            newChildren.push(child);
          }
          break;
        }
        case "movedToRun": {
          const mt = child as MovedToRun;
          if (mt.revision.id === revisionId) {
            found = true;
            // Accept: keep; Reject: remove
            if (mode === "accept") {
              newChildren.push(mt.run as ParagraphChild);
            }
          } else {
            newChildren.push(child);
          }
          break;
        }
        case "hyperlink": {
          // Hyperlinks can wrap tracked-change runs. Recurse via a
          // synthetic paragraph so the same id-matching logic applies.
          const hl = child as { type: "hyperlink"; children: readonly ParagraphChild[] } & object;
          const synth = {
            type: "paragraph",
            children: [...hl.children]
          } as unknown as Paragraph;
          processParagraph(synth);
          newChildren.push({
            ...(hl as object),
            children: synth.children
          } as ParagraphChild);
          break;
        }
        default:
          newChildren.push(child);
      }
    }
    (para as Mutable<Paragraph>).children = newChildren;

    // Paragraph property change
    if (para.properties?.propertyChange?.revision.id === revisionId) {
      found = true;
      const props: Mutable<ParagraphProperties> = { ...para.properties };
      if (mode === "reject") {
        Object.assign(props, para.properties.propertyChange.previousProperties ?? {});
      }
      delete props.propertyChange;
      (para as Mutable<Paragraph>).properties = props;
    }
  };

  const processBlocks = (blocks: readonly BodyContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        processParagraph(block);
      } else if (block.type === "table") {
        if (block.properties?.propertyChange?.revision.id === revisionId) {
          found = true;
          const props: Mutable<TableProperties> = { ...block.properties };
          if (mode === "reject") {
            Object.assign(props, block.properties.propertyChange.previousProperties ?? {});
          }
          delete props.propertyChange;
          (block as Mutable<Table>).properties = props;
        }
        const newRows: TableRow[] = [];
        for (const row of block.rows) {
          // Row insertion
          if (row.properties?.inserted?.revision.id === revisionId) {
            found = true;
            if (mode === "accept") {
              const props: Mutable<TableRowProperties> = { ...row.properties };
              delete props.inserted;
              (row as Mutable<TableRow>).properties = props;
              newRows.push(row);
            }
            // mode === "reject": skip row
            continue;
          }
          // Row deletion
          if (row.properties?.deleted?.revision.id === revisionId) {
            found = true;
            if (mode === "reject") {
              const props: Mutable<TableRowProperties> = { ...row.properties };
              delete props.deleted;
              (row as Mutable<TableRow>).properties = props;
              newRows.push(row);
            }
            // mode === "accept": skip row (delete it)
            continue;
          }
          // Row property change
          if (row.properties?.propertyChange?.revision.id === revisionId) {
            found = true;
            const props: Mutable<TableRowProperties> = { ...row.properties };
            if (mode === "reject") {
              Object.assign(props, row.properties.propertyChange.previousProperties ?? {});
            }
            delete props.propertyChange;
            (row as Mutable<TableRow>).properties = props;
          }
          // Cells
          for (const cell of row.cells) {
            if (cell.properties?.propertyChange?.revision.id === revisionId) {
              found = true;
              const props: Mutable<TableCellProperties> = { ...cell.properties };
              if (mode === "reject") {
                Object.assign(props, cell.properties.propertyChange.previousProperties ?? {});
              }
              delete props.propertyChange;
              (cell as Mutable<TableCell>).properties = props;
            }
            if (cell.properties?.cellMerge?.revision.id === revisionId) {
              found = true;
              const props: Mutable<TableCellProperties> = { ...cell.properties };
              delete props.cellMerge;
              (cell as Mutable<TableCell>).properties = props;
            }
            processBlocks(cell.content as readonly BodyContent[]);
          }
          newRows.push(row);
        }
        (block as Mutable<Table>).rows = newRows;
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table")
        );
        processBlocks(filtered as readonly BodyContent[]);
        // Inline (Run-only) SDT children: synthesise a paragraph so the
        // same per-run revision logic above runs against them.
        for (const c of block.content) {
          if (
            c &&
            typeof c === "object" &&
            !("type" in c) &&
            "content" in c &&
            Array.isArray((c as { content?: unknown }).content)
          ) {
            processParagraph({ type: "paragraph", children: [c as Run] } as Paragraph);
          }
        }
      } else if (block.type === "textBox") {
        processBlocks(block.content as readonly BodyContent[]);
      } else if (block.type === "drawingShape") {
        if (block.textContent && block.textContent.length > 0) {
          processBlocks(block.textContent as readonly BodyContent[]);
        }
      } else if (block.type === "tableOfContents") {
        if (block.cachedParagraphs && block.cachedParagraphs.length > 0) {
          processBlocks(block.cachedParagraphs as readonly BodyContent[]);
        }
      }
    }
  };

  processBlocks(doc.body as readonly BodyContent[]);

  if (doc.headers) {
    for (const [, header] of doc.headers) {
      processBlocks(header.content.children as readonly BodyContent[]);
    }
  }
  if (doc.footers) {
    for (const [, footer] of doc.footers) {
      processBlocks(footer.content.children as readonly BodyContent[]);
    }
  }
  // Notes and comments can also contain revisions; mirror the bulk APIs.
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      processBlocks(note.content as readonly BodyContent[]);
    }
  }
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      processBlocks(note.content as readonly BodyContent[]);
    }
  }
  if (doc.comments) {
    for (const c of doc.comments) {
      processBlocks(c.content as readonly BodyContent[]);
    }
  }

  return found;
}
