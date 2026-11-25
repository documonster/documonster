/**
 * Track Changes — Accept / Reject API
 *
 * Functions for accepting or rejecting tracked changes in a DocxDocument.
 */

import {
  listUnrecognizedValues,
  setAsideValue,
  UNRECOGNIZED_KEY,
  unrecognizedEntry
} from "@utils/unrecognized-values";
import { checkHandle, checkInteger } from "@word/core/argument-checks";
import {
  PARAGRAPH_MARK_REVISIONS,
  PARAGRAPH_ONLY_PROPERTIES,
  ROW_ONLY_PROPERTIES,
  SECTION_ONLY_PROPERTIES
} from "@word/core/change-records";
import type { Mutable } from "@word/core/internal-utils";
import { ownTree } from "@word/core/ownership";
import { arrangeParagraphChildren, spellOutRangeMarkers } from "@word/core/range-markers";
import { hyperlinkContent, withHyperlinkContent } from "@word/core/text-utils";
import type {
  DocxDocument,
  BodyContent,
  Hyperlink,
  Paragraph,
  ParagraphChild,
  RevisionInfo,
  Run,
  RunProperties,
  SectionProperties,
  Table,
  TableRow,
  TableCell,
  TableRowProperties,
  TableCellProperties,
  TableProperties,
  TablePropertyExceptions,
  ParagraphProperties,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  RangeMarker
} from "@word/types";

/**
 * Give `para` (in place) the children `items`, re-attaching the range markers
 * spelled out among them (see {@link arrangeParagraphChildren}).
 */
function setParagraphChildren(para: Paragraph, items: readonly ParagraphChild[]): void {
  const arranged = arrangeParagraphChildren(items);
  const target = para as Mutable<Paragraph>;
  target.children = arranged.children;
  if (arranged.rangeMarkersAfter) {
    target.rangeMarkersAfter = [...arranged.rangeMarkersAfter, ...(para.rangeMarkersAfter ?? [])];
  }
}

// =============================================================================
// Range markers between blocks
// =============================================================================

function isMoveRangeMarker(node: { type?: string }): boolean {
  return (
    node.type === "moveFromRangeStart" ||
    node.type === "moveFromRangeEnd" ||
    node.type === "moveToRangeStart" ||
    node.type === "moveToRangeEnd"
  );
}

/**
 * Remove the move-range markers between the blocks of `body` (in place), as
 * accepting or rejecting does for those inside a paragraph. Returns the number
 * removed.
 */
function dropBodyMoveRanges(body: readonly BodyContent[]): number {
  let removed = 0;
  for (const block of body) {
    removed += dropMoveRanges(block as Mutable<BodyContent>, "rangeMarkersBefore");
  }
  return removed;
}

/** Remove the move-range markers from the marker list `owner[key]` (in place). */
function dropMoveRanges<T extends object>(owner: T, key: keyof T & string): number {
  const [rest, dropped] = withoutMoveRanges(
    (owner as Record<string, readonly RangeMarker[] | undefined>)[key]
  );
  if (dropped > 0) {
    if (rest) {
      (owner as Record<string, unknown>)[key] = rest;
    } else {
      delete (owner as Record<string, unknown>)[key];
    }
  }
  return dropped;
}

/** `markers` without move-range markers, and how many were removed. */
function withoutMoveRanges(
  markers: readonly RangeMarker[] | undefined
): [readonly RangeMarker[] | undefined, number] {
  if (!markers) {
    return [undefined, 0];
  }
  const kept = markers.filter(m => !isMoveRangeMarker(m));
  return [kept.length > 0 ? kept : undefined, markers.length - kept.length];
}

/**
 * Replace the rows of `table` with `keptRows` (in place) without losing the
 * range markers of the rows taken out — a bookmark starting before a deleted
 * row must still start somewhere — by moving them before the next kept row,
 * or after the table's rows. With `dropMoves`, move-range markers are removed
 * from the table, as accepting or rejecting all revisions does inside a
 * paragraph. Returns the number of markers removed.
 */
function replaceTableRows(
  table: Mutable<Table>,
  keptRows: readonly TableRow[],
  dropMoves: boolean
): number {
  const keep = new Set(keptRows);
  let count = 0;
  let carried: RangeMarker[] = [];
  const settle = <T extends object>(owner: T, key: keyof T & string): void => {
    if (dropMoves) {
      count += dropMoveRanges(owner, key);
    }
  };
  const kept: TableRow[] = [];
  for (const row of table.rows) {
    const mrow = row as Mutable<TableRow>;
    settle(mrow, "rangeMarkersBefore");
    settle(mrow, "rangeMarkersAfter");
    for (const cell of row.cells) {
      settle(cell as Mutable<TableCell>, "rangeMarkersBefore");
    }
    if (!keep.has(row)) {
      carried.push(
        ...(row.rangeMarkersBefore ?? []),
        ...row.cells.flatMap(cell => cell.rangeMarkersBefore ?? []),
        ...(row.rangeMarkersAfter ?? [])
      );
      continue;
    }
    if (carried.length > 0) {
      mrow.rangeMarkersBefore = [...carried, ...(row.rangeMarkersBefore ?? [])];
      carried = [];
    }
    kept.push(row);
  }
  settle(table, "rangeMarkersAfter");
  if (carried.length > 0) {
    table.rangeMarkersAfter = [...carried, ...(table.rangeMarkersAfter ?? [])];
  }
  table.rows = kept;
  return count;
}

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
  checkHandle(doc, "document", "doc", "Document.acceptAllRevisions");
  ownTree(doc);
  let count = 0;
  const processBody = (body: BodyContent[], marks = true): void => {
    count += dropBodyMoveRanges(body);
    for (const block of body) {
      if (block.type === "paragraph") {
        count += acceptRevisionsInParagraph(block);
      } else if (block.type === "table") {
        // Table property change: accept keeps current, removes change
        if (block.properties?.propertyChange) {
          (block as Mutable<Table>).properties = settleChange(
            block.properties,
            "accept",
            OUTSIDE_CHANGE.table
          );
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
            (row as Mutable<TableRow>).properties = settleChange(
              row.properties,
              "accept",
              OUTSIDE_CHANGE.row
            );
            count++;
          }
          if (settleExceptionsChange(row, "accept")) {
            count++;
          }
          // Process cells
          for (const cell of row.cells) {
            // Cell property change: accept keeps current, removes change
            if (cell.properties?.propertyChange) {
              (cell as Mutable<TableCell>).properties = settleChange(
                cell.properties,
                "accept",
                OUTSIDE_CHANGE.cell
              );
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
        count += replaceTableRows(block as Mutable<Table>, newRows, true);
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table" || c.type === "sdt")
        );
        processBody(filtered as BodyContent[], false);
        eachInlineStretch(block.content as unknown[], p => {
          count += acceptRevisionsInParagraph(p);
        });
        count += settleParagraphMarks(block.content as unknown[], "accept");
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
    if (marks) {
      count += settleParagraphMarks(body, "accept");
    }
  };
  processBody(doc.body as BodyContent[]);
  if (settleDocumentSection(doc, "accept")) {
    count++;
  }
  count += dropMoveRanges(doc as Mutable<DocxDocument>, "bodyRangeMarkersAfter");

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
  checkHandle(doc, "document", "doc", "Document.rejectAllRevisions");
  ownTree(doc);
  let count = 0;
  const processBody = (body: BodyContent[], marks = true): void => {
    count += dropBodyMoveRanges(body);
    for (const block of body) {
      if (block.type === "paragraph") {
        count += rejectRevisionsInParagraph(block);
      } else if (block.type === "table") {
        // Table property change: reject restores previous properties
        if (block.properties?.propertyChange) {
          (block as Mutable<Table>).properties = settleChange(
            block.properties,
            "reject",
            OUTSIDE_CHANGE.table
          );
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
            (row as Mutable<TableRow>).properties = settleChange(
              row.properties,
              "reject",
              OUTSIDE_CHANGE.row
            );
            count++;
          }
          if (settleExceptionsChange(row, "reject")) {
            count++;
          }
          // Process cells
          for (const cell of row.cells) {
            // Cell property change: reject restores previous properties
            if (cell.properties?.propertyChange) {
              (cell as Mutable<TableCell>).properties = settleChange(
                cell.properties,
                "reject",
                OUTSIDE_CHANGE.cell
              );
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
        count += replaceTableRows(block as Mutable<Table>, newRows, true);
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table" || c.type === "sdt")
        );
        processBody(filtered as BodyContent[], false);
        eachInlineStretch(block.content as unknown[], p => {
          count += rejectRevisionsInParagraph(p);
        });
        count += settleParagraphMarks(block.content as unknown[], "reject");
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
    if (marks) {
      count += settleParagraphMarks(body, "reject");
    }
  };
  processBody(doc.body as BodyContent[]);
  if (settleDocumentSection(doc, "reject")) {
    count++;
  }
  count += dropMoveRanges(doc as Mutable<DocxDocument>, "bodyRangeMarkersAfter");

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
  let count = settleRunPropertyChanges(para, "accept");
  // Markers riding on a child (range permissions) are spelled out first so
  // they survive the child being dropped or unwrapped, then re-attached.
  const children = spellOutRangeMarkers(para.children) as ParagraphChild[];
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
          const hl = child as Hyperlink;
          const synth: Paragraph = {
            type: "paragraph",
            children: hyperlinkContent(hl)
          };
          count += acceptRevisionsInParagraph(synth);
          newChildren.push(withHyperlinkContent(hl, synth.children));
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
  setParagraphChildren(para, newChildren);

  if (settleParagraphSection(para, "accept")) {
    count++;
  }
  if (para.properties?.propertyChange) {
    (para as Mutable<Paragraph>).properties = settleChange(
      para.properties,
      "accept",
      OUTSIDE_CHANGE.paragraph
    );
    count++;
  }

  return count;
}

/** Reject revisions in a single paragraph — returns count of revisions rejected. */
function rejectRevisionsInParagraph(para: Paragraph): number {
  let count = settleRunPropertyChanges(para, "reject");
  // Markers riding on a child (range permissions) are spelled out first so
  // they survive the child being dropped or unwrapped, then re-attached.
  const children = spellOutRangeMarkers(para.children) as ParagraphChild[];
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
          const hl = child as Hyperlink;
          const synth: Paragraph = {
            type: "paragraph",
            children: hyperlinkContent(hl)
          };
          count += rejectRevisionsInParagraph(synth);
          newChildren.push(withHyperlinkContent(hl, synth.children));
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

  setParagraphChildren(para, newChildren);

  if (settleParagraphSection(para, "reject")) {
    count++;
  }
  if (para.properties?.propertyChange) {
    (para as Mutable<Paragraph>).properties = settleChange(
      para.properties,
      "reject",
      OUTSIDE_CHANGE.paragraph
    );
    count++;
  }

  return count;
}

// =============================================================================
// Single-Revision API
// =============================================================================

/**
 * The fields of each kind of properties that a rejected change keeps as they are: what the change record's previous
 * properties do not cover (a paragraph's section and mark, a row's insertion, deletion and exceptions); a cell's
 * insertion, deletion and merge, which `CT_TcPrInner` does cover but which are revisions of their own — listed and
 * settled separately, so the current one stands and a previous one is not brought back as a revision nobody listed;
 * and the model's own convenience (`rowSpan`). Everything else is the previous state: `w:pPrChange`, `w:rPrChange`,
 * `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange` and `w:tblPrExChange` each hold the whole of it (`CT_PPrBase`,
 * `CT_RPrOriginal`, `CT_TblPrBase`, `CT_TrPrBase`, `CT_TcPrInner`, `CT_TblPrExBase`). The lists are
 * `@word/core/change-records`', which the reader, the argument checks and the writer use as well.
 */
const OUTSIDE_CHANGE = {
  paragraph: PARAGRAPH_ONLY_PROPERTIES,
  run: [],
  table: [],
  row: ROW_ONLY_PROPERTIES,
  cell: ["rowSpan", "inserted", "deleted", "cellMerge"],
  exceptions: [],
  section: SECTION_ONLY_PROPERTIES
} as const satisfies {
  readonly paragraph: readonly (keyof ParagraphProperties)[];
  readonly run: readonly (keyof RunProperties)[];
  readonly table: readonly (keyof TableProperties)[];
  readonly row: readonly (keyof TableRowProperties)[];
  readonly cell: readonly (keyof TableCellProperties)[];
  readonly exceptions: readonly (keyof TablePropertyExceptions)[];
  readonly section: readonly (keyof SectionProperties)[];
};

interface ChangeHolder {
  readonly propertyChange?: {
    readonly revision: RevisionInfo;
    readonly previousProperties?: object;
  };
}

/**
 * `properties` with their change record settled: accepted, the current properties without it; rejected, the previous
 * properties in their place — not merged over them, or a property the change added would survive its rejection —
 * except for the fields in `outside`, which keep their current values and kept text. A change record a read set aside
 * inside the previous properties (they have none of their own) means nothing once they are current, and is dropped.
 */
function settleChange<T extends ChangeHolder>(
  properties: T,
  mode: "accept" | "reject",
  outside: readonly string[]
): T {
  if (mode === "accept") {
    const accepted = { ...properties } as Mutable<T>;
    delete accepted.propertyChange;
    return accepted;
  }
  return withFieldsOf(properties.propertyChange?.previousProperties ?? {}, properties, outside, [
    "propertyChange"
  ]) as T;
}

/**
 * A copy of `into` with each of `fields` — its value and any text a read kept for it — taken from `from` instead, and
 * each of `drop` left out. The kept text is rebuilt on the copy alone, so the nested objects it shares with `into`
 * and `from` are not touched.
 */
function withFieldsOf(
  into: object,
  from: object | undefined,
  fields: readonly string[],
  drop: readonly string[] = []
): Record<string, unknown> {
  const result: Record<string, unknown> = { ...into };
  delete result[UNRECOGNIZED_KEY];
  for (const field of drop) {
    delete result[field];
  }
  const keep = (holder: object | undefined, field: string): void => {
    const kept = unrecognizedEntry(holder, field);
    if (kept !== undefined) {
      const { value, ...options } = kept;
      setAsideValue(result, field, value, options);
    }
  };
  for (const { field } of listUnrecognizedValues(into)) {
    if (!fields.includes(field) && !drop.includes(field)) {
      keep(into, field);
    }
  }
  for (const field of fields) {
    delete result[field];
    const value = (from as Record<string, unknown> | undefined)?.[field];
    if (value !== undefined) {
      result[field] = value;
    }
    keep(from, field);
  }
  return result;
}

/**
 * A paragraph mark's revisions: the field each is in, the type it is listed as, and which way of settling it removes
 * the mark — joining the paragraph with the next. An inserted or moved-in mark is removed by rejecting it, a deleted
 * or moved-away one by accepting it.
 */
const MARK_SETTLING = {
  ins: { type: "insert", joinOn: "reject" },
  del: { type: "delete", joinOn: "accept" },
  moveFrom: { type: "moveFrom", joinOn: "accept" },
  moveTo: { type: "moveTo", joinOn: "reject" }
} as const satisfies Record<
  (typeof PARAGRAPH_MARK_REVISIONS)[number]["element"],
  { readonly type: RevisionEntry["type"]; readonly joinOn: "accept" | "reject" }
>;

const MARK_REVISIONS = PARAGRAPH_MARK_REVISIONS.map(({ element, field }) => ({
  field,
  ...MARK_SETTLING[element]
}));

const isParagraph = (block: unknown): block is Paragraph =>
  typeof block === "object" && block !== null && (block as { type?: unknown }).type === "paragraph";

/**
 * Settle the paragraph mark revisions in `blocks` — those with `id`, when given — and return how many there were. A
 * mark that settling removes joins its paragraph with the next one in `blocks`, the array being edited in place; with
 * no paragraph next (the end of a cell, a table after it) there is nothing to join, and only the revision goes. Nor
 * is there across a content control's edge: its content is a list of its own. The joined paragraph is checked again,
 * since the mark it now ends with is the next paragraph's.
 */
function settleParagraphMarks(blocks: unknown[], mode: "accept" | "reject", id?: number): number {
  let settled = 0;
  for (let i = 0; i < blocks.length; i++) {
    const paragraph = blocks[i];
    if (!isParagraph(paragraph) || paragraph.properties === undefined) {
      continue;
    }
    let properties: Mutable<ParagraphProperties> | undefined;
    let join = false;
    for (const { field, joinOn } of MARK_REVISIONS) {
      const revision = paragraph.properties[field];
      if (revision === undefined || (id !== undefined && revision.id !== id)) {
        continue;
      }
      properties ??= { ...paragraph.properties };
      delete properties[field];
      join ||= joinOn === mode;
      settled++;
    }
    if (properties === undefined) {
      continue;
    }
    (paragraph as Mutable<Paragraph>).properties = properties;
    const next = blocks[i + 1];
    if (join && isParagraph(next)) {
      blocks.splice(i, 2, joinParagraphs(paragraph, next));
      i--;
    }
  }
  return settled;
}

/**
 * `first` and `second` as one paragraph, as removing the mark between them makes them: the content of both, and
 * `second`'s properties, whole — its formatting, its mark, a section ending there and a formatting change pending on
 * it. That is what Word does when it settles the revision (checked against Word itself): the joined paragraph takes
 * the surviving mark's paragraph properties, and `first`'s, with any formatting change pending on them, go. The walks
 * that settle everything settle each paragraph's own changes before its mark, so only settling one revision by its id
 * meets a change on `first` that goes with it.
 */
function joinParagraphs(first: Paragraph, second: Paragraph): Paragraph {
  const joined: Mutable<Paragraph> = {
    ...first,
    properties: second.properties,
    rangeMarkersAfter: second.rangeMarkersAfter,
    children: []
  };
  if (joined.properties === undefined) {
    delete joined.properties;
  }
  if (joined.rangeMarkersAfter === undefined) {
    delete joined.rangeMarkersAfter;
  }
  setParagraphChildren(joined, [
    ...(spellOutRangeMarkers(first.children) as ParagraphChild[]),
    ...((first.rangeMarkersAfter ?? []) as ParagraphChild[]),
    ...((second.rangeMarkersBefore ?? []) as ParagraphChild[]),
    ...(spellOutRangeMarkers(second.children) as ParagraphChild[])
  ]);
  return joined;
}

/** The revisions that wrap a run, whose run may carry a property change of its own. */
const RUN_WRAPPERS: ReadonlySet<string> = new Set([
  "insertedRun",
  "deletedRun",
  "movedFromRun",
  "movedToRun"
]);

/** A content control's run-level item: a run, or a run wrapped in a revision. */
function isInline(item: unknown): boolean {
  if (typeof item !== "object" || item === null) {
    return false;
  }
  return !("type" in item) || RUN_WRAPPERS.has((item as { type: string }).type);
}

/**
 * Hand `visit` each stretch of run-level items in a content control's `content` as the paragraph they would be in,
 * and put what it leaves there in their place: their revisions are listed and settled as a paragraph's are, by every
 * walk alike. A visit that changes nothing changes nothing here either, so listing does not edit the document.
 */
function eachInlineStretch(content: unknown[], visit: (paragraph: Paragraph) => void): void {
  for (let i = 0; i < content.length;) {
    if (!isInline(content[i])) {
      i++;
      continue;
    }
    let end = i;
    while (end < content.length && isInline(content[end])) {
      end++;
    }
    const children = content.slice(i, end) as ParagraphChild[];
    const stretch: Paragraph = { type: "paragraph", children };
    visit(stretch);
    if (stretch.children === children && stretch.rangeMarkersAfter === undefined) {
      i = end;
      continue;
    }
    const settled = [...stretch.children, ...(stretch.rangeMarkersAfter ?? [])];
    content.splice(i, end - i, ...settled);
    i += settled.length;
  }
}

/**
 * The run properties directly in `para` that can carry a change record (`w:rPrChange`): each run's — on its own or
 * inside an insertion, deletion or move — and the paragraph mark's. What {@link listRevisions} lists as
 * `runPropertyChange`, and what {@link settleRunPropertyChanges} settles; a hyperlink's runs are reached through the
 * synthetic paragraph every revision walk builds for it.
 */
function* changeableRunProperties(para: Paragraph): Generator<RunProperties> {
  const mark = para.properties?.markRunProperties;
  if (mark) {
    yield mark;
  }
  for (const child of para.children) {
    const run = !("type" in child)
      ? (child as Run)
      : RUN_WRAPPERS.has((child as { type: string }).type)
        ? (child as InsertedRun).run
        : undefined;
    if (run?.properties) {
      yield run.properties;
    }
  }
}

/**
 * Settle the run property changes of {@link changeableRunProperties} — those with `id`, when given — as
 * {@link settleChange} does, and return how many there were. Done before the paragraph's other revisions, so a run
 * whose insertion or deletion then removes it has had its change counted, as it was listed. These used to be listed
 * and never settled: accepting or rejecting everything left them, and settling one by its id answered `false`.
 */
function settleRunPropertyChanges(para: Paragraph, mode: "accept" | "reject", id?: number): number {
  let settled = 0;
  const settle = (run: Run): Run => {
    const change = run.properties?.propertyChange;
    if (change === undefined || (id !== undefined && change.revision.id !== id)) {
      return run;
    }
    settled++;
    return { ...run, properties: settleChange(run.properties!, mode, OUTSIDE_CHANGE.run) };
  };
  const children = para.children.map(child => {
    if (!("type" in child)) {
      return settle(child as Run);
    }
    if (!RUN_WRAPPERS.has((child as { type: string }).type)) {
      return child;
    }
    const wrapper = child as InsertedRun;
    const run = settle(wrapper.run);
    return run === wrapper.run ? child : { ...wrapper, run };
  });
  if (settled > 0) {
    // Through the setter, as every other change to a paragraph's children is; arranging arranged children is a no-op.
    setParagraphChildren(para, children as ParagraphChild[]);
  }
  const mark = para.properties?.markRunProperties;
  const change = mark?.propertyChange;
  if (change !== undefined && (id === undefined || change.revision.id === id)) {
    (para as Mutable<Paragraph>).properties = {
      ...para.properties,
      markRunProperties: settleChange(mark!, mode, OUTSIDE_CHANGE.run)
    };
    settled++;
  }
  return settled;
}

/**
 * `section` with its change record (`w:sectPrChange`) settled as {@link settleChange} does — the one with `id`, when
 * given — or `undefined` when it has none to settle. A section ends a paragraph (in its mark's `w:pPr`) or the
 * document; both are settled by every walk.
 */
function settledSection(
  section: SectionProperties | undefined,
  mode: "accept" | "reject",
  id?: number
): SectionProperties | undefined {
  const change = section?.propertyChange;
  if (change === undefined || (id !== undefined && change.revision.id !== id)) {
    return undefined;
  }
  return settleChange(section!, mode, OUTSIDE_CHANGE.section);
}

/** Settle the change of the section `para` ends, if any — see {@link settledSection}. Whether there was one. */
function settleParagraphSection(para: Paragraph, mode: "accept" | "reject", id?: number): boolean {
  const section = settledSection(para.properties?.sectionProperties, mode, id);
  if (section === undefined) {
    return false;
  }
  (para as Mutable<Paragraph>).properties = { ...para.properties, sectionProperties: section };
  return true;
}

/** Settle the change of the document's last section, if any — see {@link settledSection}. Whether there was one. */
function settleDocumentSection(doc: DocxDocument, mode: "accept" | "reject", id?: number): boolean {
  const section = settledSection(doc.sectionProperties, mode, id);
  if (section === undefined) {
    return false;
  }
  (doc as Mutable<DocxDocument>).sectionProperties = section;
  return true;
}

/**
 * Settle the change to `row`'s table property exceptions (`w:tblPrExChange`) — the one with `id`, when given — as
 * {@link settleChange} does. Whether there was one to settle.
 */
function settleExceptionsChange(row: TableRow, mode: "accept" | "reject", id?: number): boolean {
  const exceptions = row.properties?.tblPrEx;
  const change = exceptions?.propertyChange;
  if (change === undefined || (id !== undefined && change.revision.id !== id)) {
    return false;
  }
  (row as Mutable<TableRow>).properties = {
    ...row.properties,
    tblPrEx: settleChange(exceptions!, mode, OUTSIDE_CHANGE.exceptions)
  };
  return true;
}

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
    /** A change to a row's table property exceptions (`w:tblPrExChange`). */
    | "tablePropertyExceptionsChange"
    /** A change to a section's properties (`w:sectPrChange`), on a section break or the document's last section. */
    | "sectionPropertyChange"
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
  checkHandle(doc, "document", "doc", "Query.listRevisions");
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

  const pushSection = (section: SectionProperties | undefined): void => {
    const rev = section?.propertyChange?.revision;
    if (rev) {
      push({ id: rev.id, type: "sectionPropertyChange", author: rev.author, date: rev.date });
    }
  };

  const visitParagraph = (para: Paragraph): void => {
    for (const { field, type } of MARK_REVISIONS) {
      const rev = para.properties?.[field];
      if (rev) {
        push({ id: rev.id, type, author: rev.author, date: rev.date });
      }
    }
    if (para.properties?.propertyChange) {
      const rev = para.properties.propertyChange.revision;
      push({ id: rev.id, type: "paragraphPropertyChange", author: rev.author, date: rev.date });
    }
    pushSection(para.properties?.sectionProperties);
    for (const properties of changeableRunProperties(para)) {
      if (properties.propertyChange) {
        const rev = properties.propertyChange.revision;
        push({ id: rev.id, type: "runPropertyChange", author: rev.author, date: rev.date });
      }
    }
    for (const child of para.children) {
      if (!("type" in child)) {
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
        const hl = child as Hyperlink;
        visitParagraph({
          type: "paragraph",
          children: hyperlinkContent(hl)
        } as Paragraph);
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
          if (row.properties?.tblPrEx?.propertyChange) {
            const rev = row.properties.tblPrEx.propertyChange.revision;
            push({
              id: rev.id,
              type: "tablePropertyExceptionsChange",
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
          c => "type" in c && (c.type === "paragraph" || c.type === "table" || c.type === "sdt")
        );
        visitBlocks(filtered as readonly BodyContent[]);
        eachInlineStretch(block.content as unknown[], visitParagraph);
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
  pushSection(doc.sectionProperties);

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
 * Settling a paragraph mark's insertion, deletion or move can join its paragraph with the next one. As in Word, the
 * joined paragraph has the next one's properties, a formatting change pending on them included; the first's go.
 *
 * @param doc - The document to modify.
 * @param revisionId - The revision id to accept.
 * @returns true if the revision was found and accepted, false otherwise.
 */
export function acceptRevision(doc: DocxDocument, revisionId: number): boolean {
  checkHandle(doc, "document", "doc", "Document.acceptRevision");
  checkInteger(revisionId, "revisionId", "Document.acceptRevision");
  return processSingleRevision(doc, revisionId, "accept");
}

/**
 * Reject a single revision by id (mutates in place).
 *
 * Settling a paragraph mark's insertion, deletion or move can join its paragraph with the next one. As in Word, the
 * joined paragraph has the next one's properties, a formatting change pending on them included; the first's go.
 *
 * @param doc - The document to modify.
 * @param revisionId - The revision id to reject.
 * @returns true if the revision was found and rejected, false otherwise.
 */
export function rejectRevision(doc: DocxDocument, revisionId: number): boolean {
  checkHandle(doc, "document", "doc", "Document.rejectRevision");
  checkInteger(revisionId, "revisionId", "Document.rejectRevision");
  return processSingleRevision(doc, revisionId, "reject");
}

function processSingleRevision(
  doc: DocxDocument,
  revisionId: number,
  mode: "accept" | "reject"
): boolean {
  ownTree(doc);
  let found = false;

  const processParagraph = (para: Paragraph): void => {
    if (settleRunPropertyChanges(para, mode, revisionId) > 0) {
      found = true;
    }
    const children = spellOutRangeMarkers(para.children) as ParagraphChild[];
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
          const hl = child as Hyperlink;
          const synth: Paragraph = {
            type: "paragraph",
            children: hyperlinkContent(hl)
          };
          processParagraph(synth);
          newChildren.push(withHyperlinkContent(hl, synth.children));
          break;
        }
        default:
          newChildren.push(child);
      }
    }
    setParagraphChildren(para, newChildren);

    if (settleParagraphSection(para, mode, revisionId)) {
      found = true;
    }
    if (para.properties?.propertyChange?.revision.id === revisionId) {
      found = true;
      (para as Mutable<Paragraph>).properties = settleChange(
        para.properties,
        mode,
        OUTSIDE_CHANGE.paragraph
      );
    }
  };

  const processBlocks = (blocks: readonly BodyContent[], marks = true): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        processParagraph(block);
      } else if (block.type === "table") {
        if (block.properties?.propertyChange?.revision.id === revisionId) {
          found = true;
          (block as Mutable<Table>).properties = settleChange(
            block.properties,
            mode,
            OUTSIDE_CHANGE.table
          );
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
            (row as Mutable<TableRow>).properties = settleChange(
              row.properties,
              mode,
              OUTSIDE_CHANGE.row
            );
          }
          if (settleExceptionsChange(row, mode, revisionId)) {
            found = true;
          }
          // Cells
          for (const cell of row.cells) {
            if (cell.properties?.propertyChange?.revision.id === revisionId) {
              found = true;
              (cell as Mutable<TableCell>).properties = settleChange(
                cell.properties,
                mode,
                OUTSIDE_CHANGE.cell
              );
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
        replaceTableRows(block as Mutable<Table>, newRows, false);
      } else if (block.type === "sdt") {
        const filtered = block.content.filter(
          c => "type" in c && (c.type === "paragraph" || c.type === "table" || c.type === "sdt")
        );
        processBlocks(filtered as readonly BodyContent[], false);
        if (settleParagraphMarks(block.content as unknown[], mode, revisionId) > 0) {
          found = true;
        }
        eachInlineStretch(block.content as unknown[], processParagraph);
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
    if (marks && settleParagraphMarks(blocks as unknown[], mode, revisionId) > 0) {
      found = true;
    }
  };

  processBlocks(doc.body as readonly BodyContent[]);
  if (settleDocumentSection(doc, mode, revisionId)) {
    found = true;
  }

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
