/**
 * WorksheetReader - Cross-Platform Streaming Worksheet Reader
 *
 * Works in both Node.js and Browser.
 */

import { cellGetValue, cellSetModel, cellSetValue } from "@excel/core/cell";
import type { CellData, CellModel } from "@excel/core/cell";
import type { ColumnData } from "@excel/core/column";
import { Enums } from "@excel/core/enums";
import type { RangeData } from "@excel/core/range";
import { rangeCreate, rangeExpandRow } from "@excel/core/range";
import type { RowData } from "@excel/core/row";
import { rowCreate, rowDimensions } from "@excel/core/row";
import { columnCreate, columnFromModel, rowGetCell } from "@excel/core/worksheet";
import type { Worksheet } from "@excel/core/worksheet";
import { ExcelError, ExcelStreamStateError } from "@excel/errors";
import type { InternalWorksheetOptions } from "@excel/stream/workbook-reader.browser";
import type { WorksheetState, Style } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import { slideFormula } from "@excel/utils/shared-formula";
import type { SharedStringValue } from "@excel/utils/shared-strings";
import { readCellColumn, readRowNumber } from "@excel/xlsx/xform/sheet/cell-reference";
import { CellXform, formulaResultAsValue } from "@excel/xlsx/xform/sheet/cell-xform";
import { EventEmitter } from "@utils/event-emitter";
import { SaxParser } from "@xml/sax";
import type { SaxTag } from "@xml/types";

// ============================================================================
// Internal Types
// ============================================================================

/** Column model from parsed XML */
interface ParsedColumnModel {
  min: number;
  max: number;
  width: number;
  styleId: number;
}

/** Hyperlink reference from worksheet XML */
export interface WorksheetHyperlink {
  ref: string;
  rId?: string;
  target?: string;
}

/** Events emitted during worksheet parsing */
export type WorksheetEventType = RowEvent["eventType"] | HyperlinkEvent["eventType"];

/** Row event emitted during parsing */
export interface RowEvent {
  eventType: "row";
  value: RowData;
}

/** Hyperlink event emitted during parsing */
export interface HyperlinkEvent {
  eventType: "hyperlink";
  value: WorksheetHyperlink;
}

export type WorksheetEvent = RowEvent | HyperlinkEvent;

// ============================================================================
// Public Types
// ============================================================================

/** The subset of the streaming workbook reader a worksheet reader consumes. */
export interface WorksheetReaderWorkbook {
  sharedStrings?: SharedStringValue[];
  styles: { getStyleModel(id: number): Style | null };
  properties?: { model?: { date1904?: boolean } };
  dynamicArrayCmIndices?: Set<number>;
  hasDynamicArrayMetadata?: boolean;
}

export interface WorksheetReaderOptions {
  workbook: WorksheetReaderWorkbook;
  id: number;
  iterator: AsyncIterable<unknown>;
  options?: InternalWorksheetOptions;
}

/**
 * Put a parsed `<c>` onto a streamed cell, as `rowSetModel` does for a loaded one, but keeping the cell's style
 * object its own and mutable — the streaming handle has always been a plain record its consumer may edit, where a
 * loaded cell shares a frozen container and is written through `cellOwnStyle`.
 */
function applyCellModel(cell: CellData, model: CellModel, style: Partial<Style> | undefined): void {
  // Without a style of its own the cell keeps what it inherited from its row or column, whose facets may be shared
  // snapshots — so their sharing flag is kept too, and the first write still copies them first.
  const inherited = cell.style;
  const shared = cell._sharedStyle;
  model.style = undefined;
  cellSetModel(cell, model);
  if (style) {
    cell.style = copyStyle(style) ?? {};
    cell._sharedStyle = false;
  } else {
    cell.style = inherited;
    cell._sharedStyle = shared;
  }
}

class WorksheetReader extends EventEmitter {
  workbook: WorksheetReaderWorkbook;
  id: number | string;
  sheetNo: number;
  iterator: AsyncIterable<unknown>;
  options: InternalWorksheetOptions;
  name: string;
  state?: WorksheetState;
  declare private _columns: ColumnData[];
  declare private _keys: Record<string, ColumnData>;
  declare private _dimensions: RangeData;
  hyperlinks?: Record<string, WorksheetHyperlink>;

  constructor({ workbook, id, iterator, options }: WorksheetReaderOptions) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.sheetNo = typeof id === "number" ? id : parseInt(String(id), 10);
    this.iterator = iterator;
    this.options = options || {};

    // and a name
    this.name = `Sheet${this.id}`;

    // column definitions
    this._columns = [];
    this._keys = Object.create(null) as Record<string, ColumnData>;

    // keep a record of dimensions
    this._dimensions = rangeCreate();
  }

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy(): void {
    throw new ExcelStreamStateError("destroy", "Invalid operation for a streaming reader");
  }

  // return the current dimensions of the reader
  get dimensions(): RangeData {
    return this._dimensions;
  }

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns(): ColumnData[] {
    return this._columns;
  }

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn(c: string | number): ColumnData {
    if (typeof c === "string") {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) {
        return col;
      }

      // otherwise, assume letter
      c = colCache.l2n(c);
    }
    if (c > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= c) {
        // The reader structurally masquerades as a Worksheet for the column/row
        // factories (it implements the subset they touch); a precise type would
        // require the reader to implement the full Worksheet surface.
        this._columns.push(columnCreate(this as unknown as Worksheet, n++));
      }
    }
    return this._columns[c - 1];
  }

  getColumnKey(key: string): ColumnData | undefined {
    return this._keys[key];
  }

  setColumnKey(key: string, value: ColumnData): void {
    this._keys[key] = value;
  }

  deleteColumnKey(key: string): void {
    delete this._keys[key];
  }

  eachColumnKey(f: (column: ColumnData, key: string) => void): void {
    const keys = this._keys;
    for (const key in keys) {
      f(keys[key], key);
    }
  }

  async read(): Promise<void> {
    try {
      for await (const events of this.parse()) {
        for (let i = 0; i < events.length; i++) {
          const event = events[i]!;
          this.emit(event.eventType, event.value);
        }
      }
      this.emit("finished");
    } catch (error) {
      this.emit("error", error);
    }
  }

  async *[Symbol.asyncIterator](): AsyncIterableIterator<RowData> {
    for await (const events of this.parse()) {
      for (let i = 0; i < events.length; i++) {
        const event = events[i]!;
        if (event.eventType === "row") {
          yield event.value;
        }
      }
    }
  }

  async *parse(): AsyncIterableIterator<WorksheetEvent[]> {
    const { iterator, options } = this;
    let emitSheet = false;
    let emitHyperlinks = false;
    let hyperlinks: Record<string, WorksheetHyperlink> | null = null;
    switch (options.worksheets) {
      case "emit":
        emitSheet = true;
        break;
      case "prep":
        break;
      default:
        break;
    }
    switch (options.hyperlinks) {
      case "emit":
        emitHyperlinks = true;
        break;
      case "cache":
        this.hyperlinks = hyperlinks = Object.create(null) as Record<string, WorksheetHyperlink>;
        break;
      default:
        break;
    }
    if (!emitSheet && !emitHyperlinks && !hyperlinks) {
      return;
    }

    const shouldHandleHyperlinks = emitHyperlinks || hyperlinks !== null;

    // references
    const { sharedStrings, styles, properties } = this.workbook;

    // xml position
    let inCols = false;
    let inRows = false;
    let inHyperlinks = false;

    // parse state
    let cols: ParsedColumnModel[] | null = null;
    let row: RowData | null = null;
    // Each `<c>` is parsed and reconciled by the same `CellXform` the full reader uses, so the two agree on values.
    const cellXform = new CellXform();
    /** Column of the `<c>` being read, or 0 when none is open. */
    let cellCol = 0;
    /** Groups a dependent asked for before any master was seen — see `resolveSharedFormula`. */
    const orphanedGroups = new Map<string, string>();
    const { workbook } = this;
    const date1904 = properties?.model?.date1904;
    const reconcileOptions = {
      styles,
      date1904,
      sharedStrings: sharedStrings && { getString: (index: number) => sharedStrings[index] },
      // Group index → master address and formula text, recorded by `reconcile` as each master is read.
      formulae: Object.create(null) as Record<string, string>,
      sharedFormulaText: Object.create(null) as Record<string, string>,
      // Read per cell, as before: a metadata part may arrive after the reader was created.
      get dynamicArrayCmIndices() {
        return workbook.dynamicArrayCmIndices;
      },
      get hasDynamicArrayMetadata() {
        return workbook.hasDynamicArrayMetadata;
      }
    };
    // An omitted `<row r>` / `<c r>` follows the previous one — see `cell-reference.ts`.
    let lastRowNumber = 0;
    let lastCellCol = 0;

    // Direct SAX callback mode — zero intermediate event objects.
    // We collect worksheet events per-chunk and yield them.
    let worksheetEvents: WorksheetEvent[] | null = null;

    const finishCell = (model: CellModel & Record<string, any>, col: number): void => {
      const cell = rowGetCell(row!, col);
      // The address the cell was filed under, not the attribute: it may be absent, or spelled `$C$1`.
      model.address = cell.address;
      if (model.type === Enums.ValueType.Merge) {
        // `<c r="B1"/>` — no value and no style. The cell exists, as it always did here, and holds nothing.
        return;
      }
      const styleId = model.styleId as number | undefined;
      let unresolvedSharedString: number | undefined;
      if (
        !sharedStrings &&
        model.type === Enums.ValueType.String &&
        typeof model.value === "number"
      ) {
        // Shared strings not cached: the index is all there is, handed over as `{ sharedString }`.
        unresolvedSharedString = model.value;
        model.value = undefined;
        model.type = Enums.ValueType.Null;
      }
      // `reconcile` consumes the group index, so it is read first.
      const si = model.si as string | undefined;
      cellXform.reconcile(model, reconcileOptions);
      if (si !== undefined && model.type === Enums.ValueType.Formula) {
        resolveSharedFormula(model, si);
      }
      applyCellModel(
        cell,
        model,
        styleId !== undefined ? (model.style as Partial<Style>) : undefined
      );
      if (unresolvedSharedString !== undefined) {
        cellSetValue(cell, { sharedString: unresolvedSharedString } as never);
      }
      // `<hyperlinks>` follows `<sheetData>` in a conforming sheet, so this only fires for a producer that writes it
      // first; otherwise the links reach the caller through `hyperlinks` once the sheet has been read.
      const hyperlink = hyperlinks?.[cell.address];
      if (hyperlink) {
        // Streaming-specific: stash the cell's value as `text` and attach the hyperlink. These fields are not part
        // of the standard CellData.
        const streamingCell = cell as typeof cell & {
          text?: ReturnType<typeof cellGetValue>;
          hyperlink?: WorksheetHyperlink;
        };
        streamingCell.text = cellGetValue(cell);
        cellSetValue(cell, undefined);
        streamingCell.hyperlink = hyperlink;
      }
    };

    // A shared-formula dependent gets the master's formula slid to its own address now, since the master is gone by
    // the time a caller would ask. The master comes first in a conforming sheet. A dependent whose master never
    // appears keeps its cached result — the file holds no formula text for it. A master arriving after a dependent
    // already emitted is an error: that dependent's formula can no longer be supplied. Groups live for the whole
    // sheet because a dependent may lie outside its master's declared range.
    const resolveSharedFormula = (model: CellModel & Record<string, any>, si: string): void => {
      if (model.formula) {
        const orphan = orphanedGroups.get(si);
        if (orphan !== undefined) {
          throw new ExcelError(
            `Shared formula ${si}: cell ${orphan} comes before its master ${model.address}, so its formula cannot be ` +
              "recovered by a forward-only reader. Read this workbook with Workbook.read."
          );
        }
        return;
      }
      const text = reconcileOptions.sharedFormulaText[si];
      if (model.sharedFormula && text) {
        model.formula = slideFormula(text, model.sharedFormula, model.address);
        return;
      }
      if (!orphanedGroups.has(si)) {
        orphanedGroups.set(si, model.address);
      }
      formulaResultAsValue(model);
    };

    const parser = new SaxParser({ position: false, invalidCharHandling: "skip" });

    parser.on("opentag", (node: SaxTag) => {
      if (cellCol !== 0) {
        cellXform.parseOpen(node);
        return;
      }
      if (emitSheet) {
        switch (node.name) {
          case "cols":
            inCols = true;
            cols = [];
            break;
          case "sheetData":
            inRows = true;
            break;

          case "col":
            if (inCols) {
              cols!.push({
                min: parseInt(node.attributes.min, 10),
                max: parseInt(node.attributes.max, 10),
                width: parseFloat(node.attributes.width),
                styleId: parseInt(node.attributes.style ?? "0", 10)
              });
            }
            break;

          case "row":
            if (inRows) {
              const r = readRowNumber(node.attributes.r, lastRowNumber);
              lastRowNumber = r;
              lastCellCol = 0;
              row = rowCreate(this as unknown as Worksheet, r);
              if (node.attributes.ht) {
                row.height = parseFloat(node.attributes.ht);
              }
              if (node.attributes.customHeight === "1") {
                row.customHeight = true;
              }
              if (node.attributes.s !== undefined) {
                const styleId = parseInt(node.attributes.s, 10);
                const style = styles.getStyleModel(styleId);
                if (style) {
                  row.style = copyStyle(style) ?? {};
                }
              }
            }
            break;
          case "c":
            if (row) {
              // See `cell-reference.ts`, shared with the full reader.
              cellCol = readCellColumn(node.attributes.r, row.number, lastCellCol);
              lastCellCol = cellCol;
              cellXform.parseOpen(node);
            }
            break;
          case "mergeCell":
            break;
          default:
            break;
        }
      }

      // =================================================================
      //
      if (shouldHandleHyperlinks) {
        switch (node.name) {
          case "hyperlinks":
            inHyperlinks = true;
            break;
          case "hyperlink":
            if (inHyperlinks) {
              const loc = node.attributes.location;
              const hyperlink = {
                ref: node.attributes.ref,
                rId: node.attributes["r:id"],
                // Internal links: resolve target from location attribute
                target: loc ? (loc.startsWith("#") ? loc : `#${loc}`) : undefined
              };
              if (emitHyperlinks) {
                (worksheetEvents ||= []).push({ eventType: "hyperlink", value: hyperlink });
              } else {
                hyperlinks![hyperlink.ref] = hyperlink;
              }
            }
            break;
          default:
            break;
        }
      }
    });

    parser.on("text", (text: string) => {
      // Only a cell's text is sheet data.
      if (cellCol !== 0) {
        cellXform.parseText(text);
      }
    });

    parser.on("closetag", (tag: SaxTag) => {
      if (cellCol !== 0) {
        cellXform.parseClose(tag.name);
        // Ended by its own close tag, not by `parseClose` returning false: that also answers for a child element
        // `CellXform` does not know, which would end the cell early.
        if (tag.name === "c") {
          finishCell(cellXform.model as CellModel, cellCol);
          cellCol = 0;
        }
        return;
      }
      if (emitSheet) {
        switch (tag.name) {
          case "cols":
            inCols = false;
            this._columns = columnFromModel(this as unknown as Worksheet, cols!);
            break;
          case "sheetData":
            inRows = false;
            break;

          case "row":
            if (row) {
              rangeExpandRow(this._dimensions, {
                number: row.number,
                dimensions: rowDimensions(row) ?? undefined
              });
              (worksheetEvents ||= []).push({ eventType: "row", value: row });
            }
            row = null;
            break;

          default:
            break;
        }
      }
      if (shouldHandleHyperlinks) {
        switch (tag.name) {
          case "hyperlinks":
            inHyperlinks = false;
            break;
          default:
            break;
        }
      }
    });

    // Drive the SAX parser synchronously per chunk, yield events after each chunk.
    // SAX parser.write() is synchronous: all callbacks fire within the write() call.
    // This eliminates async queue overhead entirely.
    const decoder = new TextDecoder("utf-8", { fatal: true });

    for await (const chunk of iterator) {
      const chunkStr =
        typeof chunk === "string" ? chunk : decoder.decode(chunk as Uint8Array, { stream: true });
      parser.write(chunkStr);
      // After each chunk, flush accumulated events (callbacks set worksheetEvents synchronously)
      const batch = worksheetEvents as WorksheetEvent[] | null;
      if (batch && batch.length > 0) {
        worksheetEvents = null;
        yield batch;
      }
    }

    // Flush any trailing bytes from the streaming decoder (catches truncated UTF-8)
    const trailing = decoder.decode();
    if (trailing) {
      parser.write(trailing);
    }

    parser.close();
    // Flush any remaining events
    const finalBatch = worksheetEvents as WorksheetEvent[] | null;
    if (finalBatch && finalBatch.length > 0) {
      yield finalBatch;
    }
  }
}

export { WorksheetReader };
