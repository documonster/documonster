/**
 * WorksheetReader - Cross-Platform Streaming Worksheet Reader
 *
 * Works in both Node.js and Browser.
 */

import { cellGetValue, cellNumFmt, cellSetValue } from "@excel/core/cell";
import type { ColumnData } from "@excel/core/column";
import type { RangeData } from "@excel/core/range";
import { rangeCreate, rangeExpandRow } from "@excel/core/range";
import type { RowData } from "@excel/core/row";
import { rowCreate, rowDimensions } from "@excel/core/row";
import { columnCreate, columnFromModel, rowGetCell } from "@excel/core/worksheet";
import type { Worksheet } from "@excel/core/worksheet";
import { ExcelStreamStateError } from "@excel/errors";
import type { InternalWorksheetOptions } from "@excel/stream/workbook-reader.browser";
import type { WorksheetState, CellErrorValue, Style } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import type { SharedStringValue } from "@excel/utils/shared-strings";
import { EventEmitter } from "@utils/event-emitter";
import { isDateFmt, excelToDate, decodeOoxmlEscape } from "@utils/utils";
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

/** Cell parsing state during XML processing */
interface CellParseState {
  ref: string;
  s?: number;
  t?: string;
  /** cm attribute — cell metadata index (1-indexed), used for dynamic array formulas */
  cm?: number;
  f?: { text: string };
  v?: { text: string };
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
    let c: CellParseState | null = null;
    let current: { text: string } | null = null;

    // Direct SAX callback mode — zero intermediate event objects.
    // We collect worksheet events per-chunk and yield them.
    let worksheetEvents: WorksheetEvent[] | null = null;

    const parser = new SaxParser({ position: false, invalidCharHandling: "skip" });

    parser.on("opentag", (node: SaxTag) => {
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
              const r = parseInt(node.attributes.r, 10);
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
              const styleAttr = node.attributes.s;
              const cmAttr = node.attributes.cm;
              c = {
                ref: node.attributes.r,
                s: styleAttr !== undefined ? parseInt(styleAttr, 10) : undefined,
                t: node.attributes.t,
                cm: cmAttr !== undefined ? parseInt(cmAttr, 10) : undefined
              };
            }
            break;
          case "f":
            if (c) {
              current = c.f = { text: "" };
            }
            break;
          case "v":
            if (c) {
              current = c.v = { text: "" };
            }
            break;
          case "is":
          case "t":
            if (c) {
              current = c.v = { text: "" };
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
      // only text data is for sheet values
      if (emitSheet) {
        if (current) {
          current.text += text;
        }
      }
    });

    parser.on("closetag", (tag: SaxTag) => {
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

          case "c":
            if (row && c) {
              const address = colCache.decodeAddress(c.ref);
              const cell = rowGetCell(row, address.col);
              if (c.s !== undefined) {
                const style = styles.getStyleModel(c.s);
                if (style) {
                  cell.style = copyStyle(style) ?? {};
                }
              }

              if (c.f) {
                const cellValue: {
                  formula: string;
                  result?: string | number;
                  isDynamicArray?: boolean;
                } = {
                  formula: c.f.text
                };
                if (c.v) {
                  if (c.t === "str") {
                    cellValue.result = c.v.text;
                  } else {
                    cellValue.result = parseFloat(c.v.text);
                  }
                }
                // Check if this cell is a dynamic array formula via cm → metadata mapping.
                // Uses the precise dynamicArrayCmIndices set from WorkbookReaderBase,
                // falling back to the coarser hasDynamicArrayMetadata boolean.
                if (c.cm !== undefined) {
                  const { workbook: wb } = this;
                  if (wb.dynamicArrayCmIndices) {
                    if (wb.dynamicArrayCmIndices.has(c.cm)) {
                      cellValue.isDynamicArray = true;
                    }
                  } else if (wb.hasDynamicArrayMetadata) {
                    cellValue.isDynamicArray = true;
                  }
                }
                cellSetValue(cell, cellValue);
              } else if (c.v) {
                switch (c.t) {
                  case "s": {
                    const index = parseInt(c.v.text, 10);
                    if (sharedStrings) {
                      cellSetValue(cell, sharedStrings[index]);
                    } else {
                      // Streaming format - unresolved shared string reference
                      cellSetValue(cell, {
                        sharedString: index
                      } as never);
                    }
                    break;
                  }

                  case "inlineStr":
                    // Inline strings come from <is><t>...</t></is> which uses
                    // OOXML _xHHHH_ escaping in addition to XML entities.
                    cellSetValue(
                      cell,
                      c.v.text.includes("_x") ? decodeOoxmlEscape(c.v.text) : c.v.text
                    );
                    break;
                  case "str":
                    cellSetValue(cell, c.v.text);
                    break;

                  case "e":
                    cellSetValue(cell, { error: c.v.text as CellErrorValue["error"] });
                    break;

                  case "b":
                    cellSetValue(cell, parseInt(c.v.text, 10) !== 0);
                    break;

                  default: {
                    const numFmtValue = cellNumFmt(cell);
                    const numFmtStr =
                      typeof numFmtValue === "string" ? numFmtValue : numFmtValue?.formatCode;
                    if (numFmtStr && isDateFmt(numFmtStr)) {
                      cellSetValue(
                        cell,
                        excelToDate(parseFloat(c.v.text), properties?.model?.date1904)
                      );
                    } else {
                      cellSetValue(cell, parseFloat(c.v.text));
                    }
                    break;
                  }
                }
              }
              if (hyperlinks) {
                const hyperlink = hyperlinks[c.ref];
                if (hyperlink) {
                  // Streaming-specific: stash the cell's value as `text` and
                  // attach the hyperlink so downstream processing can pick them
                  // up. These fields are not part of the standard CellData.
                  const streamingCell = cell as typeof cell & {
                    text?: ReturnType<typeof cellGetValue>;
                    hyperlink?: WorksheetHyperlink;
                  };
                  streamingCell.text = cellGetValue(cell);
                  cellSetValue(cell, undefined);
                  streamingCell.hyperlink = hyperlink;
                }
              }
              c = null;
              current = null;
            }
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
