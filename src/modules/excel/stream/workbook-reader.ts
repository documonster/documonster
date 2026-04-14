/**
 * WorkbookReader - Node.js Streaming Workbook Reader
 *
 * Extends base with file path support and temp file storage for large files.
 */

import { join } from "path";

import { HyperlinkReader } from "@excel/stream/hyperlink-reader";
import {
  WorkbookReaderBase,
  type CommonInput,
  type WorkbookReaderOptions,
  type WorksheetReadyEvent,
  WorkbookReaderOptionsSchema
} from "@excel/stream/workbook-reader.browser";
import { WorksheetReader } from "@excel/stream/worksheet-reader";
import { iterateStream } from "@excel/utils/iterate-stream";
import type { Readable } from "@stream";
import { createReadStream, createWriteStream, createTempDirSync, remove } from "@utils/fs";

// Re-export types
export type {
  WorkbookReaderOptions,
  InternalWorksheetOptions,
  SharedStringRichText,
  SharedStringValue,
  WorkbookRelationship,
  SheetMetadata,
  WorkbookModel,
  WorkbookPropertiesXform,
  ParseEventType,
  SharedStringEvent,
  WorksheetReadyEvent,
  HyperlinksEvent,
  ParseEvent
} from "@excel/stream/workbook-reader.browser";

export type NodeInput = string | CommonInput;

interface WaitingWorksheet {
  sheetNo: string;
  path: string;
  cleanup: () => void;
  writePromise: Promise<void>;
}

class WorkbookReader extends WorkbookReaderBase<
  NodeInput,
  WorksheetReader,
  HyperlinkReader,
  WaitingWorksheet
> {
  constructor(input: NodeInput, options: WorkbookReaderOptions = {}) {
    super(input as CommonInput, options, WorksheetReader, HyperlinkReader);
    this.input = input as NodeInput;
  }

  _getStream(input: NodeInput): Readable {
    if (typeof input === "string") {
      return createReadStream(input);
    }
    return super._getStream(input as CommonInput);
  }

  async _storeWaitingWorksheet(sheetNo: string, entry: any): Promise<WaitingWorksheet> {
    const tmpDir = createTempDirSync("excelts-");
    const filePath = join(tmpDir, `sheet${sheetNo}.xml`);
    const cleanup = () => {
      remove(tmpDir).catch(() => {});
    };

    const maxBytes = this._maxBufferedBytes;

    const writePromise = new Promise<void>((resolve, reject) => {
      const tempStream = createWriteStream(filePath);
      tempStream.on("error", reject);
      tempStream.on("finish", resolve);

      // Track bytes written to detect oversized waiting worksheets.
      // Use an arrow function to capture `this` for cross-sheet accumulation.
      const originalWrite = tempStream.write.bind(tempStream);
      const trackWrite = (chunk: any, ...args: any[]): boolean => {
        const size = chunk instanceof Uint8Array ? chunk.length : Buffer.byteLength(chunk);
        this._totalBufferedBytes += size;
        if (this._totalBufferedBytes > maxBytes) {
          const err = new Error(
            `Buffered worksheet temp data exceeds limit of ${maxBytes} bytes. ` +
              "The XLSX file may be malicious (adversarial ZIP entry ordering) or too large " +
              "for streaming. Increase maxBufferedWorksheetBytes if this is expected."
          );
          tempStream.destroy(err);
          reject(err);
          return false;
        }
        return originalWrite(chunk, ...args);
      };
      tempStream.write = trackWrite as typeof tempStream.write;

      entry.pipe(tempStream);
    });

    return { sheetNo, path: filePath, cleanup, writePromise };
  }

  async *_processWaitingWorksheets(
    waitingWorksheets: WaitingWorksheet[]
  ): AsyncIterableIterator<WorksheetReadyEvent<WorksheetReader>> {
    for (const ws of waitingWorksheets) {
      await ws.writePromise;
      const fileStream = createReadStream(ws.path);
      try {
        yield* this._parseWorksheet(iterateStream(fileStream), ws.sheetNo);
      } finally {
        fileStream.close();
        ws.cleanup();
      }
    }
  }
}

export { WorkbookReader, WorkbookReaderOptionsSchema };
