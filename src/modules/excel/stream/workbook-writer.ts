/**
 * WorkbookWriter - Node.js Streaming Excel Writer
 *
 * Extends base with file path support and file system image loading.
 */

import { readFileBytes, createWriteStream } from "@utils/fs";
import { WorksheetWriter } from "@excel/stream/worksheet-writer";
import { ImageError } from "@excel/errors";
import {
  WorkbookWriterBase,
  type WorkbookWriterOptions as BaseOptions,
  type WorkbookZipOptions,
  type ZlibOptions
} from "@excel/stream/workbook-writer.browser";
import { mediaPath } from "@excel/utils/ooxml-paths";

export type { WorkbookZipOptions, ZlibOptions };

// Node.js version also supports filename option for output
export interface WorkbookWriterOptions extends BaseOptions {
  /** If stream not specified, this field specifies the path to a file to write the XLSX workbook to */
  filename?: string;
}

// Interface for output stream
interface OutputStreamLike {
  emit(eventName: string | symbol, ...args: any[]): boolean;
  write(chunk: any): boolean | Promise<boolean>;
  end(): void;
  once(eventName: string | symbol, listener: (...args: any[]) => void): this;
  removeListener(eventName: string | symbol, listener: (...args: any[]) => void): this;
}

class WorkbookWriter extends WorkbookWriterBase<WorksheetWriter> {
  constructor(options: WorkbookWriterOptions = {}) {
    super(options, WorksheetWriter);
  }

  /**
   * Create output stream - supports filename option in Node.js
   */
  protected _createOutputStream(options: WorkbookWriterOptions): OutputStreamLike {
    if (options.filename) {
      return createWriteStream(options.filename);
    }
    return super._createOutputStream(options);
  }

  /**
   * Add media files - supports loading from file system
   */
  addMedia(): Promise<void[]> {
    return Promise.all(
      this.media.map(async medium => {
        if (medium.type === "image") {
          const filename = mediaPath(medium.name);
          // Node.js: support loading from file
          if (medium.filename) {
            const data = await readFileBytes(medium.filename);
            this._addFile(data, filename);
            return;
          }
          if (medium.buffer) {
            this._addFile(medium.buffer, filename);
            return;
          }
          if (medium.base64) {
            const content = medium.base64.substring(medium.base64.indexOf(",") + 1);
            this._addFile(content, filename, true);
            return;
          }
        }
        throw new ImageError("Unsupported media");
      })
    );
  }
}

export { WorkbookWriter };
