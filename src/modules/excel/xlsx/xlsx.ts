/**
 * XLSX - Node.js version with full functionality
 *
 * Extends XLSX with:
 * - readFile: Read from file path
 * - writeFile: Write to file path
 * - Constructor injects readFileAsync for filename-based media loading
 *
 * Inherited from XLSX:
 * - read: Read from stream
 * - write: Write to stream
 * - load: Load from buffer
 * - writeBuffer: Write to buffer (Uint8Array)
 * - addMedia: Supports buffer, base64, and filename (via readFileAsync)
 */

import { fileExists, readFileBytes, createReadStream, createWriteStream } from "@utils/fs";
import { XLSX as XLSXBase } from "@excel/xlsx/xlsx.browser";
import { ExcelFileError } from "@excel/errors";
import { Parse, type ZipEntry } from "@archive/unzip/stream";
import { Writable, pipeline } from "@stream";

class XLSX extends XLSXBase {
  constructor(workbook: any) {
    super(workbook);
    // Provide file reading capability for addMedia
    this.readFileAsync = (filename: string) => readFileBytes(filename);
  }

  private static async *iterateZipEntries(parser: Parse): AsyncGenerator<ZipEntry> {
    const queue: ZipEntry[] = [];
    let head = 0;
    let done = false;
    let error: unknown;
    let notify: (() => void) | null = null;

    const wake = () => {
      if (notify) {
        const fn = notify;
        notify = null;
        fn();
      }
    };

    parser.on("entry", (entry: ZipEntry) => {
      queue.push(entry);
      wake();
    });
    parser.once("error", (err: unknown) => {
      error = err;
      wake();
    });
    parser.once("close", () => {
      done = true;
      wake();
    });

    while (!done || head < queue.length) {
      if (error) {
        throw error;
      }
      if (head < queue.length) {
        const entry = queue[head++]!;
        // Periodically compact to avoid unbounded growth.
        if (head > 1024 && head > queue.length / 2) {
          queue.splice(0, head);
          head = 0;
        }
        yield entry;
        continue;
      }
      await new Promise<void>(resolve => {
        notify = resolve;
      });
    }
  }

  // ==========================================================================
  // Node.js specific: TRUE streaming read
  // ==========================================================================

  override async read(stream: any, options?: any): Promise<any> {
    const parser = new Parse();

    const swallowError = () => {
      // Prevent unhandled 'error' events from crashing the process.
      // Errors are surfaced via rejected promises.
    };

    // Always attach an error listener to avoid uncaught exceptions.
    parser.on("error", swallowError);
    if (stream && typeof stream.on === "function") {
      stream.on("error", swallowError);
    }

    // Pump incoming data into the ZIP parser without buffering the whole file.
    // NOTE: `Parse` is a Duplex; passing it directly to pipeline() can make
    // Node treat it like a transform and surface `Premature close` errors.
    // We instead pipeline the input stream into a Writable sink that forwards
    // chunks into the parser with backpressure.
    const sink = new Writable<any>({
      write(chunk: any, _encoding: any, callback: (error?: Error | null) => void) {
        try {
          const ok = (parser as any).write(chunk);
          if (ok) {
            callback();
          } else {
            (parser as any).once("drain", () => callback());
          }
        } catch (e) {
          callback(e as any);
        }
      },
      final(callback: (error?: Error | null) => void) {
        try {
          (parser as any).end();
          callback();
        } catch (e) {
          callback(e as any);
        }
      }
    });

    const onParserError = (err: unknown) => {
      try {
        sink.destroy(err as any);
      } catch {
        // ignore
      }
    };
    parser.on("error", onParserError);

    const pump = pipeline(stream, sink);

    const entries = (async function* () {
      for await (const entry of XLSX.iterateZipEntries(parser)) {
        entry.on("error", swallowError);
        const drain = async () => {
          const anyEntry = entry as any;
          if (anyEntry?.readableEnded || anyEntry?.destroyed) {
            return;
          }
          const draining = entry.autodrain();
          await draining.promise();
        };
        yield {
          name: entry.path,
          type: entry.type,
          stream: entry as any,
          drain
        };
      }
    })();

    try {
      const workbook = await this.loadFromZipEntries(entries as any, options);
      await pump;
      return workbook;
    } catch (err) {
      // Stop the ZIP parser so pipeline() can unwind promptly.
      try {
        parser.destroy();
      } catch {
        // ignore
      }

      // Ensure pump settles to avoid unhandled rejections.
      try {
        await pump;
      } catch {
        // ignore pump failures; the original parse error is more useful
      }
      throw err;
    } finally {
      try {
        parser.off("error", onParserError);
      } catch {
        // ignore
      }
      try {
        parser.off("error", swallowError);
      } catch {
        // ignore
      }
      if (stream && typeof stream.off === "function") {
        try {
          stream.off("error", swallowError);
        } catch {
          // ignore
        }
      }
    }
  }
  // ===========================================================================
  // Node.js specific: File operations
  // ===========================================================================

  async readFile(filename: string, options?: any): Promise<any> {
    if (!(await fileExists(filename))) {
      throw new ExcelFileError(filename, "read", "File not found");
    }
    const stream = createReadStream(filename);
    return this.read(stream, options);
  }

  writeFile(filename: string, options?: any): Promise<void> {
    const stream = createWriteStream(filename);

    return new Promise((resolve, reject) => {
      let settled = false;

      const cleanup = () => {
        stream.off("error", onError);
        stream.off("close", onClose);
      };

      const onError = (error: Error) => {
        if (!settled) {
          settled = true;
          cleanup();
          reject(error);
        }
      };

      // Wait for "close" (fd released) instead of "finish" — on Windows,
      // reading the file before the fd is closed can see truncated content.
      const onClose = () => {
        if (!settled) {
          settled = true;
          cleanup();
          resolve();
        }
      };

      stream.once("error", onError);
      stream.once("close", onClose);

      this.write(stream, options).catch(err => {
        if (!settled) {
          settled = true;
          reject(err);
        }
        // Ensure the underlying FD is closed on failure. Keep listeners
        // attached until the stream closes, otherwise the emitted 'error'
        // event can become an uncaught exception.
        stream.destroy(err);
      });
    });
  }
}

export { XLSX };
