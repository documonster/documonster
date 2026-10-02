/**
 * The streaming ZIP adapter every package part is written through, and the two helpers that put a part
 * into it. Shared by the XLSX writer, its chart writer and the XLSB writer.
 */
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import { StreamingZip, ZipDeflateFile } from "@archive/zip/stream";
import type { XlsxWritable } from "@excel/core/xlsx-io-types";
import { ExcelStreamStateError } from "@excel/errors";
import type { IZipWriter, StreamListener, ZipWriterOptions } from "@excel/xlsx/types";
import { base64ToUint8Array } from "@utils/utils";
import { XmlStreamWriter } from "@xml/stream-writer";

/**
 * How much text a streaming zip entry accumulates before encoding it to UTF-8.
 *
 * Measured in **characters**, not bytes, because that is what is known before encoding — 65,536 of them are
 * 64 KB of ASCII but up to 192 KB of CJK. The deflater downstream batches at 64 KB of *bytes*, so an ASCII
 * batch lands as exactly one of its batches instead of one more scrap to concatenate, and a CJK batch lands
 * as three. Either way the buffering is bounded by a constant.
 *
 * The size is not arbitrary. The encoding win saturates around 4,096 characters, but the whole-batch handoff
 * keeps paying past that: measured on a 2,000×36 table, 4,096 costs 85 ms against 73 ms at both 16,384 and
 * 65,536. 65,536 also compresses best — a 20,000-row CJK sheet comes out 0.5% smaller than at 4,096, because
 * a larger batch gives zlib more dictionary to work with.
 */
const ENTRY_TEXT_BATCH_CHARS = 65536;

class StreamingZipWriterAdapter implements IZipWriter {
  private static textEncoder = new TextEncoder();

  private readonly zip: StreamingZip;
  private readonly events: Map<string, Set<StreamListener>> = new Map();
  private pipedStream: Pick<XlsxWritable, "write" | "end"> | null = null;
  private level: number;
  private modTime: Date | undefined;
  private timestamps: ZipTimestampMode | undefined;
  private finalized = false;

  // Backpressure tracking
  private _needsDrain = false;
  private _drainResolvers: Array<() => void> = [];
  // Count of in-flight async write() calls whose backpressure result is unknown.
  // waitForDrain() must wait for this to reach 0 before checking _needsDrain.
  private _pendingWrites = 0;
  private _pendingWriteResolvers: Array<() => void> = [];
  private _drainGeneration = 0;
  private _sinkTerminated = false;
  private _sinkError: Error | null = null;
  // Compressed output that has been requested but has not reached the sink yet.
  // Browser deflate is asynchronous, so an entry's bytes are handed to
  // `pipedStream.write()` after the call that produced them has returned.
  // `waitForDrain()` joins these first, otherwise it would sample `_needsDrain`
  // before the sink ever saw the data and read a stale "no backpressure".
  private readonly _pendingOutput = new Set<Promise<void>>();

  // Buffer errors that occur before finalizeZip registers its error listener,
  // so async compression errors during writeXlsxPackage() are never silently lost.
  private _earlyError: Error | null = null;

  constructor(options?: ZipWriterOptions) {
    this.level = options?.level ?? 6;
    this.modTime = options?.modTime;
    this.timestamps = options?.timestamps;
    this.zip = new StreamingZip((err: Error | null, data: Uint8Array, final: boolean) => {
      if (err) {
        this._emit("error", err);
        return;
      }

      if (data && data.length > 0) {
        this._emit("data", data);
        if (this.pipedStream) {
          this._checkBackpressure(this.pipedStream.write(data));
        }
      }

      if (final) {
        if (this.pipedStream) {
          this.pipedStream.end();
        }
        this._emit("finish");
      }
    });
  }

  private _emit(event: string, ...args: any[]): void {
    // Buffer error events that fire before any listener is registered,
    // so finalizeZip() can surface them even if it registers late.
    if (event === "error") {
      const callbacks = this.events.get(event);
      if (!callbacks || callbacks.size === 0) {
        this._earlyError = args[0] instanceof Error ? args[0] : new Error(String(args[0]));
        return;
      }
    }
    const callbacks = this.events.get(event);
    if (!callbacks) {
      return;
    }
    for (const cb of callbacks) {
      cb(...args);
    }
  }

  /**
   * Handle backpressure from pipedStream.write().
   * Accepts both sync (boolean) and async (Promise<boolean>) return values.
   */
  private _checkBackpressure(ok: boolean | void | Promise<boolean>): void {
    if (ok && typeof (ok as PromiseLike<boolean>).then === "function") {
      this._pendingWrites++;
      const generation = this._drainGeneration;
      Promise.resolve(ok)
        .then(
          result => {
            if (!result && generation === this._drainGeneration) {
              this._needsDrain = true;
            }
          },
          () => {} // write errors surface via the stream's error event
        )
        .finally(() => {
          this._pendingWrites--;
          if (this._pendingWrites === 0) {
            const resolvers = this._pendingWriteResolvers.splice(0);
            for (const resolve of resolvers) {
              resolve();
            }
          }
        });
      return;
    }
    if (ok === false) {
      this._needsDrain = true;
    }
  }

  on(event: string, callback: StreamListener): this {
    const callbacks = this.events.get(event) || new Set<StreamListener>();
    callbacks.add(callback);
    this.events.set(event, callbacks);

    // If an error was buffered before any listener was registered, deliver it now.
    if (event === "error" && this._earlyError) {
      const err = this._earlyError;
      this._earlyError = null;
      callback(err);
    }

    return this;
  }

  once(event: string, callback: StreamListener): this {
    const wrapped: StreamListener = (...args: any[]) => {
      this.off(event, wrapped);
      callback(...args);
    };
    return this.on(event, wrapped);
  }

  off(event: string, callback: StreamListener): this {
    const callbacks = this.events.get(event);
    if (!callbacks) {
      return this;
    }
    callbacks.delete(callback);
    if (callbacks.size === 0) {
      this.events.delete(event);
    }
    return this;
  }

  pipe(stream: XlsxWritable): void {
    this.pipedStream = stream;
    // Listen for drain events to resolve backpressure waiters
    if (stream && typeof stream.on === "function") {
      stream.on("drain", () => {
        this._drainGeneration++;
        this._needsDrain = false;
        const resolvers = this._drainResolvers.splice(0);
        for (const resolve of resolvers) {
          resolve();
        }
      });
      // Forward sink errors to the zip pipeline. Without this, a user sink
      // that errors mid-write (write failure, EPIPE, abort, etc) would leave
      // `finalizeZip()` hanging forever — `writeXlsxPackage` keeps writing into a
      // dead sink, but `finalizeZip`'s 'finish'/'error' listeners on the zip
      // adapter never fire because nothing tells the adapter the sink died.
      stream.on("error", (err: Error) => {
        this._sinkTerminated = true;
        this._sinkError = err;
        this._emit("error", err);
        // Wake any backpressure waiters so writeXlsxPackage's `await zip.waitForDrain()`
        // returns instead of hanging forever for a 'drain' that will never come.
        this._needsDrain = false;
        const resolvers = this._drainResolvers.splice(0);
        for (const resolve of resolvers) {
          resolve();
        }
        const asyncResolvers = this._pendingWriteResolvers.splice(0);
        for (const resolve of asyncResolvers) {
          resolve();
        }
      });
      stream.on("close", () => {
        if (this.finalized || this._sinkTerminated) {
          return;
        }
        this._sinkTerminated = true;
        this._sinkError = new ExcelStreamStateError(
          "write",
          "destination closed before XLSX serialization finished"
        );
        this._emit("error", this._sinkError);
        this._needsDrain = false;
        const resolvers = this._drainResolvers.splice(0);
        for (const resolve of resolvers) {
          resolve();
        }
        const asyncResolvers = this._pendingWriteResolvers.splice(0);
        for (const resolve of asyncResolvers) {
          resolve();
        }
      });
    }
  }

  /**
   * Track one entry's compressed output until it has been handed to the sink.
   * Only settlement matters here: compression failures travel through the
   * `'error'` channel (`StreamingZip` reports them via the adapter callback),
   * and swallowing the rejection keeps them from surfacing twice.
   */
  private _trackOutput(output: Promise<void>): void {
    const settled = output.then(
      () => {},
      () => {}
    );
    this._pendingOutput.add(settled);
    void settled.then(() => {
      this._pendingOutput.delete(settled);
    });
  }

  /**
   * Wait for the downstream writable to drain if it signaled backpressure.
   * If any write() calls are still in-flight (returned a Promise that hasn't
   * settled), waits for all of them first so the backpressure signal isn't missed.
   */
  async waitForDrain(): Promise<void> {
    // Let every byte produced so far reach the sink before asking whether it
    // wants us to pause; see `_pendingOutput`.
    if (this._pendingOutput.size > 0) {
      await Promise.all([...this._pendingOutput]);
    }
    // Wait for all in-flight async writes to settle so _needsDrain is up to date.
    if (this._pendingWrites > 0) {
      await new Promise<void>(resolve => {
        this._pendingWriteResolvers.push(resolve);
      });
    }
    if (this._sinkError) {
      throw this._sinkError;
    }
    if (!this._needsDrain || !this.pipedStream) {
      return;
    }
    return new Promise<void>(resolve => {
      this._drainResolvers.push(resolve);
    });
  }

  append(data: any, options: { name: string; base64?: boolean }): void {
    if (this.finalized) {
      throw new ExcelStreamStateError("append", "stream already finalized");
    }

    let buffer: Uint8Array;
    if (options.base64) {
      buffer = base64ToUint8Array(typeof data === "string" ? data : String(data));
    } else if (typeof data === "string") {
      buffer = StreamingZipWriterAdapter.textEncoder.encode(data);
    } else if (data instanceof Uint8Array) {
      buffer = data;
    } else if (ArrayBuffer.isView(data)) {
      buffer = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    } else if (data instanceof ArrayBuffer) {
      buffer = new Uint8Array(data);
    } else {
      buffer = data;
    }

    const file = new ZipDeflateFile(options.name, {
      level: this.level,
      modTime: this.modTime,
      timestamps: this.timestamps
    });
    this.zip.add(file);

    this._trackOutput(file.push(buffer, true));
  }

  createEntry(name: string): { write(chunk: Uint8Array | string): void; end(): void } {
    if (this.finalized) {
      throw new ExcelStreamStateError("createEntry", "stream already finalized");
    }
    const file = new ZipDeflateFile(name, {
      level: this.level,
      modTime: this.modTime,
      timestamps: this.timestamps
    });
    this.zip.add(file);
    const encoder = StreamingZipWriterAdapter.textEncoder;
    // Text is concatenated until it is worth encoding. An `XmlStreamWriter` hands over one string per tag —
    // three per worksheet cell — and `TextEncoder.encode` costs far more per call than per character, so a
    // large table used to spend a fifth of its write time inside the encoder. Concatenating first and encoding
    // once per batch turns hundreds of thousands of calls into a handful, and hands the deflater whole batches
    // instead of a Buffer.concat over thousands of scraps.
    let pending = "";
    let closed = false;
    const flushText = (): void => {
      if (pending.length === 0) {
        return;
      }
      const bytes = encoder.encode(pending);
      pending = "";
      file.push(bytes);
    };
    return {
      write(chunk: Uint8Array | string): void {
        // Batching makes this check load-bearing. A write after `end()` used to reach a finalized
        // `ZipDeflateFile` and reject; buffered, a short one would sit in `pending` and never be flushed, so
        // the caller's bytes would vanish without a word. Silently losing part of a part is worse than either,
        // and refusing here is also what `ArchiveSink.open` already does for the buffered sink.
        if (closed) {
          throw new ExcelStreamStateError("write to zip entry", `part ${name} is already closed`);
        }
        // Bytes go through untouched. Encoding them as text would re-encode everything above 0x7F, which is
        // fine for the XML parts this used to carry exclusively and destroys a BIFF12 part. They must still
        // flush any buffered text first, or a mixed entry would come out in the wrong order.
        if (typeof chunk !== "string") {
          flushText();
          file.push(chunk);
          return;
        }
        pending += chunk;
        if (pending.length >= ENTRY_TEXT_BATCH_CHARS) {
          flushText();
        }
      },
      end: (): void => {
        // Idempotent, like `ArchiveSink.open`: a second `end()` was already harmless before batching, so this
        // keeps it that way rather than turning a redundant call into a new failure.
        if (closed) {
          return;
        }
        closed = true;
        flushText();
        this._trackOutput(file.push(new Uint8Array(0), true));
      }
    };
  }

  finalize(): void {
    if (this.finalized) {
      return;
    }
    this.finalized = true;
    this.zip.end();
  }
}

/**
 * Helper: render an xform directly to a streaming zip entry.
 * Avoids buffering the entire XML string in memory.
 * Awaits backpressure drain after each entry to respect downstream flow control.
 */
export async function renderToZip(
  zip: IZipWriter,
  path: string,
  xform: { render(xmlStream: any, model?: any): void },
  model?: any
): Promise<void> {
  const entry = zip.createEntry(path);
  const stream = new XmlStreamWriter(entry);
  xform.render(stream, model);
  entry.end();
  // Respect downstream backpressure between entries
  await zip.waitForDrain();
}

/**
 * Write one already-materialised part, then give the sink a chance to push
 * back. Every part goes through here (or `renderToZip`) so peak buffering
 * stays near a single entry instead of growing with the package: a forgotten
 * checkpoint would not corrupt anything, it would just let that part's bytes
 * queue up unthrottled.
 */
export async function appendToZip(
  zip: IZipWriter,
  data: string | Uint8Array,
  options: { name: string; base64?: boolean }
): Promise<void> {
  zip.append(data, options);
  await zip.waitForDrain();
}

/**
 * A streaming zip writer, for a caller outside this module.
 *
 * A factory rather than the class, so `StreamingZipWriterAdapter` stays private: what a caller needs is an
 * `IZipWriter`, and publishing the class would publish its constructor options and its internals along with it.
 *
 * The binary path uses this too. A ZIP is a ZIP — the same backpressure, the same entry lifecycle, the same
 * finalisation — and a second adapter for XLSB would be a second set of those to get wrong.
 */
export function createZipWriterAdapter(options?: ZipWriterOptions): IZipWriter {
  return new StreamingZipWriterAdapter(options);
}
