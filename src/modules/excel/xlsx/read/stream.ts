/**
 * Read a workbook from a stream — Node variant.
 *
 * Node has a streaming ZIP parser, so entries are parsed as they arrive instead of the whole package
 * being buffered first, which is what the browser variant (`read/stream.browser.ts`) has to do.
 */
import type { ZipEntry } from "@archive/unzip/stream";
import { Parse } from "@archive/unzip/stream";
import type { Workbook } from "@excel/core/workbook";
import {
  CompoundFileSniffer,
  containerReadError,
  createContainerBudget,
  decryptionOf
} from "@excel/utils/container-errors";
import { readXlsxEntriesInto } from "@excel/xlsx/read/package";
import type { XlsxReadOptions } from "@excel/xlsx/types";
import { Writable, pipeline } from "@stream";
import type { ReadableLike } from "@stream/types";
import { toError } from "@utils/errors";

export async function readXlsxStreamInto(
  workbook: Workbook,
  stream: ReadableLike,
  options?: XlsxReadOptions
): Promise<Workbook> {
  const parser = new Parse({ checkCrc32: options?.zip?.checkCrc32 });

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
  // An OLE compound file — an encrypted workbook, a legacy .xls — is kept away from the parser: once read, an encrypted
  // workbook is decrypted with `options.password` and its package parsed, and anything else is reported as what it is
  // (`ExcelEncryptedError`, a legacy .xls), not as a ZIP without a central directory.
  const sniffer = new CompoundFileSniffer(decryptionOf(options));
  const forward = (
    chunks: (Uint8Array | string)[],
    callback: (error?: Error | null) => void
  ): void => {
    let ok = true;
    for (const chunk of chunks) {
      ok = parser.write(chunk) && ok;
    }
    if (ok) {
      callback();
    } else {
      parser.once("drain", () => callback());
    }
  };
  const sink = new Writable({
    write(chunk: Uint8Array | string, _encoding: string, callback: (error?: Error | null) => void) {
      try {
        forward(sniffer.push(chunk), callback);
      } catch (e) {
        callback(toError(e));
      }
    },
    final(callback: (error?: Error | null) => void) {
      sniffer.end().then(
        chunks => {
          try {
            for (const chunk of chunks) {
              parser.write(chunk);
            }
            parser.end();
            callback();
          } catch (e) {
            callback(toError(e));
          }
        },
        (e: unknown) => callback(toError(e))
      );
    }
  });

  const onParserError = (err: unknown) => {
    try {
      sink.destroy(toError(err));
    } catch {
      // ignore
    }
  };
  parser.on("error", onParserError);

  const pump = pipeline(stream, sink);
  // A pipeline that fails before the parser has finished — a source that errors, a destroyed input — must stop the
  // entry iteration too: the parser never ends, so the read waited for entries that could not come and the failure
  // surfaced only as an unhandled rejection.
  let pumpError: unknown;
  pump.catch((err: unknown) => {
    pumpError = err;
    try {
      parser.destroy(toError(err));
    } catch {
      // ignore
    }
  });

  const budget = createContainerBudget(options?.zip);
  const entries = (async function* () {
    for await (const entry of iterateZipEntries(parser)) {
      entry.on("error", swallowError);
      // The container limits a buffered read applies, held as the package streams: parts counted as they
      // arrive, bytes as they inflate — parsed or drained, every byte goes through the part's meter.
      const meter = budget.enter(entry.path);
      const drain = async () => {
        if (entry.readableEnded || entry.destroyed) {
          return;
        }
        for await (const chunk of entry as AsyncIterable<Uint8Array | string>) {
          meter(chunk);
        }
      };
      yield {
        name: entry.path,
        type: entry.type,
        stream: metered(entry, meter),
        drain
      };
    }
  })();

  try {
    const result = await readXlsxEntriesInto(workbook, entries, options);
    await pump;
    return result;
  } catch (err) {
    // Taken before the parser is stopped below, which fails the pipeline in turn.
    const sourceError = pumpError;
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
      // Settled here so it is never unhandled; reported below when it is what stopped the parser.
    }
    // A failure of the source is the cause of whatever the parser then reported about the truncated input. A
    // package that is damaged or not one is reported as the buffered read reports it.
    throw sourceError ?? containerReadError(err);
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

async function* metered(
  source: AsyncIterable<Uint8Array | string>,
  meter: (chunk: Uint8Array | string) => void
): AsyncGenerator<Uint8Array | string> {
  for await (const chunk of source) {
    meter(chunk);
    yield chunk;
  }
}

async function* iterateZipEntries(parser: Parse): AsyncGenerator<ZipEntry> {
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
      throw toError(error);
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
  // The parser reports a failure and then closes; both can land before this resumes, and the failure must not be
  // mistaken for the end of a complete package.
  if (error) {
    throw toError(error);
  }
}
