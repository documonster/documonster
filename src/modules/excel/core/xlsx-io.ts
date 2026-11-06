import type { Readable } from "node:stream";

/**
 * The canonical public workbook IO surface — Node variant.
 *
 * Same shape as `xlsx-io.browser.ts`, plus the Node-only file-path `readFile` /
 * `writeFile`. `readStream` reaches the Node variant of `xlsx/read/stream`, which
 * parses ZIP entries as they arrive instead of buffering the package. Selected over the
 * browser variant via the `.browser` same-name swap at build/test time.
 */
import { checkWorkbook, type WorkbookData } from "@excel/core/workbook-core";
import {
  resolveReadFormat,
  formatFromPath,
  normalizeBytes,
  readXlsbInto,
  writeBytesToStream,
  writeXlsbBytes,
  writeXlsbToStream
} from "@excel/core/workbook-format";
import {
  checkIoOptions,
  blobBytes,
  checkReadData,
  isReadTarget,
  readWorkbookStream,
  readWorkbookWithDiagnostics,
  splitCreateOptions,
  workbookDecryption
} from "@excel/core/workbook-io-types";
import { createWorkbook } from "@excel/core/workbook.browser";
import type { XlsxReadable } from "@excel/core/xlsx-io-types";
import {
  isWebWritable,
  nodeSink,
  workbookSink,
  type WorkbookWritable
} from "@excel/core/xlsx-sink";
import type { XlsxStreamOptions } from "@excel/core/xlsx-stream";
import { createXlsxByteStream } from "@excel/core/xlsx-stream";
import type { ExcelFileError } from "@excel/errors";
import { encryptWorkbookPackage, openWorkbookPackage } from "@excel/utils/container-errors";
import { checkFilePath, fileSystemError } from "@excel/utils/file-path";
import type { XlsxReadOptions } from "@excel/xlsx/types";
import { writeXlsxBytes, writeXlsxToStream } from "@excel/xlsx/write/package";
import { createReadStream, createWriteStream } from "@utils/fs";

export type {
  WorkbookDiagnosticReadOptions,
  WorkbookReadOptions,
  WorkbookReadReport,
  WorkbookWriteOptions
} from "@excel/core/workbook-io-types";

import type {
  WorkbookCreateOptions,
  WorkbookDiagnosticReadOptions,
  WorkbookReadOptions,
  WorkbookReadReport,
  WorkbookStreamInput,
  WorkbookWriteOptions
} from "@excel/core/workbook-io-types";

/**
 * The XLSX reader, loaded on demand — the boundary `workbook-format.ts` already draws for XLSB.
 *
 * What this boundary buys depends on the bundler, and it is worth being exact about. One that drops
 * unused members of a namespace (rolldown, rspack) already leaves the reader out of a `toBuffer`-only
 * bundle, static import or not — the reader is kept out of the *writer* by `xlsx/write/` never importing
 * `xlsx/read/`, and `scripts/treeshake-verify.ts` asserts that. One that keeps a namespace object whole
 * (esbuild) retains `read` with `toBuffer`, and there this `import()` is what moves the reader into a
 * chunk fetched on first read instead of on page load. Every read entry is async, so it costs nothing.
 */
function loadXlsxReader() {
  return import("@excel/xlsx/read/package");
}

// =============================================================================
// Cross-platform flat IO functions (canonical public surface, Node binding).
// =============================================================================

/**
 * Serialize a workbook to xlsx bytes.
 *
 * Resolves to a `Uint8Array` on every platform, as every other `toBuffer` in this package does. `fs.writeFile`,
 * `res.end` and SDK upload bodies take one as they take a `Buffer`, and a caller who wants Buffer methods wraps it
 * without a copy: `Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)`.
 *
 * It is a plain `Uint8Array` at runtime too, not a `Buffer` typed as one: the two disagree on `slice` (a `Buffer`'s
 * is a view, a `Uint8Array`'s a copy), so handing out a subclass under the base class's name would let the same
 * call behave differently by platform. The view shares the package's memory; nothing is copied.
 *
 * **Timing this?** Measure with `NODE_ENV=production`. Outside production (and outside
 * vitest) every write runs the OOXML self-check, which re-parses the package it just
 * produced — measured at 38 ms against 124 ms on a 100-column × 200-row table, so a
 * development-mode figure is roughly **3.3× the real one** and is not distributed
 * across the write in the way a profile suggests. Pass `{ validate: false }` to
 * suppress it explicitly; see `XlsxWriteOptions.validate`.
 *
 * With `{ password }` the package is encrypted — Agile Encryption in an OLE compound file, what Excel writes for
 * "Encrypt with Password" — and `Workbook.read(bytes, { password })` reads it back; `encryption` tunes it.
 */
export async function toBuffer(
  wb: WorkbookData,
  options?: WorkbookWriteOptions
): Promise<Uint8Array> {
  checkWorkbook(wb);
  checkIoOptions("Workbook.toBuffer", options, "write");
  // `null` is taken as no options, as the check documents.
  options ??= undefined;
  const bytes = await packageBytes(wb, options);
  return Object.getPrototypeOf(bytes) === Uint8Array.prototype
    ? bytes
    : new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/**
 * The package `toBuffer` returns: XLSX or XLSB, encrypted with `options.password` when one is given. The encryption
 * code is loaded only then.
 */
async function packageBytes(
  wb: WorkbookData,
  options: WorkbookWriteOptions | undefined
): Promise<Uint8Array> {
  const bytes =
    options?.format === "xlsb"
      ? await writeXlsbBytes(wb, options)
      : await writeXlsxBytes(wb, options);
  return encryptWorkbookPackage(bytes, options);
}

/**
 * Read a workbook from its bytes — or from base64 text with `{ base64: true }` — and return it.
 *
 * The format is detected from the package contents unless `options.format` says otherwise. An encrypted workbook is
 * read with `options.password`, as Word's `Document.read` reads an encrypted document.
 * Detection reads the ZIP central directory; a base64 string is decoded once and the bytes
 * are reused by whichever loader runs, so an XLSX read pays for one directory scan and not
 * for a second decode.
 */
export async function read(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options?: WorkbookReadOptions & WorkbookCreateOptions
): Promise<WorkbookData>;
/**
 * @deprecated Use `read(data, options)`, which returns a new workbook. This form replaces everything
 * `wb` already holds — its sheets, names and properties — with what it reads. Removed in 1.0.
 */
export async function read(
  wb: WorkbookData,
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options?: WorkbookReadOptions
): Promise<WorkbookData>;
export async function read(
  first: WorkbookData | Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  second?:
    | Uint8Array
    | ArrayBuffer
    | ArrayBufferView
    | Blob
    | string
    | (WorkbookReadOptions & WorkbookCreateOptions),
  third?: WorkbookReadOptions
): Promise<WorkbookData> {
  if (isReadTarget(first)) {
    return readInto(
      first,
      second as Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
      third
    );
  }
  const { create, read: options } = splitCreateOptions(
    second as (WorkbookReadOptions & WorkbookCreateOptions) | undefined
  );
  return readInto(createWorkbook(create), first, options);
}

async function readInto(
  wb: WorkbookData,
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options?: WorkbookReadOptions
): Promise<WorkbookData> {
  checkIoOptions("Workbook.read", options, "read");
  data = await blobBytes(data);
  checkReadData("Workbook.read", data, options?.base64);
  const raw = normalizeBytes(data, options?.base64);
  // An encrypted workbook is decrypted with `options.password` first, so its package's format is the one detected.
  const bytes = raw && (await openWorkbookPackage(raw, workbookDecryption(options)));
  const format = resolveReadFormat(bytes, options?.format, options?.zip);
  if (format === "xlsb") {
    if (!bytes) {
      // Only reachable when a caller forces `format: "xlsb"` on a string without `base64`.
      return (await loadXlsxReader()).readXlsxInto(wb, data, options);
    }
    return readXlsbInto(wb, bytes, undefined, options);
  }
  return (await loadXlsxReader()).readXlsxInto(wb, bytes ?? data, options);
}

/**
 * Read a workbook and return what could not be recovered alongside it, as `report.workbook`.
 *
 * `read` throws away the diagnostics and `{ unsupported: "error" }` turns them into a rejection; this
 * is the third combination, and the one a converter actually wants — read the file, then report.
 */
export async function readWithDiagnostics(
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options?: WorkbookDiagnosticReadOptions & WorkbookCreateOptions
): Promise<WorkbookReadReport>;
/**
 * @deprecated Use `readWithDiagnostics(data, options)`, which reads into a new workbook. This form
 * replaces everything `wb` already holds. Removed in 1.0.
 */
export async function readWithDiagnostics(
  wb: WorkbookData,
  data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  options?: WorkbookDiagnosticReadOptions
): Promise<WorkbookReadReport>;
export async function readWithDiagnostics(
  first: WorkbookData | Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string,
  second?:
    | Uint8Array
    | ArrayBuffer
    | ArrayBufferView
    | Blob
    | string
    | (WorkbookDiagnosticReadOptions & WorkbookCreateOptions),
  third?: WorkbookDiagnosticReadOptions
): Promise<WorkbookReadReport> {
  let wb: WorkbookData;
  let data: Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string;
  let options: WorkbookDiagnosticReadOptions | undefined;
  if (isReadTarget(first)) {
    wb = first;
    data = second as Uint8Array | ArrayBuffer | ArrayBufferView | Blob | string;
    options = third;
  } else {
    const split = splitCreateOptions(
      second as (WorkbookDiagnosticReadOptions & WorkbookCreateOptions) | undefined
    );
    wb = createWorkbook(split.create);
    data = first;
    options = split.read;
  }
  // The shared body — see `readWorkbookWithDiagnostics`. This was 26 lines duplicated between the two platform
  // variants, in the one file pair where every other shared piece had already been extracted.
  return readWorkbookWithDiagnostics(wb, data, options, {
    read: readInto,
    normalizeBytes,
    resolveReadFormat
  });
}

/**
 * Read a workbook from a readable stream of its bytes — a Node `Readable`, a web `ReadableStream`, or an (async)
 * iterable of chunks — and return it. An XLSX package is read as it streams. An XLSB one is read when
 * `options.format` is `"xlsb"`: a stream cannot be inspected before it is consumed, so the package is collected and
 * read as {@link read} reads bytes.
 */
export async function readStream(
  stream: WorkbookStreamInput,
  options?: WorkbookReadOptions & WorkbookCreateOptions
): Promise<WorkbookData>;
/**
 * @deprecated Use `readStream(stream, options)`, which returns a new workbook. This form replaces
 * everything `wb` already holds. Removed in 1.0.
 */
export async function readStream(
  wb: WorkbookData,
  stream: unknown,
  options?: XlsxReadOptions
): Promise<WorkbookData>;
export async function readStream(
  first: unknown,
  second?: unknown,
  third?: XlsxReadOptions
): Promise<WorkbookData> {
  if (isReadTarget(first)) {
    return readWorkbookStream(first, second, third, { read: readInto });
  }
  const { create, read } = splitCreateOptions(
    second as (WorkbookReadOptions & WorkbookCreateOptions) | undefined
  );
  return readWorkbookStream(createWorkbook(create), first, read, { read: readInto });
}

/**
 * Write a workbook to a writable stream, resolving once the sink has accepted
 * the whole package.
 *
 * Downstream backpressure is respected, so **the sink must already be consumed**
 * — or be a terminal sink such as `fs.createWriteStream()`, an HTTP response, or
 * an upload body — *before* this call. Handing over an unconsumed intermediate
 * stream (most commonly a bare `stream.PassThrough`) deadlocks: once its buffers
 * fill (64 KB per side on Node 22+, so roughly 128 KB of output) `write()`
 * returns `false`, no `'drain'` ever fires because nothing is reading, and this
 * promise never settles. Small workbooks that fit inside those buffers appear to
 * work, which makes the failure look size-dependent.
 *
 * {@link toStream} carries no such contract — prefer it whenever the consumer is
 * not already running.
 *
 * ```ts
 * // ❌ Deadlocks above ~128 KB of output: the consumer is attached only after
 * //    this promise resolves, which never happens.
 * const passThrough = new PassThrough();
 * await Workbook.writeStream(wb, passThrough);
 * await upload(passThrough);
 *
 * // ✅ Start the consumer first, then await the producer.
 * const passThrough = new PassThrough();
 * const uploading = upload(passThrough);
 * await Workbook.writeStream(wb, passThrough);
 * await uploading;
 *
 * // ✅ Or hand out a source and let the consumer pull — no ordering contract.
 * await upload(Workbook.toStream(wb));
 *
 * // ✅ Or buffer the package in memory.
 * await upload(await Workbook.toBuffer(wb));
 * ```
 *
 * A sink that errors, or that closes before serialization finishes, rejects this
 * promise instead of hanging.
 */
export async function writeStream(
  wb: WorkbookData,
  stream: WorkbookWritable,
  options?: WorkbookWriteOptions
): Promise<void> {
  checkWorkbook(wb);
  checkIoOptions("Workbook.writeStream", options, "write");
  // `null` is taken as no options, as the check documents.
  options ??= undefined;
  // A Web `WritableStream` is driven as the Node-style writable the serializers expect; anything else is refused here.
  // A Node-style writable is checked and used in this tick, so a sink that fails straight away is observed; only a
  // Web stream waits for the adapter to load.
  const { sink, finished } = isWebWritable(stream)
    ? await workbookSink("Workbook.writeStream", stream)
    : nodeSink("Workbook.writeStream", stream);
  if (options?.password !== undefined) {
    // An encrypted package is one encrypted stream with its size up front: assembled, encrypted, then written.
    await writeBytesToStream(sink, await packageBytes(wb, options));
  } else if (options?.format === "xlsb") {
    // **Streamed part by part**, through the same `writeXlsbPackage` the buffered path uses. This used to
    // assemble the whole package and hand it over in one write, on the stated grounds that "there is nothing
    // to stream incrementally" — which was wrong: `sharedStrings.bin` is the ninth part produced for a
    // three-sheet workbook, and the eight before it are finished bytes.
    await writeXlsbToStream(wb, sink, options);
  } else {
    await writeXlsxToStream(wb, sink, options);
  }
  await finished;
}

/**
 * Serialize a workbook into a readable byte stream.
 *
 * The consumer drives production: serialization starts on the first read and
 * pauses whenever the reader stops pulling. Unlike {@link writeStream} there is
 * no ordering contract to get wrong, so an unconsumed destination cannot park
 * the writer forever.
 *
 * ```ts
 * await pipeline(Workbook.toStream(wb), createWriteStream("report.xlsx"));
 * for await (const chunk of Workbook.toStream(wb)) {
 *   // …
 * }
 * // SDKs that demand a nominal `stream.Readable` (upload bodies, `toWeb`)
 * // take it directly — no `Readable.from()` adapter, no cast:
 * const body = Workbook.toStream(wb);
 * const web = Readable.toWeb(Workbook.toStream(wb));
 * ```
 *
 * The return type refines {@link XlsxReadable} rather than replacing it: it is
 * still assignable to `XlsxReadable`, which is identical in both builds, so
 * cross-platform code written against that name keeps compiling while Node code
 * gets the nominal `stream.Readable`.
 *
 * {@link XlsxStreamOptions.highWaterMark} is not a hard memory bound:
 * backpressure is sampled after each ZIP entry, and a worksheet is rendered as
 * one entry, so peak buffering may substantially exceed the
 * configured value and is usually driven by the largest worksheet. See
 * {@link XlsxStreamOptions.highWaterMark} for how `Stream.WorkbookWriter`
 * compares.
 *
 * Do not mutate the workbook until the stream ends. A consumer that abandons the
 * stream should `destroy()` it, which releases the parked writer; serialization
 * failures surface as an `'error'` event, so consume through `stream.pipeline()`
 * or `for await` rather than a bare `.pipe()`.
 *
 * ### Leaving `for await` early
 *
 * Reading the first N chunks and `break`ing is Node's semantics, not this
 * stream's, and the two teardown paths do **not** look alike:
 *
 * ```ts
 * for await (const chunk of Workbook.toStream(wb)) break;
 * // stream.errored === AbortError: The operation was aborted
 * ```
 *
 * Node's async iterator destroys the stream with an `AbortError` on early exit
 * (`createAsyncIterator` → `destroyer`), whereas calling `destroy()` yourself
 * destroys it with no error. Nothing crashes — the iterator handles the event, so
 * an absent `'error'` listener is not an unhandled error — but a listener that
 * *is* attached receives a spurious `AbortError`, and `stream.errored` is set, so
 * code that reports either as a serialization failure will report a false one.
 *
 * Two ways to leave early *silently*, both of which still release the parked
 * writer:
 *
 * ```ts
 * // 1. destroy first, then break
 * const stream = Workbook.toStream(wb);
 * for await (const chunk of stream) {
 *   if (enough(chunk)) {
 *     stream.destroy();
 *     break;
 *   }
 * }
 *
 * // 2. opt out of the iterator's teardown — but then YOU must destroy it,
 * //    or the serializer stays parked
 * const stream = Workbook.toStream(wb);
 * try {
 *   for await (const chunk of stream.iterator({ destroyOnReturn: false })) {
 *     if (enough(chunk)) break;
 *   }
 * } finally {
 *   stream.destroy();
 * }
 * ```
 *
 * The browser build behaves the same way: its stream kernel is a port of Node's.
 */
export function toStream(
  wb: WorkbookData,
  options?: XlsxStreamOptions & WorkbookWriteOptions
): XlsxReadable & Readable {
  checkWorkbook(wb);
  checkIoOptions("Workbook.toStream", options, "stream");
  if (options?.password !== undefined) {
    // An encrypted package is assembled and encrypted before its first byte can be handed out.
    return createXlsxByteStream(
      async sink => writeBytesToStream(sink, await packageBytes(wb, options)),
      options
    ) as XlsxReadable & Readable;
  }
  if (options?.format === "xlsb") {
    // **Through the same bridge the XLSX branch uses, because the reason for not doing so was false.**
    //
    // This read: "XLSB has no incremental form: the ZIP central directory is written last and the shared-string table
    // is only complete once every worksheet has been visited". `core/xlsb-stream.ts` records that
    // argument as wrong and why — the central directory being last is how every ZIP writer works, and the shared-string
    // *part* is written after the sheets, which is where it belongs — but the conclusion was left standing here.
    //
    // What it cost, measured on an eight-sheet workbook: XLSX produced 427 chunks of at most 7 KB while XLSB produced
    // **one** chunk of 2,033 KB, so `highWaterMark` was inert and the whole package was built before a consumer saw any
    // of it. Two containers behaving that differently behind one public function is the kind of difference a caller
    // cannot see and cannot work around.
    //
    // `createXlsxByteStream` throttles the producer on `push()` returning `false`, and `writeXlsbToStream` writes part
    // by part — so the same demand-driven contract now applies to both.
    return createXlsxByteStream(
      sink => writeXlsbToStream(wb, sink, options),
      options
    ) as XlsxReadable & Readable;
  }
  // `createXlsxByteStream` is shared with the browser build, so it is typed as
  // the portable `XlsxReadable`. On Node the stream it builds is a real
  // `stream.Readable` — `@stream`'s `createReadable` constructs one — which is
  // what this narrowing states. Asserted at runtime by
  // `surface/__tests__/node-io-types.node.test.ts` ("is a byte-mode
  // stream.Readable"), so the claim cannot rot silently.
  return createXlsxByteStream(
    (sink, writeOptions) => writeXlsxToStream(wb, sink, writeOptions),
    options
  ) as XlsxReadable & Readable;
}

export type { XlsxReadable, XlsxWritable } from "@excel/core/xlsx-io-types";
export type { WorkbookWritable } from "@excel/core/xlsx-sink";
export type { XlsxReadOptions, XlsxWriteOptions } from "@excel/xlsx/types";
export type { XlsxStreamOptions } from "@excel/core/xlsx-stream";
export type { WorkbookFormat } from "@excel/core/workbook-format";

// =============================================================================
// Node-only xlsx file-path IO.
// =============================================================================

/**
 * Node-only: read a workbook from a file path and return it.
 *
 * The format comes from the extension: `.xlsb` reads the binary format, anything else keeps
 * the XLSX path — which streams the ZIP rather than buffering it, and giving that up to sniff
 * the tail of every file would be a poor trade. An explicit `options.format` overrides.
 */
export async function readFile(
  filename: string,
  options?: WorkbookReadOptions & WorkbookCreateOptions
): Promise<WorkbookData>;
/**
 * @deprecated Use `readFile(filename, options)`, which returns a new workbook. This form replaces
 * everything `wb` already holds. Removed in 1.0.
 */
export async function readFile(
  wb: WorkbookData,
  filename: string,
  options?: WorkbookReadOptions
): Promise<WorkbookData>;
export async function readFile(
  first: WorkbookData | string,
  second?: string | (WorkbookReadOptions & WorkbookCreateOptions),
  third?: WorkbookReadOptions
): Promise<WorkbookData> {
  if (isReadTarget(first)) {
    return readFileInto(first, second as string, third);
  }
  const { create, read: options } = splitCreateOptions(
    second as (WorkbookReadOptions & WorkbookCreateOptions) | undefined
  );
  return readFileInto(createWorkbook(create), first, options);
}

/** The `ExcelFileError` for a file the system would not let us read, with its error as the cause. */
function fileReadError(filename: string, error: unknown): ExcelFileError {
  return fileSystemError(filename, "read", error);
}

async function readFileInto(
  wb: WorkbookData,
  filename: string,
  options?: WorkbookReadOptions
): Promise<WorkbookData> {
  checkFilePath("Workbook.readFile", filename);
  checkIoOptions("Workbook.readFile", options, "read");
  if ((options?.format ?? formatFromPath(filename)) === "xlsb") {
    const { readFile: readFileBytes } = await import("node:fs/promises");
    let bytes: Uint8Array;
    try {
      bytes = await readFileBytes(filename);
    } catch (error) {
      throw fileReadError(filename, error);
    }
    const opened = await openWorkbookPackage(bytes, workbookDecryption(options, filename));
    // An encrypted file's package is read as the format it is, whatever the file is called.
    if (opened !== bytes && resolveReadFormat(opened, options?.format, options?.zip) === "xlsx") {
      const workbook = await (await loadXlsxReader()).readXlsxInto(wb, opened, options);
      workbook.sourceFilePath = filename;
      return workbook;
    }
    return readXlsbInto(wb, opened, filename, options);
  }
  // Asked of the file system rather than tested for existence, so its own error — `ENOENT`, `EACCES` — is kept as the
  // cause instead of being reduced to "File not found".
  const { stat } = await import("node:fs/promises");
  try {
    await stat(filename);
  } catch (error) {
    throw fileReadError(filename, error);
  }
  // The reader clears stale origin metadata for every read. Set the path only after a successful
  // parse, so a failed read leaves the previously loaded workbook's origin in place.
  const { readXlsxStreamInto } = await import("@excel/xlsx/read/stream");
  const workbook = await readXlsxStreamInto(wb, createReadStream(filename), options);
  workbook.sourceFilePath = filename;
  return workbook;
}

/**
 * Node-only: write a workbook to a file path.
 *
 * The format comes from the extension unless `options.format` says otherwise, so
 * `writeFile(wb, "report.xlsb")` writes XLSB without a second argument. A file the system will not let us write is
 * an `ExcelFileError` whose `cause` is the system's error.
 */
export async function writeFile(
  wb: WorkbookData,
  filename: string,
  options?: WorkbookWriteOptions
): Promise<void> {
  checkWorkbook(wb);
  checkFilePath("Workbook.writeFile", filename);
  checkIoOptions("Workbook.writeFile", options, "write");
  // `null` is taken as no options, as the check documents.
  options ??= undefined;
  const format = options?.format ?? formatFromPath(filename);
  if (format === "xlsb" || options?.password !== undefined) {
    // XLSB is written whole, and so is an encrypted package of either format.
    const { writeFile: writeFileBytes } = await import("node:fs/promises");
    const bytes = await packageBytes(wb, { ...options, format });
    try {
      await writeFileBytes(filename, bytes);
    } catch (error) {
      throw fileSystemError(filename, "write", error);
    }
    return;
  }
  const stream = createWriteStream(filename);
  await new Promise<void>((resolve, reject) => {
    let settled = false;

    const cleanup = () => {
      stream.off("error", onError);
      stream.off("close", onClose);
    };

    // Only the file stream's own failure is a file-system error; one from the writer is passed on as it is.
    const onError = (error: Error) => {
      if (!settled) {
        settled = true;
        cleanup();
        reject(fileSystemError(filename, "write", error));
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

    writeXlsxToStream(wb, stream, options).catch(err => {
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
