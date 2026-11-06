import { checkWorkbook, type WorkbookData } from "@excel/core/workbook-core";
import {
  resolveReadFormat,
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
/**
 * The canonical public workbook IO surface — browser variant.
 *
 * Kept out of `workbook.browser` so the serializer is not a static dependency of
 * the workbook record module (which would create a workbook ↔ xlsx import cycle).
 * Selected over `xlsx-io.ts` (Node) via the `.browser` same-name swap.
 */
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
import { encryptWorkbookPackage, openWorkbookPackage } from "@excel/utils/container-errors";
import type { XlsxParseStream, XlsxReadOptions } from "@excel/xlsx/types";
import { writeXlsxBytes, writeXlsxToStream } from "@excel/xlsx/write/package";
import { asPlainBytes } from "@utils/binary";

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
// Cross-platform flat IO functions (the canonical public surface).
// =============================================================================

/**
 * Serialize a workbook to xlsx bytes — encrypted with `{ password }` (Agile Encryption in an OLE compound file, what
 * Excel writes for "Encrypt with Password"), which `Workbook.read(bytes, { password })` reads back.
 */
export async function toBuffer(
  wb: WorkbookData,
  options?: WorkbookWriteOptions
): Promise<Uint8Array<ArrayBuffer>> {
  checkWorkbook(wb);
  checkIoOptions("Workbook.toBuffer", options, "write");
  // `null` is taken as no options, as the check documents.
  options ??= undefined;
  return asPlainBytes(await packageBytes(wb, options));
}

/** The package `toBuffer` returns, encrypted with `options.password` when one is given (the code loads only then). */
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
 * The format is detected from the package contents unless `options.format` says otherwise.
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
  if (format === "xlsb" && bytes) {
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
  stream: XlsxParseStream,
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
 * — or be a terminal sink such as an HTTP response or an upload body — *before*
 * this call. Handing over an unconsumed intermediate stream (most commonly a
 * bare `PassThrough`) deadlocks: once its buffers fill (64 KB per side, so
 * roughly 128 KB of output) `write()` returns `false`, no `'drain'` ever fires
 * because nothing is reading, and this promise never settles. Small workbooks
 * that fit inside those buffers appear to work, which makes the failure look
 * size-dependent.
 *
 * {@link toStream} carries no such contract — prefer it whenever the consumer is
 * not already running.
 *
 * ```ts
 * // ❌ Deadlocks above ~128 KB of output: the consumer is attached only after
 * //    this promise resolves, which never happens.
 * const passThrough = createPassThrough<Uint8Array>();
 * await Workbook.writeStream(wb, passThrough);
 * await upload(passThrough);
 *
 * // ✅ Start the consumer first, then await the producer.
 * const passThrough = createPassThrough<Uint8Array>();
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
    // Streamed part by part through the same writer the buffered path uses — see the Node twin for why the
    // previous "nothing to stream incrementally" was wrong.
    await writeXlsbToStream(wb, sink, options);
  } else {
    await writeXlsxToStream(wb, sink, options);
  }
  await finished;
}

/**
 * Serialize a workbook into a cross-platform readable byte stream.
 *
 * The consumer drives production: serialization starts on the first read and
 * pauses whenever the reader stops pulling. Unlike {@link writeStream} there is
 * no ordering contract to get wrong, so an unconsumed destination cannot park
 * the writer forever.
 *
 * ```ts
 * for await (const chunk of Workbook.toStream(wb)) {
 *   // …
 * }
 * await pipeline(Workbook.toStream(wb), sink);
 * const response = new Response(Readable.toWeb(Workbook.toStream(wb) as Readable));
 * ```
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
 * failures surface as an `'error'` event, so consume through `pipeline()` or
 * `for await` rather than a bare `.pipe()`.
 *
 * ### Leaving `for await` early
 *
 * As in Node, breaking out of the loop destroys the stream with an `AbortError`:
 * `errored` is set and an attached `'error'` listener receives it. To leave
 * silently, `destroy()` the stream before `break`, or iterate
 * `stream.iterator({ destroyOnReturn: false })` and `destroy()` it yourself
 * afterwards — otherwise the serializer stays parked.
 */
export function toStream(
  wb: WorkbookData,
  options?: XlsxStreamOptions & WorkbookWriteOptions
): XlsxReadable {
  checkWorkbook(wb);
  checkIoOptions("Workbook.toStream", options, "stream");
  if (options?.password !== undefined) {
    // An encrypted package is assembled and encrypted before its first byte can be handed out.
    return createXlsxByteStream(
      async sink => writeBytesToStream(sink, await packageBytes(wb, options)),
      options
    );
  }
  if (options?.format === "xlsb") {
    // The same bridge the XLSX branch uses — see the Node variant for the measurement. Keeping the two platforms on one
    // mechanism matters more here than anywhere: a difference between them is invisible to whichever one a developer
    // happens to run.
    return createXlsxByteStream(sink => writeXlsbToStream(wb, sink, options), options);
  }
  return createXlsxByteStream(
    (sink, writeOptions) => writeXlsxToStream(wb, sink, writeOptions),
    options
  );
}

export type { XlsxReadable, XlsxWritable } from "@excel/core/xlsx-io-types";
export type { WorkbookWritable } from "@excel/core/xlsx-sink";
export type { XlsxReadOptions, XlsxWriteOptions } from "@excel/xlsx/types";
export type { XlsxStreamOptions } from "@excel/core/xlsx-stream";
export type { WorkbookFormat } from "@excel/core/workbook-format";
