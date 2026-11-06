// This file is typechecked by `pnpm type` but is NOT executed by Vitest.
//
// It pins the public shape of the workbook stream surface: `toStream` must hand
// back the platform's own readable byte stream (async-iterable, pipeline-
// compatible, and on Node a nominal `stream.Readable`), `toBuffer` must promise
// a `Uint8Array` that Node sinks take as it is, and `writeStream` must accept any duck-typed
// sink while returning a bare promise.

import { createWriteStream } from "node:fs";
import { writeFile } from "node:fs/promises";
import { Readable } from "node:stream";
import { pipeline as nodePipeline } from "node:stream/promises";
import type { ReadableStream as NodeReadableStream } from "node:stream/web";

import { ExcelCsv } from "@excel/bridge/csv.node";
import { Workbook } from "@excel/index";
import type { IReadable } from "@stream";
import { pipeline, Writable } from "@stream";

declare const wb: Workbook.Handle;

const options: Workbook.XlsxStreamOptions = { highWaterMark: 1024, zip: { level: 1 } };

const source: IReadable<Uint8Array> = Workbook.toStream(wb, options);
const narrowedSource: Workbook.XlsxReadable = Workbook.toStream(wb);

// Consumers must be able to pull without a cast.
const iterable: AsyncIterable<Uint8Array> = source;

const destination = new Writable({
  write(_chunk, _encoding, callback) {
    callback();
  }
});
const completed: Promise<void> = Workbook.writeStream(wb, destination);
const piped: Promise<void> = pipeline(Workbook.toStream(wb), destination);

// Node ecosystem interop must stay cast-free: `stream.pipeline` for sinks, and
// the nominal `stream.Readable` that SDK upload bodies / `Readable.toWeb()`
// demand — handed over directly, with no `Readable.from()` adapter and no
// `as unknown as Readable` bet on an implementation detail.
const nodePiped: Promise<void> = nodePipeline(
  Workbook.toStream(wb),
  createWriteStream("report.xlsx")
);
const nodeReadable: Readable = Workbook.toStream(wb);
// `toWeb` accepting the value without an adapter is the point here; the variable
// is annotated with Node's own web-stream type because that is what `toWeb`
// returns (it is not the DOM `ReadableStream`).
const webBody: NodeReadableStream = Readable.toWeb(Workbook.toStream(wb));

// `toBuffer` resolves to a `Uint8Array` on Node as everywhere; it is not a `Buffer`,
// and what Node's own sinks take it as is the point.
declare function wantsBytes(b: Uint8Array): void;
export async function bufferSurface(): Promise<void> {
  const bytes: Uint8Array = await Workbook.toBuffer(wb);
  wantsBytes(await Workbook.toBuffer(wb));
  // @ts-expect-error a Buffer-only method is not promised
  void bytes.readUInt32LE(0);
  await writeFile("report.xlsx", bytes);
}

// `ExcelCsv.toStream` is a byte stream, like `toStream` — chunks must be bytes,
// and a `string` element type must not typecheck.
export async function csvIsBytes(): Promise<void> {
  const csv: IReadable<Uint8Array> = ExcelCsv.toStream(wb);
  const nodeCsv: Readable = ExcelCsv.toStream(wb);
  const webCsv: NodeReadableStream = Readable.toWeb(ExcelCsv.toStream(wb));
  for await (const chunk of csv) {
    const bytes: Uint8Array = chunk;
    // @ts-expect-error a CSV chunk is bytes, never a string
    const notAString: string = chunk;
    void bytes;
    void notAString;
  }
  void nodeCsv;
  void webCsv;
}

// The stream is already wired to a producer, so the writable half of the
// underlying Transform must stay off the public type: writing to it would inject
// rows into the output.
export function csvIsReadOnly(): void {
  const csv = ExcelCsv.toStream(wb);
  // @ts-expect-error a CSV read stream must not expose the formatter's `write`
  csv.write(["injected", "row"]);
  // @ts-expect-error a CSV read stream must not expose the formatter's `end`
  csv.end();
  // @ts-expect-error `_transform` is the Transform internals, not public API
  csv._transform(["x"], "utf8", () => {});
}

const duckDestination: Workbook.XlsxWritable = destination;

void narrowedSource;
void iterable;
void completed;
void piped;
void nodePiped;
void nodeReadable;
void webBody;
void duckDestination;
