/**
 * Web Streams (ReadableStream/WritableStream) example.
 *
 * This demonstrates a complete flow:
 * 1) Create a workbook
 * 2) Write it to a Web WritableStream<Uint8Array> using WorkbookWriter
 * 3) Read it back from a Web ReadableStream<Uint8Array> using WorkbookReader
 *
 * Works in:
 * - Node.js 20+ (has WHATWG ReadableStream/WritableStream)
 * - Modern browsers (when bundled)
 */

import { Stream } from "@excel/index";

async function main(): Promise<void> {
  // -------------------------------------------------------------------------
  // Write: WorkbookWriter -> WritableStream<Uint8Array>
  // -------------------------------------------------------------------------

  const outputChunks: Uint8Array[] = [];

  const writable = new WritableStream<Uint8Array>({
    write(chunk) {
      outputChunks.push(chunk);
    }
  });

  const writer = new Stream.WorkbookWriter({
    stream: writable,
    useStyles: true,
    useSharedStrings: true
  });

  const ws = writer.addWorksheet("Sheet1");
  Stream.commitRow(ws.addRow(["Name", "Score"]));
  Stream.commitRow(ws.addRow(["Alice", 98]));
  Stream.commitRow(ws.addRow(["Bob", 87]));
  ws.commit();
  await writer.commit();

  const xlsxBytes = concatUint8Arrays(outputChunks);
  console.log("Wrote bytes:", xlsxBytes.byteLength);

  // -------------------------------------------------------------------------
  // Read: ReadableStream<Uint8Array> -> WorkbookReader
  // -------------------------------------------------------------------------

  const readable = uint8ArrayToReadableStream(xlsxBytes, 32 * 1024);
  const reader = new Stream.WorkbookReader(readable, { worksheets: "emit" });

  for await (const sheet of reader) {
    console.log("Reading sheet:", sheet.name);
    for await (const row of sheet) {
      // Row.values includes a leading empty slot at index 0 in many sheet models.
      console.log("Row", row.number, Stream.rowValues(row));
    }
  }
}

function concatUint8Arrays(chunks: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const c of chunks) {
    total += c.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

function uint8ArrayToReadableStream(
  data: Uint8Array,
  chunkSize: number
): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (let i = 0; i < data.length; i += chunkSize) {
        controller.enqueue(data.slice(i, i + chunkSize));
      }
      controller.close();
    }
  });
}

main().catch(err => {
  console.error(err);
  if (typeof process !== "undefined") {
    process.exitCode = 1;
  }
});
