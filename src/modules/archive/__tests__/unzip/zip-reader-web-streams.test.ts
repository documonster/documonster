import { ZipReader } from "@archive/unzip";
import { createZip } from "@archive/zip/zip-bytes";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect } from "vitest";

async function readAllFromStream<T>(stream: ReadableStream<T>): Promise<T[]> {
  const reader = stream.getReader();
  const out: T[] = [];
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        break;
      }
      out.push(value);
    }
  } finally {
    reader.releaseLock();
  }
  return out;
}

describe("ZipReader Web Streams adapters", () => {
  it("entriesStream() should yield entries (buffer source)", async () => {
    const enc = new TextEncoder();
    const zipData = await createZip(
      [
        { name: "a.txt", data: enc.encode("A") },
        { name: "b.txt", data: enc.encode("B") }
      ],
      { level: 0, smartStore: false }
    );

    const reader = new ZipReader(zipData);

    const entries = await readAllFromStream(reader.entriesStream());
    const paths = entries.map(e => e.path).sort();
    expect(paths).toEqual(["a.txt", "b.txt"]);
  });

  it("UnzipEntry.readableStream() should expose entry bytes (streaming source)", async () => {
    const payload = new Uint8Array(200 * 1024).fill(66); // 200KiB of 'B'
    const zipData = await createZip([{ name: "big.bin", data: payload }], {
      level: 0,
      smartStore: false
    });

    const chunkSize = 1024;
    const zipStream = new ReadableStream<Uint8Array>({
      start(controller) {
        for (let i = 0; i < zipData.length; i += chunkSize) {
          controller.enqueue(zipData.subarray(i, Math.min(zipData.length, i + chunkSize)));
        }
        controller.close();
      }
    });

    const reader = new ZipReader(zipStream);
    const entryReader = reader.entriesStream().getReader();

    const first = await entryReader.read();
    expect(first.done).toBe(false);
    expect(first.value!.path).toBe("big.bin");

    const chunks = await readAllFromStream(first.value!.readableStream());
    const out = concatUint8Arrays(chunks);
    expect(new Uint8Array(out)).toEqual(payload);

    await entryReader.cancel();
    entryReader.releaseLock();
  });

  it("UnzipEntry.pipeTo(WritableStream, { preventClose }) should not close sink", async () => {
    const payload = new Uint8Array(16 * 1024).fill(67); // 16KiB of 'C'
    const zipData = await createZip([{ name: "c.bin", data: payload }], {
      level: 0,
      smartStore: false
    });

    const zipStream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(zipData);
        controller.close();
      }
    });

    const reader = new ZipReader(zipStream);
    const entryReader = reader.entriesStream().getReader();
    const { value: entry, done } = await entryReader.read();
    expect(done).toBe(false);
    expect(entry!.path).toBe("c.bin");

    let closed = false;
    const chunks: Uint8Array[] = [];
    const writable = new WritableStream<Uint8Array>({
      write(chunk) {
        chunks.push(chunk);
      },
      close() {
        closed = true;
      }
    });

    await entry!.pipeTo(writable, { preventClose: true });
    expect(closed).toBe(false);

    const out = concatUint8Arrays(chunks);
    expect(new Uint8Array(out)).toEqual(payload);

    await entryReader.cancel();
    entryReader.releaseLock();
  });
});
