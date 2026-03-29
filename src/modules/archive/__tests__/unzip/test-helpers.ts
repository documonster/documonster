/**
 * Shared test helpers for unzip tests.
 */

import { zip } from "@archive";

export function delay(ms = 0): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export async function* chunkBytes(bytes: Uint8Array, chunkSize: number): AsyncIterable<Uint8Array> {
  for (let i = 0; i < bytes.length; i += chunkSize) {
    await delay(0);
    yield bytes.subarray(i, Math.min(bytes.length, i + chunkSize));
  }
}

export async function collectStream(stream: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream) {
    chunks.push(chunk);
  }
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

export function concatChunks(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

/**
 * Create a ZIP with data descriptors by using the streaming path.
 * `zip().stream()` produces entries with FLAG_DATA_DESCRIPTOR set,
 * meaning compressedSize is unknown at local-header time.
 */
export async function createDataDescriptorZip(
  entries: Array<{ name: string; data: Uint8Array }>
): Promise<Uint8Array> {
  const z = zip({ level: 1 });
  for (const { name, data } of entries) {
    z.add(name, data);
  }
  return collectStream(z.stream());
}
