import { Archive, Crc32MismatchError } from "@archive";
import { describe, expect, it } from "vitest";

import { createDataDescriptorZip } from "./test-helpers";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const content = "the quick brown fox jumps over the lazy dog ".repeat(40);

const SIGNATURES = [
  // Local file header, central directory header, data descriptor: offset of their CRC-32
  [0x04034b50, 14],
  [0x02014b50, 16],
  [0x08074b50, 4]
] as const;

/** Change the CRC-32 an archive records for its entries, leaving the data intact. */
function corruptCrc(zip: Uint8Array): Uint8Array {
  const out = zip.slice();
  const view = new DataView(out.buffer);
  for (let i = 0; i + 4 <= out.length; i++) {
    const signature = view.getUint32(i, true);
    for (const [expected, offset] of SIGNATURES) {
      if (signature === expected && i + offset + 4 <= out.length) {
        view.setUint32(i + offset, view.getUint32(i + offset, true) ^ 0x5a5a5a5a, true);
      }
    }
  }
  return out;
}

/** Flip one byte of a stored entry's payload — the CRC is left as written. */
function corruptPayload(zip: Uint8Array): Uint8Array {
  const out = zip.slice();
  const at = new TextDecoder("latin1").decode(out).indexOf("lazy");
  out[at]! ^= 0x20;
  return out;
}

function oneShot(bytes: Uint8Array): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      for (let i = 0; i < bytes.length; i += 300) {
        controller.enqueue(bytes.slice(i, i + 300));
      }
      controller.close();
    }
  });
}

async function readAll(iterable: AsyncIterable<Uint8Array>): Promise<string> {
  const parts: string[] = [];
  for await (const chunk of iterable) {
    parts.push(decoder.decode(chunk, { stream: true }));
  }
  return parts.join("") + decoder.decode();
}

type Read = (zip: Uint8Array, checkCrc32?: false) => Promise<string>;

/** Every public way of reading the entry `a.txt`, with the opt-out threaded through. */
const readers: Record<string, Read> = {
  unpackZip: async (zip, checkCrc32) =>
    decoder.decode((await Archive.unpackZip(zip, { checkCrc32 })).get("a.txt")!.data),
  unpackZipSync: async (zip, checkCrc32) =>
    decoder.decode(Archive.unpackZipSync(zip, { checkCrc32 }).get("a.txt")!.data),
  "openZip(bytes).bytes": async (zip, checkCrc32) =>
    decoder.decode((await Archive.openZip(zip, { checkCrc32 }).bytes("a.txt"))!),
  "openZip(bytes).entries bytes": async (zip, checkCrc32) => {
    for await (const entry of Archive.openZip(zip, { checkCrc32 }).entries()) {
      return decoder.decode(await entry.bytes());
    }
    return "";
  },
  "openZip(bytes).entries stream": async (zip, checkCrc32) => {
    for await (const entry of Archive.openZip(zip, { checkCrc32 }).entries()) {
      return readAll(entry.stream());
    }
    return "";
  },
  "openZip(stream).entries bytes": async (zip, checkCrc32) => {
    let text = "";
    for await (const entry of Archive.openZip(oneShot(zip), { checkCrc32 }).entries()) {
      text += decoder.decode(await entry.bytes());
    }
    return text;
  },
  "openZip(stream).entries stream": async (zip, checkCrc32) => {
    let text = "";
    for await (const entry of Archive.openZip(oneShot(zip), { checkCrc32 }).entries()) {
      text += await readAll(entry.stream());
    }
    return text;
  },
  "RemoteZipReader.extract": async (zip, checkCrc32) => {
    const reader = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip), {
      checkCrc32
    });
    return decoder.decode((await reader.extract("a.txt"))!);
  },
  "RemoteZipReader.extractToStream": async (zip, checkCrc32) => {
    const reader = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip), {
      checkCrc32
    });
    const parts: Uint8Array[] = [];
    await reader.extractToStream(
      "a.txt",
      new WritableStream<Uint8Array>({
        write(chunk) {
          parts.push(chunk);
        }
      })
    );
    return parts.map(part => decoder.decode(part, { stream: true })).join("") + decoder.decode();
  }
};

describe("ZIP reads verify CRC-32 by default", () => {
  const archives: Record<string, () => Promise<Uint8Array>> = {
    stored: () => Archive.packZip([["a.txt", content]], { level: 0 }),
    deflated: () => Archive.packZip([["a.txt", content]], { level: 6, smartStore: false }),
    "data descriptor": () =>
      createDataDescriptorZip([{ name: "a.txt", data: encoder.encode(content) }])
  };

  for (const [kind, build] of Object.entries(archives)) {
    for (const [name, read] of Object.entries(readers)) {
      it(`${name} (${kind})`, async () => {
        const zip = await build();
        expect(await read(zip)).toBe(content);
        const corrupt = corruptCrc(zip);
        await expect(read(corrupt)).rejects.toBeInstanceOf(Crc32MismatchError);
        // The explicit opt-out still returns the data
        expect(await read(corrupt, false)).toBe(content);
      });
    }
  }

  it("a damaged stored payload is reported, not returned", async () => {
    const zip = corruptPayload(await Archive.packZip([["a.txt", content]], { level: 0 }));
    for (const read of Object.values(readers)) {
      await expect(read(zip)).rejects.toBeInstanceOf(Crc32MismatchError);
    }
  });
});
