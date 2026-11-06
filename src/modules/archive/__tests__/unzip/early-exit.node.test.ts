/**
 * Node half of the early-exit suite: the same cases over a file stream and a Node
 * `Readable`, which the reader destroys when it lets go of them — as Excel's streaming
 * workbook reader does — plus the shared Web-stream and generator cases.
 */

import { createReadStream, writeFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Readable } from "node:stream";

import { Archive } from "@archive";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import type { SourceFactory } from "./early-exit.shared";
import {
  PORTABLE_SOURCES,
  ZIP_KINDS,
  defineTarEarlyExitSuite,
  defineZipEarlyExitSuite,
  fixture,
  leaveAt,
  settleWithin,
  waitUntil
} from "./early-exit.shared";

let dir: string;
let fileCount = 0;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "archive-early-exit-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

function nodeChunks(bytes: Uint8Array, size = 64 * 1024): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let i = 0; i < bytes.length; i += size) {
    out.push(bytes.subarray(i, i + size));
  }
  return out;
}

const NODE_SOURCES: Record<string, SourceFactory> = {
  "fs.createReadStream": bytes => {
    const file = join(dir, `archive-${fileCount++}.bin`);
    writeFileSync(file, bytes);
    const stream = createReadStream(file);
    return {
      source: stream,
      async released() {
        await waitUntil(() => stream.closed);
      }
    };
  },
  "Readable.from(chunks)": bytes => {
    const stream = Readable.from(nodeChunks(bytes));
    return {
      source: stream,
      async released() {
        await waitUntil(() => stream.destroyed);
      }
    };
  }
};

const ALL_SOURCES = { ...NODE_SOURCES, ...PORTABLE_SOURCES };

defineZipEarlyExitSuite(ALL_SOURCES);
defineTarEarlyExitSuite(ALL_SOURCES);

describe("leaving the loop early releases a real file stream", () => {
  for (const [kind, build] of Object.entries(ZIP_KINDS)) {
    for (const where of ["1-a.txt", "2-big-noise.bin", "5-z.txt"]) {
      it(`${kind}: the file is closed after leaving at ${where}`, async () => {
        const f = fixture();
        const file = join(dir, `real-${fileCount++}.zip`);
        await writeFile(file, await build(f));
        const stream = createReadStream(file);
        await settleWithin(leaveAt(Archive.openZip(stream).entries(), where, "unread"));
        await settleWithin(
          new Promise<void>(resolve => (stream.closed ? resolve() : stream.once("close", resolve))),
          3000,
          "closing the file"
        );
        expect(stream.destroyed).toBe(true);
      });
    }
  }

  it("leaves no timer, handle or file descriptor behind", async () => {
    const f = fixture();
    const file = join(dir, `leak-${fileCount++}.zip`);
    await writeFile(file, await ZIP_KINDS.deflated!(f));
    // Process-wide, so the worker's own output is in the counts too: a `SimpleWriteWrap` is a write to its stdout or
    // stderr pipe still in flight, which on a loaded CI runner can still be pending when the second count is taken.
    // None of these is anything the reader opens; counting them failed the test on output the reader never wrote.
    const STDIO = new Set(["SimpleWriteWrap", "WriteWrap", "ShutdownWrap", "PipeWrap", "TTYWrap"]);
    const count = () => {
      const counts = new Map<string, number>();
      for (const kind of process.getActiveResourcesInfo()) {
        if (!STDIO.has(kind)) {
          counts.set(kind, (counts.get(kind) ?? 0) + 1);
        }
      }
      return counts;
    };
    await new Promise(resolve => setTimeout(resolve, 20));
    const before = count();
    for (const where of ["1-a.txt", "2-big-noise.bin", "3-c.txt"]) {
      const stream = createReadStream(file);
      await settleWithin(leaveAt(Archive.openZip(stream).entries(), where, "unread"));
      await waitUntil(() => stream.closed);
    }
    await new Promise(resolve => setTimeout(resolve, 20));
    for (const [kind, n] of count()) {
      expect({ kind, n }).toEqual({ kind, n: Math.min(n, before.get(kind) ?? 0) });
    }
  });
});
