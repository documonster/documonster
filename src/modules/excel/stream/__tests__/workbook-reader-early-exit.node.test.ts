/**
 * Leaving a streaming workbook read early — `break`, `return` or a `throw` from either loop —
 * settles at once and hands the caller's source back: a file stream is closed, a Web stream
 * unlocked. The sheets here are larger than the unzipper buffers per entry, which is what
 * used to leave a parser parked on an entry nobody would drain.
 */

import { createReadStream, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { Cell, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const ROWS = 12_000;

let dir: string;
let file: string;
let bytes: Uint8Array;

beforeAll(async () => {
  const wb = Workbook.create();
  for (const name of ["First", "Second", "Third"]) {
    const ws = Workbook.addWorksheet(wb, name);
    for (let r = 1; r <= ROWS; r++) {
      Cell.setValue(ws, `A${r}`, r);
      Cell.setValue(ws, `B${r}`, `${name} row ${r} with some text to make the sheet large`);
    }
  }
  bytes = await Workbook.toBuffer(wb);
  dir = mkdtempSync(join(tmpdir(), "workbook-early-exit-"));
  file = join(dir, "big.xlsx");
  writeFileSync(file, bytes);
});

afterAll(() => {
  rmSync(dir, { recursive: true, force: true });
});

function settleWithin<T>(promise: Promise<T>, ms = 10_000): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => {
      timer = setTimeout(() => reject(new Error(`did not settle within ${ms} ms`)), ms);
    })
  ]).finally(() => clearTimeout(timer));
}

async function waitUntil(test: () => boolean): Promise<void> {
  const start = Date.now();
  while (!test()) {
    if (Date.now() - start > 10_000) {
      throw new Error("condition not reached");
    }
    await new Promise(resolve => setTimeout(resolve, 5));
  }
}

type Leave = "before a row" | "during the rows" | "after the sheet" | "by throwing";
const LEAVES: Leave[] = ["before a row", "during the rows", "after the sheet", "by throwing"];

/** Read sheets up to `stopAt`, leaving there; resolves with the row counts read. */
async function leave(reader: WorkbookReader, stopAt: number, how: Leave): Promise<number[]> {
  const counts: number[] = [];
  const thrown = new Error("left");
  try {
    let index = 0;
    for await (const ws of reader) {
      const atStop = index++ === stopAt;
      if (atStop && how === "before a row") {
        break;
      }
      let rows = 0;
      for await (const _row of ws) {
        rows++;
        if (atStop && how === "during the rows" && rows === 10) {
          break;
        }
        if (atStop && how === "by throwing" && rows === 10) {
          throw thrown;
        }
      }
      counts.push(rows);
      if (atStop) {
        break;
      }
    }
  } catch (error) {
    if (error !== thrown) {
      throw error;
    }
  }
  return counts;
}

describe("WorkbookReader: leaving the loop early", () => {
  for (const [where, stopAt] of [
    ["first", 0],
    ["middle", 1],
    ["last", 2]
  ] as const) {
    for (const how of LEAVES) {
      it(`over a file stream, at the ${where} sheet, ${how}`, async () => {
        const stream = createReadStream(file);
        const counts = await settleWithin(leave(new WorkbookReader(stream, {}), stopAt, how));
        expect(counts.slice(0, stopAt)).toEqual(new Array(stopAt).fill(ROWS));
        await settleWithin(waitUntil(() => stream.closed));
      });

      it(`over a Web stream, at the ${where} sheet, ${how}`, async () => {
        let offset = 0;
        const stream = new ReadableStream<Uint8Array>(
          {
            pull(controller) {
              if (offset >= bytes.length) {
                controller.close();
                return;
              }
              controller.enqueue(bytes.slice(offset, offset + 64 * 1024));
              offset += 64 * 1024;
            }
          },
          { highWaterMark: 0 }
        );
        const counts = await settleWithin(leave(new WorkbookReader(stream, {}), stopAt, how));
        expect(counts.slice(0, stopAt)).toEqual(new Array(stopAt).fill(ROWS));
        await settleWithin(waitUntil(() => !stream.locked));
      });
    }
  }
});
