/**
 * DOCX packaging stamps ZIP entries from the document's dates, so a document
 * with dated core properties produces identical bytes on every save.
 */

import { Build, Document, Io, Streaming } from "@word/index";
import { afterEach, describe, it, expect, vi } from "vitest";

const MODIFIED = new Date("2024-03-05T06:07:08Z");
const CREATED = new Date("2023-01-02T03:04:06Z");

function buildDoc(coreProperties?: { created?: Date; modified?: Date }) {
  const h = Document.create();
  Document.addParagraph(h, "hello");
  const doc = Document.build(h);
  return coreProperties ? { ...doc, coreProperties } : doc;
}

/** DOS date/time of every local file header, as `date << 16 | time`. */
async function entryTimes(bytes: Uint8Array): Promise<number[]> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out: number[] = [];
  for (let i = 0; i + 30 <= bytes.length; i++) {
    if (view.getUint32(i, true) === 0x04034b50) {
      out.push(view.getUint16(i + 12, true) * 0x10000 + view.getUint16(i + 10, true));
    }
  }
  expect(out.length).toBeGreaterThan(3);
  return out;
}

/** The DOS date/time a ZIP writer stores for `date` (local time, two-second resolution). */
function dosApprox(date: Date): number {
  const dosDate =
    ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  const dosTime = (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1);
  return dosDate * 0x10000 + dosTime;
}

describe("DOCX deterministic ZIP timestamps", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  /**
   * Package `doc` twice with the clock an hour apart. Only `Date` is faked, so
   * the async compression still runs on real timers; an hour is far beyond the
   * two-second DOS resolution, so any read of "now" shows in the bytes.
   */
  async function packageAtTwoTimes(doc: ReturnType<typeof buildDoc>) {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date("2025-01-01T10:00:00Z"));
    const a = await Io.package(doc);
    vi.setSystemTime(new Date("2025-01-01T11:00:00Z"));
    const b = await Io.package(doc);
    return { a, b };
  }

  it("an undated document is stamped with the clock (control)", async () => {
    const { a, b } = await packageAtTwoTimes(buildDoc());
    expect(b).not.toEqual(a);
  });

  it("packages identical bytes when core properties are dated", async () => {
    const doc = buildDoc({ created: CREATED, modified: MODIFIED });
    const { a, b } = await packageAtTwoTimes(doc);
    expect(b).toEqual(a);
    for (const t of await entryTimes(a)) {
      expect(t).toBe(dosApprox(MODIFIED));
    }
  });

  it("falls back to created, and an explicit modTime wins", async () => {
    const created = await entryTimes(await Io.package(buildDoc({ created: CREATED })));
    expect(new Set(created)).toEqual(new Set([dosApprox(CREATED)]));

    const explicit = new Date("2020-06-07T08:09:10Z");
    const times = await entryTimes(
      await Io.package(buildDoc({ modified: MODIFIED }), { modTime: explicit })
    );
    expect(new Set(times)).toEqual(new Set([dosApprox(explicit)]));
  });

  it("streaming writer uses the same default", async () => {
    const stream = Streaming.createDocxStream({ coreProperties: { modified: MODIFIED } });
    stream.add(Build.textParagraph("x"));
    const times = await entryTimes(await stream.finalize());
    expect(new Set(times)).toEqual(new Set([dosApprox(MODIFIED)]));
  });
});
