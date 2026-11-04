/**
 * `splitPages()` against every PDF fixture on disk: each part must be what removing every
 * other page and writing with `toBuffer()` produces (see `split-pages.helpers.ts`).
 */
import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";

import { afterAll, beforeAll, describe, it, vi } from "vitest";

import { expectSplitMatchesReference } from "./split-pages.helpers";

const DATA = join(__dirname, "data");

function pdfFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap(entry =>
    entry.isDirectory()
      ? pdfFiles(join(dir, entry.name))
      : entry.name.endsWith(".pdf")
        ? [join(dir, entry.name)]
        : []
  );
}

beforeAll(() => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(new Date("2026-01-02T03:04:05Z"));
});
afterAll(() => {
  vi.useRealTimers();
});

describe("splitPages() on the fixture files", () => {
  it.each(pdfFiles(DATA).map(file => [relative(DATA, file), file]))("%s", async (_, file) => {
    const bytes = new Uint8Array(readFileSync(file));
    const password = file.includes(`${join("data", "encrypted")}`) ? "user" : "";
    await expectSplitMatchesReference(bytes, { password });
  });
});
