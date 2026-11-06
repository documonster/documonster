/**
 * `WordExcel.toDocx` writes a cell as its number format shows it — the text
 * `Cell.getDisplayText` and `Pdf.fromExcel` give — and the same text in every
 * timezone. Before, a date cell was written as `Date#toString()` in the
 * machine's zone ("Sat Mar 09 2024 19:00:00 GMT-0500 …" in New York).
 *
 * The zone is fixed per process, so each is a child process.
 */

import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { describe, it, expect } from "vitest";

const CHILD = fileURLToPath(new URL("./fixtures/excel-date-text.child.ts", import.meta.url));

function render(tz: string): { word: string; pdf: string } {
  const out = execFileSync(process.execPath, ["--import", "@oxc-node/core/register", CHILD], {
    env: { ...process.env, TZ: tz },
    encoding: "utf8"
  });
  return JSON.parse(out) as { word: string; pdf: string };
}

describe("WordExcel.toDocx cell text", () => {
  it("is the number format's text, identical in every timezone and as Pdf.fromExcel renders it", () => {
    const zones = ["America/New_York", "Asia/Tokyo"];
    const [a, b] = zones.map(render);
    for (const r of [a!, b!]) {
      const cells = r.word.split(/\s*\t\s*|\n/).filter(Boolean);
      expect(cells).toEqual(["2024-03-10", "31 Dec 2024 18:30", "25%", "2024-03-10"]);
      // The formatted cells as the PDF bridge renders them. (Without a number
      // format the PDF bridge falls back to `toLocaleDateString`, which is not
      // the same in every zone — so the last cell is not compared.)
      for (const cell of cells.slice(0, 3)) {
        expect(r.pdf).toContain(cell);
      }
    }
    expect(a!.word).toBe(b!.word);
  }, 60_000);
});
