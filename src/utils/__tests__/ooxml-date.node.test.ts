/**
 * `parseOoxmlDate` reads text to the same calendar value in every zone the reader may run in. It read date text as
 * local time, so the value depended on the machine: where a zone skips its midnight for daylight saving (Santiago on
 * 2019-09-08), `Sep 8 2019` read as 01:00.
 */
import { spawnSync } from "node:child_process";
import { fileURLToPath, pathToFileURL } from "node:url";

import { parseOoxmlDate } from "@utils/utils";
import { afterEach, describe, expect, it } from "vitest";

const zone = process.env.TZ;
afterEach(() => {
  if (zone === undefined) {
    delete process.env.TZ;
  } else {
    process.env.TZ = zone;
  }
});

describe("parseOoxmlDate in every zone", () => {
  it.each(["UTC", "America/Santiago", "America/Sao_Paulo", "Asia/Kolkata", "Pacific/Kiritimati"])(
    "in %s",
    tz => {
      process.env.TZ = tz;
      const read = (text: string): string => parseOoxmlDate(text).toISOString();
      expect(read("Sep 8 2019")).toBe("2019-09-08T00:00:00.000Z");
      expect(read("9/8/2019")).toBe("2019-09-08T00:00:00.000Z");
      expect(read("Nov 4 2018")).toBe("2018-11-04T00:00:00.000Z");
      expect(read("Sep 8 2019 00:30")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("1/15/2020 10:30 PM")).toBe("2020-01-15T22:30:00.000Z");
      expect(read("2019-09-08T00:30:00")).toBe("2019-09-08T00:30:00.000Z");
      // A time inside the gap is the time written, however the engine would move it.
      expect(read("2019-09-08 00:30:00")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("9/8/2019 12:30 AM")).toBe("2019-09-08T00:30:00.000Z");
      expect(read("Sep 8, 2019, 12:30:15.5 AM")).toBe("2019-09-08T00:30:15.500Z");
      // Text naming its zone names an instant.
      expect(read("Jan 15 2020 10:00 +0200")).toBe("2020-01-15T08:00:00.000Z");
      expect(read("Jan 15 2020 10:00 GMT")).toBe("2020-01-15T10:00:00.000Z");
      // A name in parentheses is a comment to the engine: the text is its calendar value, not the reader's zone's.
      expect(read("Jan 15 2020 10:00 (CEST)")).toBe("2020-01-15T10:00:00.000Z");
    }
  );
});

// A cell's text is anyone's, and the patterns that read it backtracked: two overlapping alternatives under `+` took
// 445 ms on 24 spaces before a letter (doubling with each), and `[^)]*` after each of 100,000 `(` 18 s. A pattern that
// backtracks blocks the thread it runs on, so it is run in a process of its own, ended at a deadline: in this one, a
// regression would hang the run instead of failing it.
describe("parseOoxmlDate on adversarial text", () => {
  // Each text is `before`, `unit` repeated `count` times, then `after` — built in the child: the texts themselves
  // would exceed a command line.
  const inputs: Record<
    string,
    readonly [before: string, unit: string, count: number, after: string]
  > = {
    "spaces before a letter, after a time": ["1/15/2020 10:30", " ", 100_000, "b"],
    "spaces before a letter": ["a", " ", 100_000, "b"],
    "tabs before a letter, after a time": ["1/15/2020 10:30", "\t", 100_000, "b"],
    "open parentheses": ["", "(", 100_000, ""],
    "separators and `at`": ["1/15/2020 10:30", " at ,", 50_000, ""],
    meridiems: ["10:30 ", "AM ", 100_000, ""],
    "digits and spaces": ["1", " 1", 100_000, ""],
    "a long fraction": ["2024-01-01T10:00:00.", "1", 100_000, "x"]
  };

  it("reads each in linear time", () => {
    const module = pathToFileURL(fileURLToPath(new URL("../utils.base.ts", import.meta.url))).href;
    const script = `
      const { parseOoxmlDate } = await import(${JSON.stringify(module)});
      const times = {};
      for (const [name, [before, unit, count, after]] of Object.entries(${JSON.stringify(inputs)})) {
        const text = before + unit.repeat(count) + after;
        const start = performance.now();
        parseOoxmlDate(text);
        times[name] = performance.now() - start;
      }
      console.log(JSON.stringify(times));
    `;
    // Bun runs TypeScript itself; Node is given the loader the repository's scripts run with.
    const args =
      "Bun" in globalThis
        ? ["-e", script]
        : ["--import", "@oxc-node/core/register", "--input-type=module", "-e", script];
    const result = spawnSync(process.execPath, args, { encoding: "utf8", timeout: 30_000 });
    expect(result.signal, "the parse did not finish within 30 s").toBeNull();
    expect(result.error).toBeUndefined();
    expect(result.status, result.stderr).toBe(0);
    const times = JSON.parse(result.stdout) as Record<string, number>;
    expect(Object.keys(times)).toEqual(Object.keys(inputs));
    for (const [name, ms] of Object.entries(times)) {
      expect(ms, name).toBeLessThan(2_000);
    }
  }, 60_000);
});
