/**
 * The architecture reference states facts about this repository. They have to be true.
 *
 * The architecture reference and its translations quote four counts — modules, published entry points,
 * formula functions, bridge exceptions — because the reader wants the scale. Every one
 * of them is a number only a human keeps in sync, which the reference's own last section
 * says is a number that is eventually wrong. That rule applies to the document stating
 * it, so each count is derived here from the thing it describes.
 *
 * The section and diagram counts are checked across the three languages for a different
 * reason: a translation that silently loses a section is invisible. Editing the English
 * and forgetting the other two is the likely mistake, and nothing else would catch it.
 *
 * `.node.test.ts` because it reads Markdown and source files off disk.
 */

import fs from "node:fs";
import path from "node:path";

import { listFunctionNames } from "@formula/runtime/function-registry";
import { describe, expect, it } from "vitest";

const ROOT = path.resolve(import.meta.dirname, "../../..");

/** The three editions, which must agree on structure. */
const EDITIONS = [
  "src/test/fixtures/architecture-docs/architecture.md",
  "src/test/fixtures/architecture-docs/architecture.ja.md",
  "src/test/fixtures/architecture-docs/architecture.zh.md"
] as const;

function read(file: string): string {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

/** How many production modules there are. */
function moduleCount(): number {
  return fs
    .readdirSync(path.join(ROOT, "src/modules"), { withFileTypes: true })
    .filter(entry => entry.isDirectory()).length;
}

/** How many subpaths the package publishes, excluding the manifest itself. */
function entryPointCount(): number {
  const manifest = JSON.parse(read("package.json")) as { exports: Record<string, unknown> };
  return Object.keys(manifest.exports).filter(key => key !== "./package.json").length;
}

/** How many per-file exceptions the layer checker registers. */
function bridgeExceptionCount(): number {
  const source = read("scripts/verify-layers.ts");
  const block = /const EXCEPTIONS[^{]*\{([\s\S]*?)\n\};/.exec(source);
  expect(block, "EXCEPTIONS block not found — was verify-layers.ts restructured?").not.toBeNull();
  return [...(block?.[1] ?? "").matchAll(/"src\/modules\/[^"]+":/g)].length;
}

/**
 * The counts, and how each edition spells them.
 *
 * A missing match is a failure rather than a skip: the sentence being reworded is
 * exactly the moment to confirm the number is still right.
 */
const COUNTS: readonly {
  readonly what: string;
  readonly actual: () => number;
  readonly patterns: Readonly<Record<(typeof EDITIONS)[number], RegExp>>;
}[] = [
  {
    what: "modules",
    actual: moduleCount,
    patterns: {
      "src/test/fixtures/architecture-docs/architecture.md": /\| Module count\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.ja.md": /\| モジュール数\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.zh.md": /\| 模块数\s*\|\s*(\d+)/
    }
  },
  {
    what: "published entry points",
    actual: entryPointCount,
    patterns: {
      "src/test/fixtures/architecture-docs/architecture.md":
        /\| Published entry points\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.ja.md":
        /\| 公開エントリポイント\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.zh.md": /\| 公开入口\s*\|\s*(\d+)/
    }
  },
  {
    what: "formula functions",
    actual: () => listFunctionNames().length,
    patterns: {
      "src/test/fixtures/architecture-docs/architecture.md": /\| Formula functions\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.ja.md": /\| 数式関数\s*\|\s*(\d+)/,
      "src/test/fixtures/architecture-docs/architecture.zh.md": /\| 公式函数\s*\|\s*(\d+)/
    }
  },
  {
    what: "bridge exceptions",
    actual: bridgeExceptionCount,
    patterns: {
      "src/test/fixtures/architecture-docs/architecture.md":
        /Exactly (\w+) bridge files are registered exceptions/,
      "src/test/fixtures/architecture-docs/architecture.ja.md":
        /ブリッジファイルは (\d+) つだけです/,
      "src/test/fixtures/architecture-docs/architecture.zh.md": /外的桥接文件恰好 (\d+) 个/
    }
  }
];

/** Spelled-out numerals the English edition uses in prose. */
const WORDS: Readonly<Record<string, number>> = {
  three: 3,
  four: 4,
  five: 5,
  six: 6,
  seven: 7,
  eight: 8
};

describe("the architecture documentation's counts", () => {
  it("has something to check", () => {
    // A list that quietly emptied would make every case below vacuous.
    expect(COUNTS.length).toBe(4);
    expect(moduleCount()).toBeGreaterThan(5);
    expect(entryPointCount()).toBeGreaterThan(5);
    expect(bridgeExceptionCount()).toBeGreaterThan(0);
  });

  it.each(COUNTS.flatMap(count => EDITIONS.map(edition => ({ what: count.what, edition, count }))))(
    "states the right number of $what in $edition",
    ({ edition, count }) => {
      const match = count.patterns[edition].exec(read(edition));
      expect(
        match,
        `${edition}: nothing matched ${String(count.patterns[edition])} — was it reworded?`
      ).not.toBeNull();
      const stated = match?.[1] ?? "";
      const value = WORDS[stated.toLowerCase()] ?? Number(stated);
      expect(value, `${edition} states a ${count.what} count that is not ${count.actual()}`).toBe(
        count.actual()
      );
    }
  );
});

/**
 * The ordered heading levels of an edition, with fenced code stripped first.
 *
 * Comparing the *shape* rather than the text is what makes this work across a
 * translation: the words differ by design, the outline must not. Counting `##` alone
 * saw neither a dropped `###` nor one promoted a level, so a whole subsection could go
 * missing from two of the three editions without anything noticing — which is how a
 * heading renamed in one edition only stayed invisible.
 */
function headingShape(edition: string): number[] {
  const body = read(edition).replace(/^```[\s\S]*?^```/gm, "");
  return [...body.matchAll(/^(#{2,3}) /gm)].map(match => match[1].length);
}

describe("the architecture documentation's translations", () => {
  it("keep the same section outline", () => {
    const shapes = EDITIONS.map(edition => ({ edition, shape: headingShape(edition) }));
    const first = shapes[0].shape;
    // Loose lower bounds, not the real figures: they exist to prove the comparison is not
    // vacuous, and a count restated here would be one more number to keep in sync.
    expect(first.length).toBeGreaterThan(10);
    expect(first.filter(level => level === 2).length).toBeGreaterThan(3);
    expect(first.filter(level => level === 3).length).toBeGreaterThan(3);

    for (const { edition, shape } of shapes) {
      expect(shape, `${edition} has a different heading outline`).toEqual(first);
    }
  });

  it("keep the same diagrams", () => {
    // A translation that drops a diagram loses the part of the document doing the most
    // explaining, and the prose around it still refers to it.
    const counts = EDITIONS.map(edition => ({
      edition,
      diagrams: [...read(edition).matchAll(/^```mermaid$/gm)].length
    }));
    const first = counts[0].diagrams;
    expect(first).toBeGreaterThan(0);
    for (const { edition, diagrams } of counts) {
      expect(diagrams, `${edition} has a different number of diagrams`).toBe(first);
    }
  });

  it("keep the same tables", () => {
    const counts = EDITIONS.map(edition => ({
      edition,
      // A separator row is one table, and unlike a header it cannot be confused with
      // ordinary prose containing a pipe.
      tables: [...read(edition).matchAll(/^\| *-{3,}/gm)].length
    }));
    const first = counts[0].tables;
    expect(first).toBeGreaterThan(0);
    for (const { edition, tables } of counts) {
      expect(tables, `${edition} has a different number of tables`).toBe(first);
    }
  });
});
