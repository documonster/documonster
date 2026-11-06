/**
 * The public API changes since a release, generated rather than written.
 *
 * Reads the declarations of every `exports` subpath of `documonster` and `@documonster/mcp` — the published release
 * from npm, and the current tree's build — compares them name by name, and writes:
 *
 * - `docs/api-changes-since-v<base>.md`, the exhaustive per-name list (every removal, signature change, deprecation
 *   and addition, signatures in full, plus the MCP tools' input schemas);
 * - with `--migration`, the generated parts of `MIGRATION.md` — the summary table and the deprecation appendix —
 *   between their `<!-- api-changes:… -->` markers.
 *
 * Run it before a release, after the last API change: `pnpm api:changes --migration`. It needs the current tree built
 * (`pnpm build:esm && pnpm build:packages`) and network access to fetch the base release.
 *
 * Usage:
 *   node scripts/api-changes.ts                   # against the newest release tag
 *   node scripts/api-changes.ts --base 0.21.1
 *   node scripts/api-changes.ts --base 0.21.1 --migration
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { diffApi } from "./api-changes/diff.ts";
import { dumpApi } from "./api-changes/dump.ts";
import type { ApiRecord } from "./api-changes/dump.ts";
import { fetchPackage, run } from "./api-changes/package.ts";
import { deprecationAppendix, fullReport, summaryTable } from "./api-changes/report.ts";
import { describeTools } from "./api-changes/tools.ts";

const ROOT = path.resolve(import.meta.dirname, "..");
const MCP = path.join(ROOT, "packages/mcp");
const TYPE_ROOTS = [path.join(ROOT, "node_modules/@types")];

function arg(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

function newestRelease(): string {
  try {
    return run("git", ["describe", "--tags", "--abbrev=0", "--match", "v*"], ROOT).replace(
      /^v/,
      ""
    );
  } catch {
    console.error(
      "✗ api:changes — no release tag in the local history (a shallow clone?). Fetch tags or pass --base."
    );
    process.exit(2);
  }
}

function dumpBoth(
  root: string,
  label: string,
  names: "published" | "declared" = "published"
): ApiRecord {
  return {
    ...dumpApi(root, "node", label, TYPE_ROOTS, names),
    ...dumpApi(root, "browser", label, TYPE_ROOTS, names)
  };
}

/** Replace the text between `<!-- api-changes:<name>:start -->` and its `:end` marker. */
function replaceBetween(text: string, name: string, body: string): string {
  const start = `<!-- api-changes:${name}:start -->`;
  const end = `<!-- api-changes:${name}:end -->`;
  const a = text.indexOf(start);
  const b = text.indexOf(end);
  if (a < 0 || b < a) {
    throw new Error(`api:changes — MIGRATION.md has no ${start} … ${end} markers`);
  }
  return text.slice(0, a + start.length) + "\n\n" + body.trim() + "\n\n" + text.slice(b);
}

async function main(): Promise<void> {
  const base = arg("--base") ?? newestRelease();
  const out = arg("--out") ?? path.join(ROOT, "docs", `api-changes-since-v${base}.md`);
  for (const required of [path.join(ROOT, "dist/types"), path.join(MCP, "dist/index.js")]) {
    if (!fs.existsSync(required)) {
      console.error(
        `✗ api:changes — ${path.relative(ROOT, required)} is missing. Run pnpm build:esm && pnpm build:packages.`
      );
      process.exit(2);
    }
  }

  const work = fs.mkdtempSync(path.join(os.tmpdir(), `documonster-api-changes-${base}-`));
  try {
    console.log(`api:changes — fetching documonster@${base} and @documonster/mcp@${base}`);
    const oldCore = fetchPackage(`documonster@${base}`, path.join(work, "core"));
    const oldMcp = fetchPackage(`@documonster/mcp@${base}`, path.join(work, "mcp"));
    // The old server resolves its own dependencies (its exact documonster, zod, the MCP SDK) from here.
    run(
      "npm",
      ["install", "--omit=dev", "--ignore-scripts", "--no-audit", "--no-fund", "--silent"],
      oldMcp
    );

    console.log("api:changes — reading declarations");
    const before = dumpBoth(oldCore, "documonster");
    const after = dumpBoth(ROOT, "documonster");
    const mcpBefore = dumpApi(oldMcp, "node", "@documonster/mcp", TYPE_ROOTS);
    const mcpAfter = dumpApi(MCP, "node", "@documonster/mcp", TYPE_ROOTS);
    const input = {
      base,
      before,
      after,
      diff: diffApi(before, after, {
        before: dumpBoth(oldCore, "documonster", "declared"),
        after: dumpBoth(ROOT, "documonster", "declared")
      }),
      mcpDiff: diffApi(mcpBefore, mcpAfter, {
        before: dumpApi(oldMcp, "node", "@documonster/mcp", TYPE_ROOTS, "declared"),
        after: dumpApi(MCP, "node", "@documonster/mcp", TYPE_ROOTS, "declared")
      }),
      toolsBefore: await describeTools(oldMcp),
      toolsAfter: await describeTools(MCP)
    };

    fs.mkdirSync(path.dirname(out), { recursive: true });
    fs.writeFileSync(out, fullReport(input));
    console.log(`✓ api:changes — wrote ${path.relative(ROOT, out)}`);

    if (process.argv.includes("--migration")) {
      const file = path.join(ROOT, "MIGRATION.md");
      let text = fs.readFileSync(file, "utf8");
      text = replaceBetween(text, "summary", summaryTable(input, path.relative(ROOT, out)));
      text = replaceBetween(text, "deprecations", deprecationAppendix(input));
      fs.writeFileSync(file, text);
      console.log(
        "✓ api:changes — updated the summary and the deprecation appendix in MIGRATION.md"
      );
    }
    console.log("  Run pnpm format afterwards: the tables are written unaligned.");
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

await main();
