/**
 * Compatibility gate: the previous release's examples must still type-check and run against this tree.
 *
 * Examples are the closest thing this repository has to real consumer code, and every one of them was written against
 * the public API of the release it shipped with. Running *that* release's examples against the current source asks the
 * question a renaming pass most needs answered — does code written for the last version still compile, and does it
 * still do what it did — which the current examples cannot, because they are updated with the API they exercise.
 *
 * The old examples are extracted with `git archive` into `src/modules/<module>/compat-examples/` (beside the current
 * ones, so their relative paths and the `@<module>/…` aliases resolve exactly as they did), type-checked together with
 * the tree, run one by one, and removed again, also when the run fails or is interrupted.
 *
 * Usage:
 *   node scripts/verify-compat.ts               # against the newest release tag
 *   node scripts/verify-compat.ts --ref v0.21.1
 *   node scripts/verify-compat.ts --timeout 300 # seconds per example
 *
 * Needs the tag in the local history; a shallow clone has none, and the script says so rather than passing.
 */
import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const DIR_NAME = "compat-examples";

function arg(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

function git(...args: string[]): string {
  return execFileSync("git", args, { cwd: ROOT, encoding: "utf8" }).trim();
}

function newestReleaseTag(): string {
  try {
    return git("describe", "--tags", "--abbrev=0", "--match", "v*");
  } catch {
    console.error(
      "✗ verify:compat — no release tag in the local history (a shallow clone?). Fetch tags or pass --ref."
    );
    process.exit(2);
  }
}

const ref = arg("--ref") ?? newestReleaseTag();
const timeoutMs = Number(arg("--timeout") ?? 180) * 1000;

/**
 * Old examples that are *meant* to fail now, because what they did was a bug a release fixed on purpose — each one
 * written up in MIGRATION.md. Keyed by release tag and example path; the value is text the failure must contain, so
 * the entry excuses that one failure and nothing else: the same example failing for any other reason still fails the
 * gate, and an entry whose example now runs is reported as stale.
 */
const EXPECTED_BREAKS: Record<string, Record<string, { output: string; why: string }>> = {
  "v0.21.1": {
    "src/modules/excel/compat-examples/mega-dashboard.ts": {
      output: "PivotTableError",
      why:
        "three pivot tables left at the default anchor A3 overlap; Excel opens the file but a refresh collapses " +
        "all but one (verified in Excel 16), so Pivot.add now refuses the overlap"
    }
  }
};
const expectedBreaks = EXPECTED_BREAKS[ref] ?? {};

const exampleDirs = git("ls-tree", "-d", "-r", "--name-only", ref)
  .split("\n")
  .filter(dir => /^src\/modules\/[^/]+\/examples$/.test(dir));
if (exampleDirs.length === 0) {
  console.error(`✗ verify:compat — ${ref} has no src/modules/*/examples directories.`);
  process.exit(2);
}

const placed: string[] = [];
function cleanUp(): void {
  for (const dir of placed) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
process.on("SIGINT", () => {
  cleanUp();
  process.exit(130);
});

let failed = 0;
try {
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-compat-"));
  try {
    const tarball = path.join(scratch, "examples.tar");
    execFileSync("git", ["archive", "-o", tarball, ref, ...exampleDirs], { cwd: ROOT });
    execFileSync("tar", ["xf", tarball, "-C", scratch]);
    for (const dir of exampleDirs) {
      const target = path.join(ROOT, path.dirname(dir), DIR_NAME);
      if (fs.existsSync(target)) {
        throw new Error(`${path.relative(ROOT, target)} already exists; remove it first.`);
      }
      fs.cpSync(path.join(scratch, dir), target, { recursive: true });
      placed.push(target);
    }
  } finally {
    fs.rmSync(scratch, { recursive: true, force: true });
  }

  // Type-check: the whole tree, reporting only what the old examples get wrong.
  const tsc = spawnSync("pnpm", ["exec", "tsc", "--noEmit", "-p", "tsconfig.json"], {
    cwd: ROOT,
    encoding: "utf8"
  });
  const typeErrors = `${tsc.stdout}${tsc.stderr}`
    .split("\n")
    .filter(line => line.includes(`/${DIR_NAME}/`));
  if (typeErrors.length > 0) {
    failed++;
    console.error(`✗ ${typeErrors.length} type error(s) in ${ref}'s examples:`);
    console.error(typeErrors.join("\n"));
  } else if (tsc.status !== 0) {
    // An error elsewhere in the tree, or a compiler that did not run, says nothing about the old examples — and a
    // gate that cannot tell must not pass.
    failed++;
    console.error(
      `✗ the type check failed outside ${ref}'s examples (exit ${tsc.status ?? tsc.error}):`
    );
    console.error(`${tsc.stdout}${tsc.stderr}`.trim().split("\n").slice(0, 20).join("\n"));
  }

  const scripts = placed
    .flatMap(dir =>
      fs
        .readdirSync(dir, { withFileTypes: true })
        .filter(
          entry => entry.isFile() && entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts")
        )
        .map(entry => path.join(dir, entry.name))
    )
    .sort();
  let excused = 0;
  for (const file of scripts) {
    const rel = path.relative(ROOT, file);
    const run = spawnSync("node", ["--import", "@oxc-node/core/register", file], {
      cwd: ROOT,
      encoding: "utf8",
      timeout: timeoutMs
    });
    const expected = expectedBreaks[rel];
    const output = `${run.stdout}${run.stderr}`;
    if (expected) {
      if (run.status !== 0 && !run.error && output.includes(expected.output)) {
        excused++;
        console.log(`• ${rel} fails as intended (${expected.why})`);
        continue;
      }
      failed++;
      console.error(
        run.status === 0
          ? `✗ ${rel} runs, but EXPECTED_BREAKS says it must fail — remove the stale entry`
          : `✗ ${rel} failed, but not with the expected "${expected.output}"`
      );
      console.error(output.trim().split("\n").slice(-8).join("\n"));
      continue;
    }
    if (run.status !== 0) {
      failed++;
      const reason = run.error ? run.error.message : `exit ${run.status}`;
      console.error(`✗ ${rel} (${reason})`);
      console.error(output.trim().split("\n").slice(-8).join("\n"));
    }
  }
  if (failed === 0) {
    console.log(
      `✓ verify:compat — ${scripts.length - excused} example(s) from ${ref} type-check and run` +
        (excused > 0 ? `; ${excused} fail as listed in EXPECTED_BREAKS.` : ".")
    );
  }
} finally {
  cleanUp();
}
process.exit(failed === 0 ? 0 : 1);
