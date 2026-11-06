/**
 * Run the browser suite (`vitest.browser.config.ts`) as a few shards, one after another, each in a browser of its own.
 *
 *   node scripts/run-browser-tests.ts [vitest args…]
 *
 * **Why shards.** Every test file runs in its own iframe of one Chromium page, and Chromium does not reclaim a finished
 * iframe's memory (https://issues.chromium.org/issues/530892387). The renderer's footprint therefore grows with the
 * number of files, and once the suite passed ~700 files the renderer died part-way through on the Windows runners —
 * `Browser connection was closed`, `Cannot connect to the iframe`, `Failed to fetch dynamically imported module` — with
 * every file after the crash left unrun. A shard is a fresh Vitest process and a fresh browser, so what a renderer
 * accumulates is bounded by the shard's size rather than the suite's.
 *
 * Six, not three: at ~240 files a shard the renderer still died on a Windows runner, at file 201 of 237.
 *
 * **And a garbage collection between files, on every platform.** Vitest ships the workaround for that bug — a
 * `HeapProfiler.collectGarbage` after each file — but arms it only on Linux, and there only when the disk under `TMPDIR`
 * has less than 4 GB free. The CI matrix raised the threshold, which did nothing on Windows and macOS, where the
 * renderer kept dying: the platform check comes first. So this sets all three things it needs —
 * `VITEST_CHROMIUM_GC_FORCE` past the platform check, a threshold no disk reaches past the space check, and on Windows a
 * `TMPDIR`, because the check `statfs`es `TMPDIR` or `/tmp`, which is not a path there, and a failed check skips the
 * collection. It costs about 6 ms a file. Remove it when the Chromium bug is fixed.
 *
 * Shards run sequentially: running them at once would put the same total in memory at the same time.
 * `--passWithNoTests` keeps a filtered run (`pnpm test:browser some-file`) from failing on the shards it leaves empty.
 */
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";

const SHARDS = 6;
const VITEST = join(process.cwd(), "node_modules", "vitest", "vitest.mjs");
const forwarded = process.argv.slice(2).filter(arg => arg !== "--");
const env = {
  ...process.env,
  VITEST_CHROMIUM_GC_FORCE: process.env.VITEST_CHROMIUM_GC_FORCE ?? "1",
  VITEST_CHROMIUM_GC_DISK_THRESHOLD_GB:
    process.env.VITEST_CHROMIUM_GC_DISK_THRESHOLD_GB ?? "1048576",
  TMPDIR: process.env.TMPDIR ?? tmpdir()
};

let failed = false;
for (let shard = 1; shard <= SHARDS; shard++) {
  console.log(`\n▶ browser tests, shard ${shard}/${SHARDS}`);
  const result = spawnSync(
    process.execPath,
    [
      VITEST,
      "run",
      "--config",
      "vitest.browser.config.ts",
      `--shard=${shard}/${SHARDS}`,
      "--passWithNoTests",
      ...forwarded
    ],
    { stdio: "inherit", env }
  );
  if (result.status !== 0) {
    // `error` is set when the shard never ran (Node not found, spawn refused); `status` is then null too.
    console.error(
      `✗ browser tests, shard ${shard}/${SHARDS} ${
        result.error === undefined
          ? `exited with ${result.status ?? result.signal}`
          : `could not start: ${result.error.message}`
      }`
    );
    failed = true;
  }
}
process.exit(failed ? 1 : 0);
