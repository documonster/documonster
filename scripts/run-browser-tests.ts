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
 * **A stalled shard does not hang.** In WebKit on CI a tester iframe has stopped part-way through a shard — a different
 * file each time, each of which passes on its own — while the browser stayed alive. Nothing in Vitest ends that:
 * `testTimeout` runs inside the stopped iframe, and its heartbeat (vitest-dev/vitest#10956) only notices a browser that
 * stops answering. The job then sat silent until its `timeout-minutes` cancelled it, with no error and no file named.
 * So when nothing has finished for `BROWSER_STALL_TIMEOUT_MS` — no test, no file — the shard's process tree is killed
 * and each file that never reported is listed with the test it stopped in and the last one that finished.
 *
 * The shard is then run once more in a fresh browser. The cause is not known, and failing on the first stall made every
 * such run red until someone re-ran it by hand; a retry that hid the stall would be worse. So it is never silent: a
 * stall that passes on its retry is a warning annotation and an entry in the job summary, with what the first attempt
 * was doing, and a stall that recurs or a retry that fails fails the run. The default is 180 s for WebKit on CI, where
 * it was seen — a shard takes under a minute there and its slowest file about 25 s — and off elsewhere; `0` turns it
 * off.
 *
 * Shards run sequentially: running them at once would put the same total in memory at the same time.
 * `--passWithNoTests` keeps a filtered run (`pnpm test:browser some-file`) from failing on the shards it leaves empty.
 */
import { type ChildProcess, spawn, spawnSync } from "node:child_process";
import { appendFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";

import type { BrowserProgressMessage } from "./lib/browser-progress-reporter.ts";

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

function stallTimeoutMs(): number {
  const raw = process.env.BROWSER_STALL_TIMEOUT_MS;
  if (raw === undefined || raw === "") {
    return process.env.CI && process.env.BROWSER_ENGINE === "webkit" ? 180_000 : 0;
  }
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 0) {
    console.error(`BROWSER_STALL_TIMEOUT_MS must be a whole number of milliseconds, got ${raw}`);
    process.exit(1);
  }
  return value;
}

const stallMs = stallTimeoutMs();

/**
 * Kill the shard and everything it started — Vitest, the browser and its helper processes. On POSIX the shard leads a
 * process group of its own (`detached`), so one signal reaches all of it.
 */
function killTree(child: ChildProcess): void {
  if (child.pid === undefined || child.exitCode !== null || child.signalCode !== null) {
    return;
  }
  if (process.platform === "win32") {
    spawnSync("taskkill", ["/pid", String(child.pid), "/T", "/F"], { stdio: "ignore" });
  } else {
    try {
      process.kill(-child.pid, "SIGKILL");
    } catch {
      // The group is already gone.
    }
  }
}

interface ShardResult {
  readonly status: number | null;
  readonly signal: NodeJS.Signals | null;
  readonly error?: Error;
  /** What the shard was doing when it stopped, when it stalled. */
  readonly stall?: string;
}

/** Where a file the shard holds had got to: the test running in it, and the last one that finished. */
interface FileProgress {
  running?: { readonly test: string; readonly since: number };
  lastFinished?: string;
}

/** A stall, as a report names it: each file that never finished, with the test it stopped in. */
function describeStall(
  shard: string,
  files: ReadonlyMap<string, FileProgress>,
  finished: number,
  total: number,
  now: number
): string {
  const lines = [...files].map(([file, progress]) => {
    const running = progress.running
      ? `stopped in "${progress.running.test}", started ${Math.round((now - progress.running.since) / 1000)}s ago`
      : "no test had started";
    const last =
      progress.lastFinished === undefined ? "" : `; last finished: "${progress.lastFinished}"`;
    return `\n  - ${file}: ${running}${last}`;
  });
  return (
    `browser tests, shard ${shard} stalled: nothing finished in ${stallMs / 1000}s (${finished} of ${total} files finished). ` +
    (files.size === 0
      ? "Every file had reported; the run did not end."
      : `Never reported a result:${lines.join("")}`)
  );
}

function runShard(shard: number, attempt: string): Promise<ShardResult> {
  const args = [
    VITEST,
    "run",
    "--config",
    "vitest.browser.config.ts",
    `--shard=${shard}/${SHARDS}`,
    "--passWithNoTests",
    ...forwarded
  ];
  if (stallMs === 0) {
    return new Promise(resolve => {
      const child = spawn(process.execPath, args, { stdio: "inherit", env });
      child.on("error", error => resolve({ status: null, signal: null, error }));
      child.on("exit", (status, signal) => resolve({ status, signal }));
    });
  }

  return new Promise(resolve => {
    const child = spawn(process.execPath, args, {
      stdio: ["inherit", "inherit", "inherit", "ipc"],
      env: { ...env, BROWSER_PROGRESS_IPC: "1" },
      detached: process.platform !== "win32"
    });
    const pending = new Map<string, FileProgress>();
    let total = 0;
    let stall: string | undefined;

    // The shard is in a process group of its own, so the terminal's Ctrl-C no longer reaches it.
    const forward = (signal: NodeJS.Signals) => {
      killTree(child);
      process.exit(signal === "SIGINT" ? 130 : 143);
    };
    process.once("SIGINT", forward);
    process.once("SIGTERM", forward);

    const onStall = () => {
      stall = describeStall(
        `${shard}/${SHARDS}${attempt}`,
        pending,
        total - pending.size,
        total,
        Date.now()
      );
      console.error(`\n✗ ${stall}`);
      killTree(child);
    };
    let timer = setTimeout(onStall, stallMs);
    // A test starting or ending is progress, not only a file finishing: a file whose tests go on is not stalled.
    const progress = () => {
      clearTimeout(timer);
      timer = setTimeout(onStall, stallMs);
    };

    child.on("message", (message: BrowserProgressMessage) => {
      switch (message.type) {
        case "queued":
          total++;
          pending.set(message.file, {});
          break;
        case "end":
          pending.delete(message.file);
          break;
        case "test-start": {
          const file = pending.get(message.file);
          if (file) {
            file.running = { test: message.test, since: Date.now() };
          }
          break;
        }
        case "test-end": {
          const file = pending.get(message.file);
          if (file) {
            file.running = undefined;
            file.lastFinished = message.test;
          }
          break;
        }
      }
      progress();
    });

    const settle = (result: Omit<ShardResult, "stall">) => {
      clearTimeout(timer);
      process.off("SIGINT", forward);
      process.off("SIGTERM", forward);
      resolve(stall === undefined ? result : { ...result, stall });
    };
    child.on("error", error => settle({ status: null, signal: null, error }));
    child.on("exit", (status, signal) => settle({ status, signal }));
  });
}

/** A line for GitHub's job summary, when the run is on GitHub Actions; nothing elsewhere. */
function summarize(markdown: string): void {
  const file = process.env.GITHUB_STEP_SUMMARY;
  if (file) {
    appendFileSync(file, `${markdown}\n`);
  }
}

/** A GitHub annotation, which shows on the run's page; a plain line elsewhere. */
function annotate(level: "warning" | "error", title: string, message: string): void {
  // An annotation is one line: GitHub reads `%0A` as its line break.
  const text = message.replace(/%/g, "%25").replace(/\r?\n/g, "%0A");
  console.log(
    process.env.GITHUB_ACTIONS
      ? `::${level} title=${title}::${text}`
      : `${level}: ${title}: ${message}`
  );
}

let failed = false;
for (let shard = 1; shard <= SHARDS; shard++) {
  console.log(`\n▶ browser tests, shard ${shard}/${SHARDS}`);
  let result = await runShard(shard, "");
  if (result.stall !== undefined) {
    // **Retried once, and never silently.** A tester iframe has stopped part-way through a shard in WebKit on CI, a
    // different file each time, each passing on its own; the cause is not known. Failing on the first stall made every
    // such run red, to be re-run by hand until it was not. The whole shard is run again in a fresh browser — files it
    // never reached are not in the stall report — and a pass is reported as a warning annotation and in the job
    // summary, with what the first attempt was doing, so each stall is on record. Stalling or failing again fails.
    const first = result.stall;
    console.log(
      `\n▶ browser tests, shard ${shard}/${SHARDS}, once more in a fresh browser after the stall`
    );
    result = await runShard(shard, " (retry)");
    const passed = result.stall === undefined && result.status === 0;
    annotate(
      passed ? "warning" : "error",
      passed
        ? "Browser shard stalled, passed on a retry"
        : "Browser shard stalled, and failed its retry",
      first
    );
    summarize(
      `### ${passed ? "⚠️" : "❌"} Browser tests (${process.env.BROWSER_ENGINE ?? "chromium"}), shard ${shard}/${SHARDS} stalled\n\n` +
        `\`\`\`\n${first}\n\`\`\`\n\n${passed ? "It passed when run again in a fresh browser." : "Its retry did not pass."}\n`
    );
  }
  if (result.stall !== undefined) {
    failed = true;
  } else if (result.status !== 0) {
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
