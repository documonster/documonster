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
 * **A stalled shard fails instead of hanging.** In WebKit on CI a tester iframe has stopped part-way through a shard —
 * a different file each time, each of which passes on its own — while the browser stayed alive. Nothing in Vitest ends
 * that: `testTimeout` runs inside the stopped iframe, and its heartbeat (vitest-dev/vitest#10956) only notices a browser
 * that stops answering. The job then sat silent until its `timeout-minutes` cancelled it, with no error and no file
 * named. So when no test file has finished for `BROWSER_STALL_TIMEOUT_MS`, the shard's process tree is killed, the
 * files that never reported a result are listed, and the run fails. It is not retried: a stall is a failure to
 * investigate, and a retry would turn it green. The default is 180 s for WebKit on CI, where it was seen — a shard takes
 * under a minute there and its slowest file about 25 s — and off elsewhere; `0` turns it off.
 *
 * Shards run sequentially: running them at once would put the same total in memory at the same time.
 * `--passWithNoTests` keeps a filtered run (`pnpm test:browser some-file`) from failing on the shards it leaves empty.
 */
import { type ChildProcess, spawn, spawnSync } from "node:child_process";
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
  readonly stalled: boolean;
}

function runShard(shard: number): Promise<ShardResult> {
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
      child.on("error", error => resolve({ status: null, signal: null, error, stalled: false }));
      child.on("exit", (status, signal) => resolve({ status, signal, stalled: false }));
    });
  }

  return new Promise(resolve => {
    const child = spawn(process.execPath, args, {
      stdio: ["inherit", "inherit", "inherit", "ipc"],
      env: { ...env, BROWSER_PROGRESS_IPC: "1" },
      detached: process.platform !== "win32"
    });
    const pending = new Set<string>();
    let total = 0;
    let stalled = false;

    // The shard is in a process group of its own, so the terminal's Ctrl-C no longer reaches it.
    const forward = (signal: NodeJS.Signals) => {
      killTree(child);
      process.exit(signal === "SIGINT" ? 130 : 143);
    };
    process.once("SIGINT", forward);
    process.once("SIGTERM", forward);

    const onStall = () => {
      stalled = true;
      const files = [...pending].map(file => `\n  - ${file}`).join("");
      console.error(
        `\n✗ browser tests, shard ${shard}/${SHARDS} stalled: no test file finished in ${stallMs / 1000}s ` +
          `(${total - pending.size} of ${total} finished). ` +
          (pending.size === 0
            ? "Every file had reported; the run did not end."
            : `Never reported a result:${files}`)
      );
      killTree(child);
    };
    let timer = setTimeout(onStall, stallMs);
    const progress = () => {
      clearTimeout(timer);
      timer = setTimeout(onStall, stallMs);
    };

    child.on("message", (message: BrowserProgressMessage) => {
      if (message.type === "queued") {
        total++;
        pending.add(message.file);
      } else {
        pending.delete(message.file);
      }
      progress();
    });

    const settle = (result: Omit<ShardResult, "stalled">) => {
      clearTimeout(timer);
      process.off("SIGINT", forward);
      process.off("SIGTERM", forward);
      resolve({ ...result, stalled });
    };
    child.on("error", error => settle({ status: null, signal: null, error }));
    child.on("exit", (status, signal) => settle({ status, signal }));
  });
}

let failed = false;
for (let shard = 1; shard <= SHARDS; shard++) {
  console.log(`\n▶ browser tests, shard ${shard}/${SHARDS}`);
  const result = await runShard(shard);
  if (result.stalled) {
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
