/**
 * `scripts/browser-floor/harness.cjs` has to fail, and exit, whenever the build is not what the browsers should see.
 *
 * Two ways it could pass or hang when it should fail. It starts an HTTP server before it launches the browser, and a
 * failure that skipped closing that server left the process alive with nothing to do, so the `Browser floor` CI job
 * waited for its 30-minute timeout instead of failing. And it took the bundles to expect from the directory it was
 * checking, so a bundle the build failed to write was expected by no one and the check still passed.
 *
 * These cases run it as a subprocess against a stand-in `playwright`, with a deadline far below anything a hang
 * could meet.
 */

import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { IIFE_BUNDLES } from "../../../scripts/lib/iife-bundles";

const HARNESS = path.resolve(import.meta.dirname, "../../../scripts/browser-floor/harness.cjs");

const PUBLISHED = IIFE_BUNDLES.map(({ global, file }) => ({ global, file }));

let root: string;

beforeEach(() => {
  root = mkdtempSync(path.join(tmpdir(), "browser-floor-"));
  mkdirSync(path.join(root, "dist/esm"), { recursive: true });
  mkdirSync(path.join(root, "dist/iife"), { recursive: true });
  mkdirSync(path.join(root, "workers"), { recursive: true });
  writeFileSync(path.join(root, "workers/inflate-worker.js"), "onmessage = () => {};");
  for (const { file } of PUBLISHED) {
    writeBundle(file);
  }
});

afterEach(() => {
  rmSync(root, { recursive: true, force: true });
});

function writeBundle(file: string): void {
  writeFileSync(path.join(root, "dist/iife", `documonster.${file}.iife.min.js`), "var a=1;");
}

/**
 * Run the harness with `playwright` resolving to a module whose `chromium` is `engine` (CommonJS source). The harness
 * runs from a copy beside that module: from its place in the repository, `require` would find the real Playwright.
 */
function run(
  engine: string,
  {
    bundles = PUBLISHED,
    workers = path.join(root, "workers")
  }: { bundles?: readonly { global: string; file: string }[]; workers?: string } = {}
): { status: number | null; signal: string | null; stderr: string } {
  mkdirSync(path.join(root, "node_modules/playwright"), { recursive: true });
  writeFileSync(
    path.join(root, "node_modules/playwright/index.js"),
    `"use strict";\nexports.chromium = ${engine};\n`
  );
  const harness = path.join(root, "harness.cjs");
  copyFileSync(HARNESS, harness);
  const result = spawnSync(process.execPath, [harness, path.join(root, "dist"), workers], {
    env: { ...process.env, ENGINE: "chromium", EXPECT_BUNDLES: JSON.stringify(bundles) },
    encoding: "utf8",
    timeout: 10_000
  });
  return { status: result.status, signal: result.signal, stderr: result.stderr };
}

/** An engine whose `launch` leaves a mark at `marker`, then fails — so a test can tell whether it was reached. */
function markingEngine(marker: string): string {
  return `{ launch: async () => {
    require("fs").writeFileSync(${JSON.stringify(marker)}, "");
    throw new Error("launch failed");
  } }`;
}

describe("browser-floor harness", () => {
  it("fails and exits when the browser does not launch", () => {
    const result = run(`{ launch: async () => { throw new Error("launch failed"); } }`);
    expect(result.signal).toBeNull();
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("launch failed");
  });

  it("closes the browser, fails and exits when driving it fails", () => {
    const closed = path.join(root, "closed");
    const result = run(`{
      launch: async () => ({
        version: () => "98.0",
        newPage: async () => { throw new Error("page failed"); },
        close: async () => { require("fs").writeFileSync(${JSON.stringify(closed)}, ""); }
      })
    }`);
    expect(result.signal).toBeNull();
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("page failed");
    expect(existsSync(closed)).toBe(true);
  });

  it("fails, without a browser, when a published bundle was not built", () => {
    rmSync(path.join(root, "dist/iife/documonster.utils.iife.min.js"));
    const launched = path.join(root, "launched");
    const result = run(markingEngine(launched));
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("iife/documonster.utils.iife.min.js — listed, but not built");
    expect(existsSync(launched)).toBe(false);
  });

  it("fails, without a browser, when a bundle was built that nothing publishes", () => {
    writeBundle("stray");
    const launched = path.join(root, "launched");
    const result = run(markingEngine(launched));
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("iife/documonster.stray.iife.min.js — built, but not listed");
    expect(existsSync(launched)).toBe(false);
  });

  it("fails, without a browser, when it is given no worker scripts", () => {
    const launched = path.join(root, "launched");
    for (const workers of [path.join(root, "missing"), path.join(root, "dist/esm")]) {
      const result = run(markingEngine(launched), { workers });
      expect(result.status).toBe(1);
      expect(result.stderr).toMatch(/no worker scripts/);
      expect(existsSync(launched)).toBe(false);
    }
  });

  it("refuses to run without the list of published bundles", () => {
    const result = spawnSync(process.execPath, [HARNESS, path.join(root, "dist")], {
      env: { ...process.env, ENGINE: "chromium", EXPECT_BUNDLES: "" },
      encoding: "utf8",
      timeout: 10_000
    });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain("EXPECT_BUNDLES is not set");
  });
});
