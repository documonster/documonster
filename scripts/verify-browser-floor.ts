/**
 * The browser floor, checked in the browsers themselves: Chrome 98, Firefox 102 and Safari 15.4.
 *
 * `verify-browser-syntax.ts` reads the published tree for syntax the floor cannot parse, which is fast and runs
 * on every build — but it can only refuse what it knows to look for. This runs the published tree in the floor's
 * own engines, so whatever they cannot parse or run fails here whether or not anyone thought to list it. It is
 * how the unlowered IIFE bundles were found: eight of twelve shipped class static blocks, which Safari parses
 * only from 16.4, and nothing in the static check named them at the time.
 *
 * Each engine comes from the Playwright release that shipped it, in that release's Docker image — the browser
 * builds are tied to the image's system libraries, so an old engine cannot be installed into a current one. The
 * image's reported version is checked against the floor, so a wrong pairing fails rather than testing something
 * newer. For each engine, `browser-floor/harness.cjs`:
 *
 * 1. parses every file in `dist/esm` and `dist/iife`, and every worker script shipped as a string — read here, through
 *    `lib/worker-scripts.ts` as `verify-browser-syntax.ts` reads them, and handed to the harness as files;
 * 2. checks `dist/iife` holds exactly the bundles `scripts/lib/iife-bundles.ts` publishes, and loads each as a
 *    `<script>`;
 * 3. runs every module through its bundle — a module with no check fails: a workbook round-trip (with `LET` /
 *    `LAMBDA`, which is the code 0.25.1 broke), a document, a PDF, CSV, a Markdown table, XML, a formula, a
 *    compressed ZIP, a web `ReadableStream`, a drawing as SVG and as pixels, a Mermaid diagram and base64.
 *    Safari 15.4 has no `CompressionStream`, so the ZIP there runs the pure-JavaScript DEFLATE fallback.
 *
 * Needs Docker and a built `dist/`. Usage: `node scripts/verify-browser-floor.ts [--engine <name>] [--dist <dir>]`;
 * `--dist` points it at another tree — an unpacked release, say, to show what that release did at the floor.
 */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

import { IIFE_BUNDLES } from "./lib/iife-bundles.ts";
import { readWorkerScripts } from "./lib/worker-scripts.ts";

const ROOT = path.resolve(import.meta.dirname, "..");

/** The floor, as the Playwright release whose image carries each engine at that version. */
const ENGINES: readonly {
  readonly name: string;
  readonly playwright: string;
  readonly version: string;
}[] = [
  { name: "chromium", playwright: "1.17.0", version: "98." },
  { name: "firefox", playwright: "1.24.0", version: "102." },
  { name: "webkit", playwright: "1.21.0", version: "15.4" }
];

function argValue(flag: string): string | undefined {
  const at = process.argv.indexOf(flag);
  return at === -1 ? undefined : process.argv[at + 1];
}

const only = argValue("--engine");
const DIST = path.resolve(argValue("--dist") ?? path.join(ROOT, "dist"));
const selected = ENGINES.filter(engine => only === undefined || engine.name === only);
if (selected.length === 0) {
  console.error(
    `✗ verify:browser-floor — unknown engine ${JSON.stringify(only)}; one of ${ENGINES.map(e => e.name).join(", ")}`
  );
  process.exit(1);
}

for (const dir of ["esm", "iife"]) {
  if (!fs.existsSync(path.join(DIST, dir))) {
    console.error(
      `✗ verify:browser-floor — ${path.join(DIST, dir)} does not exist; run pnpm build first`
    );
    process.exit(1);
  }
}

if (spawnSync("docker", ["version"], { stdio: "ignore" }).status !== 0) {
  console.error("✗ verify:browser-floor — needs Docker, which is not available");
  process.exit(1);
}

// The worker scripts, read here through the table `verify-browser-syntax.ts` reads, and handed to the harness as files:
// it runs on the image's Node 16, where nothing can parse the TypeScript the table and its readers are written in.
const workers = await readWorkerScripts(DIST);
if (workers.problems.length > 0) {
  console.error("✗ verify:browser-floor — the worker scripts cannot be read:");
  for (const problem of workers.problems) {
    console.error(`  ${problem}`);
  }
  process.exit(1);
}
fs.mkdirSync(path.join(ROOT, "tmp"), { recursive: true });
const WORKERS = fs.mkdtempSync(path.join(ROOT, "tmp", "browser-floor-workers-"));
for (const { name, text } of workers.scripts) {
  fs.writeFileSync(path.join(WORKERS, `${name.replace(/\W+/g, "-")}.js`), text);
}

const failed: string[] = [];
for (const engine of selected) {
  const image = `mcr.microsoft.com/playwright:v${engine.playwright}-focal`;
  const result = spawnSync(
    "docker",
    [
      "run",
      "--rm",
      "--init",
      "-e",
      `ENGINE=${engine.name}`,
      "-e",
      `EXPECT_VERSION=${engine.version}`,
      "-e",
      // The published bundles, from the table the build and the docs tests read: the harness checks `dist/iife`
      // against it, so a bundle the build did not write fails rather than being left out of what is expected.
      `EXPECT_BUNDLES=${JSON.stringify(IIFE_BUNDLES.map(({ global, file }) => ({ global, file })))}`,
      "-e",
      "NODE_PATH=/tmp/floor/node_modules",
      "-e",
      "NPM_CONFIG_UPDATE_NOTIFIER=false",
      "-v",
      `${DIST}:/dist:ro`,
      "-v",
      `${path.join(ROOT, "scripts/browser-floor")}:/harness:ro`,
      "-v",
      `${WORKERS}:/workers:ro`,
      image,
      "bash",
      "-c",
      // The release's own Playwright, installed inside the container, so the host's version is not involved.
      `mkdir -p /tmp/floor && cd /tmp/floor && npm install --silent --no-save --no-audit --no-fund playwright@${engine.playwright} >/dev/null && node /harness/harness.cjs /dist /workers`
    ],
    { stdio: "inherit" }
  );
  if (result.status !== 0) {
    failed.push(engine.name);
  }
}

fs.rmSync(WORKERS, { recursive: true, force: true });

if (failed.length > 0) {
  console.error(`✗ verify:browser-floor — failed in ${failed.join(", ")}`);
  process.exit(1);
}
console.log(`✓ verify:browser-floor — ${selected.map(e => e.name).join(", ")} at the floor`);
