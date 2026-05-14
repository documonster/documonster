import { spawnSync } from "node:child_process";
/**
 * Word Examples — Master Runner
 *
 * Runs every word example file in numeric order. Use this to regenerate the
 * `tmp/word-examples/` directory in one shot.
 *
 * Usage:
 *   pnpm exec tsx src/modules/word/examples/index.ts
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

// Discover every "NN-*.ts" file (excluding this index) sorted by numeric prefix.
const files = fs
  .readdirSync(here)
  .filter(f => /^\d{2}-.*\.ts$/.test(f))
  .sort();

if (files.length === 0) {
  console.error("No example files found.");
  process.exit(1);
}

// Clean output directory so each run produces a fresh, deterministic result.
const outDir = path.resolve(here, "../../../../tmp/word-examples");
if (fs.existsSync(outDir)) {
  fs.rmSync(outDir, { recursive: true, force: true });
}
fs.mkdirSync(outDir, { recursive: true });

console.log(`Running ${files.length} word examples → ${outDir}\n`);

const t0 = performance.now();
let passed = 0;
let failed = 0;

for (const file of files) {
  const fullPath = path.join(here, file);
  process.stdout.write(`▶ ${file} ... `);
  const start = performance.now();
  const result = spawnSync("npx", ["tsx", fullPath], {
    stdio: ["ignore", "pipe", "pipe"],
    cwd: path.resolve(here, "../../../..")
  });
  const ms = (performance.now() - start).toFixed(0);
  if (result.status === 0) {
    console.log(`✓ (${ms} ms)`);
    passed++;
  } else {
    console.log(`✗ (${ms} ms)`);
    console.log(result.stdout?.toString() ?? "");
    console.log(result.stderr?.toString() ?? "");
    failed++;
  }
}

const total = ((performance.now() - t0) / 1000).toFixed(1);
console.log(`\n${passed}/${files.length} examples passed in ${total} s`);
if (failed > 0) {
  console.error(`${failed} example(s) failed.`);
  process.exit(1);
}
